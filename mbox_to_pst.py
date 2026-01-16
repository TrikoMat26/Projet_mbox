import mailbox
import win32com.client
import os
import sys
import time
import tempfile
import logging
import json
from email.header import decode_header
from email.utils import parsedate_to_datetime, getaddresses, formataddr
import datetime
import mimetypes

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("migration.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

STATE_FILE = "migration_state.json"

def decode_mime_header(header_value):
    if not header_value:
        return ""
    try:
        decoded_parts = decode_header(header_value)
        result = []
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                # Try UTF-8 first with strict errors to trigger fallback if invalid
                try:
                    result.append(part.decode(encoding or 'utf-8', errors='strict'))
                except:
                    # Fallback to latin-1 or windows-1252 if UTF-8 fails
                    try:
                        result.append(part.decode('latin-1', errors='replace'))
                    except:
                        result.append(part.decode('utf-8', errors='replace'))
            else:
                result.append(part)
        return "".join(result)
    except:
        return str(header_value)

def set_item_properties(mail_item, date_obj, sender_name="", sender_email=""):
    """
    Uses PropertyAccessor to set the sent/received date, message flags, and SENDER info.
    """
    try:
        # Time Properties
        PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
        PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
        
        # Flags (Ready/Read)
        PR_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
        
        # Sender Properties (Force the "From" field)
        PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
        PR_SENDER_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
        PR_SENDER_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x0C1E001F"
        
        PR_SENT_REPRESENTING_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
        PR_SENT_REPRESENTING_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"
        PR_SENT_REPRESENTING_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x0064001F"

        prop_accessor = mail_item.PropertyAccessor
        
        # Set Dates
        if date_obj:
            prop_accessor.SetProperty(PR_CLIENT_SUBMIT_TIME, date_obj)
            prop_accessor.SetProperty(PR_MESSAGE_DELIVERY_TIME, date_obj)
        
        # Set Flags: MSGFLAG_READ (0x1) and clear MSGFLAG_UNSENT (0x8)
        prop_accessor.SetProperty(PR_MESSAGE_FLAGS, 1)

        # Set Sender Info manually if available
        if sender_name or sender_email:
            name = sender_name or sender_email

            prop_accessor.SetProperty(PR_SENDER_NAME, name)
            prop_accessor.SetProperty(PR_SENT_REPRESENTING_NAME, name)

            if sender_email:
                prop_accessor.SetProperty(PR_SENDER_EMAIL_ADDRESS, sender_email)
                prop_accessor.SetProperty(PR_SENDER_ADDRTYPE, "SMTP")
                prop_accessor.SetProperty(PR_SENT_REPRESENTING_EMAIL_ADDRESS, sender_email)
                prop_accessor.SetProperty(PR_SENT_REPRESENTING_ADDRTYPE, "SMTP")
        
    except Exception as e:
        # logging.warning(f"Error setting MAPI properties: {e}")
        pass

def save_state(count):
    with open(STATE_FILE, "w") as f:
        json.dump({"last_count": count}, f)

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f:
            return json.load(f).get("last_count", 0)
    return 0

def add_to_master_categories(namespace, category_names):
    """Adds categories to the Outlook Master Category List if they don't exist."""
    try:
        master_list = namespace.Categories
        existing = {cat.Name for cat in master_list}
        for name in category_names:
            if name and name not in existing:
                try:
                    # olCategoryColorNone = 0, or just let Outlook pick
                    master_list.Add(name)
                    existing.add(name)
                except: pass
    except Exception as e:
        logging.warning(f"Could not update Master Category List: {e}")

def normalize_addresses(header_value):
    if not header_value:
        return ""
    addresses = []
    for name, email in getaddresses([header_value]):
        decoded_name = decode_mime_header(name).strip()
        addresses.append(formataddr((decoded_name, email)))
    return "; ".join([addr for addr in addresses if addr.strip()])

def parse_sender(header_value):
    if not header_value:
        return "", ""
    decoded_header = decode_mime_header(header_value)
    addresses = getaddresses([header_value, decoded_header])
    if not addresses:
        name, email = "", ""
    else:
        name, email = addresses[0]
    decoded_name = decode_mime_header(name).strip()
    return decoded_name or email or decoded_header.strip(), email

def format_sender_display(sender_name, sender_email):
    if sender_email:
        return formataddr((sender_name, sender_email))
    return sender_name

def mbox_to_pst(mbox_path, pst_path, folder_name="Gmail Archive", resume=True, limit=None):
    if not os.path.exists(mbox_path):
        logging.error(f"MBOX file not found at {mbox_path}")
        return

    # Initialize Outlook
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        logging.error(f"Error connecting to Outlook: {e}. Ensure Outlook is installed.")
        return

    # Create/Open PST
    logging.info(f"Opening/Creating PST: {pst_path}")
    pst_abs_path = os.path.abspath(pst_path)
    try:
        pst_store = None
        for store in namespace.Stores:
            try:
                if store.FilePath.lower() == pst_abs_path.lower():
                    pst_store = store
                    break
            except: continue
        
        if not pst_store:
            namespace.AddStore(pst_abs_path)
            for store in namespace.Stores:
                try:
                    if store.FilePath.lower() == pst_abs_path.lower():
                        pst_store = store
                        break
                except: continue
                    
        if not pst_store:
            logging.error("Could not find or create the PST store.")
            return
            
        root_folder = pst_store.GetRootFolder()
    except Exception as e:
        logging.error(f"Error accessing PST: {e}")
        return

    # Get target folder
    try:
        target_folder = None
        for folder in root_folder.Folders:
            if folder.Name == folder_name:
                target_folder = folder
                break
        if not target_folder:
            target_folder = root_folder.Folders.Add(folder_name)
    except Exception as e:
        logging.error(f"Error creating/accessing folder '{folder_name}': {e}")
        return

    # For clearing the 'Draft' status, we will create items in a temporary storage (like default Inbox)
    # and then move them to the target PST folder. This is a common FIX for the Unsent flag.
    # We'll use the default folder for temporary creation.
    try:
        temp_folder = namespace.GetDefaultFolder(6) # 6 = olFolderInbox
    except:
        temp_folder = target_folder # Fallback

    # Master Category List Caching
    known_master_categories = set()
    try:
        for cat in namespace.Categories:
            known_master_categories.add(cat.Name)
    except: pass
    
    def ensure_categories_exist(cat_list):
        for c in cat_list:
            if c and c not in known_master_categories:
                try:
                    namespace.Categories.Add(c)
                    known_master_categories.add(c)
                except: pass

    # Create a transit folder WITHIN the PST to avoid cross-store resource issues
    # and still fix the 'Draft' status via the Move() method.
    try:
        temp_folder_name = "_Temp_Migration_"
        temp_folder = None
        for folder in root_folder.Folders:
            if folder.Name == temp_folder_name:
                temp_folder = folder
                break
        if not temp_folder:
            temp_folder = root_folder.Folders.Add(temp_folder_name)
    except Exception as e:
        logging.warning(f"Could not create temp folder in PST, using target: {e}")
        temp_folder = target_folder

    # Open MBOX
    logging.info(f"Opening MBOX file: {mbox_path}")
    mbox = mailbox.mbox(mbox_path)
    attachments_temp_dir = tempfile.TemporaryDirectory()
    
    start_at = 0
    if resume:
        start_at = load_state()
        if start_at > 0:
            logging.info(f"Resuming from message {start_at}...")

    count = 0
    errors = 0
    start_time = time.time()
    
    # Effective limit calculation
    effective_limit = (start_at + limit) if limit else None
    
    # Iterate through MBOX
    for i, message in enumerate(mbox):
        if effective_limit and i >= effective_limit:
            logging.info(f"Session limit of {limit} messages reached (Index {i}). Stopping.")
            break

        if i < start_at:
            count = i + 1
            continue
            
        mail = None # Initialize to ensure cleanup
        try:
            # Extract basic info
            subject = decode_mime_header(message['subject']) or "(No Subject)"
            sender_header = message['from'] or ""
            to_header = message['to'] or ""
            sender_name, sender_email = parse_sender(sender_header)
            to = normalize_addresses(to_header)
            
            # Date parsing
            date_val = None
            if message['date']:
                try:
                    date_val = parsedate_to_datetime(message['date'])
                except: pass

            # X-Gmail-Labels
            # USE get_all to capture multiple headers if present
            labels_headers = message.get_all('X-Gmail-Labels', [])
            categories = []
            
            for distinct_header in labels_headers:
                if distinct_header:
                    decoded = decode_mime_header(distinct_header)
                    # Split by comma
                    parts = [l.strip() for l in decoded.split(',') if l.strip()]
                    categories.extend(parts)
            
            # De-duplicate
            categories = list(set(categories))
            
            # Ensure they exist in Master List immediately
            if categories:
                ensure_categories_exist(categories)

            # Création du message dans le dossier de transit
            # Cette méthode "Create -> Move" est la seule fiable pour supprimer l'état Brouillon.
            mail = temp_folder.Items.Add(0) # 0 = olMailItem
            
            # Application des propriétés de base
            mail.Subject = subject
            mail.SentOnBehalfOfName = format_sender_display(sender_name, sender_email)
            mail.To = to
            
            if categories:
                mail.Categories = "; ".join(categories)
            
            # Corps et Pièces jointes
            body_html = ""
            body_text = ""
            if message.is_multipart():
                for part in message.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue

                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition", ""))
                    filename = part.get_filename()
                    content_id = part.get('Content-ID')
                    
                    # Determine if this part is a body or an attachment/inline image
                    # Default: Body if text and no filename/disposition
                    is_attachment = False
                    
                    if "attachment" in content_disposition:
                        is_attachment = True
                    elif filename:
                         # Has a filename, usually an attachment (even if inline)
                        is_attachment = True
                    elif content_type not in ("text/plain", "text/html"):
                        # Info not text, assume attachment (e.g. image without headers)
                        is_attachment = True
                    
                    # Handle Body
                    if not is_attachment and content_type in ("text/plain", "text/html"):
                        try:
                            payload = part.get_payload(decode=True)
                            charset = part.get_content_charset() or 'utf-8'
                            decoded = payload.decode(charset, errors='replace')
                            if content_type == "text/html":
                                body_html += decoded
                            else:
                                body_text += decoded
                        except: pass
                    
                    # Handle Attachment/Inline
                    else:
                        # Ensure we have a filename
                        if filename:
                            filename = decode_mime_header(filename)
                        else:
                            # Generate name if missing
                            ext = mimetypes.guess_extension(content_type) or ".dat"
                            filename = f"attachment_{os.urandom(4).hex()}{ext}"
                            
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                temp_path = os.path.join(attachments_temp_dir.name, filename)
                                with open(temp_path, "wb") as f:
                                    f.write(payload)
                                
                                # Add to Outlook
                                attachment = mail.Attachments.Add(temp_path, 1, 1, filename)
                                
                                # Handle Content-ID for inline images
                                # CID allows <img src="cid:foo"> to work
                                if content_id:
                                    # Remove <>
                                    cid_clean = content_id.strip('<>')
                                    try:
                                        # PR_ATTACH_CONTENT_ID = 0x3712001F
                                        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid_clean)
                                    except: pass
                                
                                try:
                                    os.remove(temp_path)
                                except OSError:
                                    pass
                                    
                        except Exception as att_err:
                            logging.warning(f"Error attached/inline {filename}: {att_err}")
            else:
                try:
                    payload = message.get_payload(decode=True)
                    charset = message.get_content_charset() or 'utf-8'
                    content = payload.decode(charset, errors='replace')
                    if message.get_content_type() == "text/html":
                        body_html = content
                    else:
                        body_text = content
                except: pass

            if body_html:
                mail.HTMLBody = body_html
            elif body_text:
                mail.Body = body_text
            
            # Forcer le format du message pour éviter le texte brut par défaut
            mail.MessageClass = "IPM.Note"
            
            # Sauvegarde initiale dans le dossier temporaire
            mail.Save()
            
            # Suppression du flag Unsent et réglage de la date via MAPI
            # On le fait AVANT le déplacement
            set_item_properties(mail, date_val, sender_name=sender_name, sender_email=sender_email)
            mail.Save()

            # DEPLACEMENT FINAL
            # Cela cristallise le statut "Envoyé" et retire définitivement "Brouillon"
            if temp_folder != target_folder:
                mail.Move(target_folder)
            
            count = i + 1
            if count % 100 == 0:
                elapsed = time.time() - start_time
                rate = (count - start_at) / elapsed if elapsed > 0 else 0
                logging.info(f"Processed {count} messages... ({rate:.2f} msgs/sec)")
                save_state(count)
            
            # Micro-pause every 10 messages to avoid resource exhaustion
            if count % 10 == 0:
                time.sleep(0.1)
                
        except Exception as e:
            errors += 1
            logging.error(f"Error processing message {i}: {e}")
            if errors > 500: # Higher threshold for 10GB
                logging.error("Too many errors, stopping.")
                break
            continue
        finally:
            # Explicitly release the COM object
            mail = None

    # Cleanup temp folder if empty
    try:
        if temp_folder.Items.Count == 0:
            temp_folder.Delete()
    except: pass

    attachments_temp_dir.cleanup()

    save_state(count)
    logging.info(f"Migration completed!")
    logging.info(f"Total messages: {count}")
    logging.info(f"Errors: {errors}")
    logging.info(f"PST: {pst_abs_path}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Migration MBOX Gmail vers Outlook PST avec Catégories")
    parser.add_argument("mbox", help="Chemin du fichier .mbox")
    parser.add_argument("pst", help="Chemin du fichier .pst de sortie")
    parser.add_argument("--folder", default="Gmail Archive", help="Nom du dossier cible dans Outlook")
    parser.add_argument("--no-resume", action="store_false", dest="resume", help="Ne pas reprendre la migration précédente")
    parser.add_argument("--limit", type=int, default=None, help="Limiter le nombre de messages à traiter (pour test)")
    
    args = parser.parse_args()
    mbox_to_pst(args.mbox, args.pst, args.folder, args.resume, args.limit)
