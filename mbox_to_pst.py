import mailbox
import win32com.client
import os
import sys
import time
import tempfile
import logging
import json
from email.header import decode_header
from email.utils import parsedate_to_datetime
import datetime

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
                try:
                    result.append(part.decode(encoding or 'utf-8', errors='replace'))
                except:
                    result.append(part.decode('latin-1', errors='replace'))
            else:
                result.append(part)
        return "".join(result)
    except:
        return str(header_value)

def set_item_properties(mail_item, date_obj):
    """
    Uses PropertyAccessor to set the sent/received date and message flags.
    Property tags:
    0x00390040: PR_CLIENT_SUBMIT_TIME
    0x0E060040: PR_MESSAGE_DELIVERY_TIME
    0x0E070003: PR_MESSAGE_FLAGS
    """
    try:
        PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
        PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
        PR_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
        
        prop_accessor = mail_item.PropertyAccessor
        
        # Set Dates
        if date_obj:
            prop_accessor.SetProperty(PR_CLIENT_SUBMIT_TIME, date_obj)
            prop_accessor.SetProperty(PR_MESSAGE_DELIVERY_TIME, date_obj)
        
        # Set Flags: MSGFLAG_READ (0x1) and clear MSGFLAG_UNSENT (0x8)
        # Value 1 means 'Read' and not 'Unsent'.
        prop_accessor.SetProperty(PR_MESSAGE_FLAGS, 1)
        
    except Exception:
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

def mbox_to_pst(mbox_path, pst_path, folder_name="Gmail Archive", resume=True):
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

    # Open MBOX
    logging.info(f"Opening MBOX file: {mbox_path}")
    mbox = mailbox.mbox(mbox_path)
    
    start_at = 0
    if resume:
        start_at = load_state()
        if start_at > 0:
            logging.info(f"Resuming from message {start_at}...")

    count = 0
    errors = 0
    start_time = time.time()
    
    # Iterate through MBOX
    # For large MBOX, we avoid loading everything.
    # We iterate and skip until start_at.
    
    for i, message in enumerate(mbox):
        if i < start_at:
            count = i + 1
            continue
            
        try:
            # Extract basic info
            subject = decode_mime_header(message['subject']) or "(No Subject)"
            sender = decode_mime_header(message['from']) or ""
            to = decode_mime_header(message['to']) or ""
            
            # Date parsing
            date_val = None
            if message['date']:
                try:
                    date_val = parsedate_to_datetime(message['date'])
                except: pass

            # X-Gmail-Labels
            labels_raw = message.get('X-Gmail-Labels', '')
            categories = []
            if labels_raw:
                # Decode headers FIRST (handles MIME encoded characters like accents)
                # This fixes labels appearing like =?UTF-8?Q?Messages...?=
                labels_decoded = decode_mime_header(labels_raw)
                # Then split by comma
                categories = [l.strip() for l in labels_decoded.split(',') if l.strip()]
                
                # Optional: Add to master list (can be slow if done every time, maybe every 100)
                if count % 100 == 0:
                    add_to_master_categories(namespace, categories)

            # Création du message dans la Boîte de réception (temp_folder)
            # Cette méthode "Create -> Move" est la seule fiable pour supprimer l'état Brouillon.
            mail = temp_folder.Items.Add(0) # 0 = olMailItem
            
            # Application des propriétés de base
            mail.Subject = subject
            mail.SentOnBehalfOfName = sender
            mail.To = to
            
            if categories:
                mail.Categories = "; ".join(categories)
            
            # Corps et Pièces jointes
            body_html = ""
            body_text = ""
            if message.is_multipart():
                for part in message.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        try:
                            payload = part.get_payload(decode=True)
                            charset = part.get_content_charset() or 'utf-8'
                            body_text += payload.decode(charset, errors='replace')
                        except: pass
                    elif content_type == "text/html" and "attachment" not in content_disposition:
                        try:
                            payload = part.get_payload(decode=True)
                            charset = part.get_content_charset() or 'utf-8'
                            body_html += payload.decode(charset, errors='replace')
                        except: pass
                    elif "attachment" in content_disposition or part.get_filename():
                        filename = part.get_filename()
                        if filename:
                            filename = decode_mime_header(filename)
                            try:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    # Create a unique temporary directory for this specific message's attachments
                                    # to ensure we can use the ACTUAL filename without collisions.
                                    with tempfile.TemporaryDirectory() as temp_dir:
                                        temp_path = os.path.join(temp_dir, filename)
                                        with open(temp_path, "wb") as f:
                                            f.write(payload)
                                        
                                        # mail.Attachments.Add(Source, Type, Position, DisplayName)
                                        # Source must be the full path. DisplayName is what Outlook shows.
                                        mail.Attachments.Add(temp_path, 1, 1, filename)
                            except Exception as att_err:
                                logging.warning(f"Erreur attachement {filename}: {att_err}")
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
            set_item_properties(mail, date_val)
            mail.Save()

            # DEPLACEMENT vers le PST cible
            # Cela cristallise le statut "Envoyé" et retire définitivement "Brouillon"
            if temp_folder != target_folder:
                mail.Move(target_folder)
            
            count = i + 1
            if count % 100 == 0:
                elapsed = time.time() - start_time
                rate = (count - start_at) / elapsed if elapsed > 0 else 0
                logging.info(f"Processed {count} messages... ({rate:.2f} msgs/sec)")
                save_state(count)
                
        except Exception as e:
            errors += 1
            logging.error(f"Error processing message {i}: {e}")
            if errors > 500: # Higher threshold for 10GB
                logging.error("Too many errors, stopping.")
                break
            continue

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
    
    args = parser.parse_args()
    mbox_to_pst(args.mbox, args.pst, args.folder, args.resume)
