import mailbox
import win32com.client
import os
import sys
import time
import tempfile
import logging
import json
from email.header import decode_header
from email.utils import parsedate_to_datetime, getaddresses, formataddr, parseaddr
import datetime
import mimetypes

# Optional: tqdm for progress bar (graceful fallback if not installed)
try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False
    tqdm = None

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
    Must be called BEFORE the first Save() to effectively clear Draft status.
    """
    try:
        prop_accessor = mail_item.PropertyAccessor
        
        # 1. FORCE CLEAR DRAFT STATUS FIRST
        # PR_MESSAGE_FLAGS (0x0E070003) -> 1 = Read, Sent.
        prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1)
        
        # PR_MESSAGE_STATUS (0x0E170003) -> 0 = Clean state
        prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E170003", 0)
        
        # PR_ICON_INDEX (0x10800003) -> 256 (Standard Unopened Mail Icon)
        try:
             prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x10800003", 256)
        except: pass

        # 2. Set Dates (Critical for display)
        if date_obj:
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x00390040", date_obj) # PR_CLIENT_SUBMIT_TIME
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E060040", date_obj) # PR_MESSAGE_DELIVERY_TIME

    except Exception as e:
        # logging.warning(f"Error setting basic MAPI flags: {e}")
        pass

    try:
         # 3. Set Sender Info (Must happen after clearing draft status for UI to respect it)
        if sender_name or sender_email:
            name = sender_name or sender_email
            email = sender_email or name
            
            # Basic props
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001F", name) # PR_SENDER_NAME
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0042001F", name) # PR_SENT_REPRESENTING_NAME
            
            if "@" in email:
                 prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001F", email) # PR_SENDER_EMAIL_ADDRESS
                 prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1E001F", "SMTP") # PR_SENDER_ADDRTYPE
                 
                 prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0065001F", email) # PR_SENT_REPRESENTING_EMAIL_ADDRESS
                 prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0064001F", "SMTP") # PR_SENT_REPRESENTING_ADDRTYPE
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

def normalize_addresses(header_value):
    if not header_value:
        return ""
    
    # Use getaddresses on the raw header converted to string.
    # We do NOT pre-decode the whole header because that can break address delimiters (commas).
    raw_values = [str(header_value)]
    
    seen = set()
    addresses = []
    
    for name, email in getaddresses(raw_values):
        if not email:
            # Sometimes getaddresses puts the whole encoded mess in 'name' if no angle brackets
            if "@" in name:
                 email = name
                 name = ""
            else:
                 continue
        
        email_clean = email.strip()
        email_lower = email_clean.lower()
        
        # Deduplication
        if email_lower in seen:
            continue
        seen.add(email_lower)
        
        # Decode the name properly
        decoded_name = ""
        if name:
            # Helper to strip surrounding quotes if they wrap an encoded word
            # e.g. "=?utf-8?..." -> =?utf-8?...
            candidate = name.strip()
            if candidate.startswith('"') and candidate.endswith('"') and "=?" in candidate:
                candidate = candidate[1:-1]
            
            decoded_name = decode_mime_header(candidate).strip()
            
            # Double-check: sometimes one pass isn't enough or it was double-encoded
            if "=?" in decoded_name:
                 decoded_name = decode_mime_header(decoded_name).strip()

        addresses.append(formataddr((decoded_name, email_clean)))
        
    return "; ".join(addresses)

def parse_sender(header_value):
    if not header_value:
        return "", ""
    
    # Use getaddresses which is more robust for headers than parseaddr
    # It handles comma-separated lists (we take the first one)
    pairs = getaddresses([str(header_value)])
    if pairs:
        name, email = pairs[0]
        decoded_name = decode_mime_header(name).strip()
        return decoded_name, email.strip()
    return "", ""

def format_sender_display(sender_name, sender_email):
    if sender_email:
        return formataddr((sender_name, sender_email))
    return sender_name

def mbox_to_pst(mbox_path, pst_path, folder_name="Gmail Archive", resume=True, limit=None):
    if not os.path.exists(mbox_path):
        logging.error(f"MBOX file not found at {mbox_path}")
        return

    pst_abs_path = os.path.abspath(pst_path)
    
    # Initialize Outlook
    init_start = time.time()
    try:
        # Connect to Outlook
        t_outlook = time.time()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print(f"[DEBUG INIT] Outlook Connection: {(time.time()-t_outlook)*1000:.1f}ms", flush=True)
    except Exception as e:
        logging.error(f"Error connecting to Outlook: {e}. Ensure Outlook is installed.")
        return

    # Create/Open PST
    t_pst = time.time()
    logging.info(f"Opening/Creating PST: {pst_abs_path}")
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
        print(f"[DEBUG INIT] PST Access: {(time.time()-t_pst)*1000:.1f}ms", flush=True)
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
    t_mbox = time.time()
    logging.info(f"Opening MBOX file: {mbox_path}")
    mbox = mailbox.mbox(mbox_path)
    print(f"[DEBUG INIT] MBOX Object Creation: {(time.time()-t_mbox)*1000:.1f}ms", flush=True)
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
    
    # File-based progress tracking
    file_size = os.path.getsize(mbox_path)
    file_size_mb = file_size / (1024 * 1024)
    
    # Open the underlying file to track position
    mbox_file = open(mbox_path, 'rb')
    
    # Setup progress bar
    progress_bar = None
    if TQDM_AVAILABLE:
        if limit:
            # With limit: use message count (we know exactly how many)
            progress_bar = tqdm(total=limit, desc="Processing", unit="msg",
                               file=sys.stderr, dynamic_ncols=True,
                               bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} msg [{elapsed}<{remaining}]')
        else:
            # Without limit: use file size in MB for smooth progress
            progress_bar = tqdm(total=int(file_size_mb), desc="Processing", unit="MB",
                               file=sys.stderr, dynamic_ncols=True,
                               bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}MB [{elapsed}<{remaining}]')
    
    # Show info about skipping if resuming
    if start_at > 0:
        logging.info(f"Seeking to message {start_at}...")
    
    # Track progress
    last_progress_mb = 0
    messages_processed = 0
    
    # Iterate through MBOX
    print(f"[DEBUG INIT] Setup complete in {(time.time()-init_start):.2f}s. Starting iteration...", flush=True)
    
    t_seek = time.time()
    for i, message in enumerate(mbox):
        # Update file-based progress during skip phase (no limit mode)
        if not limit and progress_bar and i < start_at:
            current_pos = mbox_file.tell()
            current_mb = int(current_pos / (1024 * 1024))
            if current_mb > last_progress_mb:
                progress_bar.update(current_mb - last_progress_mb)
                last_progress_mb = current_mb
        
        # Skip to resume point
        if i < start_at:
            if i > 0 and i % 1000 == 0:
                 print(f"\r[DEBUG INIT] Seeking: {i}/{start_at}...", end="", flush=True)
            continue
        
        if i == start_at and start_at > 0:
            print(f"\n[DEBUG INIT] Seek to {start_at} completed in {(time.time()-t_seek):.2f}s", flush=True)
            
        if effective_limit and i >= effective_limit:
            logging.info(f"Session limit of {limit} messages reached.")
            break

        # ===== DEBUG TIMING =====
        step_start = time.time()
        print(f"\n[DEBUG {i}] === Message {i} started ===", flush=True)
        
        mail = None # Initialize to ensure cleanup
        try:
            # Extract basic info
            t1 = time.time()
            subject = decode_mime_header(message['subject']) or "(No Subject)"
            sender_header = message['from'] or ""
            to_header = message['to'] or ""
            sender_name, sender_email = parse_sender(sender_header)
            to = normalize_addresses(to_header)
            print(f"[DEBUG {i}] Headers extraction: {(time.time()-t1)*1000:.1f}ms", flush=True)
            
            # Date parsing
            date_val = None
            if message['date']:
                try:
                    date_val = parsedate_to_datetime(message['date'])
                except: pass

            # X-Gmail-Labels
            t2 = time.time()
            labels_headers = message.get_all('X-Gmail-Labels', [])
            categories = []
            
            for distinct_header in labels_headers:
                if distinct_header:
                    decoded = decode_mime_header(distinct_header)
                    parts = [l.strip() for l in decoded.split(',') if l.strip()]
                    categories.extend(parts)
            
            categories = list(set(categories))
            
            if categories:
                ensure_categories_exist(categories)
            print(f"[DEBUG {i}] Labels/Categories: {(time.time()-t2)*1000:.1f}ms", flush=True)

            # Création du message dans le dossier de transit
            t3 = time.time()
            mail = temp_folder.Items.Add(0) # 0 = olMailItem
            print(f"[DEBUG {i}] Outlook Items.Add: {(time.time()-t3)*1000:.1f}ms", flush=True)
            
            # Application des propriétés de base
            t4 = time.time()
            mail.Subject = subject
            mail.SentOnBehalfOfName = format_sender_display(sender_name, sender_email)
            mail.To = to
            
            if categories:
                mail.Categories = "; ".join(categories)
            print(f"[DEBUG {i}] Basic properties set: {(time.time()-t4)*1000:.1f}ms", flush=True)

            # Corps et Pièces jointes
            t5 = time.time()
            body_html = ""
            body_text = ""
            attachments_count = 0
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
                                attachments_count += 1
                                
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
            print(f"[DEBUG {i}] Body+Attachments ({attachments_count} files): {(time.time()-t5)*1000:.1f}ms", flush=True)

            t6 = time.time()
            if body_html:
                mail.HTMLBody = body_html
            elif body_text:
                mail.Body = body_text
            print(f"[DEBUG {i}] Set Body content: {(time.time()-t6)*1000:.1f}ms", flush=True)
            
            # Forcer le format du message pour éviter le texte brut par défaut
            mail.MessageClass = "IPM.Note"
            try:
                mail.UnRead = False
            except Exception:
                pass
            
            # CRITICAL FIX: Set MAPI Properties (Sender, Flags, Dates) BEFORE the first Save()
            t7 = time.time()
            set_item_properties(mail, date_val, sender_name=sender_name, sender_email=sender_email)
            print(f"[DEBUG {i}] MAPI Properties: {(time.time()-t7)*1000:.1f}ms", flush=True)
            
            # Sauvegarde INITIALE & UNIQUE (Verrouille les propriétés)
            t8 = time.time()
            mail.Save()
            print(f"[DEBUG {i}] mail.Save(): {(time.time()-t8)*1000:.1f}ms", flush=True)

            # DEPLACEMENT FINAL
            t9 = time.time()
            if temp_folder != target_folder:
                mail.Move(target_folder)
            print(f"[DEBUG {i}] mail.Move(): {(time.time()-t9)*1000:.1f}ms", flush=True)
            
            print(f"[DEBUG {i}] === TOTAL: {(time.time()-step_start)*1000:.1f}ms ===", flush=True)
            
            count = i + 1
            messages_processed += 1
            
            # Update progress bar
            if progress_bar:
                if limit:
                    # Limit mode: update by message
                    progress_bar.update(1)
                else:
                    # Unlimited mode: update by file position
                    current_pos = mbox_file.tell()
                    current_mb = int(current_pos / (1024 * 1024))
                    if current_mb > last_progress_mb:
                        progress_bar.update(current_mb - last_progress_mb)
                        last_progress_mb = current_mb
            elif count % 100 == 0:
                # Fallback text logging if no tqdm
                elapsed = time.time() - start_time
                rate = (count - start_at) / elapsed if elapsed > 0 else 0
                logging.info(f"Processed {count} messages... ({rate:.2f} msgs/sec)")

            
            # Save state periodically
            if count % 100 == 0:
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

    # Close progress bar
    if progress_bar:
        progress_bar.close()
    
    # Close tracking file handle
    mbox_file.close()

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
