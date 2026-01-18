import win32com.client
import os
import sys
import time
import tempfile
import logging
import json
from email.header import decode_header
from email.utils import parsedate_to_datetime, getaddresses, formataddr, parseaddr
from email import message_from_bytes
import datetime
import mimetypes
import re

def stream_mbox(mbox_path, start_at=0, progress_callback=None):
    """
    Streaming MBOX parser that yields (message_index, file_position, message) tuples.
    
    This reads the file in chunks and parses messages on-the-fly, allowing
    real-time progress updates based on file position.
    
    Args:
        mbox_path: Path to the MBOX file
        start_at: Message index to start from (for resume support)
        progress_callback: Optional callable(file_position, file_size) for progress updates
    
    Yields:
        (message_index, file_position, email.message.Message)
    """
    file_size = os.path.getsize(mbox_path)
    mbox_from_pattern = re.compile(rb'^From .+\r?\n', re.MULTILINE)
    
    with open(mbox_path, 'rb') as f:
        message_index = 0
        current_message_data = b''
        last_progress_pos = 0
        
        # Read in chunks for efficiency
        chunk_size = 1024 * 1024  # 1 MB chunks
        buffer = b''
        
        while True:
            chunk = f.read(chunk_size)
            if not chunk and not buffer:
                break
            
            buffer += chunk
            
            # Find all "From " lines (message boundaries)
            # MBOX format: each message starts with "From <email> <date>"
            matches = list(mbox_from_pattern.finditer(buffer))
            
            if len(matches) == 0:
                # No complete message boundary found yet, keep reading
                if not chunk:  # End of file
                    # Process remaining data as last message
                    if buffer.strip():
                        if message_index >= start_at:
                            try:
                                msg = message_from_bytes(buffer)
                                yield (message_index, f.tell(), msg)
                            except Exception:
                                pass
                    break
                continue
            
            # Process all complete messages (except the last one in buffer)
            for i, match in enumerate(matches):
                if i == 0:
                    # First match - data before it belongs to previous message
                    if current_message_data:
                        if message_index >= start_at:
                            try:
                                msg = message_from_bytes(current_message_data)
                                pos = f.tell() - len(buffer) + match.start()
                                yield (message_index, pos, msg)
                            except Exception:
                                pass
                        message_index += 1
                        
                        # Progress callback during skip phase
                        if progress_callback and message_index < start_at:
                            pos = f.tell() - len(buffer) + match.start()
                            if pos - last_progress_pos > 10 * 1024 * 1024:  # Every 10MB
                                progress_callback(pos, file_size, message_index, start_at)
                                last_progress_pos = pos
                    
                    # Start new message (skip the "From " line itself)
                    if i + 1 < len(matches):
                        current_message_data = buffer[match.end():matches[i + 1].start()]
                    else:
                        current_message_data = buffer[match.end():]
                else:
                    # Complete message between this match and previous
                    if message_index >= start_at:
                        try:
                            msg = message_from_bytes(current_message_data)
                            pos = f.tell() - len(buffer) + match.start()
                            yield (message_index, pos, msg)
                        except Exception:
                            pass
                    message_index += 1
                    
                    # Progress callback during skip phase
                    if progress_callback and message_index < start_at:
                        pos = f.tell() - len(buffer) + match.start()
                        if pos - last_progress_pos > 10 * 1024 * 1024:
                            progress_callback(pos, file_size, message_index, start_at)
                            last_progress_pos = pos
                    
                    # Start new message
                    if i + 1 < len(matches):
                        current_message_data = buffer[match.end():matches[i + 1].start()]
                    else:
                        current_message_data = buffer[match.end():]
            
            # Keep only the last incomplete message in buffer
            if matches:
                buffer = buffer[matches[-1].start():]
            
            if not chunk:  # End of file
                # Process final message
                if buffer.strip():
                    # Remove the "From " line from the start
                    match = mbox_from_pattern.match(buffer)
                    if match:
                        final_data = buffer[match.end():]
                    else:
                        final_data = buffer
                    
                    if final_data.strip() and message_index >= start_at:
                        try:
                            msg = message_from_bytes(final_data)
                            yield (message_index, file_size, msg)
                        except Exception:
                            pass
                break


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
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        logging.error(f"Error connecting to Outlook: {e}. Ensure Outlook is installed.")
        return

    # Create/Open PST
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

    # Initialize state variables
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

    # Setup progress bar
    file_size = os.path.getsize(mbox_path)
    file_size_mb = file_size / (1024 * 1024)
    
    progress_bar = None
    last_progress_mb = 0
    
    # For unlimited mode: create progress bar immediately (file-based)
    if TQDM_AVAILABLE and not limit:
        progress_bar = tqdm(total=int(file_size_mb), desc="Processing", unit="MB",
                           file=sys.stderr, dynamic_ncols=True,
                           bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}MB [{elapsed}<{remaining}]')
    
    # Progress callback for the streaming parser (unlimited mode only)
    def progress_callback(pos, total, msg_idx, target_idx):
        nonlocal last_progress_mb
        if progress_bar:
            current_mb = int(pos / (1024 * 1024))
            if current_mb > last_progress_mb:
                progress_bar.update(current_mb - last_progress_mb)
                last_progress_mb = current_mb
    
    # Show info about skipping if resuming
    if start_at > 0:
        logging.info(f"Seeking to message {start_at}...")
    
    # Use streaming MBOX parser for real-time progress
    messages_processed = 0
    progress_bar_created_for_limit = False
    
    for i, file_pos, message in stream_mbox(mbox_path, start_at=0, progress_callback=progress_callback if not limit else None):
        # Update progress based on file position (for unlimited mode)
        if not limit and progress_bar:
            current_mb = int(file_pos / (1024 * 1024))
            if current_mb > last_progress_mb:
                progress_bar.update(current_mb - last_progress_mb)
                last_progress_mb = current_mb
        
        # Skip to resume point
        if i < start_at:
            continue
        
        # Create progress bar for limit mode only when processing actually starts
        if limit and TQDM_AVAILABLE and not progress_bar_created_for_limit:
            progress_bar = tqdm(total=limit, desc="Processing", unit="msg",
                               file=sys.stderr, dynamic_ncols=True,
                               bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} msgs [{elapsed}<{remaining}]')
            progress_bar_created_for_limit = True
            
        if effective_limit and i >= effective_limit:

            logging.info(f"Session limit of {limit} messages reached.")
            break

        mail = None
        try:
            # Extract headers
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

            # Création du message dans le dossier de transit
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
                    
                    is_attachment = False
                    
                    if "attachment" in content_disposition:
                        is_attachment = True
                    elif filename:
                        is_attachment = True
                    elif content_type not in ("text/plain", "text/html"):
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
                        if filename:
                            filename = decode_mime_header(filename)
                        else:
                            ext = mimetypes.guess_extension(content_type) or ".dat"
                            filename = f"attachment_{os.urandom(4).hex()}{ext}"
                            
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                temp_path = os.path.join(attachments_temp_dir.name, filename)
                                with open(temp_path, "wb") as f:
                                    f.write(payload)
                                
                                attachment = mail.Attachments.Add(temp_path, 1, 1, filename)
                                
                                if content_id:
                                    cid_clean = content_id.strip('<>')
                                    try:
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
            
            mail.MessageClass = "IPM.Note"
            try:
                mail.UnRead = False
            except Exception:
                pass
            
            # Set MAPI Properties BEFORE Save
            set_item_properties(mail, date_val, sender_name=sender_name, sender_email=sender_email)
            
            # Save & Move
            mail.Save()
            if temp_folder != target_folder:
                mail.Move(target_folder)
            
            count = i + 1
            messages_processed += 1
            
            # Update progress bar (message count for limit mode)
            if progress_bar and limit:
                progress_bar.update(1)
            elif count % 100 == 0 and not progress_bar:
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
