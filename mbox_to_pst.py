import win32com.client
import pywintypes  # Explicit import for PyInstaller
import win32timezone  # Required by pywintypes.Time()
import os
import mailbox  # Standard MBOX parser - more reliable than custom streaming
import sys
import time
import tempfile
import logging
import json
import signal
from email.header import decode_header
from email.utils import parsedate_to_datetime, getaddresses, formataddr
from email import message_from_bytes
import datetime
import mimetypes
import re

# Optional libraries for GUI and CLI progress
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False

try:
    import threading
    THREADING_AVAILABLE = True
except ImportError:
    THREADING_AVAILABLE = False

try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False
    tqdm = None

# Global state for graceful shutdown
_shutdown_requested = False
_current_count = 0

def signal_handler(signum, frame):
    """Handle Ctrl+C gracefully by flagging shutdown and saving state."""
    global _shutdown_requested
    _shutdown_requested = True
    logging.info("\n⚠ Interruption detected. Finishing current message and saving state...")

# Register signal handler (Windows compatible)
signal.signal(signal.SIGINT, signal_handler)

# Configure logging to console and file
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("migration.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

STATE_FILE = "migration_state.json"
PROBLEM_FILE = "problem_messages.json"

def log_problem_message(msg_index, subject, sender, date_str, error_type, error_detail):
    """Log a problematic message for later manual review."""
    problems = []
    if os.path.exists(PROBLEM_FILE):
        try:
            with open(PROBLEM_FILE, 'r', encoding='utf-8') as f:
                problems = json.load(f)
        except: pass
    
    # Decode sender if it's MIME-encoded
    decoded_sender = sender
    if sender:
        try:
            parts = decode_header(sender)
            decoded_parts = []
            for data, charset in parts:
                if isinstance(data, bytes):
                    decoded_parts.append(data.decode(charset or 'utf-8', errors='replace'))
                else:
                    decoded_parts.append(data)
            decoded_sender = ''.join(decoded_parts)
        except:
            decoded_sender = sender
    
    problems.append({
        "message_index": msg_index,
        "subject": subject[:100] if subject else "(No Subject)",
        "sender": decoded_sender[:100] if decoded_sender else "",
        "date": date_str or "",
        "error_type": error_type,
        "error_detail": str(error_detail)[:500],
        "logged_at": datetime.datetime.now().isoformat()
    })
    
    with open(PROBLEM_FILE, 'w', encoding='utf-8') as f:
        json.dump(problems, f, ensure_ascii=False, indent=2)

def decode_mime_header(header_value):
    if not header_value:
        return ""
    try:
        decoded_parts = decode_header(header_value)
        result = []
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                try:
                    result.append(part.decode(encoding or 'utf-8', errors='strict'))
                except:
                    try:
                        result.append(part.decode('latin-1', errors='replace'))
                    except:
                        result.append(part.decode('utf-8', errors='replace'))
            else:
                result.append(part)
        return "".join(result)
    except:
        return str(header_value)

def set_item_properties(mail_item, date_obj, sender_name="", sender_email="", references="", in_reply_to=""):
    """
    Uses PropertyAccessor to set the sent/received date, message flags, SENDER info, and threading headers.
    Must be called BEFORE the first Save() to effectively clear Draft status.
    """
    try:
        prop_accessor = mail_item.PropertyAccessor
    except Exception as e:
        logging.warning(f"Cannot get PropertyAccessor: {e}")
        return
    
    # 1. FORCE CLEAR DRAFT STATUS FIRST
    # PR_MESSAGE_FLAGS (0x0E070003) -> 1 = Read, Sent.
    try:
        prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1)
    except: pass
    
    # PR_ICON_INDEX (0x10800003) -> 256 (Standard Unopened Mail Icon)
    try:
        prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x10800003", 256)
    except: pass

    # 2. Set Dates (Critical for display)
    if date_obj:
        try:
            # Use pywintypes.Time which is the native COM date format
            pywin_date = pywintypes.Time(date_obj.timestamp())
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x00390040", pywin_date) # PR_CLIENT_SUBMIT_TIME
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E060040", pywin_date) # PR_MESSAGE_DELIVERY_TIME
        except:
            pass

    # 3. Set Sender Info
    if sender_name or sender_email:
        name = sender_name or sender_email
        email = sender_email or name
        
        try:
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001F", name) # PR_SENDER_NAME
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0042001F", name) # PR_SENT_REPRESENTING_NAME
        except: pass
        
        if "@" in email:
            try:
                prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001F", email) # PR_SENDER_EMAIL_ADDRESS
                prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1E001F", "SMTP") # PR_SENDER_ADDRTYPE
                prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0065001F", email) # PR_SENT_REPRESENTING_EMAIL_ADDRESS
                prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0064001F", "SMTP") # PR_SENT_REPRESENTING_ADDRTYPE
            except: pass

    # 4. Set Threading Headers for Conversation Grouping
    if references:
        try:
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x1039001F", references)
        except: pass
    
    if in_reply_to:
        try:
            clean_reply_to = in_reply_to.strip().strip('<>')
            prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x1042001F", clean_reply_to)
        except: pass

def save_state(count, processed_files=None, current_file=None, seen_ids=None):
    """Save the progress state including batch processed files and seen message IDs."""
    state = {
        "last_count": count,
        "processed_files": processed_files or [],
        "current_file": current_file,
        "seen_message_ids": list(seen_ids) if seen_ids else []
    }
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.warning(f"Could not save migration state: {e}")

def load_state():
    """Load the progress state."""
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
                if isinstance(state, dict):
                    last_count = state.get("last_count", 0)
                    processed_files = state.get("processed_files", [])
                    current_file = state.get("current_file", None)
                    seen_ids = set(state.get("seen_message_ids", []))
                    return last_count, processed_files, current_file, seen_ids
        except Exception as e:
            logging.warning(f"Error reading state file ({e}). Starting fresh.")
    return 0, [], None, set()

def normalize_addresses(header_value):
    if not header_value:
        return ""
    
    raw_values = [str(header_value)]
    seen = set()
    addresses = []
    
    for name, email in getaddresses(raw_values):
        if not email:
            if "@" in name:
                 email = name
                 name = ""
            else:
                 continue
        
        email_clean = email.strip()
        email_lower = email_clean.lower()
        
        if email_lower in seen:
            continue
        seen.add(email_lower)
        
        decoded_name = ""
        if name:
            candidate = name.strip()
            if candidate.startswith('"') and candidate.endswith('"') and "=?" in candidate:
                candidate = candidate[1:-1]
            
            decoded_name = decode_mime_header(candidate).strip()
            if "=?" in decoded_name:
                 decoded_name = decode_mime_header(decoded_name).strip()

        addresses.append(formataddr((decoded_name, email_clean)))
        
    return "; ".join(addresses)

def parse_sender(header_value):
    if not header_value:
        return "", ""
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

def mbox_to_pst(mbox_path, pst_path, folder_name="Gmail Archive", resume=True, limit=None, progress_callback=None):
    """
    Main function to process MBOX file(s) and migrate them into an Outlook PST file.
    Supports directories, seen message-ID tracking across resumes, and early binding COM acceleration.
    """
    global _shutdown_requested
    
    # 1. Expand list of MBOX files (Batch Mode)
    mbox_files = []
    if os.path.isdir(mbox_path):
        for file in sorted(os.listdir(mbox_path)):
            if file.lower().endswith('.mbox'):
                mbox_files.append(os.path.join(mbox_path, file))
        if not mbox_files:
            logging.error(f"No .mbox files found in directory: {mbox_path}")
            return
        logging.info(f"Batch mode: Found {len(mbox_files)} MBOX files in directory.")
    else:
        if not os.path.exists(mbox_path):
            logging.error(f"MBOX file not found at {mbox_path}")
            return
        mbox_files = [mbox_path]

    pst_abs_path = os.path.abspath(pst_path)
    
    # 2. Connect to Outlook (Early Binding COM Optimization)
    try:
        try:
            logging.info("Connecting to Outlook (Early Binding)...")
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        except Exception as eb_err:
            logging.warning(f"Early binding failed: {eb_err}. Falling back to standard late binding.")
            outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        logging.error(f"Error connecting to Outlook: {e}. Ensure Outlook is installed.")
        return

    # 3. Create/Open PST
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

    # 4. Get target folder
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

    # 5. Create a transit folder inside the PST to fix Draft status via Move
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
        logging.warning(f"Could not create temp transit folder inside PST, using target folder: {e}")
        temp_folder = target_folder

    # 6. Cache/Add Outlook Categories
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

    # 7. Initialize State and Resumption Points
    attachments_temp_dir = tempfile.TemporaryDirectory()
    
    if not resume:
        # Clear state file if resume is False
        if os.path.exists(STATE_FILE):
            try:
                os.remove(STATE_FILE)
            except: pass
        start_at, processed_files, current_file, seen_message_ids = 0, [], None, set()
    else:
        start_at, processed_files, current_file, seen_message_ids = load_state()
        if start_at > 0 or processed_files:
            logging.info(f"Resuming progress: {len(processed_files)} MBOX files processed, current file: '{current_file}' at index {start_at}.")

    total_duplicates = 0
    total_errors = 0
    total_processed = 0
    start_time = time.time()
    
    skip_files = False
    if resume and current_file:
        current_file_abs = os.path.abspath(current_file)
        if current_file_abs in [os.path.abspath(f) for f in mbox_files]:
            skip_files = True

    # 8. Main loop over each MBOX file
    for file_idx, mbox_file in enumerate(mbox_files):
        mbox_file_abs = os.path.abspath(mbox_file)
        
        # Skip fully processed files
        if mbox_file_abs in processed_files:
            logging.info(f"Skipping already processed MBOX: {os.path.basename(mbox_file)}")
            continue
            
        # Fast-forward to resume file
        if skip_files:
            if mbox_file_abs == os.path.abspath(current_file):
                skip_files = False
                file_start_at = start_at
            else:
                logging.info(f"Skipping {os.path.basename(mbox_file)} (already migrated)")
                continue
        else:
            file_start_at = 0
            
        logging.info(f"Processing MBOX file [{file_idx+1}/{len(mbox_files)}]: {os.path.basename(mbox_file)}")
        current_file = mbox_file_abs
        save_state(file_start_at, processed_files, current_file, seen_message_ids)
        
        try:
            mbox = mailbox.mbox(mbox_file)
            total_messages = len(mbox)
        except Exception as e:
            logging.error(f"Could not open MBOX file {os.path.basename(mbox_file)}: {e}")
            total_errors += 1
            processed_files.append(mbox_file_abs)
            continue
            
        logging.info(f"Found {total_messages} messages. Starting migration from message index {file_start_at}...")
        
        count = file_start_at
        progress_bar = None
        progress_bar_created = False
        
        for i, message in enumerate(mbox):
            if i < file_start_at:
                continue
                
            # If user set a limit, check if we reached it
            if limit and (count - file_start_at) >= limit:
                logging.info(f"Limit of {limit} messages reached for this session.")
                break
            
            if _shutdown_requested:
                logging.info(f"Graceful shutdown triggered. Saving state at message {count}...")
                save_state(count, processed_files, current_file, seen_message_ids)
                break

            # CLI Progress Bar
            if TQDM_AVAILABLE and not progress_bar_created and not progress_callback:
                desc_text = f"MBOX {file_idx+1}/{len(mbox_files)}"
                p_total = limit if limit else (total_messages - file_start_at)
                progress_bar = tqdm(total=p_total, desc=desc_text, unit="msg", file=sys.stderr, dynamic_ncols=True, leave=False)
                progress_bar_created = True

            mail = None
            try:
                # Deduplication by Message-ID (Persisted in state)
                message_id = message.get('Message-ID', '') or message.get('Message-Id', '')
                if message_id:
                    message_id = message_id.strip()
                    if message_id in seen_message_ids:
                        total_duplicates += 1
                        count = i + 1
                        if progress_bar:
                            progress_bar.update(1)
                        continue
                    seen_message_ids.add(message_id)
                
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
                    except:
                        pass

                # Conversation Threading headers
                references = message.get('References', '') or ''
                in_reply_to = message.get('In-Reply-To', '') or ''

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

                # Create draft message in temp folder
                mail = temp_folder.Items.Add(0) # 0 = olMailItem
                
                mail.Subject = subject
                mail.SentOnBehalfOfName = format_sender_display(sender_name, sender_email)
                mail.To = to

                if categories:
                    mail.Categories = "; ".join(categories)

                body_html = ""
                body_text = ""
                
                # Parse bodies and attachments
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
                        
                        # Handle Body text
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
                        
                        # Handle Attachment or Inline image
                        else:
                            if filename:
                                filename = decode_mime_header(filename)
                            else:
                                ext = mimetypes.guess_extension(content_type) or ".dat"
                                filename = f"attachment_{os.urandom(4).hex()}{ext}"
                            
                            filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
                                
                            try:
                                payload = part.get_payload(decode=True)
                                if payload is None:
                                    # Fallback payload parsing
                                    raw_payload = part.get_payload(decode=False)
                                    transfer_encoding = (part.get('Content-Transfer-Encoding') or '').lower().strip()
                                    
                                    if raw_payload:
                                        if transfer_encoding == 'base64':
                                            import base64
                                            try:
                                                if isinstance(raw_payload, str):
                                                    raw_payload = raw_payload.encode('ascii', errors='ignore')
                                                payload = base64.b64decode(raw_payload)
                                            except Exception as b64_err:
                                                logging.debug(f"Base64 decode failed for {filename}: {b64_err}")
                                        elif transfer_encoding == 'quoted-printable':
                                            import quopri
                                            try:
                                                if isinstance(raw_payload, str):
                                                    raw_payload = raw_payload.encode('ascii', errors='ignore')
                                                payload = quopri.decodestring(raw_payload)
                                            except Exception as qp_err:
                                                logging.debug(f"Quoted QP decode failed for {filename}: {qp_err}")
                                        elif transfer_encoding in ('7bit', '8bit', 'binary', ''):
                                            if isinstance(raw_payload, str):
                                                payload = raw_payload.encode('utf-8', errors='replace')
                                            else:
                                                payload = raw_payload
                                
                                if payload:
                                    import uuid
                                    base_name, ext = os.path.splitext(filename)
                                    unique_filename = f"{base_name}_{uuid.uuid4().hex[:8]}{ext}"
                                    temp_path = os.path.join(attachments_temp_dir.name, unique_filename)
                                    
                                    with open(temp_path, "wb") as f_attach:
                                        f_attach.write(payload)
                                        f_attach.flush()
                                        os.fsync(f_attach.fileno())
                                    
                                    written_size = os.path.getsize(temp_path)
                                    if written_size > 0:
                                        # Position 0 for inline CIDs, Position 1 for attachments
                                        position = 0 if content_id else 1
                                        attachment = mail.Attachments.Add(temp_path, 1, position, filename)
                                        
                                        # Set Full MAPI inline properties if Content-ID exists
                                        if content_id:
                                            cid_clean = content_id.strip('<>')
                                            try:
                                                # 1. PR_ATTACH_CONTENT_ID (0x3712001F)
                                                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid_clean)
                                                # 2. PR_ATTACH_FLAGS (0x37140003) -> 4 (ATT_MHTML_REF)
                                                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003", 4)
                                                # 3. PR_ATTACH_MIME_TAG (0x370E001F)
                                                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", content_type or "image/jpeg")
                                                # 4. PR_ATTACHMENT_HIDDEN (0x7FFE000B) -> True (hides it from standard attachment panel)
                                                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True)
                                            except Exception as prop_err:
                                                logging.debug(f"Failed setting inline attachment properties for {filename}: {prop_err}")
                                        
                                        try:
                                            os.remove(temp_path)
                                        except OSError:
                                            pass
                                    else:
                                        logging.warning(f"Empty attachment skipped: {filename}")
                                else:
                                    logging.warning(f"No payload extracted for attachment: {filename}")
                            except Exception as att_err:
                                logging.warning(f"Attachment error [{filename}]: {att_err}")
                                date_str = str(message['date']) if message['date'] else ""
                                log_problem_message(i, subject, sender_header, date_str, 
                                                  "attachment_error", f"{filename}: {att_err}")
                else:
                    # Simple singlepart message
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
                
                # Set MAPI Properties BEFORE Save to effectively clear Draft status
                set_item_properties(mail, date_val, sender_name=sender_name, sender_email=sender_email,
                                    references=references, in_reply_to=in_reply_to)
                
                # Save & Move (forces Outlook to materialize and clear MSGFLAG_UNSENT)
                mail.Save()
                if temp_folder != target_folder:
                    mail.Move(target_folder)
                
                count = i + 1
                total_processed += 1
                
                # Report progress
                if progress_callback:
                    current_done = count - file_start_at
                    total_to_do = limit if limit else (total_messages - file_start_at)
                    progress_callback(current_done, total_to_do, f"Fichier {file_idx+1}/{len(mbox_files)} : {os.path.basename(mbox_file)} - {count}/{total_messages} msg")
                elif progress_bar:
                    progress_bar.update(1)
                elif count % 100 == 0:
                    elapsed = time.time() - start_time
                    rate = total_processed / elapsed if elapsed > 0 else 0
                    logging.info(f"Processed {count}/{total_messages} messages... ({rate:.2f} msgs/sec)")
                
                # Periodic save state
                if count % 100 == 0:
                    save_state(count, processed_files, current_file, seen_message_ids)
                
                # Throttle slightly to avoid resource exhaustion
                if count % 10 == 0:
                    time.sleep(0.1)
                    
            except Exception as e:
                total_errors += 1
                logging.error(f"Error processing message {i}: {e}")
                if total_errors > 500:
                    logging.error("Too many errors, stopping migration process.")
                    break
                continue
            finally:
                mail = None

        if progress_bar:
            progress_bar.close()
            
        if _shutdown_requested:
            break
            
        # File completed successfully
        processed_files.append(mbox_file_abs)
        current_file = None
        start_at = 0
        save_state(0, processed_files, current_file, seen_message_ids)
        logging.info(f"Completed MBOX file: {os.path.basename(mbox_file)}")

    # 9. Final Cleanup and Logging
    try:
        if temp_folder.Items.Count == 0:
            temp_folder.Delete()
    except: pass

    attachments_temp_dir.cleanup()
    
    if not _shutdown_requested:
        # Clear state file upon clean successful completion of all files
        if os.path.exists(STATE_FILE):
            try:
                os.remove(STATE_FILE)
            except: pass
        logging.info(f"Migration completed successfully!")
    else:
        logging.info(f"Migration suspended.")

    logging.info(f"Total messages migrated: {total_processed}")
    logging.info(f"Duplicates skipped: {total_duplicates}")
    logging.info(f"Errors encountered: {total_errors}")
    logging.info(f"PST File: {pst_abs_path}")

# ==============================================================================
# GUI - TKINTER INTERFACE
# ==============================================================================

class TextHandler(logging.Handler):
    """Custom logging handler to redirect logging to a Tkinter Text widget."""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert('end', msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.yview('end')
        # Thread-safe appending using the after() method
        self.text_widget.after(0, append)

def run_gui():
    """Build and launch the Tkinter Graphical User Interface."""
    if not GUI_AVAILABLE:
        print("Error: Tkinter library is not available. Cannot run in GUI mode.")
        sys.exit(1)
        
    window = tk.Tk()
    window.title("Migration MBOX Gmail vers Outlook PST")
    window.geometry("850x650")
    window.configure(bg="#1e1e24")
    
    # Style styling
    style = ttk.Style()
    style.theme_use('clam')
    
    style.configure(".", background="#1e1e24", foreground="#ffffff")
    style.configure("TFrame", background="#1e1e24")
    style.configure("TLabel", background="#1e1e24", foreground="#ffffff", font=("Segoe UI", 10))
    style.configure("TButton", background="#5c6bc0", foreground="#ffffff", font=("Segoe UI", 10, "bold"), borderwidth=0)
    style.map("TButton", background=[("active", "#3f51b5"), ("disabled", "#555555")])
    style.configure("TCheckbutton", background="#1e1e24", foreground="#ffffff")
    style.configure("Horizontal.TProgressbar", thickness=15, troughcolor="#2a2a35", background="#00acc1")
    style.configure("TEntry", fieldbackground="#2a2a35", foreground="#ffffff", insertcolor="#ffffff", bordercolor="#424250", lightcolor="#424250")
    
    # Title Frame
    header_frame = tk.Frame(window, bg="#2a2a35", height=80)
    header_frame.pack(fill="x", padx=0, pady=0)
    
    header_label = tk.Label(header_frame, text="Migration Gmail (MBOX) vers Outlook PST", 
                            font=("Segoe UI", 16, "bold"), fg="#ffffff", bg="#2a2a35")
    header_label.pack(anchor="w", padx=20, pady=(15, 2))
    
    subtitle_label = tk.Label(header_frame, text="Convertisseur haute performance avec support des catégories et des dossiers", 
                              font=("Segoe UI", 9, "italic"), fg="#b0bec5", bg="#2a2a35")
    subtitle_label.pack(anchor="w", padx=20, pady=(0, 15))
    
    # Form layout
    form_frame = ttk.Frame(window, padding=20)
    form_frame.pack(fill="x")
    
    # Row 0: MBOX Selection
    ttk.Label(form_frame, text="Source MBOX (Fichier ou Dossier) :").grid(row=0, column=0, sticky="w", pady=6)
    mbox_path_var = tk.StringVar()
    mbox_entry = ttk.Entry(form_frame, textvariable=mbox_path_var, width=65)
    mbox_entry.grid(row=0, column=1, padx=10, pady=6)
    
    def select_mbox_file():
        path = filedialog.askopenfilename(filetypes=[("Fichiers MBOX", "*.mbox"), ("Tous les fichiers", "*.*")])
        if path: mbox_path_var.set(path)
        
    def select_mbox_dir():
        path = filedialog.askdirectory()
        if path: mbox_path_var.set(path)
        
    mbox_btn_frame = ttk.Frame(form_frame)
    mbox_btn_frame.grid(row=0, column=2, sticky="w", pady=6)
    ttk.Button(mbox_btn_frame, text="Fichier...", command=select_mbox_file, width=8).pack(side="left", padx=2)
    ttk.Button(mbox_btn_frame, text="Dossier...", command=select_mbox_dir, width=8).pack(side="left", padx=2)
    
    # Row 1: PST Selection
    ttk.Label(form_frame, text="Fichier PST de destination :").grid(row=1, column=0, sticky="w", pady=6)
    pst_path_var = tk.StringVar()
    pst_entry = ttk.Entry(form_frame, textvariable=pst_path_var, width=65)
    pst_entry.grid(row=1, column=1, padx=10, pady=6)
    
    def select_pst():
        path = filedialog.asksaveasfilename(defaultextension=".pst", filetypes=[("Fichiers PST Outlook", "*.pst"), ("Tous les fichiers", "*.*")])
        if path: pst_path_var.set(path)
        
    ttk.Button(form_frame, text="Parcourir...", command=select_pst, width=17).grid(row=1, column=2, sticky="w", pady=6)
    
    # Row 2: Target Folder Name
    ttk.Label(form_frame, text="Dossier racine cible dans Outlook :").grid(row=2, column=0, sticky="w", pady=6)
    folder_var = tk.StringVar(value="Gmail Archive")
    folder_entry = ttk.Entry(form_frame, textvariable=folder_var, width=30)
    folder_entry.grid(row=2, column=1, sticky="w", padx=10, pady=6)
    
    # Row 3: Session options
    options_frame = ttk.Frame(form_frame)
    options_frame.grid(row=3, column=1, columnspan=2, sticky="w", pady=6)
    
    resume_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(options_frame, text="Reprendre la migration interrompue (Resume)", variable=resume_var).pack(side="left", padx=(10, 20))
    
    ttk.Label(options_frame, text="Limite de messages (optionnelle) :").pack(side="left")
    limit_var = tk.StringVar()
    limit_entry = ttk.Entry(options_frame, textvariable=limit_var, width=10)
    limit_entry.pack(side="left", padx=10)
    
    # Execution Log frame
    log_frame = tk.Frame(window, bg="#111115")
    log_frame.pack(fill="both", expand=True, padx=20, pady=10)
    
    log_label = tk.Label(log_frame, text="Journal de migration (migration.log)", font=("Segoe UI", 9, "bold"), fg="#b0bec5", bg="#111115")
    log_label.pack(anchor="w", padx=15, pady=5)
    
    log_text = tk.Text(log_frame, bg="#111115", fg="#cfd8dc", font=("Consolas", 9), state="disabled", wrap="word", borderwidth=0)
    log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
    
    # Redirect root logger calls to the Tkinter Text widget
    text_handler = TextHandler(log_text)
    text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.getLogger().addHandler(text_handler)
    
    # Progress Indicators
    progress_frame = ttk.Frame(window, padding=20)
    progress_frame.pack(fill="x")
    
    status_var = tk.StringVar(value="Prêt à démarrer la migration.")
    status_label = ttk.Label(progress_frame, textvariable=status_var, font=("Segoe UI", 9, "bold"), foreground="#b0bec5")
    status_label.pack(anchor="w", pady=(0, 4))
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100, style="Horizontal.TProgressbar")
    progress_bar.pack(fill="x", pady=5)
    
    # Actions
    btn_frame = ttk.Frame(progress_frame)
    btn_frame.pack(pady=10)
    
    btn_start = ttk.Button(btn_frame, text="DÉMARRER LA MIGRATION", width=25)
    btn_start.pack(side="left", padx=10)
    
    btn_stop = ttk.Button(btn_frame, text="METTRE EN PAUSE / ARRÊTER", state="disabled", width=25)
    btn_stop.pack(side="left", padx=10)
    
    # Thread Callback updates
    def update_progress(current, total, status_msg):
        def ui_update():
            pct = (current / total) * 100 if total > 0 else 0
            progress_var.set(pct)
            status_var.set(f"{status_msg} ({pct:.1f}%)")
        window.after(0, ui_update)
        
    def run_migration_process():
        global _shutdown_requested
        _shutdown_requested = False
        
        mbox_path = mbox_path_var.get().strip()
        pst_path = pst_path_var.get().strip()
        folder_name = folder_var.get().strip()
        resume = resume_var.get()
        
        limit_val = None
        if limit_var.get().strip():
            try:
                limit_val = int(limit_var.get().strip())
            except ValueError:
                window.after(0, lambda: messagebox.showerror("Erreur de Saisie", "La limite de messages doit être un entier numérique."))
                reset_inputs()
                return

        if not mbox_path or not pst_path:
            window.after(0, lambda: messagebox.showerror("Champs Requis", "Veuillez spécifier le fichier/dossier MBOX et le fichier PST."))
            reset_inputs()
            return
            
        logging.info("Starting background migration thread...")
        
        try:
            mbox_to_pst(mbox_path, pst_path, folder_name, resume, limit_val, progress_callback=update_progress)
        except Exception as thr_err:
            logging.critical(f"Fatal migration error: {thr_err}", exc_info=True)
            window.after(0, lambda: messagebox.showerror("Erreur Critique", f"Une erreur critique est survenue dans la migration :\n{thr_err}"))
            
        def clean_up_finished():
            btn_start.configure(state="normal")
            btn_stop.configure(state="disabled")
            mbox_entry.configure(state="normal")
            pst_entry.configure(state="normal")
            folder_entry.configure(state="normal")
            limit_entry.configure(state="normal")
            
            if _shutdown_requested:
                status_var.set("Migration suspendue. Progression sauvegardée.")
                messagebox.showinfo("Migration suspendue", "La migration a été arrêtée proprement. Vous pourrez la reprendre au même endroit.")
            else:
                progress_var.set(100)
                status_var.set("Migration complétée avec succès !")
                messagebox.showinfo("Migration Terminée", "Félicitations, la conversion de vos messages est terminée !")
                
        window.after(0, clean_up_finished)
        
    def reset_inputs():
        btn_start.configure(state="normal")
        btn_stop.configure(state="disabled")
        mbox_entry.configure(state="normal")
        pst_entry.configure(state="normal")
        folder_entry.configure(state="normal")
        limit_entry.configure(state="normal")
        status_var.set("Prêt.")
        
    def handle_start():
        btn_start.configure(state="disabled")
        btn_stop.configure(state="normal")
        mbox_entry.configure(state="disabled")
        pst_entry.configure(state="disabled")
        folder_entry.configure(state="disabled")
        limit_entry.configure(state="disabled")
        
        status_var.set("Préparation de la migration...")
        progress_var.set(0)
        
        if not THREADING_AVAILABLE:
            messagebox.showerror("Threading non disponible", "Le package threading est manquant, impossible de lancer le processus.")
            reset_inputs()
            return
            
        thread = threading.Thread(target=run_migration_process, daemon=True)
        thread.start()
        
    def handle_stop():
        global _shutdown_requested
        _shutdown_requested = True
        status_var.set("Demande d'interruption envoyée...")
        btn_stop.configure(state="disabled")
        
    btn_start.configure(command=handle_start)
    btn_stop.configure(command=handle_stop)
    
    # Center window
    window.update_idletasks()
    w = window.winfo_width()
    h = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (w // 2)
    y = (window.winfo_screenheight() // 2) - (h // 2)
    window.geometry(f'{w}x{h}+{x}+{y}')
    
    window.mainloop()

# ==============================================================================
# ENTRY POINT
# ==============================================================================

if __name__ == "__main__":
    # If no arguments provided, or if --gui argument is present, launch the graphical UI
    if len(sys.argv) == 1:
        run_gui()
    else:
        import argparse
        parser = argparse.ArgumentParser(description="Migration MBOX Gmail vers Outlook PST avec Catégories")
        parser.add_argument("mbox", nargs="?", default=None, help="Chemin du fichier .mbox ou dossier de fichiers .mbox")
        parser.add_argument("pst", nargs="?", default=None, help="Chemin du fichier .pst de destination")
        parser.add_argument("--folder", default="Gmail Archive", help="Nom du dossier racine cible dans Outlook")
        parser.add_argument("--no-resume", action="store_false", dest="resume", help="Ne pas reprendre la migration précédente")
        parser.add_argument("--limit", type=int, default=None, help="Limiter le nombre de messages à traiter (pour test)")
        parser.add_argument("--gui", action="store_true", help="Lancer l'interface graphique utilisateur")
        
        args = parser.parse_args()
        
        if args.gui or (not args.mbox or not args.pst):
            run_gui()
        else:
            mbox_to_pst(args.mbox, args.pst, args.folder, args.resume, args.limit)
