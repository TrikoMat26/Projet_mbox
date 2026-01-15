import mailbox
import win32com.client
import os
import sys
import time
from email.utils import parsedate_to_datetime
import tempfile

def mbox_to_pst(mbox_path, pst_path, folder_name="Gmail Archive"):
    if not os.path.exists(mbox_path):
        print(f"Error: MBOX file not found at {mbox_path}")
        return

    # Initialize Outlook
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        return

    # Create/Open PST
    print(f"Opening/Creating PST: {pst_path}")
    try:
        # Check if already added
        pst_store = None
        for store in namespace.Stores:
            if store.FilePath.lower() == os.path.abspath(pst_path).lower():
                pst_store = store
                break
        
        if not pst_store:
            namespace.AddStore(os.path.abspath(pst_path))
            # Find it again
            for store in namespace.Stores:
                if store.FilePath.lower() == os.path.abspath(pst_path).lower():
                    pst_store = store
                    break
        
        if not pst_store:
             # Try finding by name if AddStore returned but store list is slow to update
             time.sleep(1)
             for store in namespace.Stores:
                if store.FilePath.lower() == os.path.abspath(pst_path).lower():
                    pst_store = store
                    break
                    
        root_folder = pst_store.GetRootFolder()
    except Exception as e:
        print(f"Error creating PST: {e}")
        return

    # Get or create target folder
    try:
        target_folder = None
        for folder in root_folder.Folders:
            if folder.Name == folder_name:
                target_folder = folder
                break
        if not target_folder:
            target_folder = root_folder.Folders.Add(folder_name)
    except Exception as e:
        print(f"Error creating folder: {e}")
        return

    # Open MBOX
    print(f"Parsing MBOX: {mbox_path}")
    mbox = mailbox.mbox(mbox_path)
    
    count = 0
    start_time = time.time()
    
    for message in mbox:
        try:
            # Extract basic info
            subject = message['subject'] or "(No Subject)"
            sender = message['from'] or ""
            to = message['to'] or ""
            date_str = message['date']
            
            # X-Gmail-Labels
            labels = message.get('X-Gmail-Labels', '')
            categories = ""
            if labels:
                # Clean labels (ignore system labels if desired, or keep all)
                # Gmail labels in MBOX are usually comma separated
                label_list = [l.strip() for l in labels.split(',')]
                categories = ", ".join(label_list)

            # Create MailItem
            mail = target_folder.Items.Add(0) # 0 = olMailItem
            mail.Subject = subject
            mail.SentOnBehalfOfName = sender
            mail.To = to
            
            # Date is tricky via COM (read-only for existing items, but can be set for new ones? No, usually it's set on Send)
            # However, when importing, we might want to preserve it.
            # For PST import, we often have to use MAPI properties to override SentTime.
            
            # Categories
            if categories:
                mail.Categories = categories
            
            # Body handling
            body_html = ""
            body_text = ""
            
            if message.is_multipart():
                for part in message.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        body_text += part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='replace')
                    elif content_type == "text/html" and "attachment" not in content_disposition:
                        body_html += part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='replace')
                    elif "attachment" in content_disposition:
                        # Handle attachment
                        filename = part.get_filename()
                        if filename:
                            with tempfile.NamedTemporaryFile(delete=False) as tf:
                                tf.write(part.get_payload(decode=True))
                                temp_path = tf.name
                            mail.Attachments.Add(temp_path, 1, 1, filename)
                            os.unlink(temp_path)
            else:
                payload = message.get_payload(decode=True).decode(message.get_content_charset() or 'utf-8', errors='replace')
                if message.get_content_type() == "text/html":
                    body_html = payload
                else:
                    body_text = payload

            if body_html:
                mail.HTMLBody = body_html
            else:
                mail.Body = body_text
            
            mail.Save()
            
            count += 1
            if count % 100 == 0:
                elapsed = time.time() - start_time
                rate = count / elapsed
                print(f"Processed {count} messages... ({rate:.2f} msgs/sec)")
                
        except Exception as e:
            print(f"Error processing message {count}: {e}")
            continue

    print(f"Finished! Total messages: {count}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python mbox_to_pst.py <input.mbox> <output.pst>")
    else:
        mbox_to_pst(sys.argv[1], sys.argv[2])
