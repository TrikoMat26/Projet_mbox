import win32com.client
import os
import tempfile
from email.message import EmailMessage

def test_eml_import():
    pst_path = r"E:\Sauveguarde_Messages_GMAIL\Takeout\Mail\archive_outlook.pst" # Update if needed from user history
    print(f"Testing EML Import into: {pst_path}")

    # 1. Create a dummy EML file in a safe temp location
    fd, eml_path = tempfile.mkstemp(suffix=".eml")
    os.close(fd) # Close file descriptor so we can write/read via name
    
    msg = EmailMessage()
    msg["Subject"] = "Test EML Import - Should contain SENDER"
    msg["From"] = "Sender Name <sender@example.com>"
    msg["To"] = "Recipient <recipient@example.com>"
    msg["Date"] = "Mon, 01 Oct 2023 10:00:00 +0000"
    msg.set_content("This is a test message imported via EML format. It should NOT look like a draft.")
    
    with open(eml_path, "wb") as f:
        f.write(msg.as_bytes())
    
    print(f"Created temporary EML: {eml_path}")
    print(f"File exists: {os.path.exists(eml_path)}")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Open PST
        pst_store = None
        for store in namespace.Stores:
            if store.FilePath.lower() == pst_path.lower():
                pst_store = store
                break
        
        if not pst_store:
            namespace.AddStore(pst_path)
            pst_store = namespace.Stores.Item(namespace.Stores.Count) # Get last added
            
        root = pst_store.GetRootFolder()
        
        # 2. Use OpenSharedItem
        print("Opening via OpenSharedItem...")
        # Note: OpenSharedItem returns a MailItem but it is not saved yet (or is 'sent' state depending on source)
        mail = namespace.OpenSharedItem(eml_path)
        
        print(f"Item Class: {mail.MessageClass}")
        print(f"Sent: {mail.Sent}")
        
        # 3. Save and Move
        # We need to move it to the PST root or a folder
        mail.Save() # Saves to default Drafts/Inbox potentially? Or is it just in memory linked to file?
        
        # Usually OpenSharedItem creates an item in a temporary store context.
        # Moving it is the key.
        
        target_folder = root.Folders.Add("Test_EML_Import") if "Test_EML_Import" not in [f.Name for f in root.Folders] else root.Folders["Test_EML_Import"]
        
        mail.Move(target_folder)
        print("Moved to 'Test_EML_Import' folder in PST.")
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if os.path.exists(eml_path):
            os.remove(eml_path)

if __name__ == "__main__":
    test_eml_import()
