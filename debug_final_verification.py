import win32com.client
import os
import datetime
import time

def test_final_verification():
    pst_path = r"E:\Sauveguarde_Messages_GMAIL\Takeout\Mail\archive_outlook.pst"
    folder_name = "Test_Final_Fix"
    
    print(f"Testing Final Fix in: {pst_path}")
    
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
            pst_store = namespace.Stores.Item(namespace.Stores.Count)
            
        root = pst_store.GetRootFolder()
        
        # Ensure Target Folder Exists
        try:
            target_folder = root.Folders(folder_name)
        except:
            target_folder = root.Folders.Add(folder_name)
            
        print(f"Target Folder: {target_folder.Name}")
        
        # Use a temporary folder (Inbox or Drafts of the PST if possible, or Default Inbox)
        # Using Default Inbox to simulate "New Email" creation context
        temp_creation_folder = namespace.GetDefaultFolder(6) # Inbox
        
        print("Creating Item with PROPS BEFORE SAVE...")
        mail = temp_creation_folder.Items.Add("IPM.Note")
        mail.Subject = "Final Verification - Props BEFORE Save"
        mail.Body = "This email should be 'Received' and have a sender."
        
        # --- CRITICAL SECTION ---
        # Set ALL Properties explicitly BEFORE the first Save()
        
        # 1. Clear UNSENT flag (Draft)
        # PR_MESSAGE_FLAGS = 1 (Read)
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1)
        
        # 2. Reset Icon
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x10800003", -1)
        
        # 3. Set Sender
        sender_name = "Test Sender"
        sender_email = "test.sender@example.com"
        
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001F", sender_name) # PR_SENDER_NAME
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001F", sender_email) # PR_SENDER_EMAIL_ADDRESS
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1E001F", "SMTP")       # PR_SENDER_ADDRTYPE
        
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0042001F", sender_name) # PR_SENT_REPRESENTING_NAME
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0065001F", sender_email) # PR_SENT_REPRESENTING_EMAIL_ADDRESS
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0064001F", "SMTP")       # PR_SENT_REPRESENTING_ADDRTYPE
        
        # 4. Set Time
        now = datetime.datetime.now()
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x00390040", now)
        mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E060040", now)
        
        # --- SAVE ---
        print("Saving (First Time)...")
        mail.Save()
        
        # --- MOVE ---
        print("Moving to target folder...")
        my_item = mail.Move(target_folder)
        
        print(f"Move complete. Item Subject: {my_item.Subject}")
        print(f"Please check folder: {folder_name}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_final_verification()
