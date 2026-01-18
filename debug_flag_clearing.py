import win32com.client
import os
import datetime

def test_flag_clearing():
    pst_path = r"E:\Sauveguarde_Messages_GMAIL\Takeout\Mail\archive_outlook.pst"
    folder_name = "Test_Flags_Debug"
    
    print(f"Testing Flag Clearing in: {pst_path}")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Open PST
        pst_store = None
        for store in namespace.Stores:
            if store.FilePath.lower() == pst_path.lower():
                pst_store = store
                print(f"Found PST: {store.DisplayName}")
                break
        
        if not pst_store:
            print("PST not found in session, adding it...")
            namespace.AddStore(pst_path)
            pst_store = namespace.Stores.Item(namespace.Stores.Count)
            
        root = pst_store.GetRootFolder()
        
        # Create Test Folder
        try:
            target_folder = root.Folders(folder_name)
        except:
            target_folder = root.Folders.Add(folder_name)
            
        print(f"Target Folder: {target_folder.Name}")
        
        # TEST 1: Standard (Save then Props) - Mimics current code (roughly)
        print("Creating Item 1: Save First (Control)")
        mail1 = target_folder.Items.Add("IPM.Note")
        mail1.Subject = "Test 1 - Save THEN Props"
        mail1.Save() # First save locks flags?
        
        try:
            # Set Flags
            mail1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1) # MSGFLAG_READ
            # Set Sender
            mail1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001F", "Sender 1")
            mail1.Save()
        except Exception as e:
            print(f"Error Item 1: {e}")

        # TEST 2: Props BEFORE Save
        print("Creating Item 2: Props BEFORE Save (Experiment)")
        mail2 = target_folder.Items.Add("IPM.Note")
        mail2.Subject = "Test 2 - Props BEFORE Save"
        # NO SAVE HERE
        
        try:
            # Set Flags
            # PR_MESSAGE_FLAGS = 1 (Read, Sent)
            mail2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1)
            # PR_ICON_INDEX = -1
            mail2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x10800003", -1)
            
            # Set Sender
            mail2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001F", "Sender 2")
            mail2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001F", "sender2@test.com")
            mail2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1E001F", "SMTP")
            
            # NOW Save
            mail2.Save()
        except Exception as e:
            print(f"Error Item 2: {e}")
            
        print("Test Complete. Please check Outlook folder 'Test_Flags_Debug'.")
        print("Item 1 should be Draft. Item 2 should hopefully be Sent/Received.")

    except Exception as e:
        print(f"Global Error: {e}")

if __name__ == "__main__":
    test_flag_clearing()
