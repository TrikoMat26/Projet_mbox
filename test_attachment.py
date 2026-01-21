"""
Test script to debug Outlook inline image CID rendering.
Sets all required MAPI properties for proper inline display.
"""
import win32com.client
import os

# Path to the extracted image from debug script
IMAGE_PATH = r"E:\Sauveguarde_Messages_GMAIL\debug_output\SAV STAGO.jpg"
PST_PATH = r"E:\Sauveguarde_Messages_GMAIL\test_attachment.pst"

def test_attachment():
    print("Starting Outlook COM test with full MAPI properties...")
    
    if not os.path.exists(IMAGE_PATH):
        print(f"ERROR: Image not found at {IMAGE_PATH}")
        return
    
    original_size = os.path.getsize(IMAGE_PATH)
    print(f"Original image size: {original_size:,} bytes")
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    pst_abs = os.path.abspath(PST_PATH)
    try:
        namespace.AddStore(pst_abs)
    except:
        pass
    
    pst_store = None
    for store in namespace.Stores:
        try:
            if store.FilePath.lower() == pst_abs.lower():
                pst_store = store
                break
        except:
            continue
    
    if not pst_store:
        print("ERROR: Could not find PST store")
        return
    
    root = pst_store.GetRootFolder()
    
    test_folder = None
    for f in root.Folders:
        if f.Name == "Test Attachments":
            test_folder = f
            break
    if not test_folder:
        test_folder = root.Folders.Add("Test Attachments")
    
    try:
        temp_folder = namespace.GetDefaultFolder(6)
    except:
        temp_folder = test_folder
    
    mail = temp_folder.Items.Add(0)
    mail.Subject = "Test Inline Image - Full MAPI Properties"
    
    print("Adding attachment with full MAPI inline properties...")
    
    # Add attachment - Type 1 = olByValue (copy), Position 0 for inline
    attachment = mail.Attachments.Add(IMAGE_PATH, 1, 0, "SAV_STAGO.jpg")
    
    cid = "testimage123"
    prop_accessor = attachment.PropertyAccessor
    
    # 1. PR_ATTACH_CONTENT_ID (0x3712001F) - Required for CID reference
    try:
        prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
        print(f"  PR_ATTACH_CONTENT_ID set: {cid}")
    except Exception as e:
        print(f"  PR_ATTACH_CONTENT_ID error: {e}")
    
    # 2. PR_ATTACH_FLAGS (0x37140003) - Set ATT_MHTML_REF (0x4) for inline
    try:
        prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003", 4)
        print("  PR_ATTACH_FLAGS set: 4 (ATT_MHTML_REF)")
    except Exception as e:
        print(f"  PR_ATTACH_FLAGS error: {e}")
    
    # 3. PR_ATTACH_MIME_TAG (0x370E001F) - MIME type
    try:
        prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/jpeg")
        print("  PR_ATTACH_MIME_TAG set: image/jpeg")
    except Exception as e:
        print(f"  PR_ATTACH_MIME_TAG error: {e}")
    
    # 4. PR_ATTACHMENT_HIDDEN (0x7FFE000B) - Should be False for visible inline
    try:
        prop_accessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", False)
        print("  PR_ATTACHMENT_HIDDEN set: False")
    except Exception as e:
        print(f"  PR_ATTACHMENT_HIDDEN error: {e}")
    
    # Set HTML body AFTER attachment
    mail.HTMLBody = f"""<html>
<head><meta charset="utf-8"></head>
<body>
<p>Test image with full MAPI properties (should be ~1.6MB):</p>
<img src="cid:{cid}" alt="Test Image" style="max-width:100%;">
<p>If you see the image above, CID rendering works!</p>
</body>
</html>"""
    
    mail.Save()
    print("Mail saved")
    
    mail = mail.Move(test_folder)
    print("Mail moved to test folder")
    
    print(f"\n=== TEST COMPLETE ===")
    print(f"Check the mail 'Test Inline Image - Full MAPI Properties' in Outlook")

if __name__ == "__main__":
    test_attachment()
