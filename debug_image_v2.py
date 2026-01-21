"""
Debug script using standard mailbox library to compare extraction methods.
"""
import mailbox
import os
import re
from email.header import decode_header

TARGET_SUBJECT = "Re: sav - stago"
TARGET_SENDER = "krikor.kayzakian"
MBOX_PATH = r"E:\Sauveguarde_Messages_GMAIL\Tous les messages, y compris ceux du dossier Spam -002.mbox"
OUTPUT_DIR = r"E:\Sauveguarde_Messages_GMAIL\debug_output_v2"

os.makedirs(OUTPUT_DIR, exist_ok=True)

def decode_mime_header(header_value):
    if not header_value:
        return ""
    try:
        decoded = decode_header(header_value)
        result = []
        for part, encoding in decoded:
            if isinstance(part, bytes):
                result.append(part.decode(encoding or 'utf-8', errors='replace'))
            else:
                result.append(part)
        return "".join(result)
    except:
        return str(header_value)

def main():
    print(f"Opening MBOX with standard mailbox library: {MBOX_PATH}")
    print(f"Looking for: {TARGET_SUBJECT}")
    print()
    
    mbox = mailbox.mbox(MBOX_PATH)
    
    found = False
    for i, message in enumerate(mbox):
        if i % 1000 == 0:
            print(f"Scanned {i} messages...")
        
        subject = decode_mime_header(message.get('subject', ''))
        sender = message.get('from', '')
        
        if TARGET_SUBJECT.lower() in subject.lower() and TARGET_SENDER in sender.lower():
            print(f"\n{'='*60}")
            print(f"FOUND at index {i}")
            print(f"Subject: {subject}")
            print(f"From: {sender}")
            print(f"{'='*60}\n")
            
            # Save raw message
            raw_path = os.path.join(OUTPUT_DIR, "raw_message_v2.eml")
            with open(raw_path, 'wb') as f:
                f.write(bytes(message))
            print(f"Raw message saved: {raw_path}")
            print(f"Raw size: {os.path.getsize(raw_path):,} bytes")
            
            # Extract all parts
            extract_parts(message)
            found = True
            break
    
    mbox.close()
    
    if not found:
        print("Message not found!")
    else:
        print(f"\nOutput saved to: {OUTPUT_DIR}")

def extract_parts(message, prefix=""):
    if message.is_multipart():
        for i, part in enumerate(message.get_payload()):
            extract_parts(part, f"{prefix}part{i+1}_")
    else:
        content_type = message.get_content_type()
        filename = message.get_filename()
        content_id = message.get('Content-ID')
        transfer_encoding = message.get('Content-Transfer-Encoding', 'none')
        
        print(f"{prefix}Type: {content_type}")
        if filename:
            print(f"{prefix}  Filename: {filename}")
        if content_id:
            print(f"{prefix}  CID: {content_id}")
        print(f"{prefix}  Encoding: {transfer_encoding}")
        
        # Get raw payload first
        raw_payload = message.get_payload(decode=False)
        raw_size = len(raw_payload) if raw_payload else 0
        print(f"{prefix}  Raw payload: {raw_size} chars")
        
        # Decode
        payload = message.get_payload(decode=True)
        if payload:
            print(f"{prefix}  Decoded: {len(payload):,} bytes")
            
            if content_type.startswith('image/'):
                ext = content_type.split('/')[1]
                safe_name = filename or f"image.{ext}"
                safe_name = re.sub(r'[<>:"/\\|?*]', '_', safe_name)
                out_path = os.path.join(OUTPUT_DIR, f"{prefix}{safe_name}")
                with open(out_path, 'wb') as f:
                    f.write(payload)
                print(f"{prefix}  SAVED: {out_path}")
                
                # Verify it's a valid JPEG
                if payload[:2] == b'\xff\xd8':
                    print(f"{prefix}  ✓ Valid JPEG header")
                else:
                    print(f"{prefix}  ✗ INVALID JPEG header!")
                    print(f"{prefix}    First 20 bytes: {payload[:20]}")
        else:
            print(f"{prefix}  ✗ DECODE FAILED!")
        print()

if __name__ == "__main__":
    main()
