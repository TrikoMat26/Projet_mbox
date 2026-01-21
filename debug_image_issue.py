"""
Debug script to analyze a specific email with truncated inline images.
This will extract the raw MIME structure and image data for investigation.
"""
import os
import re
from email import message_from_bytes
from email.header import decode_header
import base64

# Target email to find
TARGET_SUBJECT = "Re: sav - stago"
TARGET_DATE = "2024"  # Partial match
TARGET_SENDER = "krikor.kayzakian"

MBOX_PATH = r"E:\Sauveguarde_Messages_GMAIL\Tous les messages, y compris ceux du dossier Spam -002.mbox"
OUTPUT_DIR = r"E:\Sauveguarde_Messages_GMAIL\debug_output"

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

def analyze_mbox():
    print(f"Searching in: {MBOX_PATH}")
    print(f"Looking for: {TARGET_SUBJECT}")
    
    mbox_from_pattern = re.compile(rb'^From \S+ .+\r?\n', re.MULTILINE)
    
    with open(MBOX_PATH, 'rb') as f:
        buffer = b''
        chunk_size = 10 * 1024 * 1024  # 10 MB
        message_index = 0
        current_message_data = b''
        found = False
        
        while not found:
            chunk = f.read(chunk_size)
            if not chunk and not buffer:
                break
            
            buffer += chunk
            matches = list(mbox_from_pattern.finditer(buffer))
            
            if len(matches) == 0:
                if not chunk:
                    # Last message
                    if buffer.strip():
                        current_message_data = buffer
                        try:
                            msg = message_from_bytes(current_message_data)
                            found = check_and_analyze(msg, message_index, current_message_data)
                        except:
                            pass
                    break
                continue
            
            for i, match in enumerate(matches):
                if i == 0:
                    if current_message_data:
                        try:
                            msg = message_from_bytes(current_message_data)
                            found = check_and_analyze(msg, message_index, current_message_data)
                            if found:
                                break
                        except Exception as e:
                            print(f"Parse error: {e}")
                        message_index += 1
                    
                    if i + 1 < len(matches):
                        current_message_data = buffer[match.end():matches[i + 1].start()]
                    else:
                        current_message_data = buffer[match.end():]
                else:
                    try:
                        msg = message_from_bytes(current_message_data)
                        found = check_and_analyze(msg, message_index, current_message_data)
                        if found:
                            break
                    except:
                        pass
                    message_index += 1
                    
                    if i + 1 < len(matches):
                        current_message_data = buffer[match.end():matches[i + 1].start()]
                    else:
                        current_message_data = buffer[match.end():]
            
            if found:
                break
                
            # Keep only the last partial message
            if matches:
                buffer = buffer[matches[-1].start():]
            
            if message_index % 1000 == 0:
                print(f"Scanned {message_index} messages...")
    
    if not found:
        print("Target message not found!")

def check_and_analyze(msg, index, raw_data):
    subject = decode_mime_header(msg.get('subject', ''))
    sender = msg.get('from', '')
    date = msg.get('date', '')
    
    if TARGET_SUBJECT.lower() in subject.lower() and TARGET_SENDER in sender.lower():
        print(f"\n{'='*60}")
        print(f"FOUND TARGET MESSAGE at index {index}")
        print(f"Subject: {subject}")
        print(f"From: {sender}")
        print(f"Date: {date}")
        print(f"Raw message size: {len(raw_data)} bytes")
        print(f"{'='*60}\n")
        
        # Save raw message
        raw_path = os.path.join(OUTPUT_DIR, "raw_message.eml")
        with open(raw_path, 'wb') as f:
            f.write(raw_data)
        print(f"Saved raw message to: {raw_path}")
        
        # Analyze MIME structure
        analyze_mime_structure(msg)
        return True
    return False

def analyze_mime_structure(msg, level=0):
    indent = "  " * level
    content_type = msg.get_content_type()
    filename = msg.get_filename()
    content_id = msg.get('Content-ID')
    transfer_encoding = msg.get('Content-Transfer-Encoding', 'none')
    disposition = msg.get('Content-Disposition', 'none')
    
    print(f"{indent}Part: {content_type}")
    if filename:
        print(f"{indent}  Filename: {filename}")
    if content_id:
        print(f"{indent}  Content-ID: {content_id}")
    print(f"{indent}  Transfer-Encoding: {transfer_encoding}")
    print(f"{indent}  Disposition: {disposition}")
    
    if msg.is_multipart():
        for i, part in enumerate(msg.get_payload()):
            print(f"{indent}  --- Subpart {i+1} ---")
            analyze_mime_structure(part, level + 1)
    else:
        # Get payload info
        payload = msg.get_payload(decode=False)
        if isinstance(payload, bytes):
            print(f"{indent}  Payload size (raw bytes): {len(payload)}")
        elif isinstance(payload, str):
            print(f"{indent}  Payload size (raw string): {len(payload)} chars")
        
        # Try to decode
        decoded = msg.get_payload(decode=True)
        if decoded:
            print(f"{indent}  Decoded size: {len(decoded)} bytes")
            
            # If it's an image, save it
            if content_type.startswith('image/'):
                ext = content_type.split('/')[1]
                safe_filename = filename or f"image_{id(msg)}.{ext}"
                safe_filename = re.sub(r'[<>:"/\\|?*]', '_', safe_filename)
                img_path = os.path.join(OUTPUT_DIR, safe_filename)
                with open(img_path, 'wb') as f:
                    f.write(decoded)
                print(f"{indent}  SAVED IMAGE: {img_path}")
        else:
            print(f"{indent}  WARNING: Could not decode payload!")
            # Try manual decoding
            if transfer_encoding.lower() == 'base64' and payload:
                try:
                    if isinstance(payload, str):
                        payload_bytes = payload.encode('ascii', errors='ignore')
                    else:
                        payload_bytes = payload
                    manual_decoded = base64.b64decode(payload_bytes)
                    print(f"{indent}  Manual base64 decode: {len(manual_decoded)} bytes")
                except Exception as e:
                    print(f"{indent}  Manual decode FAILED: {e}")

if __name__ == "__main__":
    analyze_mbox()
    print(f"\nOutput files saved to: {OUTPUT_DIR}")
