import win32com.client
import os

# --- CONFIGURATION ---
# What word must be in the subject? (e.g., "Resume", "HR-2025")
TARGET_CODE = "Resume" 
# ---------------------

# 1. Setup paths
output_folder = os.path.join(os.getcwd(), "Resumes")
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

print(f"--- STARTING BULK SCAN FOR '{TARGET_CODE}' ---")

# 2. Connect to Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) # Inbox
    messages = inbox.Items
    
    # Sort by newest first
    messages.Sort("[ReceivedTime]", True)

    print(f"Connected to Inbox. Scanning {messages.Count} total emails...")
    
    count = 0
    found_emails = 0

    # 3. Loop through ALL emails (No time limit)
    for message in messages:
        try:
            subject = message.Subject
            
            # Check if TARGET_CODE is in the subject (Case Insensitive)
            if TARGET_CODE.lower() in subject.lower():
                found_emails += 1
                print(f"Found match: '{subject}'...")
                
                attachments = message.Attachments
                if attachments.Count > 0:
                    pdf_found = False
                    for i in range(1, attachments.Count + 1):
                        attachment = attachments.Item(i)
                        
                        # Check for PDF
                        if attachment.FileName.lower().endswith(".pdf"):
                            save_path = os.path.join(output_folder, attachment.FileName)
                            # Handle duplicate names (add a number if file exists)
                            if os.path.exists(save_path):
                                base, ext = os.path.splitext(attachment.FileName)
                                save_path = os.path.join(output_folder, f"{base}_{count}{ext}")

                            attachment.SaveAsFile(save_path)
                            print(f"   -> SAVED: {attachment.FileName}")
                            pdf_found = True
                            count += 1
                    
                    if not pdf_found:
                        print("   -> WARNING: Email found, but attachment was NOT a PDF.")
                else:
                    print("   -> WARNING: Email found, but NO attachments.")

        except Exception as e:
            # Skip items that aren't emails (like calendar invites)
            continue

    print(f"------------------------------------------------")
    print(f"Scan Complete.")
    print(f"Emails with '{TARGET_CODE}': {found_emails}")
    print(f"Resumes Saved: {count}")

except Exception as main_error:
    print(f"CRITICAL ERROR: {main_error}")