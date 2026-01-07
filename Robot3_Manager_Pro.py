import pandas as pd
import datetime
import os
import win32com.client
from docxtpl import DocxTemplate

# 1. SETUP & CONFIGURATION
EXCEL_FILE = "Final_Report.xlsx"
TEMPLATE_FILE = "Offer_Template.docx"
TEST_LINK = "Paste Google form here" # <--- Put your link here

# 2. DEFINING THE RULES
# High Score (8-10) -> Get Offer
# Medium Score (6-7) -> Get Test Link
# Low Score (0-5) -> Get Rejection
HIGH_THRESHOLD = 8
MEDIUM_THRESHOLD = 6 

print("Starting Robot 3: The Advanced Manager...")

# 3. Load Data
if not os.path.exists(EXCEL_FILE):
    print("Error: Report file not found. Run Robot 2 first!")
    exit()

df = pd.read_csv(EXCEL_FILE) if EXCEL_FILE.endswith('csv') else pd.read_excel(EXCEL_FILE)
outlook = win32com.client.Dispatch("Outlook.Application")

# 4. Process Each Candidate
for index, row in df.iterrows():
    name = row['Name']
    email = row['Email']
    score = row['Score']
    
    print(f"Processing: {name} (Score: {score})...")
    
    try:
        mail = outlook.CreateItem(0)
        mail.To = email
        
        # === SCENARIO 1: HIGH SCORE (OFFER) ===
        if score >= HIGH_THRESHOLD:
            print(f"   -> Result: STAR CANDIDATE. Generating Offer...")
            
            # Create Word Doc
            doc = DocxTemplate(TEMPLATE_FILE)
            context = {'Name': name, 'Score': score, 'today_date': datetime.date.today().strftime("%B %d, %Y")}
            doc.render(context)
            
            # Save File
            output_filename = f"Offer_{name.replace(' ', '_')}.docx"
            output_path = os.path.join(os.getcwd(), output_filename)
            doc.save(output_path)
            
            # Email Content
            mail.Subject = "Congratulations! Job Offer Enclosed"
            mail.Body = f"Dear {name},\n\nWe were impressed by your profile (Score: {score}/10).\nPlease find your official offer letter attached.\n\nBest,\nHR Team"
            mail.Attachments.Add(output_path)
            mail.Send()
            print("   -> Offer Email Sent!")

        # === SCENARIO 2: MEDIUM SCORE (SEND TEST) ===
        elif score >= MEDIUM_THRESHOLD:
            print(f"   -> Result: POTENTIAL. Sending Test Link...")
            
            mail.Subject = "Next Steps: Technical Assessment"
            mail.Body = f"""Dear {name},

Thank you for your application. Your profile looks promising.

To proceed to the next stage, please complete our technical assessment at the link below:
{TEST_LINK}

Good luck!
The HR Team"""
            mail.Send()
            print("   -> Test Invite Sent!")

        # === SCENARIO 3: LOW SCORE (REJECT) ===
        else:
            print(f"   -> Result: REJECT. Sending polite notice...")
            
            mail.Subject = "Update on your application"
            mail.Body = f"Dear {name},\n\nThank you for applying. Unfortunately, we have decided to move forward with other candidates who more closely match our specific requirements at this time.\n\nWe wish you the best in your search.\n\nSincerely,\nHR Team"
            mail.Send()
            print("   -> Rejection Email Sent.")

    except Exception as e:
        print(f"   -> ERROR sending email: {e}")

print("\nManager Robot Finished.")