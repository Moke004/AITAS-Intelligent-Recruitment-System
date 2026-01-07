import pandas as pd
import win32com.client
import os

# --- CONFIGURATION ---
# 1. Download your Google Form results as an Excel file
# 2. Save it in your folder as "Quiz_Results.xlsx"
INPUT_FILE = "Quiz_Results.xlsx"
PASS_MARK = 70
# ---------------------

print("--- ROBOT 4: THE FINAL JUDGE ---")

if not os.path.exists(INPUT_FILE):
    print(f"Error: Please download the Google Sheet as '{INPUT_FILE}' first.")
    exit()

# Load Data
df = pd.read_excel(INPUT_FILE)
outlook = win32com.client.Dispatch("Outlook.Application")

print(f"Checking {len(df)} quiz submissions...")

for index, row in df.iterrows():
    # Google Forms columns are usually "Email Address" and "Score"
    email = row.get('Email Address')
    raw_score = row.get('Score')
    name = row.get('Name') # Make sure your Google Form asks for "Name"!

    # Clean up the score (e.g., turn "80 / 100" into just 80)
    try:
        if isinstance(raw_score, str):
            score = int(raw_score.split('/')[0].strip())
        else:
            score = int(raw_score)
    except:
        continue

    # DECISION TIME
    try:
        mail = outlook.CreateItem(0)
        mail.To = email
        
        if score >= PASS_MARK:
            print(f"   -> HIRED: {name} (Score: {score})")
            
            # 1. Email the Candidate
            mail.Subject = "Welcome to the Team! (Offer Confirmed)"
            mail.Body = f"""Dear {name},

Congratulations! You scored {score} on our technical assessment, which is above our pass mark of {PASS_MARK}.

We are thrilled to offer you the position. 
Please reply to this email to schedule your first day.

Welcome aboard!
"""
            mail.Send()
            
            # 2. Email the Team (You)
            team_mail = outlook.CreateItem(0)
            team_mail.To = "your-email@example.com" # <--- Change this
            team_mail.Subject = f"NEW HIRE ALERT: {name}"
            team_mail.Body = f"{name} just passed the test with a score of {score}. Please prepare their laptop."
            team_mail.Send()
            
        else:
            print(f"   -> REJECTED: {name} (Score: {score})")
            
            mail.Subject = "Update on your application"
            mail.Body = f"""Dear {name},

Thank you for taking our technical assessment. 
Unfortunately, your score of {score} did not meet our threshold of {PASS_MARK} for this specific role.

We will keep your resume on file for future openings.
"""
            mail.Send()

    except Exception as e:
        print(f"   -> Error emailing {email}: {e}")

print("Done! All emails sent.")