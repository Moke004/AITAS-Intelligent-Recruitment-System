import os
import pdfplumber
import pandas as pd
import time
import json
from google import genai
from google.genai import types

# --- CONFIGURATION ---
API_KEY = "PASTE_YOUR_KEY_HERE"  # <--- Check your key is here

FOLDER_PATH = os.path.join(os.getcwd(), "Resumes")
OUTPUT_FILE = "Final_Report.xlsx"
# ---------------------

def ask_gemini_sdk(client, text_content):
    # --- UPDATED PROMPT FOR QA ROLE ---
    prompt = f"""
    You are an expert Technical Recruiter hiring for a 'Quality Assurance (QA) Analyst' role.
    
    Analyze the resume text below.
    Extract the Candidate Name and Email.
    
    Then, give a 'Score' from 1 to 10 based on these QA criteria:
    - Experience with Testing (Manual or Automated).
    - Knowledge of tools like Selenium, JIRA, Cypress, or Postman.
    - Attention to detail and bug reporting skills.
    - Knowledge of the Software Development Life Cycle (SDLC).
    
    Return the answer ONLY as a JSON object like this:
    {{
        "Name": "John Doe",
        "Email": "john@example.com",
        "Score": 8,
        "Reason": "Has strong experience with Selenium and JIRA."
    }}
    
    Resume Text:
    {text_content}
    """
    
    try:
        response = client.models.generate_content(
            model='gemini-flash-latest', 
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema={
                    "type": "OBJECT",
                    "properties": {
                        "Name": {"type": "STRING"},
                        "Email": {"type": "STRING"},
                        "Score": {"type": "INTEGER"},
                        "Reason": {"type": "STRING"}
                    }
                }
            )
        )
        return json.loads(response.text)
        
    except Exception as e:
        print(f"   -> AI Error: {e}")
        return None

# --- MAIN PROCESS ---
print("--- STARTING ROBOT 2: QA SPECIALIST ---")
results = []

client = genai.Client(api_key=API_KEY)

if os.path.exists(FOLDER_PATH):
    files = [f for f in os.listdir(FOLDER_PATH) if f.lower().endswith('.pdf')]
    print(f"Found {len(files)} resumes.")

    for filename in files:
        print(f"Processing: {filename}...")
        filepath = os.path.join(FOLDER_PATH, filename)
        
        full_text = ""
        try:
            with pdfplumber.open(filepath) as pdf:
                for page in pdf.pages:
                    extract = page.extract_text()
                    if extract: full_text += extract + "\n"
        except:
            print("   -> Error reading PDF.")
            continue

        if len(full_text) > 50: 
            data = ask_gemini_sdk(client, full_text[:4000])
            
            if data:
                data['Filename'] = filename
                results.append(data)
                print(f"   -> SUCCESS! QA Score: {data.get('Score')} ({data.get('Name')})")
                time.sleep(10) 
        else:
            print("   -> PDF text was empty.")

    if results:
        df = pd.DataFrame(results)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\nReport saved to: {OUTPUT_FILE}")
    else:
        print("\nNo results.")
else:
    print("Error: 'Resumes' folder not found.")