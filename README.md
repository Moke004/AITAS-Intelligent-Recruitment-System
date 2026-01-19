# AITAS: Automated Intelligent Talent Acquisition System

**Master's Project - Intelligent Business Processes**
*TechFlow Solutions Case Study*

## 📌 Overview
AITAS is a Python-based automation pipeline designed to reduce recruitment screening time by 95%. It orchestrates 4 distinct robots to download resumes, analyze them using Google Gemini AI, and automate decision-making (Hiring vs. Testing vs. Rejection).

## 🚀 Features
- **Robot 1 (Collector):** Auto-downloads PDF resumes from Outlook.
- **Robot 2 (Analyst):** Uses **Google Gemini 2.0 Flash** to grade candidates (1-10) on QA skills (Selenium, JIRA, etc.) and returns structured JSON.
- **Robot 3 (Manager):** Auto-generates contract offers (Word) or sends Google Form test links based on AI score.
- **Robot 4 (Onboarder):** Processes test results and triggers onboarding emails.

## 🛠️ Tech Stack
- **Language:** Python 3.14
- **AI Engine:** Google Gemini API (Generative AI)
- **Integration:** Microsoft Outlook (Win32Com), Excel (Pandas)
- **Docs:** `docxtpl` for dynamic contract generation

## ⚙️ Setup & Installation

1. **Clone the repo:**
   ```bash
   git clone [https://github.com/YOUR_USERNAME/AITAS-System.git](https://github.com/YOUR_USERNAME/AITAS-System.git)
   
To cite this content, please use:
   ```bash
   @misc{AITAS,
    author       = {Moke Dara},
    title        = {AITAS},
    howpublished = 
    year         = {2026}
   }
