# 📬 Bulk Email Automation Tool — Quick Start Guide

Welcome! This tool allows you to easily extract email lists and securely send personalized bulk emails using Python and an Excel spreadsheet. Follow this simple guide to set up your files and run the automation engine.

---

## 📁 1. Project Directory Layout

Ensure your project folder contains these files in the root directory:
* `requirements.txt` — Holds the required libraries.
* `send_emails.py` — The core script that handles email dispatch.
* `emails.xlsx` — Your Excel file tracking the email database.

---

## 📊 2. Preparing the Spreadsheet (`emails.xlsx`)

The script relies on an Excel sheet named exactly `emails.xlsx`. Before running the tool, create or modify your spreadsheet with the following layout:

* **Column A (Cell A1):** Must be titled exactly **`Email Address`**.
* **Column B (Cell B1):** Must be titled exactly **`Status`**.
* **Rows below Column A:** Paste the target email addresses you want to reach.
* **Rows below Column B:** You **must** type **`Pending`** next to every email address you want to send an email to.

### Data Structure Blueprint:
```text
+-------------------------+---------+
| Email Address           | Status  |
+-------------------------+---------+
| john.doe@example.com    | Pending |
| jane.smith@domain.org   | Pending |
+-------------------------+---------+

Google Account Configuration
To allow the Python script to log into your Gmail account, your regular email password will not work due to Google's security blocks. You must generate a temporary 16-character App Password:
Go to your Google Account Settings (https://myaccount.google.com/security).
Turn ON 2-Step Verification.
Go to the App Passwords creation section (https://myaccount.google.com/apppasswords).
Generate a new App Password and name it something recognizable like "Bulk Email Engine".
Copy the generated 16-character password.
Update the Script:
Open the send_emails.py file in a text editor and enter your email address and generated App Password in the configuration block at the top:
Python
GMAIL_ADDRESS      = "your.email@gmail.com"     # Your Gmail address
GMAIL_APP_PASSWORD = "abcd efgh ijkl mnop"    # Your 16-character App Password (remove spaces)
GOOGLE_FORM_LINK   = "[https://forms.gle/your-link](https://forms.gle/your-link)" # Optional: Paste your form/feedback link
💻 4. Terminal Commands to Execute
Open your Terminal (macOS/Linux) or Command Prompt/PowerShell (Windows) and run these exact commands in order to execute the bulk send:
Bash
# 1. Navigate to your project folder workspace
cd /path/to/your/email_atomation

# 2. Create a clean virtual environment sandbox
python3 -m venv .venv

# 3. Activate the virtual environment
# On Mac/Linux:
source .venv/bin/activate
# On Windows (PowerShell):
.\.venv\Scripts\activate

# 4. Install the mandatory dependencies
pip install -r requirements.txt

# 5. Kick off the bulk automated sending process
python send_emails.py
