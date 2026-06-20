# Email Tool — Extract & Send

A minimal web tool to extract emails from PDFs/Excel files and send personalised emails via Gmail.

---

## Quick Start (5 steps)

### Step 1 — Install dependencies
```bash
pip install pypdf pandas openpyxl flask flask-cors
```

### Step 2 — Set up your Gmail App Password
1. Enable 2-Step Verification at https://myaccount.google.com/security
2. Go to https://myaccount.google.com/apppasswords
3. Create a new App Password (name it "Email Tool")
4. Copy the 16-character password

### Step 3 — Configure server.py
Open `server.py` and edit the top block:
```python
GMAIL_ADDRESS      = "your@gmail.com"
GMAIL_APP_PASSWORD = "abcd efgh ijkl mnop"   # your 16-char app password
DAILY_LIMIT        = 500
DELAY_SECONDS      = 5
```

### Step 4 — Start the server
```bash
python server.py
```
You'll see: `🚀  Email Tool server running → http://localhost:5000`

### Step 5 — Open the web UI
Open `index.html` directly in your browser (double-click it), or serve it:
```bash
# Option A: just double-click index.html in your file manager
# Option B: Python simple server
python -m http.server 8080
# then go to http://localhost:8080
```

---

## How to use

### Extract tab
- Drop a PDF or Excel file onto the upload zone
- Emails are extracted automatically and added to `emails.xlsx`
- Duplicates and already-sent emails are never added again

### Compose & Send tab
- Write your subject and message body
- Use `{name}` — auto-filled from the recipient's email username
- Use `{form_link}` — replaced with your Google Form URL
- Click "Start sending" — progress updates live on screen
- If the script stops, re-run it — already-sent emails are skipped

### Email List tab
- See every email with its current status (Pending / Sent / Failed)
- Filter and search
- Download the Excel file at any time

---

## Command-line scripts (standalone, no web UI needed)

```bash
# Extract emails from a PDF
python extract.py --input myfile.pdf --output emails.xlsx

# Send emails
python send_emails.py
```

---

## File structure
```
email-tool/
├── index.html        ← Web UI (open in browser)
├── server.py         ← Flask backend (run this)
├── extract.py        ← Standalone PDF extractor
├── send_emails.py    ← Standalone email sender
├── requirements.txt  ← pip dependencies
└── emails.xlsx       ← Created automatically
```

## Safety notes
- Gmail personal accounts allow ~500 emails/day — the tool enforces this limit
- A 5-second delay between sends mimics human behaviour and reduces spam risk
- Each sent email is saved to Excel immediately — no duplicates ever
- Failed emails are marked "Failed: reason" and can be retried
