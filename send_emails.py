"""
send_emails.py — Smart Bulk Email Sender
=====================================================
Reads the Excel file ("emails.xlsx"), sends a personalised email
to every "Pending" address (up to DAILY_LIMIT per run), and marks each
row "Sent" or "Failed: <reason>" immediately after attempting.

HOW TO RUN:
    python send_emails.py
"""

# ──────────────────────────────────────────────────────────────────
#  CONFIGURATION  — Update your Google credentials here!
# ──────────────────────────────────────────────────────────────────
GMAIL_ADDRESS      = " Gmail address"   # Your Gmail address
GMAIL_APP_PASSWORD = "Google App Password"    # Your 16-char Google App Password
EXCEL_FILE         = "emails.xlsx"              # Your master Excel file
DAILY_LIMIT        = 500                        # Max emails per script run
DELAY_SECONDS      = 5                          # Pause between sends (to prevent spam blocks)

# ── Email Content Configuration ───────────────────────────────────
EMAIL_SUBJECT      = "Help Us Build Something You Loveee 👀✨"
GOOGLE_FORM_LINK   = "https://forms.gle/UpN9xLL7JEoyS41R6"   # ← paste your form link

# Fixed layout payload string structure
EMAIL_BODY_TEMPLATE = """\

ENTER THE MESSAGE HERE 


"""
# ──────────────────────────────────────────────────────────────────

import time
import smtplib
import traceback
from email.message import EmailMessage
from pathlib import Path
import pandas as pd

def username_from_email(email: str) -> str:
    """Derive a clean name from the email address front string."""
    local = email.split("@")[0]
    friendly = local.replace(".", " ").replace("_", " ").replace("-", " ").title()
    return friendly.strip()

def save_df(df: pd.DataFrame, path: Path) -> None:
    """Save changes immediately to disk so no progress is lost."""
    df.to_excel(path, index=False)

def main():
    excel_path = Path(EXCEL_FILE)

    if not excel_path.exists():
        print(f"❌  Excel file not found: {EXCEL_FILE}")
        print("    Please ensure emails.xlsx exists or upload files via the UI first.")
        return

    df = pd.read_excel(excel_path)

    # Validate that columns exist
    if "Email Address" not in df.columns or "Status" not in df.columns:
        print("❌  Excel file is missing required columns: 'Email Address' and 'Status'")
        return

    # Filter to Pending rows
    pending_mask = df["Status"].astype(str).str.strip().str.lower() == "pending"
    pending_df   = df[pending_mask]

    if pending_df.empty:
        print("🎉  All done! No 'Pending' emails remain in your spreadsheet.")
        return

    # Cap at your daily batch limit
    queue = pending_df.head(DAILY_LIMIT)
    total = len(queue)
    print(f"\n📬  {len(pending_df)} pending email(s) found.")
    print(f"📤  Bulk sending started → Processing a batch of {total} email(s).\n")

    sent_count   = 0
    failed_count = 0

    try:
        print("🔐  Connecting to Gmail Secure SMTP Network…")
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
            print("✅  Authentication Successful. Starting bulk dispatch...\n")

            for idx, row in queue.iterrows():
                # Strictly sanitize strings to strip hidden control linebreaks
                recipient = str(row["Email Address"]).strip().replace("\r", "").replace("\n", "")
                name      = username_from_email(recipient)
                body      = EMAIL_BODY_TEMPLATE.format(name=name, form_link=GOOGLE_FORM_LINK)

                msg = EmailMessage()
                # Enforce clean single lines for headers
                msg["From"]    = GMAIL_ADDRESS.strip().replace("\r", "").replace("\n", "")
                msg["To"]      = recipient
                msg["Subject"] = EMAIL_SUBJECT.strip().replace("\r", "").replace("\n", "")
                
                # Explicitly pass structural body content safely
                msg.set_content(body)

                try:
                    smtp.send_message(msg)
                    df.at[idx, "Status"] = "Sent"
                    save_df(df, excel_path)
                    sent_count += 1
                    print(f"  ✅  [{sent_count}/{total}] Sent successfully to → {recipient}")

                    # Anti-spam throttle delay
                    if sent_count < total:
                        time.sleep(DELAY_SECONDS)

                except Exception as e:
                    reason = str(e)[:100]
                    df.at[idx, "Status"] = f"Failed: {reason}"
                    save_df(df, excel_path)
                    failed_count += 1
                    print(f"  ❌  Dispatch Failed to → {recipient} | Reason: {reason}")

    except smtplib.SMTPAuthenticationError:
        print("\n❌  Gmail Login Denied.")
        print("    Verify your GMAIL_APP_PASSWORD is a valid 16-character App Password.")
        return
    except Exception as e:
        print(f"\n❌  Critical Pipeline Error: {e}")
        traceback.print_exc()
        return

    # Summary Generation
    remaining = (df["Status"].astype(str).str.strip().str.lower() == "pending").sum()
    print(f"\n──────────────────────────────────────")
    print(f"🚀  Batch Complete!")
    print(f"✅  Successfully Sent: {sent_count}")
    print(f"❌  Failed Attempts:   {failed_count}")
    print(f"⏳  Remaining Queue:   {remaining}")
    if remaining > 0:
        print(f"    Run this script again tomorrow to process your next batch.")


if __name__ == "__main__":
    main()









