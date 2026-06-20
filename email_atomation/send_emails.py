"""
send_emails.py — Script 2: Smart Daily Email Sender
=====================================================
Reads the Excel file produced by extract.py, sends a personalised email
to every "Pending" address (up to DAILY_LIMIT per run), and marks each
row "Sent" or "Failed: <reason>" immediately after attempting — so a crash
or Ctrl-C never loses progress or causes duplicate sends.

HOW TO RUN:
    python send_emails.py

BEFORE RUNNING:
  1. Enable 2-Step Verification on your Google account.
  2. Go to https://myaccount.google.com/apppasswords and generate an App Password.
  3. Paste the 16-character password (no spaces) into GMAIL_APP_PASSWORD below.
"""

# ──────────────────────────────────────────────────────────────────
#  CONFIGURATION  — edit ONLY this block, nothing else needs to change
# ──────────────────────────────────────────────────────────────────
GMAIL_ADDRESS    = "mrk21creates@gmail.com"   # Your Gmail address
GMAIL_APP_PASSWORD = "xxxx xxxx xxxx xxxx"    # 16-char Google App Password
EXCEL_FILE       = "emails.xlsx"              # Excel file from extract.py
DAILY_LIMIT      = 500                        # Max emails per script run
DELAY_SECONDS    = 5                          # Pause between sends (be gentle!)

# ── Email content ──────────────────────────────────────────────────
EMAIL_SUBJECT    = "We'd love your feedback!"

# Use {name} where you want the recipient's name/username to appear.
# {name} is auto-extracted from the part before the @ in their email address.
# Use {form_link} where the Google Form URL should appear.
EMAIL_BODY_TEMPLATE = """\
Hi {name},

I hope this message finds you well.

We're gathering feedback and would love to hear from you. It only takes
2–3 minutes — please fill in the short form below:

  {form_link}

Thank you so much for your time. Your input genuinely matters to us.

Warm regards,
mrk
"""

GOOGLE_FORM_LINK = "https://forms.gle/YOUR_FORM_LINK_HERE"   # ← paste your form link
# ──────────────────────────────────────────────────────────────────

import time
import smtplib
import traceback
from email.message import EmailMessage
from pathlib import Path

import pandas as pd


# ─── Helper ───────────────────────────────────────────────────────

def username_from_email(email: str) -> str:
    """
    Derive a friendly display name from the email address.
    'john.doe@example.com'  →  'John Doe'
    'mrk21creates@gmail.com' →  'Mrk21creates'
    """
    local = email.split("@")[0]
    # Replace separators with spaces, title-case each word
    friendly = local.replace(".", " ").replace("_", " ").replace("-", " ").title()
    return friendly.strip()


def save_df(df: pd.DataFrame, path: Path) -> None:
    """Write the DataFrame back to Excel. Called after every send attempt."""
    df.to_excel(path, index=False)


# ─── Main ─────────────────────────────────────────────────────────

def main():
    excel_path = Path(EXCEL_FILE)

    if not excel_path.exists():
        print(f"❌  Excel file not found: {EXCEL_FILE}")
        print("    Run extract.py first to generate it.")
        return

    df = pd.read_excel(excel_path)

    # Validate expected columns
    required_cols = {"Email Address", "Status"}
    if not required_cols.issubset(df.columns):
        print(f"❌  Excel file is missing columns. Expected: {required_cols}")
        return

    # Filter to only Pending rows
    pending_mask = df["Status"].str.strip().str.lower() == "pending"
    pending_df   = df[pending_mask]

    if pending_df.empty:
        print("🎉  All done! No 'Pending' emails remain in the spreadsheet.")
        return

    # Cap at daily limit
    queue = pending_df.head(DAILY_LIMIT)
    total = len(queue)
    print(f"\n📬  {len(pending_df)} pending email(s) found.")
    print(f"📤  Sending up to {DAILY_LIMIT} today → processing {total} email(s).\n")

    sent_count   = 0
    failed_count = 0

    # Open one persistent SMTP-SSL connection for all emails
    try:
        print("🔐  Connecting to Gmail SMTP…")
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
            print("✅  Login successful.\n")

            for idx, row in queue.iterrows():
                recipient = str(row["Email Address"]).strip()
                name      = username_from_email(recipient)
                body      = EMAIL_BODY_TEMPLATE.format(
                    name      = name,
                    form_link = GOOGLE_FORM_LINK
                )

                msg = EmailMessage()
                msg["From"]    = GMAIL_ADDRESS
                msg["To"]      = recipient
                msg["Subject"] = EMAIL_SUBJECT
                msg.set_content(body)

                try:
                    smtp.send_message(msg)
                    df.at[idx, "Status"] = "Sent"
                    save_df(df, excel_path)         # ← saved immediately
                    sent_count += 1
                    print(f"  ✅  [{sent_count}/{total}]  Sent → {recipient}")

                    # Polite delay between sends
                    if sent_count < total:
                        time.sleep(DELAY_SECONDS)

                except smtplib.SMTPRecipientsRefused as e:
                    reason = f"Recipient refused: {e}"
                    df.at[idx, "Status"] = f"Failed: {reason}"
                    save_df(df, excel_path)
                    failed_count += 1
                    print(f"  ❌  Failed → {recipient} | {reason}")

                except Exception as e:
                    reason = str(e)[:120]           # cap length for Excel cell
                    df.at[idx, "Status"] = f"Failed: {reason}"
                    save_df(df, excel_path)
                    failed_count += 1
                    print(f"  ❌  Failed → {recipient} | {reason}")

    except smtplib.SMTPAuthenticationError:
        print("\n❌  Gmail authentication failed.")
        print("    Make sure GMAIL_APP_PASSWORD is a valid Google App Password.")
        print("    Guide: https://myaccount.google.com/apppasswords")
        return

    except Exception as e:
        print(f"\n❌  Unexpected error: {e}")
        traceback.print_exc()
        return

    # ── Summary ───────────────────────────────────────────────────
    remaining = (df["Status"].str.strip().str.lower() == "pending").sum()
    print(f"\n──────────────────────────────────────")
    print(f"✅  Sent:    {sent_count}")
    print(f"❌  Failed:  {failed_count}")
    print(f"⏳  Remaining pending: {remaining}")
    if remaining > 0:
        print(f"    Run the script again tomorrow to send the next batch.")
    else:
        print("🎉  All emails have been processed!")


if __name__ == "__main__":
    main()
