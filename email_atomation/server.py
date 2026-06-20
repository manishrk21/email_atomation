"""
server.py — Flask backend for the Email Tool web UI
=====================================================
Exposes REST endpoints that the single-page frontend calls.
All state lives in a single Excel file on disk.

HOW TO RUN:
    pip install flask flask-cors pypdf pandas openpyxl
    python server.py
Then open  http://localhost:5000  in your browser.
"""

# ──────────────────────────────────────────────────────────────────
#  CONFIGURATION
# ──────────────────────────────────────────────────────────────────
GMAIL_ADDRESS      = "YOUR_EMAIL"
GMAIL_APP_PASSWORD = "xxxx xxxx xxxx xxxx"    # 16-char Google App Password
DAILY_LIMIT        = 500
DELAY_SECONDS      = 5
EXCEL_FILE         = "emails.xlsx"
UPLOAD_FOLDER      = "uploads"
# ──────────────────────────────────────────────────────────────────

import re
import os
import time
import smtplib
import threading
import traceback
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

EMAIL_REGEX = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
    re.IGNORECASE
)

# Live progress shared between the sending thread and the /status endpoint
send_progress = {
    "running": False,
    "sent": 0,
    "failed": 0,
    "total": 0,
    "log": [],
    "done": False,
    "error": None,
}


# ─── Helpers ──────────────────────────────────────────────────────

def username_from_email(email: str) -> str:
    local = email.split("@")[0]
    return local.replace(".", " ").replace("_", " ").replace("-", " ").title().strip()


def excel_path() -> Path:
    return Path(EXCEL_FILE)


def load_df() -> pd.DataFrame:
    path = excel_path()
    if path.exists():
        return pd.read_excel(path)
    return pd.DataFrame(columns=["Email Address", "Status"])


def save_df(df: pd.DataFrame) -> None:
    df.to_excel(excel_path(), index=False)


def extract_emails_from_pdf(file_path: str) -> set[str]:
    from pypdf import PdfReader
    reader = PdfReader(file_path)
    text = "".join(page.extract_text() or "" for page in reader.pages)
    return set(EMAIL_REGEX.findall(text))


def extract_emails_from_excel(file_path: str) -> set[str]:
    df = pd.read_excel(file_path)
    emails: set[str] = set()
    for col in df.columns:
        for cell in df[col].astype(str):
            emails.update(EMAIL_REGEX.findall(cell))
    return emails


def merge_emails(new_emails: set[str]) -> tuple[int, int]:
    """Merge new emails into the master Excel. Returns (added, total)."""
    df = load_df()
    existing = set(df["Email Address"].str.lower()) if not df.empty else set()
    fresh = sorted(e for e in new_emails if e.lower() not in existing)
    if fresh:
        new_rows = pd.DataFrame({"Email Address": fresh, "Status": "Pending"})
        df = pd.concat([df, new_rows], ignore_index=True)
        save_df(df)
    return len(fresh), len(df)


# ─── Sending thread ───────────────────────────────────────────────

def send_thread(subject: str, body_template: str, form_link: str):
    global send_progress
    send_progress.update({"running": True, "sent": 0, "failed": 0,
                          "log": [], "done": False, "error": None})

    df = load_df()
    pending_mask = df["Status"].str.strip().str.lower() == "pending"
    queue = df[pending_mask].head(DAILY_LIMIT)
    send_progress["total"] = len(queue)

    if queue.empty:
        send_progress["running"] = False
        send_progress["done"] = True
        send_progress["error"] = "no_pending"
        return

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)

            for idx, row in queue.iterrows():
                recipient = str(row["Email Address"]).strip()
                name      = username_from_email(recipient)
                body      = body_template.replace("{name}", name)\
                                         .replace("{form_link}", form_link)

                msg = EmailMessage()
                msg["From"]    = GMAIL_ADDRESS
                msg["To"]      = recipient
                msg["Subject"] = subject
                msg.set_content(body)

                try:
                    smtp.send_message(msg)
                    df.at[idx, "Status"] = "Sent"
                    save_df(df)
                    send_progress["sent"] += 1
                    send_progress["log"].append({"email": recipient, "status": "sent"})
                    time.sleep(DELAY_SECONDS)

                except Exception as e:
                    reason = str(e)[:100]
                    df.at[idx, "Status"] = f"Failed: {reason}"
                    save_df(df)
                    send_progress["failed"] += 1
                    send_progress["log"].append({"email": recipient, "status": "failed", "reason": reason})

    except smtplib.SMTPAuthenticationError:
        send_progress["error"] = "auth_failed"
    except Exception as e:
        send_progress["error"] = str(e)
        traceback.print_exc()
    finally:
        send_progress["running"] = False
        send_progress["done"]    = True


# ─── API Routes ───────────────────────────────────────────────────

@app.route("/api/upload", methods=["POST"])
def upload_file():
    """Accept a PDF or Excel file, extract emails, merge into master list."""
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename"}), 400

    filename = f.filename
    saved_path = os.path.join(UPLOAD_FOLDER, filename)
    f.save(saved_path)

    ext = Path(filename).suffix.lower()
    try:
        if ext == ".pdf":
            emails = extract_emails_from_pdf(saved_path)
        elif ext in (".xlsx", ".xls"):
            emails = extract_emails_from_excel(saved_path)
        else:
            return jsonify({"error": f"Unsupported file type: {ext}"}), 400

        added, total = merge_emails(emails)
        return jsonify({
            "found": len(emails),
            "added": added,
            "total_in_list": total,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/emails", methods=["GET"])
def get_emails():
    """Return the full email list with statuses."""
    df = load_df()
    stats = df["Status"].value_counts().to_dict() if not df.empty else {}
    records = df.to_dict(orient="records") if not df.empty else []
    return jsonify({"emails": records, "stats": stats})


@app.route("/api/send", methods=["POST"])
def send_emails():
    """Kick off background sending thread."""
    global send_progress
    if send_progress["running"]:
        return jsonify({"error": "Already sending. Check /api/status."}), 409

    data    = request.get_json() or {}
    subject = data.get("subject", "").strip()
    body    = data.get("body", "").strip()
    link    = data.get("form_link", "").strip()

    if not subject or not body:
        return jsonify({"error": "Subject and body are required."}), 400

    t = threading.Thread(target=send_thread, args=(subject, body, link), daemon=True)
    t.start()
    return jsonify({"message": "Sending started."})


@app.route("/api/status", methods=["GET"])
def send_status():
    return jsonify(send_progress)


@app.route("/api/download", methods=["GET"])
def download_excel():
    if not excel_path().exists():
        return jsonify({"error": "No Excel file yet."}), 404
    return send_file(str(excel_path()), as_attachment=True)


@app.route("/api/reset-pending", methods=["POST"])
def reset_failed():
    """Reset all 'Failed' rows back to 'Pending' for retry."""
    df = load_df()
    mask = df["Status"].str.startswith("Failed")
    df.loc[mask, "Status"] = "Pending"
    save_df(df)
    count = mask.sum()
    return jsonify({"reset": int(count)})


@app.route("/api/clear", methods=["POST"])
def clear_all():
    """Wipe the entire email list."""
    save_df(pd.DataFrame(columns=["Email Address", "Status"]))
    return jsonify({"message": "Cleared."})


if __name__ == "__main__":
    print("🚀  Email Tool server running → http://localhost:5000")
    app.run(debug=False, port=5000)
