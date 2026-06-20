"""
extract.py — Script 1: PDF / Excel Email Extractor
====================================================
Reads a PDF or existing Excel file, pulls out every valid email address,
deduplicates them, and saves the result to an Excel spreadsheet with
two columns: "Email Address" and "Status" (default: "Pending").

HOW TO RUN:
    python extract.py --input your_file.pdf --output emails.xlsx
    python extract.py --input existing_emails.xlsx --output emails.xlsx  (merges)
"""

# ──────────────────────────────────────────────
#  CONFIGURATION  (edit these defaults if needed)
# ──────────────────────────────────────────────
DEFAULT_INPUT_FILE  = "input.pdf"          # Path to your PDF or Excel file
DEFAULT_OUTPUT_FILE = "emails.xlsx"        # Where extracted emails will be saved
# ──────────────────────────────────────────────

import re
import sys
import argparse
import pandas as pd
from pathlib import Path

# Email regex — RFC-5321-ish, catches virtually all real addresses
EMAIL_REGEX = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
    re.IGNORECASE
)


def extract_from_pdf(pdf_path: Path) -> set[str]:
    """Pull every unique email from a PDF using pypdf."""
    try:
        from pypdf import PdfReader
    except ImportError:
        print("❌  pypdf not installed. Run:  pip install pypdf")
        sys.exit(1)

    print(f"📄  Reading PDF: {pdf_path}")
    reader = PdfReader(str(pdf_path))
    all_text = ""
    for page in reader.pages:
        all_text += page.extract_text() or ""

    emails = set(EMAIL_REGEX.findall(all_text))
    print(f"✅  Found {len(emails)} unique email(s) in PDF.")
    return emails


def extract_from_excel(xlsx_path: Path) -> set[str]:
    """Pull every unique email from an existing Excel sheet."""
    print(f"📊  Reading Excel: {xlsx_path}")
    df = pd.read_excel(xlsx_path)

    emails: set[str] = set()
    # Search every cell in every column for email-like strings
    for col in df.columns:
        for cell in df[col].astype(str):
            found = EMAIL_REGEX.findall(cell)
            emails.update(found)

    print(f"✅  Found {len(emails)} unique email(s) in Excel.")
    return emails


def merge_with_existing(output_path: Path, new_emails: set[str]) -> pd.DataFrame:
    """
    If the output file already exists, keep its rows intact (preserving
    'Sent' / 'Failed' statuses) and only add truly new addresses.
    """
    if output_path.exists():
        existing_df = pd.read_excel(output_path)
        existing_emails = set(existing_df["Email Address"].str.lower())
        fresh = {e for e in new_emails if e.lower() not in existing_emails}
        print(f"ℹ️   {len(fresh)} new email(s) will be added to existing list.")
        new_rows = pd.DataFrame({
            "Email Address": sorted(fresh),
            "Status": "Pending"
        })
        return pd.concat([existing_df, new_rows], ignore_index=True)
    else:
        return pd.DataFrame({
            "Email Address": sorted(new_emails),
            "Status": "Pending"
        })


def main():
    parser = argparse.ArgumentParser(description="Extract emails from PDF or Excel.")
    parser.add_argument("--input",  default=DEFAULT_INPUT_FILE,  help="Input file (.pdf or .xlsx)")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_FILE, help="Output Excel file (.xlsx)")
    args = parser.parse_args()

    input_path  = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        print(f"❌  File not found: {input_path}")
        sys.exit(1)

    suffix = input_path.suffix.lower()
    if suffix == ".pdf":
        emails = extract_from_pdf(input_path)
    elif suffix in (".xlsx", ".xls"):
        emails = extract_from_excel(input_path)
    else:
        print(f"❌  Unsupported file type '{suffix}'. Use .pdf or .xlsx")
        sys.exit(1)

    if not emails:
        print("⚠️   No email addresses found. Check your input file.")
        sys.exit(0)

    df = merge_with_existing(output_path, emails)
    df.to_excel(output_path, index=False)
    print(f"\n🎉  Saved {len(df)} total email(s) to '{output_path}'")
    pending = (df["Status"] == "Pending").sum()
    print(f"📬  {pending} email(s) are ready to send (Status = Pending).")


if __name__ == "__main__":
    main()
