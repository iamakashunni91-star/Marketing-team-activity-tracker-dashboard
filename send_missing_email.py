"""
Send the daily "missing activity updates" email.
================================================
Reads missing_yesterday.json (written by dashboard.py) and emails the
list of people who did not log activity for the previous working day.

Runs only on the 5 PM IST scheduled run (the workflow gates this).
All credentials come from GitHub Secrets — nothing is hardcoded here.

Required GitHub Secrets:
  SMTP_SERVER    e.g. smtp.gmail.com   (Gmail)  or  smtp.office365.com (Outlook)
  SMTP_PORT      e.g. 587
  SMTP_USER      the sending email address
  SMTP_PASSWORD  an APP PASSWORD (not your normal login password)
  EMAIL_TO       the private address that should receive the list
"""

import os
import json
import smtplib
import sys
from email.message import EmailMessage

JSON_FILE = "missing_yesterday.json"


def main():
    # If the check didn't run (e.g. weekend) there is no file — exit quietly.
    if not os.path.exists(JSON_FILE):
        print("No missing_yesterday.json found (weekend or check skipped). Nothing to send.")
        return

    with open(JSON_FILE, encoding="utf-8") as f:
        data = json.load(f)

    for_date = data.get("for_date", "yesterday")
    missing = data.get("missing_names", [])
    total = data.get("total_roster", 0)
    checked_on = data.get("checked_on", "")

    # Build the email body
    if missing:
        lines = "\n".join(f"  • {name}" for name in missing)
        body = (
            f"Activity Tracker — Missing Updates\n"
            f"===================================\n\n"
            f"Report date: {for_date}\n"
            f"Checked: {checked_on}\n\n"
            f"{len(missing)} of {total} did NOT log activity for {for_date}:\n\n"
            f"{lines}\n\n"
            f"— Automated check from the Marketing Dashboard\n"
        )
        subject = f"[Activity Tracker] {len(missing)} missing updates for {for_date}"
    else:
        body = (
            f"Activity Tracker — Missing Updates\n"
            f"===================================\n\n"
            f"Report date: {for_date}\n"
            f"Checked: {checked_on}\n\n"
            f"All {total} team members logged activity for {for_date}. ✓\n\n"
            f"— Automated check from the Marketing Dashboard\n"
        )
        subject = f"[Activity Tracker] All updated for {for_date} ✓"

    # Read SMTP config from environment (GitHub Secrets)
    server = os.environ.get("SMTP_SERVER", "")
    port = int(os.environ.get("SMTP_PORT", "587"))
    user = os.environ.get("SMTP_USER", "")
    password = os.environ.get("SMTP_PASSWORD", "")
    to_addr = os.environ.get("EMAIL_TO", "")

    if not all([server, user, password, to_addr]):
        print("ERROR: Missing SMTP configuration. Check the GitHub Secrets.")
        sys.exit(1)

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = user
    # EMAIL_TO may contain multiple addresses separated by commas
    recipients = [addr.strip() for addr in to_addr.split(",") if addr.strip()]
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

    # Send via STARTTLS (standard for port 587)
    with smtplib.SMTP(server, port) as smtp:
        smtp.starttls()
        smtp.login(user, password)
        smtp.send_message(msg)

    print(f"Email sent to {', '.join(recipients)}: {subject}")


if __name__ == "__main__":
    main()
