"""
Send individual branded "please update your tracker" reminder emails.
====================================================================
Reads missing_yesterday.json (written by dashboard.py) and sends ONE
branded HTML reminder to each person who did not log activity for the
previous working day. Akash and Kannan are CC'd on every email.

Runs only on the scheduled midday run / manual run (workflow-gated).
Credentials come from GitHub Secrets — nothing hardcoded.

Required GitHub Secrets:
  SMTP_SERVER    smtp.gmail.com
  SMTP_PORT      587
  SMTP_USER      waymarketingteam@gmail.com
  SMTP_PASSWORD  the 16-char Gmail app password
  (EMAIL_TO is no longer used for the summary; recipients are per-person)
"""

import os
import json
import smtplib
import sys
from email.message import EmailMessage

JSON_FILE = "missing_yesterday.json"

# Managers CC'd on every reminder
CC_LIST = ["akash.rs@way.com", "kannan.dhananjayan@way.com"]

# Name -> work email. "Bajan" alias included. Naitik/Kiran/Akash excluded as reminder targets.
NAME_TO_EMAIL = {
    "Gokul Nath": "iamakashunni91@gmail.com",
}


def first_name(full_name: str) -> str:
    return full_name.strip().split()[0] if full_name.strip() else full_name


def build_html(name: str, for_date: str) -> str:
    """Branded Way HTML email — mint #0BEFBA on dark green, Inter font.
    Inline styles + table layout for email-client compatibility."""
    fn = first_name(name)
    return f"""\
<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#0A1510;font-family:'Inter',Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#0A1510;padding:28px 16px;">
    <tr>
      <td align="center">
        <table role="presentation" width="520" cellpadding="0" cellspacing="0" style="max-width:520px;width:100%;background:#122016;border-radius:12px;overflow:hidden;border:1px solid #23422F;">

          <!-- Header bar: black with mint bottom accent -->
          <tr>
            <td style="background:#000000;border-bottom:3px solid #0BEFBA;padding:22px 28px;">
              <div style="font-size:11px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:#0BEFBA;">
                Way Marketing Team
              </div>
              <div style="font-size:20px;font-weight:800;color:#ffffff;margin-top:6px;">
                Activity Tracker Reminder
              </div>
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:28px;">
              <p style="font-size:15px;color:#E8F5EE;margin:0 0 16px;">Hi {fn},</p>

              <p style="font-size:14px;line-height:1.6;color:#C4D6CC;margin:0 0 20px;">
                Our records show that your activity log for
                <span style="color:#0BEFBA;font-weight:700;">{for_date}</span>
                has not been updated in the tracker yet.
              </p>

              <!-- Callout box -->
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#0D1A0F;border-left:4px solid #0BEFBA;border-radius:6px;margin:0 0 22px;">
                <tr>
                  <td style="padding:14px 18px;font-size:14px;color:#ffffff;font-weight:600;">
                    Please update the data at the earliest to ensure compliance.
                  </td>
                </tr>
              </table>

              <p style="font-size:13px;line-height:1.6;color:#8FA89A;margin:0;">
                If you have already updated it, kindly disregard this message.
              </p>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#0D1A0F;border-top:1px solid #23422F;padding:16px 28px;">
              <div style="font-size:11px;color:#6B8577;">
                Automated reminder &middot; Way Marketing Team
              </div>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>"""


def send_one(smtp, sender, to_email, name, for_date):
    msg = EmailMessage()
    msg["Subject"] = f"Reminder — Activity tracker pending for {for_date}"
    msg["From"] = sender
    msg["To"] = to_email
    msg["Cc"] = ", ".join(CC_LIST)
    # plain-text fallback
    msg.set_content(
        f"Hi {first_name(name)},\n\n"
        f"Your activity log for {for_date} has not been updated in the tracker yet. "
        f"Please update the data at the earliest to ensure compliance.\n\n"
        f"If you have already updated it, kindly disregard this message.\n\n"
        f"— Way Marketing Team"
    )
    msg.add_alternative(build_html(name, for_date), subtype="html")
    all_recipients = [to_email] + CC_LIST
    smtp.send_message(msg, from_addr=sender, to_addrs=all_recipients)


def main():
    if not os.path.exists(JSON_FILE):
        print("No missing_yesterday.json found (weekend or check skipped). Nothing to send.")
        return

    with open(JSON_FILE, encoding="utf-8") as f:
        data = json.load(f)

    for_date = data.get("for_date", "yesterday")
    missing = data.get("missing_names", [])

    if not missing:
        print(f"All team members logged activity for {for_date}. No reminders needed.")
        return

    server = os.environ.get("SMTP_SERVER", "")
    port = int(os.environ.get("SMTP_PORT", "587"))
    user = os.environ.get("SMTP_USER", "")
    password = os.environ.get("SMTP_PASSWORD", "")

    if not all([server, user, password]):
        print("ERROR: Missing SMTP configuration. Check the GitHub Secrets.")
        sys.exit(1)

    sent, skipped = [], []
    with smtplib.SMTP(server, port) as smtp:
        smtp.starttls()
        smtp.login(user, password)
        for name in missing:
            to_email = NAME_TO_EMAIL.get(name)
            if not to_email:
                # Unknown name (typed manually / not in map) — do NOT guess.
                skipped.append(name)
                continue
            send_one(smtp, user, to_email, name, for_date)
            sent.append(f"{name} <{to_email}>")

    print(f"Reminders sent for {for_date} ({len(sent)}):")
    for s in sent:
        print("  ->", s)
    if skipped:
        print(f"\nNOT emailed (name not recognised — check manually): {', '.join(skipped)}")
        print("These names were in the tracker's missing list but not in the email map.")


if __name__ == "__main__":
    main()
