"""
Send individual branded reminder emails + one manager summary.
=============================================================
Reads missing_yesterday.json (written by dashboard.py) and:
  1. Sends ONE branded reminder to each missing person (no CC).
  2. Sends ONE summary email to Akash + Kannan only, listing everyone
     reminded — or "all clear" if nobody was missing.

Runs on the scheduled midday run / manual run (workflow-gated).
Credentials come from GitHub Secrets — nothing hardcoded.

Required GitHub Secrets:
  SMTP_SERVER    smtp.gmail.com
  SMTP_PORT      587
  SMTP_USER      waymarketingteam@gmail.com
  SMTP_PASSWORD  the 16-char Gmail app password
"""

import os
import json
import smtplib
import sys
from email.message import EmailMessage

JSON_FILE = "missing_yesterday.json"

# Manager summary recipients
MANAGERS = ["akash.rs@way.com", "kannan.dhananjayan@way.com"]

# Name -> work email. "Bajan" alias included.
NAME_TO_EMAIL = {
    "Ajay Singh": "ajay.singh@way.com",
    "Anna Mary": "anna.theresa@way.com",
    "Archa Ullas": "archa.ullas@way.com",
    "Arun Mahadev": "arun.mahadev@way.com",
    "Arun Nath J": "arun.nath@way.com",
    "Bajan": "bajan.bj@way.com",
    "Bajan BJ": "bajan.bj@way.com",
    "Balavignesh P": "balavignesh.p@way.com",
    "Devika Sheeja": "devika.sheeja@way.com",
    "Fanny Dorris": "fanny.dorris@way.com",
    "Gautham S": "gautham.s@way.com",
    "Gokul Nath": "gokul.g@way.com",
    "Haripriya L": "haripriya.l@way.com",
    "Jofia Joseph": "jofia.joseph@way.com",
    "Kulwinder Singh": "kulwinder.singh@way.com",
    "Naveen PC": "naveen.pc@way.com",
    "Priya Kumari": "priya.kumari@way.com",
    "Rajeswari Menon": "rajeswari.menon@way.com",
    "Savitha Vasanthan": "savitha.vasanthan@way.com",
    "Seethal vargheese": "seethal.varghese@way.com",
    "Shilpa Sara": "shilpa.sam@way.com",
    "Sneha S": "sneha.subhanath@way.com",
    "Sreejith SL": "sreejith.sl@way.com",
}


def first_name(full_name: str) -> str:
    return full_name.strip().split()[0] if full_name.strip() else full_name


# ═══════════════════════════════════════════════════════════════════════
# INDIVIDUAL REMINDER — sent to each missing person, no CC
# ═══════════════════════════════════════════════════════════════════════
def build_reminder_html(name: str, for_date: str) -> str:
    fn = first_name(name)
    return f"""<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#0A1510;font-family:'Inter',Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#0A1510;padding:28px 16px;">
    <tr><td align="center">
      <table role="presentation" width="520" cellpadding="0" cellspacing="0" style="max-width:520px;width:100%;background:#122016;border-radius:12px;overflow:hidden;border:1px solid #23422F;">
        <tr><td style="background:#000000;border-bottom:3px solid #0BEFBA;padding:22px 28px;">
          <div style="font-size:11px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:#0BEFBA;">Way Marketing Team</div>
          <div style="font-size:20px;font-weight:800;color:#ffffff;margin-top:6px;">Activity Tracker Reminder</div>
        </td></tr>
        <tr><td style="padding:28px;">
          <p style="font-size:15px;color:#E8F5EE;margin:0 0 16px;">Hi {fn},</p>
          <p style="font-size:14px;line-height:1.6;color:#C4D6CC;margin:0 0 20px;">
            Our records show that your activity log for
            <span style="color:#0BEFBA;font-weight:700;">{for_date}</span>
            has not been updated in the tracker yet.
          </p>
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#0D1A0F;border-left:4px solid #0BEFBA;border-radius:6px;margin:0 0 22px;">
            <tr><td style="padding:14px 18px;font-size:14px;color:#ffffff;font-weight:600;">
              Please update the tracker before 12 PM daily, to ensure governance compliance and to meet organisational objectives.
            </td></tr>
          </table>
          <p style="font-size:13px;line-height:1.6;color:#8FA89A;margin:0;">If you have already updated it, kindly disregard this message.</p>
        </td></tr>
        <tr><td style="background:#0D1A0F;border-top:1px solid #23422F;padding:16px 28px;">
          <div style="font-size:11px;color:#6B8577;">Automated reminder &middot; Way Marketing Team</div>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>"""


def send_reminder(smtp, sender, to_email, name, for_date):
    msg = EmailMessage()
    msg["Subject"] = f"Reminder — Activity tracker pending for {for_date}"
    msg["From"] = sender
    msg["To"] = to_email
    msg.set_content(
        f"Hi {first_name(name)},\n\n"
        f"Your activity log for {for_date} has not been updated in the tracker yet. "
        f"Please update the tracker before 12 PM daily, to ensure governance "
        f"compliance and to meet organisational objectives.\n\n"
        f"If you have already updated it, kindly disregard this message.\n\n"
        f"— Way Marketing Team"
    )
    msg.add_alternative(build_reminder_html(name, for_date), subtype="html")
    smtp.send_message(msg, from_addr=sender, to_addrs=[to_email])


# ═══════════════════════════════════════════════════════════════════════
# MANAGER SUMMARY — sent once daily to Akash + Kannan only
# ═══════════════════════════════════════════════════════════════════════
def build_summary_html(for_date: str, missing_names: list, total_roster: int) -> str:
    all_clear = len(missing_names) == 0
    if all_clear:
        headline = "All Clear ✓"
        subline = f"All {total_roster} team members logged activity for {for_date}."
        callout = f"<span style='color:#0BEFBA;'>All {total_roster} logged.</span> No reminders sent."
        list_html = ""
    else:
        headline = f"{len(missing_names)} of {total_roster} Missed"
        subline = f"Individual reminders sent to the {len(missing_names)} team members below:"
        callout = f"<span style='color:#F4C542;'>{len(missing_names)} of {total_roster} did not log activity.</span> Individual reminders sent."
        list_items = "".join(
            f"<tr><td style='padding:6px 0;border-bottom:1px solid #1E3527;color:#E8F5EE;font-size:14px;'>&bull; {n}</td></tr>"
            for n in missing_names
        )
        list_html = f"""
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px;">
            {list_items}
          </table>"""

    return f"""<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#0A1510;font-family:'Inter',Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#0A1510;padding:28px 16px;">
    <tr><td align="center">
      <table role="presentation" width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#122016;border-radius:12px;overflow:hidden;border:1px solid #23422F;">
        <tr><td style="background:#000000;border-bottom:3px solid #0BEFBA;padding:22px 28px;">
          <div style="font-size:11px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:#0BEFBA;">Way Marketing Team &middot; Daily Compliance Summary</div>
          <div style="font-size:20px;font-weight:800;color:#ffffff;margin-top:6px;">{headline}</div>
          <div style="font-size:12px;color:#8FA89A;margin-top:4px;">Report for {for_date}</div>
        </td></tr>
        <tr><td style="padding:24px 28px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#0D1A0F;border-left:4px solid {'#0BEFBA' if all_clear else '#F4C542'};border-radius:6px;margin-bottom:16px;">
            <tr><td style="padding:14px 18px;font-size:14px;color:#ffffff;font-weight:600;">
              {callout}
            </td></tr>
          </table>
          <p style="font-size:13px;color:#C4D6CC;margin:0;">{subline}</p>
          {list_html}
        </td></tr>
        <tr><td style="background:#0D1A0F;border-top:1px solid #23422F;padding:16px 28px;">
          <div style="font-size:11px;color:#6B8577;">Automated daily compliance summary &middot; Way Marketing Team</div>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>"""


def send_summary(smtp, sender, for_date, missing_names, total_roster):
    all_clear = len(missing_names) == 0
    if all_clear:
        subject = f"[Tracker] ✓ All Clear — {for_date}"
        body = (f"All {total_roster} team members logged activity for {for_date}. "
                f"No reminders sent.\n\n— Way Marketing Team")
    else:
        subject = f"[Tracker] {len(missing_names)} missed — {for_date}"
        body = (f"{len(missing_names)} of {total_roster} did NOT log activity for {for_date}.\n\n"
                + "\n".join(f"  • {n}" for n in missing_names)
                + "\n\nIndividual reminders have been sent to each.\n\n— Way Marketing Team")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(MANAGERS)
    msg.set_content(body)
    msg.add_alternative(build_summary_html(for_date, missing_names, total_roster), subtype="html")
    smtp.send_message(msg, from_addr=sender, to_addrs=MANAGERS)


# ═══════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════
def main():
    if not os.path.exists(JSON_FILE):
        print("No missing_yesterday.json found (weekend or check skipped). Nothing to send.")
        return

    with open(JSON_FILE, encoding="utf-8") as f:
        data = json.load(f)

    for_date = data.get("for_date", "yesterday")
    missing = data.get("missing_names", [])
    total_roster = data.get("total_roster", 0)

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

        # 1. Individual reminders (only if there are missing people)
        for name in missing:
            to_email = NAME_TO_EMAIL.get(name)
            if not to_email:
                # Unknown name (typed manually / not in map) — don't guess.
                skipped.append(name)
                continue
            send_reminder(smtp, user, to_email, name, for_date)
            sent.append(f"{name} <{to_email}>")

        # 2. Manager summary (always sent — all-clear or missing list)
        send_summary(smtp, user, for_date, missing, total_roster)

    if missing:
        print(f"Reminders sent for {for_date} ({len(sent)}):")
        for s in sent:
            print("  ->", s)
    else:
        print(f"All {total_roster} logged for {for_date}. No individual reminders sent.")

    print(f"\nSummary email sent to: {', '.join(MANAGERS)}")

    if skipped:
        print(f"\nNOT emailed (name not in map — check manually): {', '.join(skipped)}")


if __name__ == "__main__":
    main()
