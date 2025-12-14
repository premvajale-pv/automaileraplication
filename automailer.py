import smtplib
import pandas as pd
import time
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import config


def get_smtp(email):
    domain = email.split("@")[-1].lower()

    smtp_map = {
        "gmail.com": ("smtp.gmail.com", 587),
        "outlook.com": ("smtp.office365.com", 587),
        "hotmail.com": ("smtp.office365.com", 587),
        "live.com": ("smtp.office365.com", 587),
        "yahoo.com": ("smtp.mail.yahoo.com", 587)
    }

    return smtp_map.get(domain, (f"mail.{domain}", 587))


def build_email(name, to_email):
    msg = MIMEMultipart()
    msg["From"] = f"{config.FROM_NAME} <{config.EMAIL_ADDRESS}>"
    msg["To"] = to_email
    msg["Subject"] = f"Proposal for {name}"

    body = f"""
Hi {name},

I hope you are doing well.

I’m reaching out to share a short proposal that can help
enhance your brand visibility and generate new opportunities.

I’d be happy to connect and explore how this could benefit you.

Best regards,
{config.FROM_NAME}
{config.COMPANY_NAME}
{config.SIGNATURE_EMAIL}

If you prefer not to receive future emails, please reply "unsubscribe".
"""
    msg.attach(MIMEText(body, "plain"))
    return msg


def main():
    df = pd.read_excel(config.EXCEL_FILE)

    sent_today = 0
    smtp_host, smtp_port = get_smtp(config.EMAIL_ADDRESS)

    server = smtplib.SMTP(smtp_host, smtp_port)
    server.starttls()
    server.login(config.EMAIL_ADDRESS, config.EMAIL_PASSWORD)

    print("✅ SMTP Connected")

    for i, row in df.iterrows():

        if sent_today >= config.DAILY_LIMIT:
            print("⚠️ Daily limit reached")
            break

        if pd.notna(row["Sent_Date"]):
            continue

        try:
            msg = build_email(row["Name"], row["Email"])
            server.send_message(msg)

            df.at[i, "Sent_Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df.at[i, "Status"] = "Sent"
            df.to_excel(config.EXCEL_FILE, index=False)

            sent_today += 1
            print(f"📨 Sent to {row['Email']}")

            time.sleep(config.DELAY_SECONDS)

        except Exception as e:
            df.at[i, "Status"] = f"Failed: {e}"
            df.to_excel(config.EXCEL_FILE, index=False)
            print(f"❌ Failed {row['Email']}")

    server.quit()
    print("✅ Campaign Completed")


if __name__ == "__main__":
    main()
