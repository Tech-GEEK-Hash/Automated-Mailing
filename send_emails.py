import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

# --- Your Credentials and Email Details ---
sender_email = "sponsorship.instruo@gmail.com"
sender_password = "yawicpjuchlhcjkk"

# --- Email Content ---
email_subject = "Collaboration Proposal: INSTRUO 14 × {name}"

email_body = """
Dear {name} Team,

We are delighted to invite {name} to collaborate with INSTRUO 14, the annual techno-management festival of IIEST Shibpur, scheduled from October 31st to November 2nd, 2025.

INSTRUO is among Eastern India’s largest technical festivals, attracting over 10,000 students from premier institutes in Kolkata, along with top participants from IITs, NITs, and other leading universities across the country. It serves as a dynamic platform for innovation, knowledge exchange, and industry–academia partnerships.

With your expertise and leadership in your industry, {name} perfectly aligns with the spirit of technological and educational excellence that we champion. Our flagship events, spanning technology, mobility, energy, education, media, and innovation, are among the festival’s biggest highlights, offering a unique opportunity for your brand to engage directly with talented young engineers, innovators, and future leaders.

We would be glad to explore collaboration avenues that match your vision, and would love to schedule a brief call to discuss the possibilities further.

Warm regards,
Aayush Sarkar
Sponsorship Associate
Team INSTRUO
Contact: 7367999558"""

# Path to your attachment (keep file in same folder or give full path)
attachment_path = "INSTRUO_Brochure.pdf"  # Change filename if needed

def send_personalized_emails():
    # Load Excel contacts
    try:
        df = pd.read_excel("contacts.xlsx")
    except FileNotFoundError:
        print("Error: 'contacts.xlsx' not found.")
        return
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    if 'Name' not in df.columns or 'Email' not in df.columns:
        print("Error: Excel file must contain 'Name' and 'Email' columns.")
        return

    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, sender_password)
            print("Logged in successfully...")

            for _, row in df.iterrows():
                receiver_name = row['Name']
                receiver_email = row['Email']

                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = receiver_email
                message["Subject"] = email_subject.format(name=receiver_name)

                # Personalize body
                message.attach(MIMEText(email_body.format(name=receiver_name), "plain"))

                # Attach file if exists
                if os.path.exists(attachment_path):
                    with open(attachment_path, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename={os.path.basename(attachment_path)}",
                    )
                    message.attach(part)
                else:
                    print(f"⚠️ Attachment '{attachment_path}' not found. Sending email without it.")

                # Send email
                server.sendmail(sender_email, receiver_email, message.as_string())
                print(f"✅ Email sent to {receiver_name} ({receiver_email})")

    except smtplib.SMTPAuthenticationError:
        print("❌ Authentication failed. Check your email and app password.")
    except Exception as e:
        print(f"❌ An error occurred: {e}")

if __name__ == "__main__":
    send_personalized_emails()
