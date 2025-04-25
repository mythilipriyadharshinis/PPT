import smtplib
from email.message import EmailMessage
import os

def send_email_with_attachment(to_email, subject, body, attachment_path):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = "mythilipriyadharshinis@gmail.com"  # Your email
    msg["To"] = to_email
    msg.set_content(body)

    # Add attachment
    with open(attachment_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

    # Send email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:  # Use correct SMTP
        smtp.login("youremail@example.com", "yourpassword")  # Use env vars for safety
        smtp.send_message(msg)
        print(f"📨 Email sent to {to_email}")
