import pandas as pd
from datetime import datetime
import schedule
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from table3 import generate_ppt  # Your existing report generator
from db import get_master_and_client_ids_from_sql
from db import fetch_smtp_details
import os



'''
# 🔹 TEMPORARY SMTP CONFIG — You can later fetch these from DB
SMTP_HOST = "email-smtp.us-west-2.amazonaws.com"
SMTP_PORT = 587
SMTP_USERNAME = "AKIARCGI7GAWDQ6WW75F"
SMTP_PASSWORD = "BEYwNyBC2eFOvUHfJL2yRZKZc7vK0LWF9yCC04kcdQiG"
'''
SENDER_EMAIL = "mythili@hummingbirdindia.com"  # This is the actual sender shown to recipients

# 🔹 1. Load master-client relationships from Excel
'''
def get_master_and_client_ids_from_excel(path="report_clients.xlsx"):
    df = pd.read_excel(path)
    master_to_clients = {}

    for _, row in df.iterrows():
        master_id = int(row['master_client_id'])
        client_id = row['client_id']

        if pd.isna(client_id):
            master_to_clients.setdefault(master_id, [])
        else:
            master_to_clients.setdefault(master_id, []).append(int(client_id))

    return master_to_clients
'''

# 🔹 2. Function to send email
def send_email(recipient_email, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach the email body as plain text
    msg.attach(MIMEText(body, 'plain'))

    # Attach the PowerPoint file
    part = MIMEBase('application', 'octet-stream')
    with open(attachment_path, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
    msg.attach(part)
    
    smtp = fetch_smtp_details()
    
    if smtp:
      SMTP_HOST = smtp['SMTP_HOST']
      SMTP_PORT = smtp['SMTP_PORT']
      SMTP_USERNAME = smtp['SMTP_USERNAME']
      SMTP_PASSWORD= smtp['SMTP_PASSWORD']
    
    # Send the email
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
            print(f"✅ Email sent to {recipient_email}")
    except Exception as e:
        print(f"❌ Failed to send email to {recipient_email}: {e}")

# 🔹 3. Common report runner
def run_excel_based_scheduler(booking_month):
    print(f"\n📅 Generating reports for: {booking_month}")
    master_client_map = get_master_and_client_ids_from_sql()

    for master_id, client_ids in master_client_map.items():
        if client_ids:
            for client_id in client_ids:
                print(f"📊 CLIENT Report → MasterID: {master_id}, ClientID: {client_id}")
                ppt_path = generate_ppt(master_id, booking_month, client_id=client_id, is_client_level=True)

                if ppt_path is None:
                    print(f"⚠️ Skipping email: Report not generated for Client {client_id} (MasterID: {master_id})")
                    continue

                recipient_email = get_client_email(master_id, client_id)
                send_email(recipient_email, f"Monthly Report for Client {client_id}", "Hi,\nPlease find attached the monthly report.", ppt_path)
        else:
            print(f"📊 MASTERCLIENT Report → MasterID: {master_id}")
            ppt_path = generate_ppt(master_id, booking_month, is_client_level=False)

            if ppt_path is None:
                print(f"⚠️ Skipping email: Report not generated for Master Client {master_id}")
                continue

            recipient_email = get_master_client_email(master_id)
            send_email(recipient_email, f"Monthly Report for Master Client {master_id}", "Hi,\nPlease find attached the monthly report.", ppt_path)

# 🔹 4. Temporary email mappings (you'll replace this with DB queries later)
def get_client_email(master_id, client_id):
    # Temporary: Replace this with DB logic
    return "mythilipriyadharshinis@gmail.com"

def get_master_client_email(master_id):
    # Temporary: Replace this with DB logic
    return "mythilipriyadharshinis@gmail.com"

# 🔹 5. Monthly scheduler
def schedule_monthly_job():
    today = datetime.now()
    next_month = today.replace(day=28) + pd.DateOffset(days=4)
    last_day = (next_month - pd.DateOffset(days=next_month.day)).day

    if today.day == last_day:
        booking_month = today.strftime("%Y-%m")
        run_excel_based_scheduler(booking_month)

# 🔹 6. Test scheduler: run every 2 minutes using current month
schedule.every().day.at("10:00").do(schedule_monthly_job)
schedule.every(15).minutes.do(lambda: run_excel_based_scheduler(datetime.now().strftime("%Y-%m")))

# 🔹 7. Manual test with specific month
def run_test_scheduler(booking_month):
    print("\n🧪 Running test scheduler for month:", booking_month)
    run_excel_based_scheduler(booking_month)
    


# 🔹 8. MAIN
if __name__ == "__main__":
    run_test_scheduler("2024-03")

    '''#Uncomment below to enable background scheduling
     print("✅ Schedulers running...\n- Monthly at 10:00 AM\n- Test every 2 minutes\n")
     while True:
         schedule.run_pending()
         time.sleep(60)'''
