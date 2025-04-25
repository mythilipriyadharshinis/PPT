import pandas as pd
from table3 import generate_ppt
from email_sender import send_email_with_attachment
import datetime
import os

# Read client list
clients_df = pd.read_csv("clients.csv")

# Use today's date to decide booking month
today = datetime.date.today()
booking_month = f"{today.year}-{str(today.month-1).zfill(2)}"  # Previous month

for _, row in clients_df.iterrows():
    master_id = row['master_client_id']
    client_id = row['client_id']
    email = row['email']

    output_file = f"reports/booking_report_{client_id}_{booking_month}.pptx"
    os.makedirs("reports", exist_ok=True)

    # Generate report
    try:
        generate_ppt(master_id, booking_month, client_id, output_file)

        # Send email
        send_email_with_attachment(
            email,
            subject=f"Monthly Booking Report - {booking_month}",
            body="Please find attached your monthly booking report. Let us know if you have any questions.",
            attachment_path=output_file
        )
    except Exception as e:
        print(f"❌ Failed for client {client_id}: {e}")
