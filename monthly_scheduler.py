import schedule
import time
from datetime import datetime
from table3 import generate_ppt

# 🎯 Define your Master Clients and their respective Clients
master_clients = [
    {"id": 969, "clients": ["2631"]},
    {"id": 383, "clients": ["1741"]},
    {"id":969}
    # Add more as needed
]

def get_last_day_of_month():
    from calendar import monthrange
    today = datetime.now()
    return monthrange(today.year, today.month)[1]

def run_monthly_report_generation():
    today = datetime.now()
    current_month = today.strftime("%Y-%m")

    print(f"\n🚀 Running report generation for {current_month}")

    for master in master_clients:
        master_id = master["id"]

        # 1️⃣ Generate Master-level report
        print(f"📊 Generating Master report for {master_id}")
        generate_ppt(master_id, current_month, is_client_level=False)

        # 2️⃣ Generate Client-level reports
        for client_id in master["clients"]:
            print(f"📄 Generating Client report for {client_id} under {master_id}")
            generate_ppt(master_id, current_month, client_id=client_id, is_client_level=True)

    print("✅ All reports for the month generated!\n")

# ⏰ Schedule to run every day at 11:55 PM
schedule.every().day.at("23:55").do(
    #lambda: run_monthly_report_generation() if datetime.now().day == get_last_day_of_month() else None
    lambda: run_monthly_report_generation() if True else None
)

# 🔁 Keep checking every minute
print("🔁 Scheduler started. Waiting for end-of-month...")
while True:
    schedule.run_pending()
    time.sleep(60)
