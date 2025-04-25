import pandas as pd
from datetime import datetime
import schedule
import time
from table3 import generate_ppt  # Your existing report generator

# 🔹 1. Load master-client relationships from Excel
def get_master_and_client_ids_from_excel(path="report_clients.xlsx"):
    df = pd.read_excel(path)
    master_to_clients = {}

    for _, row in df.iterrows():
        master_id = int(row['master_client_id'])
        client_id = (row['client_id'])

        if pd.isna(client_id):
            master_to_clients.setdefault(master_id, [])
        else:
            master_to_clients.setdefault(master_id, []).append(int(client_id))

    return master_to_clients

# 🔹 2. Common report runner (used by both schedulers)
def run_excel_based_scheduler(booking_month):
    print(f"\n📅 Generating reports for: {booking_month}")
    
    master_client_map = get_master_and_client_ids_from_excel()

    for master_id, client_ids in master_client_map.items():
        if client_ids:
            for client_id in client_ids:
                print(f"📊 CLIENT Report → MasterID: {master_id}, ClientID: {client_id}")
                generate_ppt(master_id, booking_month, client_id=client_id, is_client_level=True)
        else:
            print(f"📊 MASTERCLIENT Report → MasterID: {master_id}")
            generate_ppt(master_id, booking_month, is_client_level=False)

# 🔹 3. Monthly scheduler (runs at 1 AM on last day of the month)
def schedule_monthly_job():
    today = datetime.now()
    next_month = today.replace(day=28) + pd.DateOffset(days=4)
    last_day = (next_month - pd.DateOffset(days=next_month.day)).day

    if today.day == last_day:
        booking_month = today.strftime("%Y-%m")
        run_excel_based_scheduler(booking_month)

# 🔹 4. Test scheduler: run every 2 minutes using current month
schedule.every().day.at("10:00").do(schedule_monthly_job)
schedule.every(2).minutes.do(lambda: run_excel_based_scheduler(datetime.now().strftime("%Y-%m")))

# 🔹 5. Manual test with specific month (e.g., "2024-11")
def run_test_scheduler(booking_month):
    print("\n🧪 Running test scheduler for month:", booking_month)
    run_excel_based_scheduler(booking_month)

# 🔹 6. MAIN
if __name__ == "__main__":
    # ✅ To test manually with a specific month, call this:
    run_test_scheduler("2024-03")

    # ✅ To enable background scheduling, uncomment below:
    # print("✅ Schedulers running...\n- Monthly at 1:00 AM\n- Test every 2 minutes\n")
    # while True:
    #     schedule.run_pending()
    #     time.sleep(60)
