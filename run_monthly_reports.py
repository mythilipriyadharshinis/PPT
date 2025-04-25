from fetch_clients import fetch_all_master_client_ids
from table3 import generate_ppt
import sys

def run_all_reports(booking_month):
    master_client_ids = fetch_all_master_client_ids()
    if not master_client_ids:
        print("⚠️ No active Master Client IDs found.")
        return

    print(f"📊 Generating reports for {len(master_client_ids)} master clients for {booking_month}...\n")

    for master_client_id in master_client_ids:
        try:
            print(f"🛠️ Generating report for MasterClientID: {master_client_id}")
            generate_ppt(master_client_id, booking_month)
            print(f"✅ Done with MasterClientID: {master_client_id}\n")
        except Exception as e:
            print(f"❌ Failed to generate report for MasterClientID {master_client_id}: {e}\n")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("⚠️ Usage: python run_monthly_reports.py YYYY-MM")
    else:
        booking_month = sys.argv[1]
        run_all_reports(booking_month)
