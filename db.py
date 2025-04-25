import pymssql
import pandas as pd

# 🔐 Database credentials (replace with your actual credentials)
DB_SERVER = "52.172.98.46"
DB_DATABASE = "TestDB_24.12.2024"
DB_USERNAME = "Maximus"
DB_PASSWORD = "H#rm0n!ous@123"


def fetch_data_from_sp(master_client_id, booking_month, client_id, is_client_level=False):
    """Fetch multiple result sets from the stored procedure GetBookingStatistics."""
    try:
        conn = pymssql.connect(server=DB_SERVER, user=DB_USERNAME, password=DB_PASSWORD, database=DB_DATABASE)
        cursor = conn.cursor()

        # Execute stored procedure
        if is_client_level:
            cursor.execute("EXEC GetBookingStatistics_clientwise %s, %s, %s", (master_client_id, booking_month, client_id))
        else:
            cursor.execute("EXEC GetBookingStatistics %s, %s", (master_client_id, booking_month))

        data_frames = []

        while True:
            rows = cursor.fetchall()
            if not rows:
                break

            columns = [desc[0] for desc in cursor.description]
            df = pd.DataFrame(rows, columns=columns)
            data_frames.append(df)

            if not cursor.nextset():
                break

        cursor.close()
        conn.close()

        return data_frames

    except Exception as e:
        print(f"⚠️ Error: {e}")
        return None


def get_master_and_client_ids_from_sql():
    query = """
        SELECT 
            CASE 
                WHEN p.MasterClientId = 0 THEN c.MasterClientId 
                ELSE p.MasterClientId 
            END AS MasterClientId,
            p.ClientId
        FROM WRBHBPowerBIUser p
        LEFT JOIN WRBHBClientManagement c ON c.Id = p.ClientId;
    """
    try:
        conn = pymssql.connect(server=DB_SERVER, user=DB_USERNAME, password=DB_PASSWORD, database=DB_DATABASE)
        df = pd.read_sql(query, conn)
        conn.close()

        master_to_clients = {}
        for _, row in df.iterrows():
            master_id = row['MasterClientId']
            client_id = row['ClientId']

            if pd.isna(master_id):
                continue

            master_id = int(master_id)

            if pd.isna(client_id):
                master_to_clients.setdefault(master_id, [])
            else:
                master_to_clients.setdefault(master_id, []).append(int(client_id))

        return master_to_clients

    except Exception as e:
        print(f"⚠️ Error fetching master-client IDs: {e}")
        return None


def fetch_smtp_details(action='SMTP', Str1='', Id=0):
    """Fetch SMTP details from the stored procedure GetSmtpDetails."""
    try:
        conn = pymssql.connect(server=DB_SERVER, user=DB_USERNAME, password=DB_PASSWORD, database=DB_DATABASE)
        cursor = conn.cursor()

        # Execute stored procedure with proper parameter tuple
        cursor.execute("EXEC SP_SMTPMailSetting_Help %s, %s, %s", (action, Str1, Id))

        row = cursor.fetchone()
        cursor.close()
        conn.close()

        if row:
            # Access by index since `pymssql` returns tuples (not objects like pyodbc)
            return {
                "SMTP_HOST": row[0],  # Adjust these indices based on your column order
                "SMTP_PORT": row[1],
                "SMTP_USERNAME": row[2],
                "SMTP_PASSWORD": row[3]
            }
        else:
            print("⚠️ No SMTP details found.")
            return None

    except Exception as e:
        print(f"⚠️ Error fetching SMTP details: {e}")
        return None


# 🔹 Test Block
if __name__ == "__main__":
    smtp_details = fetch_smtp_details()
    if smtp_details:
        print("SMTP Details:", smtp_details)
    else:
        print("Failed to fetch SMTP details.")
