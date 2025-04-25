# fetch_clients.py

import pyodbc

# Database Connection Configuration
DB_SERVER = "52.172.98.46"
DB_DATABASE = "TestDB_24.12.2024"
DB_USERNAME = "Maximus"
DB_PASSWORD = "H#rm0n!ous@123"

def fetch_all_master_client_ids():
    """Fetch all distinct active MasterClientIDs from the Clients table."""
    try:
        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};SERVER={DB_SERVER};DATABASE={DB_DATABASE};UID={DB_USERNAME};PWD={DB_PASSWORD}",
            autocommit=True
        )
        cursor = conn.cursor()

        query = """
        SELECT DISTINCT MasterClientID
        FROM wrbhbpowerbiuser
        WHERE IsActive = 1 AND MasterClientID IS NOT NULL
        """

        cursor.execute(query)
        rows = cursor.fetchall()
        master_client_ids = [row[0] for row in rows]

        cursor.close()
        conn.close()

        return master_client_ids

    except Exception as e:
        print(f"⚠️ Error fetching master client IDs: {e}")
        return []
