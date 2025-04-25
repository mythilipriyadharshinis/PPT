import pyodbc
import pandas as pd

# Database Connection Configuration
DB_SERVER = "52.172.98.46"
DB_DATABASE = "TestDB_24.12.2024"
DB_USERNAME = "Maximus"
DB_PASSWORD = "H#rm0n!ous@123"


def fetch_data_from_sp(master_client_id, booking_month, client_id, is_client_level=False):
    """Fetch multiple result sets from the stored procedure GetBookingStatistics."""
    try:
            
        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};SERVER={DB_SERVER};DATABASE={DB_DATABASE};UID={DB_USERNAME};PWD={DB_PASSWORD}",
            autocommit=True
        )
        cursor = conn.cursor()

        # Execute stored procedure
        if is_client_level:
            
          cursor.execute("EXEC GetBookingStatistics_clientwise ?, ?, ?", master_client_id, booking_month, client_id)
        
        else:
          
          cursor.execute("EXEC GetBookingStatistics ?, ?", master_client_id, booking_month)  
         
        data_frames = []  # List to store multiple DataFrames

        while True:
            rows = cursor.fetchall()
            if not rows:  
                break  # No more result sets
            
            columns = [desc[0] for desc in cursor.description]
            df = pd.DataFrame.from_records(rows, columns=columns)
            data_frames.append(df)

            if not cursor.nextset():  
                break  # Move to the next result set

        cursor.close()
        conn.close()

        return data_frames  # Returns a list of DataFrames

    except Exception as e:
        print(f"⚠️ Error: {e}")
        return None  # Return None if an error occurs
    
def get_master_and_client_ids_from_sql():
    query = """
        SELECT 
    CASE 
        WHEN p.MasterClientId = 0 THEN c.MasterClientId 
        ELSE p.MasterClientId 
    END AS MasterClientId,
    p.ClientId
    FROM 
    WRBHBPowerBIUser p
    LEFT JOIN 
    WRBHBClientManagement c ON c.Id = p.ClientId;
    """
     # Connect and load data
    conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};SERVER={DB_SERVER};DATABASE={DB_DATABASE};UID={DB_USERNAME};PWD={DB_PASSWORD}",
            autocommit=True
        )
    df = pd.read_sql(query, conn)
    conn.close()

    # Build master-client mapping
    master_to_clients = {}
    for _, row in df.iterrows():
        master_id = (row['MasterClientId'])
        client_id = row['ClientId']
        
        if pd.isna(master_id):
            continue
        
        master_id=int(master_id)

        if pd.isna(client_id):
            master_to_clients.setdefault(master_id, [])
        else:
            master_to_clients.setdefault(master_id, []).append(int(client_id))

    return master_to_clients

def fetch_smtp_details(action='SMTP', Str1='',Id=0):
    """Fetch SMTP details from the stored procedure GetSmtpDetails."""
    try:
        # Connect to the database
        conn = pyodbc.connect(
            f"DRIVER={{SQL Server}};SERVER={DB_SERVER};DATABASE={DB_DATABASE};UID={DB_USERNAME};PWD={DB_PASSWORD}",
            autocommit=True
        )
        cursor = conn.cursor()

        # Execute the stored procedure
        cursor.execute("EXEC SP_SMTPMailSetting_Help ?, ?, ?",action,Str1,Id)

        # Fetch the result
        smtp_details = cursor.fetchone()

        # Close the connection
        cursor.close()
        conn.close()

        # Return SMTP details
        if smtp_details:
            return {
                "SMTP_HOST": smtp_details.Host,
                "SMTP_PORT": smtp_details.Port,
                "SMTP_USERNAME": smtp_details.CredentialsUserName,
                "SMTP_PASSWORD": smtp_details.CredentialsPassword
            }
        else:
            print("⚠️ No SMTP details found.")
            return 
            
    except Exception as e:
        print(f"⚠️ Error fetching SMTP details: {e}")
        return None
    

# 🔹 Test Function
if __name__ == "__main__":
    #dfs = fetch_data_from_sp(383, 1741,'2024-09', True)  # Test with sample IDs
    #if dfs:
       # for i, df in enumerate(dfs):
            #print(f"\nDataFrame {i+1}:")
            #print(df.head())  # Print first few rows of each result set
    
    
    
    '''result = get_master_and_client_ids_from_sql()
    print("Master to Clients Mapping:\n", result)'''
    
    smtp_details = fetch_smtp_details()
    if smtp_details:
        print("SMTP Details:", smtp_details)
    else:
        print("Failed to fetch SMTP details.")
