import matplotlib.pyplot as plt
import pandas as pd
from db import fetch_data_from_sp

def generate_table(df, title, filename):
    """Creates and saves a formatted table as a PNG image."""
    if df.empty:
        print(f"⚠️ No data available for {title}")
        return None

    fig, ax = plt.subplots(figsize=(8, len(df) * 0.6 + 1))
    
    # Set Title Styling
    ax.set_title(title, fontsize=14, fontweight="bold", pad=10, color="darkblue")

    # Hide axes
    ax.axis("tight")
    ax.axis("off")

    # Format Data as Table
    column_labels = df.columns.tolist()
    table_data = df.values.tolist()

    table = ax.table(cellText=table_data, 
                     colLabels=column_labels, 
                     cellLoc="center", 
                     loc="center",
                     colColours=["#0D47A1"] * len(column_labels))  # Header Color

    # Style Table
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.2)

    # Format Header Row
    for key, cell in table._cells.items():
        row, col = key
        if row == 0:  # Header row
            cell.set_fontsize(11)
            cell.set_text_props(weight="bold", color="white")  # White text
            cell.set_height(0.05)  # Adjust row height
        else:
            cell.set_facecolor("#E3F2FD")  # Light Blue Background for Data Rows
            cell.set_height(0.04)

    # Save the table as an image
    table_filename = f"{filename}.png"
    plt.savefig(table_filename, bbox_inches="tight", dpi=300, transparent=True)
    plt.close()
    
    print(f"✅ Table saved as {table_filename}")
    return table_filename

def generate_tables(master_client_id, client_id):
    """Fetches stored procedure data and generates styled tables."""
    dfs = fetch_data_from_sp(master_client_id, client_id)

    if not dfs or len(dfs) < 3:
        print("⚠️ No sufficient data fetched from stored procedure.")
        return None

    table_files = {}

    # Table 1: Client-wise Booking Statistics
    client_df = dfs[0][["ClientName", "total_guests", "total_bookings", "total_spent"]]
    table_files["client_table"] = generate_table(client_df, "Client Booking Statistics", "client_booking_table")

    # Table 2: Traveller Profile
    tp_df=dfs[1][["ClientName","guest_volume_below_2000","guest_volume_2000_5000","guest_volume_above_5000"]]
    table_files["tp_df"] = generate_table(tp_df,"Traveller Profile","Traveller_profile_table")

    # Table 2: City-wise Booking Statistics
    city_df = dfs[3][["city", "total_guests", "total_bookings", "total_spent", "total_roomnights"]]
    table_files["city_table"] = generate_table(city_df, "City Booking Statistics", "city_booking_table")

    # Table 3: Hotel-wise Booking Statistics
    hotel_df = dfs[4][["PropertyName", "total_guests", "total_bookings", "total_spent", "total_roomnights"]]
    table_files["hotel_table"] = generate_table(hotel_df, "Hotel Booking Statistics", "hotel_booking_table")

    return table_files

if __name__ == "__main__":
    master_client_id = 969  # Replace with valid ID
    client_id = 803  # Replace with valid Client ID for filtering
    table_files = generate_tables(master_client_id, client_id)
    print(f"Generated Tables: {table_files}")
