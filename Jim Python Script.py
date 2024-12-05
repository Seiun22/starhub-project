from pathlib import Path
import win32com.client
import datetime as dt
import pandas as pd
import getpass
import pyodbc

# Creating output folder in python script's directory
script_dir = Path(__file__).parent
output_dir = script_dir / "Daily BAU"
output_dir.mkdir(parents=True, exist_ok=True)

# Connecting to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Set default folder
source_inbox = outlook.GetDefaultFolder(6)

# Access specific subfolders
subfolder_name = "Daily BAU"  # Specify your target folder name
try:
    target_folder = source_inbox.Folders.Item(subfolder_name)
    print(f"Accessing folder: {target_folder.Name}")
except Exception as e:
    print(f"Error: Could not find subfolder '{subfolder_name}' under Inbox.")
    target_folder = None

# Process emails
if target_folder:
    # Cache today's date
    today = dt.datetime.now()
    today_date = today.strftime("%d%m%y")

    # Create a new subfolder with today's date in "output_dir/Daily BAU"
    daily_output_folder = output_dir / subfolder_name / today_date
    daily_output_folder.mkdir(parents=True, exist_ok=True)

    processed_count = 0  # Counter for processed messages
    attachment_count = 0  # Counter for downloaded attachments

    for msg in target_folder.Items:
        try:
            # Filter today's emails
            received_date = msg.ReceivedTime.date()
            if received_date == today.date():
                subject = msg.Subject or "No Subject"
                print(f"Processing today's message: {subject} (Received: {received_date})")
                processed_count += 1

                # Save attachments to the dated folder
                for attachment in msg.Attachments:
                    attachment_path = daily_output_folder / attachment.FileName
                    attachment.SaveAsFile(str(attachment_path))
                    print(f"Saved attachment: {attachment_path}")
                    attachment_count += 1
        except Exception as e:
            print(f"Error processing message: {e}")

    # Count check
    print(f"Total messages processed today: {processed_count}")
    print(f"Total attachments downloaded today: {attachment_count}")

# 2. Process and clean data
try:
    # Load csv into df
    service_df = pd.read_csv(daily_output_folder / 'Raw Service.csv')
    customer_df = pd.read_csv(daily_output_folder / 'Raw Customer.csv')
    order_df = pd.read_csv(daily_output_folder / 'Raw Orders.csv')
    active_df = pd.read_csv(daily_output_folder / 'Raw Active.csv')

    # Data Cleanup
    for df in [service_df, customer_df, order_df, active_df]:
        df['REPORT_DATE'] = pd.to_datetime(df['REPORT_DATE'], dayfirst=True)
        df.loc[:, df.columns != 'REPORT_DATE'] = df.loc[:, df.columns != 'REPORT_DATE'].astype('object')
        df[df.select_dtypes(include=['object']).columns] = df.select_dtypes(include=['object']).apply(str.strip)
        df.drop_duplicates(inplace=True)
        df.dropna(inplace=True)

    # Merge Order with Service on REPORT_DATE and SERVICE_ID
    merged_df = pd.merge(order_df, service_df, on=['REPORT_DATE', 'SERVICE_ID'], how='inner')

    # Sort by REPORT_DATE to facilitate merge_asof
    customer_df = customer_df.sort_values(by=['REPORT_DATE', 'CUSTOMER_ID'])
    active_df = active_df.sort_values(by=['REPORT_DATE', 'CUSTOMER_ID'])

    # Perform merge_asof to align REPORT_DATE from customer_df with the nearest preceding report_date in active_df
    customer_active_df = pd.merge_asof(
        customer_df,
        active_df,
        on='REPORT_DATE',
        by='CUSTOMER_ID',
        direction='backward'
    )

    # Merge the result with the merged_df on REPORT_DATE and SERVICE_ID
    final_df = pd.merge(merged_df, customer_active_df, on=['REPORT_DATE', 'SERVICE_ID'], how='inner')

    # Clean up final_df
    final_df = final_df.drop(columns=['SERVICE_NAME_y', 'SUBSCRIPTION_STATUS'])
    final_df = final_df.rename(columns={'SERVICE_NAME_x': 'SERVICE_NAME'})
    final_df = final_df[['REPORT_DATE', 'SERVICE_ID', 'CUSTOMER_ID', 'CUSTOMER_SEGMENT_FLAG', 'CUSTOMER_GENDER', 'CUSTOMER_NATIONALITY',
                         'ORDER_TYPE', 'ORDER_TYPE_L2', 'SERVICE_NAME', 'SERVICE']]
    final_df = final_df.sort_values(by=['REPORT_DATE', 'CUSTOMER_ID', 'SERVICE_ID'])

    # Save the processed dfs to a new CSV file
    final_df.to_csv(daily_output_folder / 'processed_order.csv', index=False)
    print("Data processing complete. 'processed_order.csv' saved.")
    active_df.to_csv(daily_output_folder / 'processed_active.csv', index=False)
    print("Data processing complete. 'processed_active.csv' saved.")

except FileNotFoundError as e:
    print(f"Error: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")

# 3. Database connection and data insertion
try:
    # Prompt for password
    password = getpass.getpass(prompt='Please enter database password: ')

    # DB Connection
    print("Connecting to the database...")
    conn_str = (
        "DRIVER={NetezzaSQL};"
        "SERVER=your_server;"
        "PORT=5480;"  # The default Netezza port is 5480
        "DATABASE=your_database;"
        "UID=your_username;"
        f"PWD={password};"
    )

    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    print("Database connection established.")

    # Create tables if not exists
    print("Creating tables if they do not exist...")
    create_orders_table = """
    CREATE TABLE IF NOT EXISTS PROCESSED_ORDER_DAILY (
        REPORT_DATE DATE,
        SERVICE_ID VARCHAR(32),
        CUSTOMER_ID VARCHAR(32),
        CUSTOMER_SEGMENT_FLAG VARCHAR(32),
        CUSTOMER_GENDER VARCHAR(32),
        CUSTOMER_NATIONALITY VARCHAR(32),
        ORDER_TYPE VARCHAR(32),
        ORDER_TYPE_L2 VARCHAR(32),
        SERVICE_NAME VARCHAR(32),
        SERVICE VARCHAR(32),
    )
    """
    create_active_table = """
    CREATE TABLE IF NOT EXISTS PROCESSED_ACTIVE_DAILY (
        REPORT_DATE DATE,
        CUSTOMER_ID VARCHAR(32),
        SERVICE_ID VARCHAR(32),
        SERVICE_NAME VARCHAR(32),
        SUBSCRIPTION_STATUS VARCHAR(32)
    )
    """
    cursor.execute(create_orders_table)
    cursor.execute(create_active_table)
    conn.commit()
    print("Tables created or verified successfully.")

    # Define function to write df to db
    def write_to_db(df_orders, df_active, conn):
        cursor = conn.cursor()

        # Insert data into PROCESSED_ORDER_DAILY
        print("Inserting data into PROCESSED_ORDER_DAILY...")
        for index, row in df_orders.iterrows():
            placeholders = ", ".join(["?" for _ in row])
            insert_sql = f"INSERT INTO PROCESSED_ORDER_DAILY VALUES ({placeholders})"
            cursor.execute(insert_sql, tuple(row))
        conn.commit()
        print("Data inserted into PROCESSED_ORDER_DAILY successfully.")

        # Insert data into PROCESSED_ACTIVE_DAILY
        print("Inserting data into PROCESSED_ACTIVE_DAILY...")
        for index, row in df_active.iterrows():
            placeholders = ", ".join(["?" for _ in row])
            insert_sql = f"INSERT INTO PROCESSED_ACTIVE_DAILY VALUES ({placeholders})"
            cursor.execute(insert_sql, tuple(row))
        conn.commit()
        print("Data inserted into PROCESSED_ACTIVE_DAILY successfully.")

        cursor.close()

    # Execute function to write df to db
    write_to_db(final_df, active_df, conn)

except pyodbc.Error as e:
    print(f"Database error: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
finally:
    if 'conn' in locals():
        conn.close()
        print("Database connection closed.")
