import os
import pandas as pd
from sqlalchemy import create_engine
from concurrent.futures import ThreadPoolExecutor

# Athena connection details
region = 'us-east-1'
s3_staging_dir = 's3://aws-athena-query-results-us-east-1-770418278010/query-results/'
database = 'raw'

# Create Athena engine
engine = create_engine(f'awsathena+rest://@athena.{region}.amazonaws.com:443/{database}?s3_staging_dir={s3_staging_dir}')

# Read the date from the 'date' file
try:
    with open('date.txt', 'r') as f:
        query_date = f.readline().strip()  # Read the single line (e.g., '2025-01-02')
        year, month, day = query_date.split('-')
except FileNotFoundError:
    print("The 'date.txt' file was not found.")
    exit(1)
except ValueError:
    print("The date in 'date.txt' is not in the correct format (YYYY-MM-DD).")
    exit(1)
except Exception as e:
    print(f"Error reading date file: {e}")
    exit(1)

# Read the cust_code and device names from the text file
device_file = 'devices_list.txt'
try:
    with open(device_file, 'r') as f:
        lines = [line.strip() for line in f if line.strip()]  # Remove empty lines and strip whitespace
        cust_code = lines[0]  # First line contains the cust_code
        device_names = lines[1:]  # Remaining lines contain device names
except FileNotFoundError:
    print(f"The device list file '{device_file}' was not found.")
    exit(1)
except Exception as e:
    print(f"Error reading device file: {e}")
    exit(1)

# SQL query template with placeholders for year, month, day, cust_code, and device_name
query_template = """
SELECT value, name
FROM raw."4sight_raw_sensors"
WHERE 
    (substr(name, 1, 3) = 'CDL' OR substr(name, 1, 5) = 'J1939')
    AND device IN ('{device_name}')
    AND cust_code = '{cust_code}'
    AND year = '{year}'
    AND month = '{month}'
    AND day = '{day}';
"""

# Function to execute a query for a device and save the results
def execute_device_query(device_name):
    query = query_template.format(
        device_name=device_name,
        cust_code=cust_code,
        year=year,
        month=month,
        day=day
    )
    try:
        print(f"Executing query for {device_name} on {query_date}...")
        df = pd.read_sql(query, engine)
        
        # Save results for each device
        output_folder = 'parquet'
        os.makedirs(output_folder, exist_ok=True)
        file_path = os.path.join(output_folder, f"{device_name}.parquet")
        df.to_parquet(file_path, index=False, compression='snappy')  # Added compression
        print(f"Results for {device_name} saved to {file_path}")
        return file_path
    except Exception as e:
        print(f"Error executing query for {device_name}: {e}")
        return None

# Execute queries in parallel for each device
with ThreadPoolExecutor(max_workers=4) as executor:  # Adjust workers based on system capacity
    futures = {executor.submit(execute_device_query, device): device for device in device_names}
    
    # Wait for all threads to complete and collect results
    successes = []
    failures = []

    for future in futures:
        device_name = futures[future]
        try:
            result = future.result()
            if result:
                successes.append(device_name)
        except Exception as e:
            failures.append(device_name)
            print(f"Error with {device_name}: {e}")

# Print summary of results
print("\nSummary:")
print(f"Successful queries: {len(successes)}")
print(f"Failed queries: {len(failures)}")
if failures:
    print(f"Failed devices: {', '.join(failures)}")