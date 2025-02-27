import pandas as pd
import os

# Define folder paths
parquet_folder = "parquet"
raw_folder = "raw"
input_file_path = "input_file.txt"

# Ensure the output folder exists
os.makedirs(raw_folder, exist_ok=True)

# Read the filename from input_file.txt
with open(input_file_path, "r") as f:
    filename = f.read().strip()

parquet_path = os.path.join(parquet_folder, filename)

# Check if the file exists
if not os.path.isfile(parquet_path):
    print(f"Error: File {filename} not found in {parquet_folder}.")
else:
    # Read the Parquet file
    df = pd.read_parquet(parquet_path)

    # Remove columns that contain the word "RPM" (case-insensitive)
    df_filtered = df.loc[:, ~df.columns.str.contains("RPM", case=False, na=False)]

    # Define output TXT path (tab-separated)
    txt_filename = os.path.splitext(filename)[0] + ".txt"
    txt_path = os.path.join(raw_folder, txt_filename)

    # Save as TXT file (tab-separated)
    df_filtered.to_csv(txt_path, sep="\t", index=False)

    print(f"Converted {filename} to {txt_filename} (excluding 'RPM' columns) and saved in {raw_folder}.")
