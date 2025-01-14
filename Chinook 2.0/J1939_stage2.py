import pandas as pd
from openpyxl import load_workbook, Workbook
import os

# Paths to the files
data_file_path = 'excel_outputs/Format_temp.xlsx'  # Correct path to the data file
j1939_limits_file = 'j1939_limit.xlsx'  # Correct path to the J1939 limits file

try:
    # Load the data file and the limits file into DataFrames
    data_df = pd.read_excel(data_file_path)
    limits_df = pd.read_excel(j1939_limits_file)

    # Print the column names in the limits file for debugging
    print("Limits file columns:", limits_df.columns)

    # Ensure numeric columns are properly formatted
    data_df['value'] = pd.to_numeric(data_df['value'], errors='coerce')  # Convert 'value' column to numeric
    limits_df['min_value'] = pd.to_numeric(limits_df['min_value'], errors='coerce')  # Convert 'min_value' column to numeric
    limits_df['max_value'] = pd.to_numeric(limits_df['max_value'], errors='coerce')  # Convert 'max_value' column to numeric

    # Drop rows with NaN values in the relevant columns
    data_df.dropna(subset=['value'], inplace=True)
    limits_df.dropna(subset=['min_value', 'max_value'], inplace=True)

    # Check if required columns are present
    if 'name' not in limits_df.columns or 'min_value' not in limits_df.columns or 'max_value' not in limits_df.columns:
        raise ValueError("Columns 'name', 'min_value', or 'max_value' are missing in the limits file.")

    # Flag to check if any cells were marked
    out_of_bounds_values = []

    # Assuming the "name" column in the data file corresponds to the "name" column in the limits file
    for idx, row in data_df.iterrows():
        # Get the name and value from the data file
        name = row['name']
        value = row['value']

        # Find the corresponding row in the limits file
        limit_row = limits_df[limits_df['name'] == name]

        # If there is a matching limit row
        if not limit_row.empty:
            min_value = limit_row['min_value'].values[0]
            max_value = limit_row['max_value'].values[0]

            # Check if the value is outside the limits (including negative numbers)
            if value < min_value or value > max_value:
                out_of_bounds_values.append(row)  # Collect the out-of-bounds row

    # Save the out-of-bounds values to a new Excel file
    if out_of_bounds_values:
        print("No out-of-bounds values found.")
        out_of_bounds_df = pd.DataFrame(out_of_bounds_values)
        out_of_bounds_file = 'excel_outputs/J1939_out_of_bounds.xlsx'
        out_of_bounds_df.to_excel(out_of_bounds_file, index=False)
    else:
        print("No out-of-bounds values found.")
        out_of_bounds_df = pd.DataFrame(out_of_bounds_values)
        out_of_bounds_file = 'excel_outputs/J1939_out_of_bounds.xlsx'
        out_of_bounds_df.to_excel(out_of_bounds_file, index=False)

except Exception as e:
    print(f"An error occurred: {e}")

# Step 5: Identifying non-duplicates in the 'name' column

# If temp2.xlsx doesn't exist, create it with headers
if not os.path.exists(data_file_path):
    wb = Workbook()  # Create a new workbook
    ws = wb.active
    ws.append(['name', 'value', 'duplicate_count'])  # Adding headers to the new sheet
    wb.save(data_file_path)  # Save the new workbook

# Load the Excel file into a DataFrame
df = pd.read_excel(data_file_path)

# Identify non-duplicates in the 'name' column
non_duplicates = ~df['name'].duplicated(keep=False)  # Boolean Series for non-duplicates

# Save the non-duplicates to a new Excel file
if non_duplicates.any():
    non_duplicates_values = df[non_duplicates]
    non_duplicates_file = 'excel_outputs/J1939_non_duplicates.xlsx'
    non_duplicates_values.to_excel(non_duplicates_file, index=False)
    print(f"Non-duplicate values saved to: {non_duplicates_file}")
else:
    non_duplicates_values = df[non_duplicates]
    non_duplicates_file = 'excel_outputs/J1939_non_duplicates.xlsx'
    non_duplicates_values.to_excel(non_duplicates_file, index=False)
    print(f"Non-duplicate values saved to: {non_duplicates_file}")
