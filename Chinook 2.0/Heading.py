import pandas as pd
import os

# File paths
input_txt = "input_file.txt"
vehicle_xlsx = "Vehicle_details.xlsx"
output_folder = "excel_outputs"
output_xlsx = os.path.join(output_folder, "Heading.xlsx")

# Read the search term from the text file (keep case sensitivity)
with open(input_txt, "r", encoding="utf-8") as file:
    search_term = file.readline().strip()  # Read first line and clean spaces

# Load the Excel sheet
df = pd.read_excel(vehicle_xlsx, sheet_name=0, engine="openpyxl", dtype=str)

# Trim spaces from all cells but keep original case
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Find the row where Column A matches the search term exactly
match = df[df.iloc[:, 0] == search_term]

# Extract data if a match is found
if not match.empty:
    # Extract values from Columns B-E
    row_data = match.iloc[:, 1:6].values.flatten().tolist()
    
    # Get headers for Columns B-E
    output_headers = df.columns[1:6]  # Get headers for B-E only
    
    # Create output DataFrame (without the search term)
    output_df = pd.DataFrame([row_data], columns=output_headers)
    
    # Save to Excel
    output_df.to_excel(output_xlsx, index=False, engine="openpyxl")
    print(f"Match found! Data saved in '{output_xlsx}'")
else:
    print(f"No match found for '{search_term}'. Output file will be empty.")


