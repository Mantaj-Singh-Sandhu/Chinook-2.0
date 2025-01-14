import pandas as pd
from pathlib import Path

# Paths to the data files
data_file_path = Path('excel_outputs/Format_temp.xlsx')  # Path to the data file
j1939_limits_file_path = Path('j1939_limit.xlsx')  # Path to the J1939 limits file

# Output file paths
combined_file_path = Path('excel_outputs/combined_statistics_J1939.xlsx')
merged_file_path = Path('excel_outputs/merged_combined_statistics_ordered_J1939.xlsx')

try:
    # Step 1: Read the data file and validate columns
    df = pd.read_excel(data_file_path)
    required_columns = {'name', 'duplicate_count', 'value'}
    if not required_columns.issubset(df.columns):
        raise ValueError(f"Columns {required_columns - set(df.columns)} are missing in the data file.")

    # Group by 'name' and calculate statistics
    grouped_df = df.groupby('name').agg(
        duplicate_count_sum=('duplicate_count', 'sum'),
        value_min=('value', 'min'),
        value_avg=('value', 'mean'),
        value_max=('value', 'max')
    ).reset_index()

    # Ensure output file is created even if the DataFrame is empty
    if grouped_df.empty:
        grouped_df = pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max'])
    grouped_df.to_excel(combined_file_path, index=False)
    print(f"Grouped and statistics data saved to: {combined_file_path}")

    # Step 2: Read the limits file and validate columns
    limits_df = pd.read_excel(j1939_limits_file_path)
    if 'name' not in limits_df.columns:
        raise ValueError("The 'name' column is missing in the limits file.")

    # Step 3: Merge the limits data with the aggregated data
    merged_df = pd.merge(limits_df[['name']], grouped_df, on='name', how='left')

    # Ensure output file is created even if the DataFrame is empty
    if merged_df.empty:
        merged_df = pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max'])
    merged_df.to_excel(merged_file_path, index=False)
    print(f"Merged and ordered data saved to: {merged_file_path}")

except FileNotFoundError as fnf_error:
    print(f"File not found: {fnf_error}")
    # Create empty output files if source files are missing
    pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max']).to_excel(combined_file_path, index=False)
    pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max']).to_excel(merged_file_path, index=False)
except ValueError as val_error:
    print(f"Value error: {val_error}")
    # Create empty output files if validation fails
    pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max']).to_excel(combined_file_path, index=False)
    pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max']).to_excel(merged_file_path, index=False)
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    # Create empty output files in case of unexpected errors
    pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max']).to_excel(combined_file_path, index=False)
    pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max']).to_excel(merged_file_path, index=False)


