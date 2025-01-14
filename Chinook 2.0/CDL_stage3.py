import pandas as pd
import os

def save_empty_file(file_path, columns):
    """Save an empty DataFrame with the specified columns."""
    empty_df = pd.DataFrame(columns=columns)
    empty_df.to_excel(file_path, index=False)
    print(f"Empty file created and saved to: {file_path}")

try:
    # Paths to data files
    data_file_path = 'excel_outputs/Format_temp-CDL.xlsx'
    CDL_limits_file_path = 'CDL_limit.xlsx'
    combined_file_path_CDL = 'excel_outputs/combined_statistics_CDL.xlsx'
    merged_file_path_CDL = 'excel_outputs/merged_combined_statistics_ordered_CDL.xlsx'

    # Step 1: Read and process the data file
    if not os.path.exists(data_file_path):
        print(f"Data file not found: {data_file_path}")
        save_empty_file(combined_file_path_CDL, ['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max'])
    else:
        df = pd.read_excel(data_file_path)

        required_columns = ['name', 'duplicate_count', 'value']
        if not all(col in df.columns for col in required_columns):
            raise ValueError(f"Missing required columns: {set(required_columns) - set(df.columns)}")

        grouped_df = df.groupby('name').agg(
            duplicate_count_sum=('duplicate_count', 'sum'),
            value_min=('value', 'min'),
            value_avg=('value', 'mean'),
            value_max=('value', 'max')
        ).reset_index()

        if grouped_df.empty:
            save_empty_file(combined_file_path_CDL, ['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max'])
        else:
            grouped_df.to_excel(combined_file_path_CDL, index=False)
            print(f"Grouped and statistics data saved to: {combined_file_path_CDL}")

    # Step 2: Read and process the limits file
    if not os.path.exists(CDL_limits_file_path):
        print(f"Limits file not found: {CDL_limits_file_path}")
        save_empty_file(merged_file_path_CDL, ['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max'])
    else:
        limits_df = pd.read_excel(CDL_limits_file_path)

        if 'name' not in limits_df.columns:
            raise ValueError("The 'name' column is missing in the limits file.")

        # Merge aggregated data with limits
        if os.path.exists(combined_file_path_CDL):
            grouped_df = pd.read_excel(combined_file_path_CDL)
        else:
            grouped_df = pd.DataFrame(columns=['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max'])

        merged_df = pd.merge(limits_df[['name']], grouped_df, on='name', how='left')

        if merged_df.empty:
            save_empty_file(merged_file_path_CDL, ['name', 'duplicate_count_sum', 'value_min', 'value_avg', 'value_max'])
        else:
            merged_df.to_excel(merged_file_path_CDL, index=False)
            print(f"Merged and ordered data saved to: {merged_file_path_CDL}")

except Exception as e:
    print(f"An error occurred: {e}")

