import pandas as pd
import os

# Directory containing Excel files
directory = 'C:/Users/Miha Rupar/Desktop/python/jernej-diplomska/Species-finder/results'

# Initialize an empty DataFrame
merged_df = pd.DataFrame()

# Iterate over all Excel files in the directory
for file in os.listdir(directory):
    if file.endswith(".xlsx"):  # Process only Excel files
        file_path = os.path.join(directory, file)
        print(f"Processing file: {file_path}")

        # Load the Excel file and iterate through all sheets
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)

            # If merged_df is empty, initialize it with the first DataFrame
            if merged_df.empty:
                merged_df = df
            else:
                # Merge DataFrames on index (rows), keeping all columns
                merged_df = pd.merge(merged_df, df, left_index=True, right_index=True, how='outer')

# Save the merged table to a new Excel file
output_file = 'C:/Users/Miha Rupar/Desktop/python/jernej-diplomska/Species-finder/merged_result.xlsx'
merged_df.to_excel(output_file, sheet_name='Merged')

print(f'Tables successfully merged and saved to {output_file}!')
