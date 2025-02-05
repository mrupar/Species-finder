import pandas as pd
# Load the uploaded Excel file
file_name = 'drakskobler_in_rozman_2021_2'
file_path = '/mnt/c/Users/jakob/Downloads/'+file_name+'.xlsx'

# Initialize an empty DataFrame for the combined data
combined_data = pd.DataFrame()

# Load all sheet names
sheet_names = pd.ExcelFile(file_path).sheet_names


# Function to remove rows that are half empty
def remove_half_empty_rows(df):
    threshold = len(df.columns) / 2
    return df.dropna(thresh=threshold)


# Process sheets in pairs to combine them
for i in range(0, len(sheet_names), 2):
    # Load sheets
    sheet1 = pd.read_excel(file_path, sheet_name=sheet_names[i])
    sheet2 = pd.read_excel(file_path, sheet_name=sheet_names[i + 1]) if i + 1 < len(sheet_names) else None

    # Remove rows that are half empty
    # sheet1 = remove_half_empty_rows(sheet1)
    #if sheet2 is not None:
        #sheet2 = remove_half_empty_rows(sheet2)


    # Ensure unique column names for both sheets
    #sheet1.columns = [f"{col}_1" for col in sheet1.columns]
    if sheet2 is not None:
        # sheet2.columns = [f"{col}_2" for col in sheet2.columns]
        # Combine the two sheets side by side (columns)
        combined_pair = pd.concat([sheet1, sheet2], axis=0)
    else:
        combined_pair = sheet1  # If there's no pair, use just the one sheet

    # Append the combined pair below the existing combined data
    combined_data = pd.concat([combined_data, combined_pair], axis=0, ignore_index=True)

# Save the combined data to a new Excel file
output_path = file_name+'_merged.xlsx'
combined_data.to_excel(output_path, index=False)
output_path

