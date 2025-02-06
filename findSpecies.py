import pandas as pd
from openpyxl import load_workbook
import os

directory = 'C:/Users/Miha Rupar/Desktop/python/jernej-diplomska/Species-finder/exceli_iz_clankov'

filename = 'Dakskobler_2015_merged.xlsx'
file_path = os.path.join(directory, filename)
sheets = pd.ExcelFile(file_path).sheet_names

# Load result file path
file_path_result = 'C:/Users/Miha Rupar/Desktop/python/jernej-diplomska/Species-finder/result.xlsx'


def extract_tables(file, sheet, header_df, data_df):
    df = pd.read_excel(file, sheet_name=sheet)

    # Find all instances of 'primula auricula'
    primula_auricula_all = df[df.iloc[:, 1].str.contains("primula auricula", case=False, na=False)]
    if primula_auricula_all.empty:
        raise Exception('No primula auricula tables found')

    # Get first instance of 'primula auricula'
    primula_auricula = primula_auricula_all.iloc[0]
    
    # tables in the sheet
    tables_n = 0

    for index, row in df.iterrows():
        # Skip rows with all '.' or NaN
        if row.iloc[2:].replace('.', pd.NA).isna().all() or \
            'številka popisa' in str(row.iloc[1]).lower() or \
                'štev. popisa' in str(row.iloc[1]).lower():
            continue
        if tables_n == 0 or ('(' in str(row.iloc[1]) and index > primula_auricula.name and not 'e' in str(row.iloc[2]).lower()):

            if tables_n < len(primula_auricula_all):
                primula_auricula = primula_auricula_all.iloc[tables_n]
                tables_n += 1

                # Remove columns containing only '.' or NaN
                filtered_primula = primula_auricula[primula_auricula != "."].dropna()

                # Remove last two columns
                filtered_primula = filtered_primula.iloc[:-2]

                # Remove first three columns
                filtered_primula = filtered_primula.iloc[2:]

            else:
                raise Exception('More tables found than expected.')
            continue

        # Skip empty rows
        if pd.isna(row.iloc[1]) or row.iloc[1] is None:
            continue

        # Store header rows separately
        #? Sometimes my genius is... it's almost frightening
        if '(' in str(row.iloc[1]) and not 'e' in str(row.iloc[2]).lower():
            if tables_n > 0:
                if row.iloc[2] is None or pd.isna(row.iloc[2]):
                    unique_key = str(row.iloc[1]).strip()
                else:
                    unique_key = f'{row.iloc[1]} {row.iloc[2]}'.strip()
                new_values = row[row.index.isin(filtered_primula.index)].to_frame().T  # Convert to DataFrame
                new_values.index = [unique_key]  # Set key as row name
                # Append to existing entry if the key exists
                if unique_key in header_df.index:
                    header_df.loc[unique_key] = header_df.loc[unique_key].combine_first(new_values.iloc[0])
                elif header_df.empty:
                    header_df = new_values
                else:
                    header_df = pd.concat([header_df, new_values])
            else:
                raise Exception('No table initialized')

        else:
            if tables_n > 0:
                # Check if row contains auricula
                has_auricula = any(str(row[col]).strip() in ['r', '+', '1', '2', '3', '4', '5'] for col in filtered_primula.index)

                if has_auricula:
                    unique_key = f'{row.iloc[1]} {row.iloc[2]}'.strip()
                    new_values = row[row.index.isin(filtered_primula.index)].to_frame().T  # Convert to DataFrame
                    new_values.index = [unique_key]  # Set unique key as row name

                    # Append to existing entry if the key exists
                    if unique_key in data_df.index:
                        data_df.loc[unique_key] = data_df.loc[unique_key].combine_first(new_values.iloc[0])
                    elif data_df.empty:
                        data_df = new_values
                    else:
                        data_df = pd.concat([data_df, new_values])
                else:
                    continue
            else:
                raise Exception('No table initialized')

    return header_df, data_df


header = pd.DataFrame()
data = pd.DataFrame()

for file in os.listdir(directory):
    if file.endswith(".xlsx"):
        print(f"Processing file: {file}")
        file_path = os.path.join(directory, file)
        sheets = pd.ExcelFile(file_path).sheet_names

        for sheet in sheets:
            header, data = extract_tables(file=file_path, sheet=sheet, header_df=header, data_df=data)

# Combine header and data
new_table = pd.concat([header, data], axis=0)

if os.path.exists(file_path_result):
    with pd.ExcelWriter(file_path_result, mode='a', engine='openpyxl', if_sheet_exists="replace") as writer:
        book = load_workbook(file_path_result)

        if "Table" in book.sheetnames:
            existing_df = pd.read_excel(file_path_result, sheet_name="Table", index_col=0)

            new_table = existing_df.combine_first(new_table)

        new_table.to_excel(writer, sheet_name="Table", index=True)
else:
    new_table.to_excel(file_path_result, sheet_name="Table", index=True)

print("All files processed successfully!")
