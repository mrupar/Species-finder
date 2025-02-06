import pandas as pd
from openpyxl import load_workbook

# Load the Excel file
#! MUST DELETE ALL NON-HEADER 'ZAPOREDNA ŠTEVILKA POPISA' ROWS
filename = 'Dakskobler_2015_merged.xlsx'
file_path = 'C:/Users/Miha Rupar/Desktop/python/Species-finder/exceli_iz_clankov/' + filename
sheets = pd.ExcelFile(file_path).sheet_names
print('Working on a file named:', filename)

# Load result file path
file_path_result = 'C:/Users/Miha Rupar/Desktop/python/Species-finder/tmp.xlsx'


def extract_tables(sheet):
    df = pd.read_excel(file_path, sheet_name=sheet)

    # Find all instances of 'primula auricula'
    primula_auricula_all = df[df.iloc[:, 1].str.contains("primula auricula", case=False, na=False)]
    tables_n = 0

    # DataFrames for storage
    header_df = pd.DataFrame()
    data_df = pd.DataFrame()

    for index, row in df.iterrows():
        if index == 0 or ('zaporedna številka popisa' in str(row.iloc[1]).lower() and index > 1):
            if tables_n < len(primula_auricula_all):
                primula_auricula = primula_auricula_all.iloc[tables_n]
                tables_n += 1

                # Remove columns containing only '.' or NaN
                filtered_primula = primula_auricula[primula_auricula != "."].dropna()

                # Remove last two columns (if necessary)
                filtered_primula = filtered_primula.iloc[:-2]

            else:
                raise Exception('No more primula auricula tables found')
            continue

        if pd.isna(row[1]) or row[1] is None:
            continue

        # Store header rows separately
        if pd.isna(row.iloc[-1]) or row.iloc[-1] is None:
            if tables_n > 0:
                key = str(row[1]).strip()
                new_values = row[row.index.isin(filtered_primula.index)].to_frame().T  # Convert to DataFrame
                new_values.index = [key]  # Set key as row name

                if key in header_df.index:
                    header_df.loc[key] = header_df.loc[key].combine_first(new_values.iloc[0])
                else:
                    header_df = pd.concat([header_df, new_values])
            else:
                raise Exception('No table initialized')

        else:
            if tables_n > 0:
                has_auricula = any(row[col] in ['r', '+', 1, 2, 3, 4, 5, '1', '2', '3', '4', '5'] for col in filtered_primula.index)

                if has_auricula:
                    unique_key = f'{row[1]} {row[2]}'.strip()
                    new_values = row[row.index.isin(filtered_primula.index)].to_frame().T  # Convert to DataFrame
                    new_values.index = [unique_key]  # Set unique key as row name

                    # Append to existing entry if the key exists
                    if unique_key in data_df.index:
                        data_df.loc[unique_key] = data_df.loc[unique_key].combine_first(new_values.iloc[0])
                    else:
                        data_df = pd.concat([data_df, new_values])
                else:
                    print('No auricula with:', row[1])
            else:
                raise Exception('No table initialized')

    return header_df, data_df


for sheet in sheets:
    header_df, data_df = extract_tables(sheet)

    # Combine header and data safely
    new_table = pd.concat([header_df, data_df], axis=0)

    # Save to Excel and append correctly
    try:
        with pd.ExcelWriter(file_path_result, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            book = load_workbook(file_path_result)
            
            if "Table" in book.sheetnames:
                existing_df = pd.read_excel(file_path_result, sheet_name="Table", index_col=0)

                # Merge new data with existing data
                new_table = existing_df.combine_first(new_table)

            new_table.to_excel(writer, sheet_name="Table")
    except FileNotFoundError:
        new_table.to_excel(file_path_result, sheet_name="Table")
