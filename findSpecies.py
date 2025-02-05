import pandas as pd
import math

# Load the Excel file
filename = 'Dakskobler_2015_merged.xlsx'
file_path = '/mnt/c/Users/jakob/Downloads/JERNEJ-dipl/exceli_iz_clankov/'+filename
sheets = pd.ExcelFile(file_path).sheet_names
print('Working on a file named:', filename)

# load excel where we write result
file_path_result = '/mnt/c/Users/jakob/Downloads/JERNEJ-dipl/tabela_za_nadaljevat_avrikelj.xlsx'
sheets_result = pd.ExcelFile(file_path_result).sheet_names
result_df = pd.read_excel(file_path_result, sheet_name=sheets_result[0], index_col=0)

# Function to process each sheet
def process_sheet(sheet):
    df = pd.read_excel(file_path, sheet_name=sheet, header=None)
    results = set()

    # Detect header rows by identifying rows where column 1 has non-empty values
    header_rows = df[df[1]=='Coordinate GK X (D-48)'].index
    if header_rows.empty:
        header_rows = df[df[1]=='Coordinate (Koordinate) GK Y (D-48)'].index
    print('Found headers at indexes: ', header_rows)

    # Iterate through tables defined by header rows
    for t in range(len(header_rows)):
        start = header_rows[t]
        if t + 1 < len(header_rows):
            end = header_rows[t+1]
        else:
            end = len(df)

        print('\n------------------------|Table: ', t, ' Range from: ', start, ' to: ', end, '---------------------------')

        table = df.iloc[start+1:end]
        coordinate_gk_row = start # Get the coordinate gk row for the current table
        # print(table)
        # 2.open new most right row in sheets_result 
        coordinate_x_row = result_df.iloc[21]
        length_col = len(result_df.iloc[0])

        # Find row of Primula auricula, handling potential errors if not found
        try:
            row_index_in_table = table[table.iloc[:, 1].astype(str) == 'Primula auricula'].index[0]
            row_index_in_table = row_index_in_table - start - 1
            # print(row_index_in_table)
            # Get the absolute row index in the original DataFrame
            row_index = table.index[row_index_in_table]
            target_row = table.loc[row_index]
            # Get the total number of columns
            num_columns = len(target_row)

            # List to store the names of columns that meet the criteria
            matching_columns = []

            # Iterate through each cell in the target row (which is a Series)
            for i, (column_name, cell_value) in enumerate(target_row.items()):
                # add primula symbols
                
                # Stop the loop if we reach the last two columns
                if i >= num_columns - 2:
                    break
                # Check if the cell value (converted to string) contains any of the target characters
                if isinstance(cell_value, str):  # Ensure it's a string to avoid errors
                    if ('r' in cell_value or '+' in cell_value or '1' in cell_value or '2' in cell_value or '3' in cell_value or '4' in cell_value or '5' in cell_value) and column_name != 0 and column_name != 1 and column_name != 2:
                        matching_columns.append(column_name)
            # Print the list of matching columns
            print('Found columns (stolpce) of Primula auricula at indexes: ', matching_columns)

            # loop through all the columns and save the names in 1 row
            for column_name in matching_columns:
                #column_name -= 1
                print(column_name)
                length_col += 1
                if column_name in table.columns:  # Ensure the column exists in the table
                    column_data = table[column_name]
                    for idx, cell_value in column_data.items():
                        # found a match in a given column
                        if isinstance(cell_value, str) and ('r' in cell_value or '+' in cell_value or '1' in cell_value or '2' in cell_value or '3' in cell_value or '4' in cell_value or '5' in cell_value):
                            # 1.get species name and save the symbol they have
                            species_name = table.loc[idx, 1]
                            species_ind = table.loc[idx, 2]
                            # rename Primula to Primula auricula s.str
                            if species_name == 'Primula auricula':
                                species_name = 'Primula auricula s.str'
                            # remove header (non relevant) names found in the table
                            remove_words = ["Number of relevé", "Database", "Elevation in m", "Aspect (Lega)", "Slope in degrees (Nagib v stopinjah)", "Successive number of relevé (Zaporedna številka popisa)", "Parent material (Matična podlaga)", "Soil (Tla)", "Stoniness in % (Kamnitost v %)", "Upper tree layer (Zgornja drevesna plast)", "Lower tree layer (Spodnja drevesna plast)", "Shrub layer (Grmovna plast)", "Herb layer (Zeliščna plast)", "Moss layer (Mahovna plast)", "Maximum diameter of trees", "Maximum height of tress", "Number of species (Število vrst)", "Relevé area (Velikost popisne ploskve)", "Date of taking relevé (Datum popisa)", "Locality (Nahajališče)", "Quadrant (Kvadrant)", "Coordinate GK Y"]
                            if not any(word in species_name for word in remove_words):
                                results.add(species_name)
                                species_name = str(species_name).strip()
                                species_symbol = cell_value
                                # on the first column write name of the file and column_name
                                result_df.loc[(result_df.index[0]), length_col] = filename
                                result_df.loc[(result_df.index[1]), length_col] = str(column_name)
                                # 1.1 save the value from 'Coordinate GK X (D-48)' and use its value in the same row as header
                                coordinate_gk_value = df.iloc[coordinate_gk_row, column_name]
                                #print(f"  - Found match for species: {species_name}, symbol: {species_symbol}, under column index: {idx}, Coordinate GK: {coordinate_gk_value}")
                                #print(species_ind)
                                # 3.find and match species name's on the most left and paste saved symbol from before to most right row. If there doesn't exsists a specie name, create it in the last row.
                                if species_name in result_df.index:
                                    if len(result_df.loc[species_name]) != len(result_df.loc['Pseudofumaria alba']):
                                        # index = (result_df.loc[species_name].index[0])
                                        # print('Index: ', index)

                                        # boolean indexing
                                        #1. Exact Match
                                        result_df.loc[(result_df.index == species_name) & (result_df[result_df.columns[0]] == species_ind), length_col] = species_symbol

                                        #2. Partial Match (handle "E3" matching "E3a", "E3b")
                                        if species_ind.startswith('E') and species_ind[-1].isdigit(): # only if the species_ind is in the form E[digits]
                                            partial_match_condition = (result_df.index == species_name) & (result_df[result_df.columns[0]].str.startswith(species_ind))
                                            print(type(species_ind))
                                            print(result_df[result_df.columns[0]])
                                            # TO JE TREBA DOKONCAT
                                            #partial_match_condition_2 = (result_df.index == species_name) & (species_ind.startswith(result_df[result_df.columns[0]].str))
                                            if partial_match_condition.any(): #or partial_match_condition_2.any():
                                                result_df.loc[(result_df.index == species_name) & (result_df[result_df.columns[0]].str.startswith(species_ind)), length_col] = species_symbol
                                            else:
                                                result_df.loc[species_name+' '] = {(result_df.columns[0]): species_ind, length_col: species_symbol}
                                                print(f"  - No E matches! Will add new species to table: {species_name}, {species_ind}, symbol: {species_symbol}, at row index: {species_name}, col index: {length_col}, Cell value: {cell_value}")
                                                continue

                                        # matching_rows_df = result_df.loc[species_name]
                                        # for index, row in matching_rows_df.iterrows():
                                        #     if species_ind in row.iloc[0]:
                                        #         #print(row.iloc[0])
                                        #         #print(species_ind)
                                        #         row.loc[length_col] = species_symbol
                                                
                                        print(f"  - Found match for species: {species_name}, {species_ind}, symbol: {species_symbol}, at row index: {species_name}, col index: {length_col}, Cell value: {cell_value}")

                                    else:
                                        # checks for NaN values (not a number) that why we have ==
                                        if species_name in result_df.index and (result_df.loc[species_name].iloc[0] == result_df.loc[species_name].iloc[0]):
                                            if species_ind in result_df.loc[species_name].iloc[0]:
                                                # NAJDU JE MATCH IN DODAL VREDNOST CISTO NA DESNO
                                                result_df.loc[species_name, length_col] = species_symbol
                                                print(f"  - Found match for species: {species_name}, {species_ind}, symbol: {species_symbol}, at row index: {species_name}, col index: {length_col}, Cell value: {cell_value}")
                                            else:
                                                length_row = len(result_df.index)
                                                # ADD new e1, e2, e3 ...
                                                #result_df.loc[species_name] = {len(result_df.iloc[1]): species_ind, length_col: species_symbol}
                                                #print(result_df.columns[0])
                                                # DODA NOVO RASTLINO V EXCELL
                                                result_df.loc[species_name] = {(result_df.columns[0]): species_ind, length_col: species_symbol}
                                                print(f"  - Will add new species to table: {species_name}, {species_ind}, symbol: {species_symbol}, at row index: {species_name}, col index: {length_col}, Cell value: {cell_value}")
                                
                                        else:
                                            length_row = len(result_df.index)
                                            # ADD new e1, e2, e3 ...
                                            result_df.loc[species_name] = {(result_df.columns[0]): species_ind, length_col: species_symbol}
                                            #result_df.loc[species_name] = {'sloji':species_ind}
                                            print(f"  - Will add new species to table: {species_name}, {species_ind}, symbol: {species_symbol}, at row index: {species_name}, col index: {length_col}, Cell value: {cell_value}")
                                else:
                                    length_row = len(result_df.index)
                                    # ADD new e1, e2, e3 ...
                                    result_df.loc[species_name] = {(result_df.columns[0]): species_ind, length_col: species_symbol}
                                    print(f"  - Will add new species to table: {species_name}, {species_ind}, symbol: {species_symbol}, at row index: {species_name}, col index: {length_col}, Cell value: {cell_value}")
                            else:
                                continue


                            # species_name_str = str(species_name).strip() # Convert to string and strip whitespace
                            # coordinate_gk_value_str = str(coordinate_gk_value).strip()

                            # if species_name_str not in result_df.index:
                            #     print(f"    - Species '{species_name_str}' not found in result sheet, appending.")
                            #     result_df.loc[species_name_str] = pd.Series(name=species_name_str) # Add new row if species not found

                            # if coordinate_gk_value_str not in result_df.columns:
                            #     print(f"    - Coordinate GK '{coordinate_gk_value_str}' not found in result sheet columns, appending.")
                            #     result_df[coordinate_gk_value_str] = None # Add new column if coordinate GK not found

                            # current_value = result_df.loc[species_name_str, coordinate_gk_value_str]
                            # if pd.isna(current_value) or current_value == '':
                            #     result_df.loc[species_name_str, coordinate_gk_value_str] = cell_value
                            #     print(f"    - Updated result sheet: Species '{species_name_str}', Coordinate GK '{coordinate_gk_value_str}', Symbol '{cell_value}'")
                            # else:
                            #     print(f"    - Result sheet already has value '{current_value}' for Species '{species_name_str}', Coordinate GK '{coordinate_gk_value_str}'. Skipping update.")
            
            

            # testing
            # print(results)
                     


        except IndexError:
            print(f"Smth went wrong in the table starting at row {start} and ending {end} in sheet {sheet}")

    return results

# Process all sheets
final_results = set()
for sheet in sheets:
    sheet_results = process_sheet(sheet)
    if sheet_results:  # Only add if the set is not empty
        final_results.update(sheet_results)
        print(final_results)
        
# save new excell
result_df.to_excel(file_path_result)

print("\nNumber of species that match:")
print(len(final_results))