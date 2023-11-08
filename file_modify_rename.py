import pandas as pd
import random
import re
import numpy as np

def process_excel_file(input_excel_path, output_excel_path):
    # Read the Excel file
    df = pd.read_excel(input_excel_path)
    # Remove the last 4 characters from the 'File Name' column
    modify_removing_extension = df['File Name'] = df['File Name'].str[:-4]
    #modify_removing_extension= df['Modified File Name'] = df['File Name'].str[:-4]
    #modify_removing_extension= df['File Name'].str[:-4].iloc[index]
    # Check for library rows
    library_rows = df['Full File Name'].str.contains(r'\\Library\\|\\Libraries\\|\\Content Center Files\\', case=False, na=False)
    # Initialize 'modify_file_item_ID' column
    df['modify_file_item_ID'] = None

    def clean_file_name(file_name):
        if isinstance(file_name, str):
            # Define the list of patterns to remove
            patterns_to_remove = [
                (r'\([^)]*\)', ''),     # Remove text within parentheses
                (r'\[[^\]]*\]', ''),    # Remove text within brackets
                ('function key', ''),   # Remove 'function key'
                ('ÇÃ', ''),             # Remove 'ÇÃ'
                ('Copyof', ''),         # Remove 'Copyof'
                ('Copy of', ''),         # Remove 'Copyof'
                ('Ç Ã', ''),            # Remove 'Ç Ã'
                (',', ''),              # Remove commas
                ('DEL', ''),            # Remove 'DEL'
                ('NULL', ''),           # Remove 'NULL'
                ('ISO', ''),            # Remove 'ISO'
                ('x', ''),              # Remove 'x'
                ('test', ''),           # Remove 'test'
                ('old', ''),            # Remove 'old'
                ('move', ''),           # Remove 'move'
                (' ', ''),              # Remove spaces
            ]

            # Apply each pattern to the file name
            for pattern, replacement in patterns_to_remove:
                file_name = re.sub(re.escape(pattern), replacement, file_name)

            # Remove any double spaces that may have been created
            file_name = re.sub(r' +', ' ', file_name)
            # Remove trailing periods
            file_name = file_name.rstrip('.')
            file_name = file_name.rstrip('_')

            # Remove special characters individually
            special_characters = r'[!@#$%^&*(){}[\]><\/|?+|]'
            file_name = re.sub('[' + re.escape(special_characters) + ']', '', file_name)
            file_name = file_name.upper()
        

        return file_name

    # Process each file name

    for index, row in enumerate(df['File Name']):
        cleaned_file_name = clean_file_name(row)
        if isinstance(cleaned_file_name, str) and library_rows.iloc[index]:
            if modify_removing_extension.iloc[index] == cleaned_file_name and len(cleaned_file_name) > 18:
                new_file_name = cleaned_file_name[:13] + '_DNU'
                print(f'File {index}: Not Modified (len>18 and DNU): {new_file_name}')
            elif modify_removing_extension.iloc[index] != cleaned_file_name:
                new_file_name = cleaned_file_name[:13] + '_DNU'
                print(f'File {index}: Not Modified (len>18 and DNU): {new_file_name}')
            else:
                # Handle the case where cleaned_file_name is not alpha
                new_file_name = cleaned_file_name  # Or specify another action as needed
        else:
            if isinstance(cleaned_file_name, str):
                if modify_removing_extension.iloc[index] == cleaned_file_name and len(cleaned_file_name) > 18 :
                    new_file_name = cleaned_file_name[:13]
                    new_file_name = new_file_name.rstrip('-') + '-P01'
                    print(f'File {index}: Not Modified (len>18 and P01): {new_file_name}')
                elif modify_removing_extension.iloc[index] != cleaned_file_name :
                    new_file_name = cleaned_file_name[:13]
                    new_file_name = new_file_name.rstrip('-') + '-P01'
                    print(f'File {index}: Not Modified (len>18 and P01): {new_file_name}')
                else:
                    # Handle the case where cleaned_file_name is not alpha
                    new_file_name = cleaned_file_name  # Or specify another action as needed

        print(f'File {index}: {new_file_name}')
        df.at[index, 'modify_file_item_ID'] = new_file_name
     # Export the DataFrame to the new Excel file

    def process_file(index, cleaned_file_name, library_row):
        new_file_name = None  # Initialize new_file_name

        if isinstance(cleaned_file_name, str):
            clean_file_name_for_alfa = cleaned_file_name.replace('_', '')
            if clean_file_name_for_alfa.isalpha():
                if library_row:
                    random_digits = ''.join([str(random.randint(0, 9)) for _ in range(7)])
                    new_file_name = f"MM{str(len(cleaned_file_name))}{random_digits}_DNU"
                    print(f'File {index}: Not Modified (is alpha): {new_file_name}')
                    print('Alfa:', clean_file_name_for_alfa)
                else:
                    random_digits = ''.join([str(random.randint(0, 9)) for _ in range(7)])
                    new_file_name = f"MM{str(len(clean_file_name_for_alfa))}{random_digits}-P01"
                    print(f'File {index}: Not Modified (is alpha): {new_file_name}')
                    print('Alfa:', clean_file_name_for_alfa)

        if new_file_name is not None:
            print(f'File {index}: {new_file_name}')
            df.at[index, 'modify_file_item_ID'] = new_file_name

    # Process each file name
    for index, row in enumerate(df['File Name']):
        cleaned_file_name = clean_file_name(row)
        library_row = library_rows.iloc[index]
        process_file(index, cleaned_file_name, library_row)


    df.to_excel(output_excel_path, index=False)

# Usage example:
input_path = r'D:\PYDATAANALYSIS\analysis\Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
process_excel_file(input_path, output_path)
