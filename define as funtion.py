import pandas as pd
import random
import re
import numpy as np

def process_excel_file(input_excel_path, output_excel_path):
    # Read the Excel file
    df = pd.read_excel(input_excel_path)
    
    # Remove the last 4 characters from the 'File Name' column

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
    def process_file(index, cleaned_file_name, library_row, modify_removing_extension, df):
        new_file_name = None
    
        if isinstance(cleaned_file_name, str) and library_row:
            if modify_removing_extension.iloc[index].iloc[index] == cleaned_file_name and len(cleaned_file_name) > 18:
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

            if new_file_name is not None:
                print(f'File {index}: {new_file_name}')
                df.at[index, 'modify_file_item_ID'] = new_file_name


    # Process each file name
    for index, row in enumerate(df['File Name']):
        cleaned_file_name = clean_file_name(row)
        library_row = library_rows.iloc[index]
        #modify_removing_extension.iloc[index] = df['File Name'].str[:-4].iloc[index]
        modify_removing_extension =  df['File Name'] = df['File Name'].str[:-4]
        process_file(index, cleaned_file_name, library_row, modify_removing_extension, df)


    # Export the DataFrame to the new Excel file
    df.to_excel(output_excel_path, index=False)

# Usage example:
input_path = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
output_path = r'D:\PYDATAANALYSIS\analysis\Modified_Data.xlsx'
process_excel_file(input_path, output_path)