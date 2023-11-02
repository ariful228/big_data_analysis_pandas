#Bismillahir Rahmanir Raheem

#Big data analaysis
#Script developed by Ariful Islam

import pandas as pd
print("Start Working!")
try:
    # Read the data from Excel into DataFrames from Sheet1, Sheet2, and Sheet3
    df1 = pd.read_excel('analysis/EDM_Vault_Migration_Mining_Screen_Pilot_request.xlsx', sheet_name='Sheet1')
    df2 = pd.read_excel('analysis/EDM_Vault_Migration_Mining_Screen_Pilot_request.xlsx', sheet_name='Sheet2')
    

    def extractDefault(df1, df2):
        # Create an empty list to store the results
        results = []

        for index, row in df1.iterrows():
            # Extract the 'Part Number' and 'SAP_Material' from the current row, and convert them to strings
            part_number = str(row['Part Number']).strip()
            Full_file_name = str(row['Full File Name']).strip()
            File_Name = str(row['File Name']).strip()

            # Find matching rows in df2 based on 'SAP_Document' and 'SAP_Material'
            matching_rows_df2 = df2[(df2['SAP_Document'].str.strip() == part_number) |
                                   (df2['SAP_Material'].str.strip() == part_number) |
                                   (df2['File Name'].str.strip() == File_Name) |
                                   (df2['Full File Name'].str.strip() == Full_file_name) |
                                   (df2['Part Number'].str.strip() == part_number)] 
                                   

            if not matching_rows_df2.empty:
                # Take the first matching row from df2
                matching_row_df2 = matching_rows_df2.iloc[0]
            else:
                matching_row_df2 = pd.Series()  # Create an empty Series if no match is found in df2

            # Rename columns in matching rows to avoid duplicates
            matching_row_df2 = matching_row_df2.add_prefix('df2_')

            # Merge the current row from df1, matching rows from df2 
            merged_row = pd.concat([row, matching_row_df2])

            # Append the merged row to the results list
            results.append(merged_row)

        # Create a new DataFrame from the results list
        merged_df = pd.concat(results, axis=1).T

        # Save the merged_df to an Excel file
        merged_df.to_excel('analysis\Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx', index=False)  # Change 'output.xlsx' to your desired file name

    # Call the function to extract and save the results
    extractDefault(df1, df2)

except FileNotFoundError as e:
    print(f"Error: File not found - {e}")
except Exception as e:
    print(f"An error occurred: {e}")
print("Revision and SAP Material merged Done!")


print("File name modify and renaming statr!")
#tc&file revision and SAP material marge done!!
#>>>>>>Item id and file name processing<<<<<<<<
#import pandas as pd
import random
import re
import numpy as np

def process_excel_file(input_excel_path, output_excel_path):
        df = pd.read_excel(input_excel_path)
        #input_file_path = r'analysis/Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
        #output_file_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
        #process_dataframe(input_file_path, output_file_path)
        # Read the Excel file
        #df = pd.read_excel(input_excel_path)
        modify_removing_extension= df['file_name2'] = df['File Name'].apply(lambda x: str(x))
        modify_removing_extension = df['File Name'] = df['File Name'].str[:-4]
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

        df['File Name'] = df['file_name2']
        df.drop('file_name2', axis=1, inplace=True)

        #df = pd.read_excel(input_path)
        #library_rows = df['Full File Name'].str.contains(r'\\Library\\|\\Libraries\\|\\Content Center Files\\', case=False, na=False)
        # Define a function to apply the updated conditions and create the 'Item ID' column
        def create_item_id(row):
                if pd.notna(row['df2_SAP_Material']):
                    return row['df2_SAP_Material']
                if pd.notna(row['df2_Rename_Item_ID']):
                    return row['df2_Rename_Item_ID']
                if row['Part Number'] == row['df2_SAP_Document']:
                    return row['df2_SAP_Material']
                if library_rows[row.name]:
                    if row['Part Number'] != row['modify_file_item_ID'] and len(str(row['Part Number'])) < 18:
                        return row['Part Number']
                    if len(str(row['Part Number'])) > 18:
                        return row['modify_file_item_ID']
                    if row['Part Number'] == row['modify_file_item_ID']:
                        return row['Part Number']
                else:
                    if row['Part Number'] != row['modify_file_item_ID']:
                        return row['modify_file_item_ID']
                    elif row['Part Number'] == row['modify_file_item_ID']:
                        return row['Part Number']
                return None  # Default case

        # Apply the function to create the 'Item ID' column
        df['Item ID'] = df.apply(create_item_id, axis=1)
        df.to_excel(output_excel_path, index=False)
        # Usage example:
        input_path = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
        output_path = r'D:\PYDATAANALYSIS\analysis\Modified_Data.xlsx'
        process_excel_file(input_path, output_path)
        print("File modify and renaming Done!")



        print("Start item revision processing")
        #>>>>>>>Item id and file name processing DOEN<<<<<<
        #START revision processing********
        #import pandas as pd
        # Define the function to update 'Revision Number (iProperty)'
        def update_revision_number(row):
            tc_revision = row['df2_TC_REV']
            file_revision = row['df2_Revision']
            bl_revision = row['Revision Number (iProperty)']
            file_name = str(row['File Name'])  # Convert to string
            modified_file_name = file_name[:-4]

            if pd.notna(tc_revision):
                # If tc_revision is not NaN, return tc_revision
                return tc_revision

            elif pd.notna(file_revision):
                # If tc_revision is NaN but 'df2_Revision' is not NaN, return file_revision
                return file_revision

            # Additional conditions for 'Revision Number (iProperty)'
            else:
                if row['File Extension'] in ['dwg', 'ipt', 'iam'] and pd.notna(row['Part Number (iProperty)']):
                    bl_revision = row['Revision Number (iProperty)']

                    if pd.notna(bl_revision):
                        matching_rows = df[df['Part Number'] == row['Part Number']]
                        df.loc[matching_rows.index, 'Revision Number (iProperty)'] = bl_revision

                    if pd.notna(bl_revision):
                        matching_rows = df[df['File Name'].str[:-4] == modified_file_name]
                        df.loc[matching_rows.index, 'Revision Number (iProperty)'] = bl_revision

                # Return the final value for 'Revision Number (iProperty)'
                return bl_revision

        # Load the Excel file into a DataFrame
        #file_path = r'D:\PYDATAANALYSIS\analysis\Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
        #df = pd.read_excel(file_path)
        # Apply the `update_revision_number` function to update 'Revision Number (iProperty)'
        df['Revision Number (iProperty)'] = df.apply(update_revision_number, axis=1)
        # Save the updated DataFrame back to the Excel file
        #output_path = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
        #df.to_excel(output_path, index=False)
        print("Item revision processing Done!!")


        print("Start Item type change processing")
        # item revision set Done<<<<<<<<<<
        #Start Item type change
        #import pandas as pd
        # Load the Excel file into a DataFrame
        #file_path = r'E:\Dataanalysis\Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
        #df = pd.read_excel(file_path)
        #Item type
        #import pandas as pd
        # Load the Excel file into a DataFrame
        #file_path = r'D:\PYDATAANALYSIS\analysis\Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
        #df = pd.read_excel(file_path)
        # Define a function to update 'Item Type' based on conditions                                                                                               
        def update_item_type(row):
            # 1st step: Update 'Item Type' to 'Document' for DWG files meeting the conditions
            if row['File Extension'] == 'dwg' and row['Part Number'] in df['df2_SAP_Document'] and row['df2_SAP_Document'] != row['df2_SAP_Material'] and row['Part Number'] != row['df2_SAP_Material']:
                return 'Document'
            
            # 2nd step: Update 'Item Type' to 'Document' for Excel files
            if row['File Extension'] == 'excel':
                return 'Document'

            # 3rd step: Update 'Item Type' to 'Item' for all other files
            return 'Item'

        # Apply the update_item_type function to update 'Item Type' column
        df['Item Type'] = df.apply(update_item_type, axis=1)
        #output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx' 
        # Save the updated DataFrame back to the Excel file
        #df.to_excel(output_path, index=False)
        print("Item type changing Done!")


        print("Start Metadata processing")       
        #************************************
        #Metadata process
        #import pandas as pd
        # Load the Excel file into a DataFrame
        #input_path = r'analysis/Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
        #output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
        #df = pd.read_excel(input_path)
        # Define a function to collect values following the specified order
        def collect_values(row):
            if not pd.isna(row['Description']):
                return row['Description']
            elif not pd.isna(row['Item Rev. Name']):
                return row['Item Rev. Name']
            elif not pd.isna(row['Part Number']):
                return row['Part Number']
            else:
                return None  # Set to null if all columns are null

        # Apply the function to create a new 'Item_rev_Name' column
        df['Item_rev_Name'] = df.apply(collect_values, axis=1)
        # Save the updated DataFrame to an Excel file
        #df.to_excel(output_path, index=False)
        print("Item_rev_Name column added and DataFrame saved to Excel.")

# Process the DataFrame
print('Metadata updating done!>>>>>')


print("Start Renaming duplicate item ID")
### DONE!!!!!!!!!! for Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx

print('Start item duplicate renaming')
#import pandas as pd
from collections import Counter

def process_dataframe(input_file, output_file):
    try:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(input_file)
        # Define input and output file paths
        input_file_path = r'Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
        output_file_path = r'Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
        # Process the DataFrame
        process_dataframe(input_file_path, output_file_path)

        # Step 1: Count the duplicates for each 'Item ID' and create a new column 'Duplicates Count'
        df['Item ID'] = df['Item ID'].astype(str)
        df['Duplicates Count'] = df.groupby('Item ID')['Item ID'].transform('count')

        # Identify library rows
        library_rows = df['Full File Name'].str.contains(r'\\Library\\|\\Libraries\\|\\Content Center Files\\', case=False, na=False)

        # Step 2: Identify duplicates with a count of 3
        duplicates_3 = (df['Duplicates Count'] >= 2)
        #duplicates_4 = (df['Duplicates Count'] >= 4)

        # Check if the combined extensions in each group form "iamiptdwg"
        specific_extensions_set_1 = df[duplicates_3].groupby('Item ID')['File Extension'].transform(lambda x: ''.join(sorted(x)).lower() in ('iptipt', 'iamiam', 'idwidw', 'dwgdwg', 'iptiptipt', 'iamiamiam', 'dwgdwgdwg','iptiptiptipt', 'iamiamiamiam', 'dwgdwgdwgdwg', 'iptiptiptiptipt', 'iamiamiamiamiam', 'dwgdwgdwgdwgdwg', 'iptiptiptiptiptipt', 'iamiamiamiamiamiam', 'dwgdwgdwgdwgdwgdwg', 'iptiptiptiptiptiptipt', 'iamiamiamiamiamiamiam', 'dwgdwgdwgdwgdwgdwgdwg', 'iptiptiptiptiptiptiptipt', 'iamiamiamiamiamiamiamiam', 'dwgdwgdwgdwgdwgdwgdwgdwg', 'iptiptiptiptiptiptiptiptipt', 'iamiamiamiamiamiamiamiamiam', 'dwgdwgdwgdwgdwgdwgdwgdwgdwg',))
        # Create a Counter object to count the same type of output
        output_counter = Counter()
        # Check if at least one of the extensions is 'ipt' among the duplicates
        ipt_extension = df[duplicates_3]['File Extension'].str.lower() == 'ipt'
        specific_extensions_set_2 = df[duplicates_3].groupby('Item ID')['File Extension'].transform(lambda x: ''.join(sorted(x)).lower() in ('iamiptdwg', 'iptiamdwg', 'dwgiptiam', 'dwgiamipt','iamiptidw', 'iptiamidw', 'idwiptiam', 'idwiamipt'))
                
      
        def rename_item_id(row):
            new_final_item_id = row['Item ID']  # Initialize with the original 'Item ID'
            # Duplicate count 3
            if (row['Duplicates Count'] >= 2):
                if specific_extensions_set_1[row.name]:
                    output_counter[new_final_item_id] += 1
                    serial_number = output_counter[new_final_item_id]

                    if not library_rows[row.name]:
                        if serial_number == 1:
                            new_final_item_id = row['Item ID'].replace('-P01', '')
                            print('iptiptipt_1st count action:', new_final_item_id)
                        if serial_number >= 2:
                            serial_number -= 1
                            new_final_item_id = f"{row['Item ID'][:14]}-P{serial_number:02}"
                            print('iptiptipt_2nd count action:', new_final_item_id)
                    else:
                        if serial_number == 1:
                            new_final_item_id = row['Item ID'].replace('_DNU', '')
                            print('iptiptipt/iptipt_1st count action without _DNU:', new_final_item_id)
                        if serial_number >= 2:
                            new_final_item_id = row['Item ID'][:13] + '_DNU1'
                            serial_number -= 1
                            new_final_item_id = f"{row['Item ID'][:13]}_DNU{serial_number:01}"
                            print('iptiptipt/iptipt_2nd count action with _DNU:', new_final_item_id)
                else:
                    if specific_extensions_set_2[row.name]:
                       output_counter[new_final_item_id] += 1
                       serial_number = output_counter[new_final_item_id]
                       if ipt_extension[row.name]:
                        # Create 'Final Item ID' by modifying 'Item ID'
                          new_final_item_id = row['Item ID'][:14] + '-P01'
                          print('Last-iptiamdwg', new_final_item_id)

            return new_final_item_id

        # Apply renaming for 'Final Item ID' for the appropriate rows
        df['Final Item ID'] = df.apply(rename_item_id, axis=1)

        # Save the updated DataFrame back to a new Excel file
        df.to_excel(output_file, index=False)
        print("Data processing and saving completed successfully.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

#Dataanalysis done!
#Start exporting *****
print("Renaming duplicate item status done")



print("Start Exporting")
#import pandas as pd
# Specify the XLSX file path
xlsx_file = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
# List of columns to export
columns_to_export = ["Full File Name", "Final Item ID", "Revision Number (iProperty)", "Item Type", "Item_rev_Name"]
try:
    # Read the XLSX file into a pandas DataFrame
    df = pd.read_excel(xlsx_file)

    # Select only the specified columns
    df = df[columns_to_export]

    # Rename columns to match the desired output (optional)
    df = df.rename(columns={
        "Full File Name": "Full File Name",
        "Final Item ID": "Final Item ID",
        "Revision Number (iProperty)": "Revision Number (iProperty)",
        "Item Type": "Item Type",
        "Item_rev_Name": "Item_rev_Name"
    })

    # Specify the CSV file name
    csv_file = "data.csv"

    # Write the data to the CSV file using a pipe delimiter
    df.to_csv(csv_file, sep="|", index=False)

    print(f"Data has been exported to {csv_file}")

except Exception as e:
    print(f"An error occurred: {e}")

#export csv done!
print('xml changer export done!')


print('Start exporting IPS template')
#start exporting IPS template
##(8) ### ips_item_create_template
# Select the desired columns
# List of columns to export
columns_to_export = ["Item Type","Final Item ID", "Revision Number (iProperty)", "Item_rev_Name"]

try:
    # Read the XLSX file into a pandas DataFrame
    df = pd.read_excel(xlsx_file)

    # Select only the specified columns
    df = df[columns_to_export]

    # Rename columns to match the desired output (optional)
    df = df.rename(columns={
        "Item Type": "Item Type",
        "Final Item ID": "Final Item ID",
        "Revision Number (iProperty)": "Revision Number (iProperty)",
        "Item_rev_Name": "Item_rev_Name"
    })

    # Specify the CSV file name
    csv_file = "IPS_Item_create_template.csv"

    # Write the data to the CSV file using a pipe delimiter
    df.to_csv(csv_file, sep="~", index=False)

    print(f"Data has been exported to {csv_file}")

except Exception as e:
    print(f"An error occurred: {e}")
print('Exporting IPS template Done!')
#Exporting done

# Define the function to filter rows with empty 'Revision Number (iProperty)' and save to Excel
def save_empty_revision_to_excel(input_df, output_path):
    file_path = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
    # Load your data into a DataFrame
    df = pd.read_excel(file_path)
    # Define the output file path for the new Excel file
    output_path2 = r'D:\PYDATAANALYSIS\analysis\Revision_missing_file.xlsx'
    # Call the function to filter and save the data
    save_empty_revision_to_excel(df, output_path2)

    empty_revision_rows = input_df[input_df['Revision Number (iProperty)'].isnull()]
    file_names = empty_revision_rows['File Name']
    # Create a new DataFrame with just the 'File Name'
    new_df = pd.DataFrame({'File Name': file_names})
    # Save the new DataFrame to the Excel file
    new_df.to_excel(output_path, index=False)
    print("File Names with empty 'Revision Number (iProperty)' copied to a new Excel file.")


print("Thank you for using this tool! Develoved  by ***Ariful Islam***")
print('please check revision missing file list')