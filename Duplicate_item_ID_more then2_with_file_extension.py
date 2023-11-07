import pandas as pd
from collections import Counter

def process_dataframe(input_file, output_file):
    try:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(input_file)

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

# Define input and output file paths
input_file_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
output_file_path = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'

# Process the DataFrame
process_dataframe(input_file_path, output_file_path)
