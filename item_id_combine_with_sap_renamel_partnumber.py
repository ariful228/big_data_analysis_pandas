import pandas as pd

def extract_default(input_path, output_excel_path):
    print("Start working with combining sap_rename_partNumber!")
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(input_path)

        # Create a boolean Series to identify library rows
        library_rows = df['Full File Name'].str.contains(r'\\Library\\|\\Libraries\\|\\Content Center Files\\', case=False, na=False)

        # Define a function to create the 'Item ID' column based on specified conditions
        def create_item_id(row):
            if pd.notna(row['df2_SAP_Material']):
                return row['df2_SAP_Material']
            
            if pd.notna(row['df2_Rename_Item_ID']):
                return row['df2_Rename_Item_ID']
            
            if row['Part Number'] == row['df2_SAP_Document']:
                if pd.notna(row['df2_SAP_Material']):
                   return row['df2_SAP_Material']
                else:
                    return row['Part Number']
            
            if library_rows[row.name]:
                if row['Part Number'] != row['modify_file_item_ID'] and len(str(row['Part Number'])) < 18:
                    return row['Part Number']
                if len(str(row['Part Number'])) > 18:
                    return row['modify_file_item_ID']
                if pd.Series([row['Part Number']]).equals(pd.Series([row['modify_file_item_ID']])):
                    return row['Part Number']
            else:
                if row['Part Number'] != row['modify_file_item_ID']:
                    return row['modify_file_item_ID']
                if pd.Series([row['Part Number']]).equals(pd.Series([row['modify_file_item_ID']])):
                    return row['Part Number']
            return None  # Default case

        # Apply the function to create the 'Item ID' column
        df['Item ID'] = df.apply(create_item_id, axis=1)

        # Save the updated DataFrame to a new Excel file
        df.to_excel(output_excel_path, index=False)

        print("File modify and renaming Done!")

    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

    print("Revision and SAP Material merged Done!")

# Usage example:
input_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
extract_default(input_path, output_path)