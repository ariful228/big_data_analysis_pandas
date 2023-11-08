       
import pandas as pd
# Load the Excel file into a DataFrame
file_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
df = pd.read_excel(file_path)

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
output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx' 
# Save the updated DataFrame back to the Excel file
df.to_excel(output_path, index=False)
print("Item type changing Done!")
print("Start Metadata processing")
