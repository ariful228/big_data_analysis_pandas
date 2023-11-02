import pandas as pd

# Define the paths for input and output Excel files
input_path = r'analysis/Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'

# Read the input Excel file into a DataFrame
df = pd.read_excel(input_path)

# Define the `update_revision_number` function
def update_revision_number(row):
    tc_revision = row['df2_TC_REV']
    file_revision = row['df2_Revision']
    bl_revision = row['Revision Number (iProperty)']
    file_name = str(row['File Name'])  # Convert to string
    modified_file_name = file_name[:-4]

    # Check if 'Revision Number (iProperty)' is not null
    if pd.notna(bl_revision):
        return bl_revision

    # Handle the 'dwg' files
    if row['File Extension'] == 'dwg':
        # Find matching rows by 'Part Number' or 'File Name' condition
        matching_rows = df[(df['Part Number'] == row['Part Number']) | (df['File Name'].str[:-4] == modified_file_name)]
        if pd.notna(bl_revision):
            # Update 'File_Item_Revision' based on 'bl_revision'
            df.loc[matching_rows.index, 'File_Item_Revision'] = bl_revision
            print('DWG Part:', bl_revision)
        return bl_revision

    # Handle other file extensions
    if pd.notna(tc_revision):
        return tc_revision
    if pd.notna(file_revision):
        return file_revision

    return bl_revision

# Apply the `update_revision_number` function to update the 'File_Item_Revision' column
df['File_Item_Revision'] = df.apply(update_revision_number, axis=1)

# Save the updated DataFrame to a new Excel file
df.to_excel(output_path, index=False)
