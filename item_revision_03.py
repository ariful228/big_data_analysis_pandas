import pandas as pd

# Define your `update_revision_number` function
def update_revision_number(row, df):
    tc_revision = row['df2_TC_REV']
    file_revision = row['df2_Revision']
    bl_revision = row['Revision Number (iProperty)']
    file_name = str(row['File Name'])  # Convert to string
    modified_file_name = file_name[:-4]

    if pd.notna(file_revision):
        return file_revision

    elif row['File Extension'] in ['idw', 'dwg', 'ipt', 'iam'] and pd.notna(row['df2_Revision']):
        if pd.notna(file_revision):
            matching_rows = df[df['Part Number'] == row['Part Number']]
            df.loc[matching_rows.index, 'df2_Revision'] = file_revision
            print('DWG_part_from_file:', file_revision)
            return file_revision
        
        elif pd.notna(bl_revision):
            matching_rows = df[df['File Name'].str[:-4] == modified_file_name]
            df.loc[matching_rows.index, 'df2_Revision'] = file_revision
            print('DWG_modify_from_file:', file_revision)
            return file_revision
        

    elif row['File Extension'] in ['idw', 'dwg', 'ipt', 'iam'] and pd.notna(row['Part Number (iProperty)']):
        if not pd.notna(bl_revision):
            matching_rows = df[df['Part Number'] == row['Part Number']]
            df.loc[matching_rows.index, 'Revision Number (iProperty)'] = bl_revision
            print('DWG_part:', bl_revision)
            return bl_revision
        
        elif pd.notna(bl_revision):
            matching_rows = df[df['File Name'].str[:-4] == modified_file_name]
            df.loc[matching_rows.index, 'Revision Number (iProperty)'] = bl_revision
            print('DWG_Modify:', bl_revision)
            return bl_revision
    else:
            if pd.notna(tc_revision):
                return tc_revision

    return bl_revision

# Load the DataFrame from your Excel file
input_path = r'analysis/Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
df = pd.read_excel(input_path)

# Apply the update_revision_number function to update the 'Revision Number (iProperty)' column
df['Revision Number (iProperty)'] = df.apply(lambda row: update_revision_number(row, df), axis=1)

# Save the updated DataFrame to a new Excel file
output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
df.to_excel(output_path, index=False)
