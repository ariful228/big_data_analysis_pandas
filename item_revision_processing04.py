import pandas as pd

# Save the updated DataFrame to a new Excel file
output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
# Load the DataFrame from your Excel file
input_path = r'analysis/Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
df = pd.read_excel(input_path)

# Define your `update_revision_number` function here
def update_revision_number(index, row):
    tc_revision = row['df2_TC_REV']
    file_revision = row['df2_Revision']
    bl_revision = row['Revision Number (iProperty)']
    file_name = str(row['File Name'])  # Convert to string
    modified_file_name = file_name[:-4]

    if pd.notna(tc_revision):
        # If tc_revision is not NaN, update 'Revision Number (iProperty)' with tc_revision
        df.at[index, 'Revision Number (iProperty)'] = tc_revision
    elif pd.notna(file_revision):
        # If tc_revision is NaN but 'df2_Revision' is not NaN, update 'Revision Number (iProperty)' with 'df2_Revision'
        df.at[index, 'Revision Number (iProperty)'] = file_revision

                # You can apply additional conditions to 'Revision Number (iProperty)' here if needed
    else:
         if row['File Extension'] == 'dwg' or row['File Extension'] == 'ipt' or row['File Extension'] == 'iam' and pd.notna(row['Part Number (iProperty)']):
                bl_revision = row['Revision Number (iProperty)']

                if pd.notna(bl_revision):
                    matching_rows = df[df['Part Number'] == row['Part Number']]
                    df.loc[matching_rows.index, 'Revision Number (iProperty)'] = bl_revision

                if pd.notna(bl_revision):
                    matching_rows = df[df['File Name'].str[:-4] == modified_file_name]
                    df.loc[matching_rows.index, 'Revision Number (iProperty)'] = bl_revision                    


df['Revision Number (iProperty)'] = df.apply(update_revision_number, axis=1)
df.to_excel(output_path, index=False)
# Apply the update_revision_number function to update the 'Revision Number (iProperty)' column


