import pandas as pd

def process_item_revisions(input_file_path, output_file_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_file_path)

    # Iterate through the rows and update 'Revision Number (iProperty)' based on tc_revision or revision
    for index, row in df.iterrows():
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

    # Save the updated DataFrame back to the Excel file
    df.to_excel(output_file_path, index=False)
    print("Item revision processing Done!!")

# Call the function to process item revisions
input_file_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
output_file_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
process_item_revisions(input_file_path, output_file_path)

# You can add more code here for further processing if needed
print("Item type change processing")
