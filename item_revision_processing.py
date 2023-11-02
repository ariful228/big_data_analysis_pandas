print("Start item revision processing")
#>>>>>>>Item id and file name processing DOEN<<<<<<
#START revision processing********
import pandas as pd
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
file_path = r'D:\PYDATAANALYSIS\analysis\Dataanalysis_merged_with_tc_rev_and_sap_material.xlsx'
df = pd.read_excel(file_path)
# Apply the `update_revision_number` function to update 'Revision Number (iProperty)'
df['Revision Number (iProperty)'] = df.apply(update_revision_number, axis=1)
# Save the updated DataFrame back to the Excel file
output_path = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
df.to_excel(output_path, index=False)
print("Item revision processing Done!!")
        