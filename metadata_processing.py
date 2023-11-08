import pandas as pd
# Load the Excel file into a DataFrame
input_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
output_path = r'D:\PYDATAANALYSIS\analysis\Item_id_&_revision_material_ready_now_duplicate_renaming.xlsx'
df = pd.read_excel(input_path)
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
df.to_excel(output_path, index=False)
print("Item_rev_Name column added and DataFrame saved to Excel.")
