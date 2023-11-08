import pandas as pd

print("Listed for item revision missing files")
# Define the function to filter rows with empty 'Revision Number (iProperty)' and save to Excel
def save_empty_revision_to_excel(input_df, output_path):
    file_path = r'D:\PYDATAANALYSIS\Big_Data_Analysis\Output\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
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
print("File Names with empty 'Revision' copied to a new Excel file.")

