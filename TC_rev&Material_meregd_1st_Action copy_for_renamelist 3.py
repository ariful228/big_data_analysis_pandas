
import pandas as pd

try:
    # Read the data from Excel into DataFrames from Sheet1, Sheet2, and Sheet3
    df1 = pd.read_excel('analysis/EDM_Vault_Migration_Mining_Screen_Pilot_request.xlsx', sheet_name='Sheet1')
    df2 = pd.read_excel('analysis/EDM_Vault_Migration_Mining_Screen_Pilot_request.xlsx', sheet_name='Sheet2')
    

    def extractDefault(df1, df2):
        # Create an empty list to store the results
        results = []

        for index, row in df1.iterrows():
            # Extract the 'Part Number' and 'SAP_Material' from the current row, and convert them to strings
            part_number = str(row['Part Number']).strip()
            Full_file_name = str(row['Full File Name']).strip()
            File_Name = str(row['File Name']).strip()

            # Find matching rows in df2 based on 'SAP_Document' and 'SAP_Material'
            matching_rows_df2 = df2[(df2['SAP_Document'].str.strip() == part_number) |
                                   (df2['SAP_Material'].str.strip() == part_number) |
                                   (df2['File Name'].str.strip() == File_Name) |
                                   (df2['Full File Name'].str.strip() == Full_file_name) |
                                   (df2['Part Number'].str.strip() == part_number)] 
                                   

            if not matching_rows_df2.empty:
                # Take the first matching row from df2
                matching_row_df2 = matching_rows_df2.iloc[0]
            else:
                matching_row_df2 = pd.Series()  # Create an empty Series if no match is found in df2

            # Rename columns in matching rows to avoid duplicates
            matching_row_df2 = matching_row_df2.add_prefix('df2_')

            # Merge the current row from df1, matching rows from df2 
            merged_row = pd.concat([row, matching_row_df2])

            # Append the merged row to the results list
            results.append(merged_row)

        # Create a new DataFrame from the results list
        merged_df = pd.concat(results, axis=1).T

        # Save the merged_df to an Excel file
        merged_df.to_excel('output.xlsx', index=False)  # Change 'output.xlsx' to your desired file name

    # Call the function to extract and save the results
    extractDefault(df1, df2)

except FileNotFoundError as e:
    print(f"Error: File not found - {e}")
except Exception as e:
    print(f"An error occurred: {e}")
