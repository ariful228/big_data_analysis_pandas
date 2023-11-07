import pandas as pd

def read_excel_file(xlsx_file, columns_to_export):
    try:
        df = pd.read_excel(xlsx_file)
        df = df[columns_to_export]
        df = df.rename(columns={
            "Full File Name": "Full File Name",
            "Final Item ID": "Final Item ID",
            "Revision Number (iProperty)": "Revision Number (iProperty)",
            "Item Type": "Item Type",
            "Item_rev_Name": "Item_rev_Name"
        })
        return df
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return None

def export_to_csv(df, csv_file):
    try:
        df.to_csv(csv_file, sep="|", index=False)
        print(f"Data has been exported to {csv_file}")
    except Exception as e:
        print(f"An error occurred while exporting to CSV: {e}")

def main():
    xlsx_file = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
    columns_to_export = ["Full File Name", "Final Item ID", "Revision Number (iProperty)", "Item Type", "Item_rev_Name"]
    csv_file = "data.csv"

    df = read_excel_file(xlsx_file, columns_to_export)
    if df is not None:
        export_to_csv(df, csv_file)

if __name__ == "__main__":
    main()
