import pandas as pd
def export_IPS_template(xlsx_file, columns_to_export, csv_file):
    try:
        df = pd.read_excel(xlsx_file)
        df = df[columns_to_export]
        df = df.rename(columns={
            "Item Type": "Item Type",
            "Final Item ID": "Final Item ID",
            "Revision Number (iProperty)": "Revision Number (iProperty)",
            "Item_rev_Name": "Item_rev_Name"
        })
        df.to_csv(csv_file, sep="~", index=False)
        print(f"Data has been exported to {csv_file}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    xlsx_file = r'D:\PYDATAANALYSIS\analysis\Bulkloader_analysis_ready_for_xml_changer_&_IPS.xlsx'
    columns_to_export = ["Item Type", "Final Item ID", "Revision Number (iProperty)", "Item_rev_Name"]
    csv_file = "IPS_Item_create_template.csv"

    export_IPS_template(xlsx_file, columns_to_export, csv_file)
