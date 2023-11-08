# Import the necessary modules

def status(line):
    print(line)
import sys
print(sys.path)
def main():
    from bl_sap_tc_data_marged import extractDefault
    status("1 >>>> done")
    from file_modify_rename import process_excel_file
    status(2)
    from item_id_combine_with_sap_renamel_partnumber import extract_default
    status(3)
    from items_revisions_with_man_tc_global import process_item_revisions
    status(4)
    from item_type_change import update_item_type
    status(5)
    from metadata_processing import collect_values
    status(6)
    from duplicate_item_ID_more_then2_with_file_extension import process_dataframe
    status(7)
    from revision_mising_file_list import save_empty_revision_to_excel
    status(8)
    from export_ips_template import export_IPS_template
    status(9)
    from export_csv_for_xml_changer import process_excel_to_csv
    status('10 & DONE')

    # Call the functions sequentially
    #extractDefault()
    process_excel_file()
    extract_default()
    process_item_revisions()
    update_item_type()
    collect_values()
    process_dataframe()
    save_empty_revision_to_excel()
    process_excel_to_csv()
    export_IPS_template()
    return 
main()
# Perform further processing with the results if needed
