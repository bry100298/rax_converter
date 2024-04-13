import os
import shutil
import pandas as pd
from xml.etree import ElementTree as ET
import time

# Variable dictionary mapping company folders to company names
company_names = {
    'MRSG': "METRO RETAIL STORES GROUP, INC"
}

# Function to convert XML to Excel
def xml_to_excel(xml_file, parent_dir):
    # Parse XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Define paths
    inbound_folder = os.path.join(parent_dir, 'Inbound')
    outbound_folder = os.path.join(parent_dir, 'Outbound')
    inbound_outbound_folder = os.path.join(parent_dir, 'Inbound', 'Outbound')
    archive_folder = os.path.join(parent_dir, 'Archive')
    error_folder = os.path.join(parent_dir, 'Error')

    # Check if filename starts with "_var_"
    if not os.path.basename(xml_file).startswith('_var_'):
        company_folder = os.path.basename(os.path.dirname(xml_file))
        error_company_folder = os.path.join(error_folder, company_folder)
        os.makedirs(error_company_folder, exist_ok=True)
        # Move file to Error Folder
        print(f"Moving '{os.path.basename(xml_file)}' to Error Folder - File name doesn't start with '_var_'")
        shutil.move(xml_file, os.path.join(error_company_folder, os.path.basename(xml_file)))
        return

    # Extract data from XML and populate DataFrame
    data = []
    for invoice_paid in root.findall('.//INVOICES-PAID'):
        invoice_detail = invoice_paid.find('INVOICE-DETAIL')
        company_folder = os.path.basename(os.path.dirname(xml_file))
        EDI_Customer = company_names.get(company_folder, '')
        ASNAME = root.find('.//ASNAME').text
        INVOICE_TYPE = invoice_detail.find('INVOICE_TYPE').text
        EDI_TransType = None
        AIINV = invoice_detail.find('AIINV').text
        AIAMT = invoice_detail.find('AIAMT').text
        WITHHOLDINGTAX = invoice_paid.find('.//DEDUCTIONS/WITHHOLDINGTAX').text
        AITONT = invoice_detail.find('AITONT').text
        AICHQ = invoice_detail.find('AICHQ').text
        BATCH_AMOUNT = invoice_detail.find('BATCH_AMOUNT').text
        
        # Find all INVOICE-DISCREPANCY elements
        for invoice_discrepancy in root.findall(".//INVOICE-DISCREPANCY"):
            # Find the corresponding DISCREPANCY-DETAIL element
            discrepancy_detail = invoice_discrepancy.find("DISCREPANCY-DETAIL")
            if discrepancy_detail is not None and discrepancy_detail.find("AIINV").text == AIINV:
                # Found the matching discrepancy detail
                PONUMB = discrepancy_detail.find('PONUMB').text
                break  # Exit the loop once a match is found
        else:
            # No matching discrepancy detail found
            PONUMB = None

        data.append([EDI_Customer, ASNAME, INVOICE_TYPE, EDI_TransType, PONUMB, AIINV, AIAMT, None, WITHHOLDINGTAX, AITONT, AICHQ, None, BATCH_AMOUNT])

    # Create DataFrame
    df = pd.DataFrame(data, columns=['EDI_Customer', 'EDI_Company', 'EDI_DocType', 'EDI_TransType', 'EDI_PORef', 'EDI_InvRef', 'EDI_Gross', 'EDI_Discount', 'EDI_EWT', 'EDI_Net', 'EDI_RARef', 'EDI_RADate', 'EDI_RAAmt'])

    # Create Excel file path
    company_folder = os.path.basename(os.path.dirname(xml_file))
    excel_folder = os.path.join(inbound_outbound_folder, company_folder)
    os.makedirs(excel_folder, exist_ok=True)
    excel_file = os.path.join(excel_folder, os.path.basename(xml_file).replace('.xml', '.xlsx'))

    # Write DataFrame to Excel
    df.to_excel(excel_file, index=False)

    # Create Archive Folder if not exists
    archive_excel_folder = os.path.join(archive_folder, 'excel', company_folder)
    os.makedirs(archive_excel_folder, exist_ok=True)

    # Create Archive Folder if not exists
    archive_xml_folder = os.path.join(archive_folder, 'xml', company_folder)
    os.makedirs(archive_xml_folder, exist_ok=True)

    # Copy Excel file to Archive excel Folder
    shutil.copy(excel_file, os.path.join(archive_excel_folder, os.path.basename(excel_file)))

    # Move XML file to Archive xml Folder
    shutil.move(xml_file, os.path.join(archive_xml_folder, os.path.basename(xml_file)))

    # Move Excel file to Outbound Folder
    shutil.move(excel_file, os.path.join(outbound_folder, company_folder))

# Main function
def main():
    # Get parent directory of the script
    # script_dir = os.path.dirname(__file__)
    # parent_dir = os.path.join(script_dir, 'SM_Group')
    parent_dir = 'Metro_Gaisano'

    # Iterate over XML files in Inbound Folder
    inbound_dir = os.path.join(parent_dir, 'Inbound')
    for root_folder, dirs, files in os.walk(inbound_dir):
        for file in files:
            if file.endswith('.xml'):
                xml_file = os.path.join(root_folder, file)
                xml_to_excel(xml_file, parent_dir)

    time.sleep(5)  # Wait for 5 seconds before exiting

if __name__ == "__main__":
    main()
