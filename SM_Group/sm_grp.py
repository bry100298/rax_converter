import os
import shutil
import pandas as pd
from xml.etree import ElementTree as ET
import time

# Function to convert XML to Excel
def xml_to_excel(xml_file, inbound_folder, outbound_folder, inbound_outbound_folder, archive_folder, error_folder):
    # Parse XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Check if filename starts with "RA" (case sensitive)
    if not os.path.basename(xml_file).startswith('RA'):
        # Move file to Error Folder
        print(f"Moving '{os.path.basename(xml_file)}' to Error Folder - File name doesn't start with 'RA'")
        shutil.move(xml_file, os.path.join(error_folder, 'Error'))
        return

    # Extract data from XML
    data = []
    for article in root.findall('.//article'):
        trans_code = article.find('transCode').text
        EDI_TransType = None
        po_number = article.find('poNumber').text
        doc_ref = article.find('docRef').text
        EDI_DocRef = None
        EDI_DocDesc = None
        gross_amount = article.find('grossAmount').text
        EDI_VAT = None
        EDI_EWT = None
        netAmount = article.find('netAmount').text
        check_number = root.find('.//checkNumber').text
        check_date = root.find('.//checkDate').text
        #netPayable = root.find('//netPayable').text

        # Extract netPayable directly from the root
        netPayable_elem = root.find('.//netPayable')
        netPayable = netPayable_elem.text if netPayable_elem is not None else None

        # Append to data list
        data.append([trans_code, EDI_TransType, po_number, doc_ref, EDI_DocRef, EDI_DocDesc, gross_amount, EDI_VAT, EDI_EWT, netAmount, check_number, check_date, netPayable])

    # Create DataFrame
    df = pd.DataFrame(data, columns=['EDI_DocType', 'EDI_TransType', 'EDI_PORef', 'EDI_InvRef', 'EDI_DocRef', 'EDI_DocDesc', 'EDI_Gross', 'EDI_VAT', 'EDI_EWT', 'EDI_Net', 'EDI_RARef', 'EDI_RADate', 'EDI_RAAmt'])

    # Create Excel file path
    company_folder = os.path.basename(os.path.dirname(xml_file))
    excel_folder = os.path.join(inbound_outbound_folder, company_folder)
    if not os.path.exists(excel_folder):
        os.makedirs(excel_folder)
    excel_file = os.path.join(excel_folder, os.path.basename(xml_file).replace('.xml', '.xlsx'))

    # Write DataFrame to Excel
    df.to_excel(excel_file, index=False)

    # Create Archive Folder if not exists
    archive_excel_folder = os.path.join(archive_folder, 'excel', company_folder)
    if not os.path.exists(archive_excel_folder):
        os.makedirs(archive_excel_folder)

    # Copy Excel file to Archive excel Folder
    shutil.copy(excel_file, os.path.join(archive_excel_folder, os.path.basename(excel_file)))

    # Move XML file to Archive xml Folder
    shutil.move(xml_file, os.path.join(archive_folder, 'xml', company_folder))

    # Move Excel file to Outbound Folder
    shutil.move(excel_file, os.path.join(outbound_folder, company_folder))

# Main function
def main():
    # Path settings
    inbound_folder = 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Inbound'
    outbound_folder = 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Outbound'
    inbound_outbound_folder = 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Inbound/Outbound'
    archive_folder = 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Archive'
    error_folder = 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Error'

    # Iterate over XML files in Inbound Folder
    for root_folder, dirs, files in os.walk(inbound_folder):
        for file in files:
            if file.endswith('.xml'):
                xml_file = os.path.join(root_folder, file)
                xml_to_excel(xml_file, inbound_folder, outbound_folder, inbound_outbound_folder, archive_folder, error_folder)

    time.sleep(5)  # Wait for 5 seconds before exiting

if __name__ == "__main__":
    main()
