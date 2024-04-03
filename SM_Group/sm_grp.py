import os
import shutil
import pandas as pd
from xml.etree import ElementTree as ET
import time

# Function to convert XML to Excel
def xml_to_excel(xml_file):
    # Parse XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Check if filename starts with "RA" (case sensitive)
    if not os.path.basename(xml_file).startswith('RA'):
        # Move file to Error Folder
        shutil.move(xml_file, 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Error')
        return

    # Check if XML file meets conditions
    company_name = root.find('.//companyName')
    if company_name is None or company_name.text.strip() != 'SANFORD MARKETING CORPORATION':
        # Move file to Error Folder
        shutil.move(xml_file, 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Error')
        return

    # Extract data from XML
    data = []
    for article in root.findall('.//article'):
        trans_code = article.find('transCode').text
        po_number = article.find('poNumber').text
        doc_ref = article.find('docRef').text
        gross_amount = article.find('grossAmount').text
        net_payable = article.find('netAmount').text
        check_number = root.find('.//checkNumber').text
        check_date = root.find('.//checkDate').text

        # Append to data list
        data.append([trans_code, po_number, doc_ref, gross_amount, net_payable, check_number, check_date])

    # Create DataFrame
    df = pd.DataFrame(data, columns=['transCode', 'poNumber', 'docRef', 'grossAmount', 'EDI_Net', 'checkNumber', 'checkDate'])

    # Create Excel file path
    excel_file = os.path.join('C:/Users/User/Documents/Project/rax_converter/SM_Group/Outbound', os.path.basename(xml_file).replace('.xml', '.xlsx'))

    # Write DataFrame to Excel
    df.to_excel(excel_file, index=False)

    # Move XML file to Archive Folder
    shutil.move(xml_file, 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Archive')

# Main function
def main():
    # Path to Inbound Folder
    inbound_folder = 'C:/Users/User/Documents/Project/rax_converter/SM_Group/Inbound'

    # Iterate over XML files in Inbound Folder
    for file in os.listdir(inbound_folder):
        if file.endswith('.xml'):
            xml_file = os.path.join(inbound_folder, file)
            xml_to_excel(xml_file)
        time.sleep(15)  # Wait for 15 seconds before processing next file

if __name__ == "__main__":
    main()
