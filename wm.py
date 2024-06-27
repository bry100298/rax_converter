import os
import shutil
import pandas as pd
from xml.etree import ElementTree as ET
import time

# Variable array or object mapping company folders to company names
company_names = {
    'WMSI': "WALTERMART SUPERMARKET INC. "
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

    # Check if filename starts with "RA" (case sensitive)
    if not os.path.basename(xml_file).startswith('RA'):
        company_folder = os.path.basename(os.path.dirname(xml_file))
        error_company_folder = os.path.join(error_folder, 'Error', company_folder)
        if not os.path.exists(error_company_folder):
            os.makedirs(error_company_folder)
        # Move file to Error Folder
        print(f"Moving '{os.path.basename(xml_file)}' to Error Folder - File name doesn't start with 'RA'")
        shutil.move(xml_file, os.path.join(error_folder, 'Error', company_folder))
        return

    # Extract data from XML
    data = []
    vendor_name = root.find('.//vendor_name').text
    ra_number = root.find('.//remittance_advice_number').text
    payment_date = root.find('.//payment_date').text
    total_amount_elem = root.find('.//total_amount')
    total_amount = total_amount_elem.text.replace(',', '') if total_amount_elem is not None else None # Remove commas

    for article in root.findall('.//article'):
        company_folder = os.path.basename(os.path.dirname(xml_file))
        # trans_code = article.find('Document_Type').text
        trans_code = article.find('Document_Type').text.strip() if article.find('Document_Type') is not None else None
        # po_number = article.find('Remarks').text
        po_number = article.find('Remarks').text.split('#', 1)[-1].strip() if article.find('Remarks') is not None else None
        # remarks_element = article.find('Remarks')
        # po_number = remarks_element.text.split('#', 1)[-1].strip() if remarks_element is not None else None
        doc_number = article.find('Document_Number').text
        gross_amount = article.find('Invoice_Amount').text.replace(',', '')  # Remove commas
        wTax = article.find('.//WTAX').text.replace(',', '') # Remove commas
        net_payable = article.find('Net_Payable').text.replace(',', '') # Remove commas

        # Append to data list
        data.append([company_names[company_folder], vendor_name, trans_code, None, po_number, doc_number, gross_amount, None, wTax, net_payable, ra_number, payment_date, total_amount])

    # Create DataFrame
    df = pd.DataFrame(data, columns=['EDI_Customer', 'EDI_Company', 'EDI_DocType', 'EDI_TransType', 'EDI_PORef', 'EDI_InvRef', 'EDI_Gross', 'EDI_Discount', 'EDI_EWT', 'EDI_Net', 'EDI_RARef', 'EDI_RADate', 'EDI_RAAmt'])

    # Convert columns to numeric
    numeric_columns = ['EDI_PORef', 'EDI_InvRef', 'EDI_Gross', 'EDI_Discount', 'EDI_EWT', 'EDI_Net', 'EDI_RARef', 'EDI_RAAmt']
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce').fillna(0)  # Coerce errors to NaN and fill NaNs with 0
    
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

    # Create Archive Folder if not exists
    archive_xml_folder = os.path.join(archive_folder, 'xml', company_folder)
    if not os.path.exists(archive_xml_folder):
        os.makedirs(archive_xml_folder)

    # Copy Excel file to Archive excel Folder
    shutil.copy(excel_file, os.path.join(archive_excel_folder, os.path.basename(excel_file)))

    # Move XML file to Archive xml Folder
    shutil.move(xml_file, os.path.join(archive_folder, 'xml', company_folder))

    # Move Excel file to Outbound Folder
    shutil.move(excel_file, os.path.join(outbound_folder, company_folder))

# Main function
def main():
    # Get parent directory of the script
    # script_dir = os.path.dirname(__file__)
    # parent_dir = os.path.join(script_dir, 'SM_Group')
    parent_dir = 'Waltermart'

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
