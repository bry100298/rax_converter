import os
import shutil
import pandas as pd
from bs4 import BeautifulSoup
import time
import datetime

# Variable dictionary mapping company folders to company names
company_names = {
    'PPCI': "PUREGOLD PRICE CLUB INC.",
    'AGRI': "AYAGOLD RETAILERS, INC."
}

# Function to convert HTML to Excel
def html_to_excel(html_file, parent_dir):
    # Read HTML file
    with open(html_file, 'r') as f:
        html_content = f.read()

    # Parse HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Extract data from HTML and populate DataFrame
    data = []

    # Extract EDI_Company
    payee_tag = soup.find('th', string='Payee:')
    if payee_tag:
        # EDI_Company = payee_tag.find_next('td').text.strip()
        payee_text = payee_tag.find_next('td').text.strip()
        # Find the indices of the asterisks
        asterisk_indices = [i for i, char in enumerate(payee_text) if char == '*']
        print("Asterisk indices:", asterisk_indices)
        if len(asterisk_indices) >= 2:
            # Get the substring between the two asterisks
            start_index = asterisk_indices[-2] + 1  # Index after the second last asterisk
            end_index = asterisk_indices[-1]  # Index of the last asterisk
            company_name = payee_text[start_index:end_index].strip()
            EDI_Company = company_name
            print("Extracted company name:", EDI_Company)
        else:
            EDI_Company = None
    else:
        EDI_Company = None

    # Extract EDI_Gross
    # gross_tag = soup.find('th', string='Amount:')
    # if gross_tag:
    #     EDI_Gross = gross_tag.find_next('td').text.strip()
    # else:
    #     EDI_Gross = None

    # Extract EDI_RARef
    ra_ref_tag = soup.find('th', string='Transaction No.:')
    if ra_ref_tag:
        EDI_RARef = ra_ref_tag.find_next('td').text.strip()
    else:
        EDI_RARef = None

    # Extract EDI_RADate
    # radate_tag = soup.find('th', string='Check Date:')
    # if radate_tag:
    #     EDI_RADate = radate_tag.find_next('td').text.strip()
    # else:
    #     EDI_RADate = None

    # Extract EDI_RADate
    radate_tag = soup.find('th', string='Check Date:')
    if radate_tag:
        EDI_RADate_text = radate_tag.find_next('td').text.strip()
        # Convert the extracted date string to a datetime object
        try:
            date_obj = datetime.datetime.strptime(EDI_RADate_text, "%Y-%m-%d")
            # Format the datetime object as mm-dd-yyyy
            EDI_RADate = date_obj.strftime("%m-%d-%Y")
        except ValueError:
            EDI_RADate = None
    else:
        EDI_RADate = None

    # Extract EDI_RAAmt
    ra_amt_tag = soup.find('th', string='Amount:')
    if ra_amt_tag:
        EDI_RAAmt = ra_amt_tag.find_next('td').text.strip()
        # Remove asterisks and 'PHP' from EDI_RAAmt
        EDI_RAAmt = EDI_RAAmt.replace('*', '').replace('PHP', '').strip()
    else:
        EDI_RAAmt = None


    for invoice_row in soup.find_all('tr'):
        first_cell_text = invoice_row.find('font').text.strip() if invoice_row.find('font') else None
        if first_cell_text != "Loc":  # Skip rows where the first cell contains "Loc"
            cells = invoice_row.find_all('font')
            if len(cells) == 9:  # Assuming each row has 9 cells
                company_folder = os.path.basename(os.path.dirname(html_file))
                EDI_Customer = company_names.get(company_folder, '')
                # EDI_Company = cells[0].text.strip()  # Payee
                EDI_DocType = cells[7].text.strip()  # Description
                EDI_TransType = None
                EDI_PORef = None
                EDI_InvRef = cells[1].text.strip()  # Invoice Number
                EDI_Gross = cells[3].text.strip()  # Original/Bal Amount
                EDI_Discount = None
                EDI_EWT = cells[4].text.strip()  # WHT Amount
                EDI_Net = cells[5].text.strip()  # Paid Amount (NET)
                # EDI_RARef = cells[6].text.strip()  # Transaction No.
                # EDI_RADate = cells[3].text.strip()  # PostDate
                # EDI_RAAmt = cells[2].text.strip()  # Amount (assuming it's the same as Original/Bal Amount)
                
                data.append([EDI_Customer, EDI_Company, EDI_DocType, EDI_TransType, EDI_PORef, EDI_InvRef, EDI_Gross, EDI_Discount, EDI_EWT, EDI_Net, EDI_RARef, EDI_RADate, EDI_RAAmt])

    # Create DataFrame
    df = pd.DataFrame(data, columns=['EDI_Customer', 'EDI_Company', 'EDI_DocType', 'EDI_TransType', 'EDI_PORef', 'EDI_InvRef', 'EDI_Gross', 'EDI_Discount', 'EDI_EWT', 'EDI_Net', 'EDI_RARef', 'EDI_RADate', 'EDI_RAAmt'])

    # Create Excel file path
    company_folder = os.path.basename(os.path.dirname(html_file))
    excel_folder = os.path.join(parent_dir, 'Inbound', 'Outbound', company_folder)
    os.makedirs(excel_folder, exist_ok=True)
    excel_file = os.path.join(excel_folder, os.path.basename(html_file).replace('.html', '.xlsx'))

    # Write DataFrame to Excel
    df.to_excel(excel_file, index=False)

    # Create Archive Folder if not exists
    archive_excel_folder = os.path.join(parent_dir, 'Archive', 'excel', company_folder)
    os.makedirs(archive_excel_folder, exist_ok=True)

    # Create Archive Folder if not exists
    archive_html_folder = os.path.join(parent_dir, 'Archive', 'html', company_folder)
    os.makedirs(archive_html_folder, exist_ok=True)

    # Copy Excel file to Archive excel Folder
    shutil.copy(excel_file, os.path.join(archive_excel_folder, os.path.basename(excel_file)))

    # Move HTML file to Archive html Folder
    shutil.move(html_file, os.path.join(archive_html_folder, os.path.basename(html_file)))

    # Move Excel file to Outbound Folder
    shutil.move(excel_file, os.path.join(parent_dir, 'Outbound', company_folder))

# Main function
def main():
    parent_dir = 'Puregold'

    # Iterate over HTML files in Inbound Folder
    inbound_dir = os.path.join(parent_dir, 'Inbound')
    for root_folder, dirs, files in os.walk(inbound_dir):
        for file in files:
            if file.endswith('.html'):
                html_file = os.path.join(root_folder, file)
                html_to_excel(html_file, parent_dir)

    time.sleep(5)  # Wait for 5 seconds before exiting

if __name__ == "__main__":
    main()
