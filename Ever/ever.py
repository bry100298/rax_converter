import os
import shutil
import pandas as pd
from bs4 import BeautifulSoup
import time
import pdfplumber  # Import pdfplumber for PDF to HTML conversion

# Variable dictionary mapping company folders to company names
company_names = {
    'EVRP': "EVER PLUS SUPERSTORE, INC"
}

# Function to convert PDF to HTML
def pdf_to_html(pdf_file, inbound_html_dir):
    # Create HTML file path
    html_file = os.path.join(inbound_html_dir, os.path.basename(pdf_file).replace('.pdf', '.html'))

    # Convert PDF to HTML
    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        # Write extracted text to HTML file
        with open(html_file, 'w') as f:
            f.write(text)

    return html_file

# Function to convert HTML to Excel
def html_to_excel(html_file, parent_dir):
    # Read HTML file
    with open(html_file, 'r') as f:
        html_content = f.read()

    # Parse HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Extract data from HTML and populate DataFrame
    data = []
    for invoice_row in soup.find_all('tr'):
        first_cell_text = invoice_row.find('font').text.strip() if invoice_row.find('font') else None
        if first_cell_text != "Loc":  # Skip rows where the first cell contains "Loc"
            cells = invoice_row.find_all('font')
            if len(cells) == 9:  # Assuming each row has 9 cells
                company_folder = os.path.basename(os.path.dirname(html_file))
                EDI_Customer = company_names.get(company_folder, '')
                EDI_Company = None
                EDI_DocType = "Vendor Name"
                EDI_TransType = None
                EDI_PORef = None
                EDI_InvRef = cells[1].text.strip()  # Vendor Code
                EDI_Gross = cells[6].text.strip()  # APID
                EDI_Discount = None
                EDI_EWT = None
                EDI_Net = None
                EDI_RARef = cells[5].text.strip()  # CV No
                EDI_RADate = cells[8].text.strip()  # Payment Date
                EDI_RAAmt = cells[7].text.strip()  # Total:
                
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
    archive_excel_folder = os.path.join(parent_dir, 'Archive', 'excel')
    os.makedirs(archive_excel_folder, exist_ok=True)

    # Create Archive Folder if not exists
    archive_html_folder = os.path.join(parent_dir, 'Archive', company_folder)
    os.makedirs(archive_html_folder, exist_ok=True)

    # Copy Excel file to Archive excel Folder
    shutil.copy(excel_file, os.path.join(archive_excel_folder, os.path.basename(excel_file)))

    # Move HTML file to Archive html Folder
    shutil.move(html_file, os.path.join(archive_html_folder, os.path.basename(html_file)))

    # Move Excel file to Outbound Folder
    shutil.move(excel_file, os.path.join(parent_dir, 'Outbound', company_folder))

# Main function
def main():
    parent_dir = 'Ever'
    inbound_html_dir = os.path.join(parent_dir, 'Inbound', 'HTML')  # HTML files will be stored here

    # Iterate over PDF files in Ever directory
    for root_folder, dirs, files in os.walk(parent_dir):
        for file in files:
            if file.endswith('.pdf'):
                pdf_file = os.path.join(root_folder, file)
                # Convert PDF to HTML
                html_file = pdf_to_html(pdf_file, inbound_html_dir)
                # Convert HTML to Excel
                html_to_excel(html_file, parent_dir)

    time.sleep(5)  # Wait for 5 seconds before exiting

if __name__ == "__main__":
    main()
