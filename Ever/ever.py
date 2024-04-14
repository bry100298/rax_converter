import os
import shutil
import pandas as pd
from bs4 import BeautifulSoup
import time
import pdfplumber
import tabula

# Variable dictionary mapping company folders to company names
company_names = {
    'EVRP': "EVER PLUS SUPERSTORE, INC"
}

# Function to convert PDF to HTML
def pdf_to_html(pdf_file, output_folder):
    # Generate HTML file path
    html_file = os.path.join(output_folder, os.path.basename(pdf_file).replace('.pdf', '.html'))

    # Extract tables from PDF
    tables = tabula.read_pdf(pdf_file, pages='all')

    # Concatenate tables into a single DataFrame
    concatenated_df = pd.concat(tables)

    # Convert DataFrame to HTML
    html_content = concatenated_df.to_html(index=False)

    # Write HTML content to file
    with open(html_file, 'w') as f:
        f.write(html_content)

    return html_file

# Function to convert HTML to Excel
def html_to_excel(html_file, parent_dir, pdf_file):
    # Read HTML file
    with open(html_file, 'r') as f:
        html_content = f.read()

    # Parse HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Extract data from HTML and populate DataFrame
    data = []
    for invoice_row in soup.find_all('tr'):
        cells = invoice_row.find_all('td')
        if len(cells) == 7:  # Assuming each row has 7 cells
            company_folder = os.path.basename(os.path.dirname(pdf_file))
            EDI_Customer = company_names[company_folder]
            EDI_Company = None
            EDI_DocType = cells[1].text.strip()  # Vendor Name
            EDI_TransType = None
            EDI_PORef = None
            EDI_InvRef = cells[2].text.strip()  # Vendor Code
            EDI_Gross = cells[3].text.strip()  # APID
            EDI_Discount = None
            EDI_EWT = None
            EDI_Net = None
            EDI_RARef = cells[4].text.strip()  # CV No
            EDI_RADate = cells[5].text.strip()  # Payment Date
            EDI_RAAmt = cells[6].text.strip()  # Total

            data.append([EDI_Customer, EDI_Company, EDI_DocType, EDI_TransType, EDI_PORef, EDI_InvRef, EDI_Gross, EDI_Discount, EDI_EWT, EDI_Net, EDI_RARef, EDI_RADate, EDI_RAAmt])

    # Create DataFrame
    df = pd.DataFrame(data, columns=['EDI_Customer', 'EDI_Company', 'EDI_DocType', 'EDI_TransType', 'EDI_PORef', 'EDI_InvRef', 'EDI_Gross', 'EDI_Discount', 'EDI_EWT', 'EDI_Net', 'EDI_RARef', 'EDI_RADate', 'EDI_RAAmt'])

    # Create Excel file path
    company_folder = os.path.basename(os.path.dirname(pdf_file))
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

    # Create Archive Folder if not exists
    archive_pdf_folder = os.path.join(parent_dir, 'Archive', 'pdf', company_folder)
    os.makedirs(archive_pdf_folder, exist_ok=True)
    archive_folder = os.path.join(parent_dir, 'Archive')

    # Move HTML file to Archive html Folder
    shutil.move(html_file, os.path.join(archive_html_folder, os.path.basename(html_file)))

    # Move PDF file to Archive pdf Folder
    # shutil.move(pdf_file, os.path.join(archive_folder, 'pdf', company_folder))
    shutil.move(pdf_file, os.path.join(archive_pdf_folder, os.path.basename(pdf_file)))

    # Copy Excel file to Archive excel Folder
    shutil.copy(excel_file, os.path.join(archive_excel_folder, os.path.basename(excel_file)))

    # Move Excel file to Outbound Folder
    shutil.move(excel_file, os.path.join(parent_dir, 'Outbound', company_folder))

# Main function
def main():
    parent_dir = 'Ever'

    # Iterate over PDF files in Inbound Folder
    inbound_dir = os.path.join(parent_dir, 'Inbound')
    for root_folder, dirs, files in os.walk(inbound_dir):
        for file in files:
            if file.endswith('.pdf'):
                pdf_file = os.path.join(root_folder, file)
                company_folder = os.path.basename(root_folder)
                # Convert PDF to HTML in the company folder
                html_file = pdf_to_html(pdf_file, os.path.join(inbound_dir, 'html'))
                # Process HTML to Excel
                html_to_excel(html_file, parent_dir, pdf_file)



    time.sleep(5)  # Wait for 5 seconds before exiting

if __name__ == "__main__":
    main()
