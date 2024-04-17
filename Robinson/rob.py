import pandas as pd
from datetime import datetime
import os
import shutil
from openpyxl import load_workbook

# Define parent directory
parent_dir = 'Robinson'

# Define the company names
company_names = {
    'ROBS': "ROBINSONS SUPERMARKET",
    'ROBD': "ROBINSONS DEPARTMENT STORE"
}

def clean_column_names(df):
    # Remove trailing white spaces from column names
    df.columns = df.columns.str.strip()

def move_to_archive(source_folder, target_folder):
    # Create the target directory if it doesn't exist
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # Move files from source to target directory
    for root, dirs, files in os.walk(source_folder):
        for file in files:
            source_file = os.path.join(root, file)
            target_file = os.path.join(target_folder, file)
            shutil.move(source_file, target_file)

def merge_excel_files_robd(company_code):
    company_folder = company_names.get(company_code)
    if company_folder is None:
        print(f"Company code '{company_code}' not found.")
        return
    
    # Construct paths relative to the parent directory
    inbound_folder = os.path.join(parent_dir, "Inbound", company_code)
    merged_folder = os.path.join(parent_dir, "Inbound", "Merged", company_code)
    archive_folder = os.path.join(parent_dir, "Archive", "excel", "Original", company_code, datetime.now().strftime('%Y%m%d%H%M%S'))
    
    # Construct full paths for Excel files
    summary_file_path = os.path.join(inbound_folder, "Outright Summary of Payments Date.xls")
    advice_file_path = os.path.join(inbound_folder, "Outright Payment Advice Date.xls")
    
    # Check if files exist
    if not os.path.exists(summary_file_path):
        print(f"Summary file not found at: {summary_file_path}")
        return
    if not os.path.exists(advice_file_path):
        print(f"Advice file not found at: {advice_file_path}")
        return
    
    # Read Excel files for the given company
    summary_file = pd.read_excel(summary_file_path, skiprows=12)
    advice_file = pd.read_excel(advice_file_path, skiprows=14)
    
    # Clean column names
    clean_column_names(summary_file)
    clean_column_names(advice_file)
    
    # Print data types of columns for debugging
    print(f"Summary File Data Types for {company_code}:")
    print(summary_file.dtypes)
    print(f"\nAdvice File Data Types for {company_code}:")
    print(advice_file.dtypes)
    
    # Merge "Cheque Amount" from summary to advice file
    advice_file["Cheque Amount"] = summary_file["Cheque Amount"]
    
    # Save the merged DataFrame to a new Excel file
    save_path = os.path.join(merged_folder, f"opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    advice_file.to_excel(save_path, index=False)
    print(f"Merged {company_code} files saved to: {save_path}")

    # Move original files to archive
    move_to_archive(inbound_folder, archive_folder)
    print("Original files moved to archive.")

def merge_excel_files_robs(company_code):
    company_folder = company_names.get(company_code)
    if company_folder is None:
        print(f"Company code '{company_code}' not found.")
        return
    
    # Construct paths relative to the parent directory
    inbound_folder = os.path.join(parent_dir, "Inbound", company_code)
    merged_folder = os.path.join(parent_dir, "Inbound", "Merged", company_code)
    archive_folder = os.path.join(parent_dir, "Archive", "excel", "Original", company_code, datetime.now().strftime('%Y%m%d%H%M%S'))
    
    # Construct full paths for Excel files
    summary_file_path = os.path.join(inbound_folder, "Outright Summary of Payments Day.xlsx")
    advice_file_path = os.path.join(inbound_folder, "Outright Payment Advice Day.xlsx")
    
    # Check if files exist
    if not os.path.exists(summary_file_path):
        print(f"Summary file not found at: {summary_file_path}")
        return
    if not os.path.exists(advice_file_path):
        print(f"Advice file not found at: {advice_file_path}")
        return
    
    # Read Excel files for the given company
    summary_file = pd.read_excel(summary_file_path, skiprows=13)
    advice_file = pd.read_excel(advice_file_path, skiprows=15)
    
    # Clean column names
    clean_column_names(summary_file)
    clean_column_names(advice_file)
    
    # Print data types of columns for debugging
    print(f"Summary File Data Types for {company_code}:")
    print(summary_file.dtypes)
    print(f"\nAdvice File Data Types for {company_code}:")
    print(advice_file.dtypes)
    
    # Merge "Cheque Amount" from summary to advice file
    advice_file["Cheque Amount"] = summary_file["Cheque Amount"]
    
    # Save the merged DataFrame to a new Excel file
    save_path = os.path.join(merged_folder, f"opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    advice_file.to_excel(save_path, index=False)
    print(f"Merged {company_code} files saved to: {save_path}")

    # Move original files to archive
    move_to_archive(inbound_folder, archive_folder)
    print("Original files moved to archive.")


def generate_inbound_outbound_excel(company_folder, company_names):
    # Get the company name based on the folder
    EDI_Customer = company_names[company_folder]

    # Load the data from the existing Excel file
    excel_dir = os.path.join(parent_dir, 'Inbound', 'Merged', company_folder)
    excel_files = [f for f in os.listdir(excel_dir) if f.endswith('.xlsx')]  # Filter out non-Excel files
    if len(excel_files) == 0:
        print("No Excel files found in the specified directory.")
        return
    
    # Assuming there's only one Excel file in the directory
    excel_file = excel_files[0]
    excel_path = os.path.join(excel_dir, excel_file)
    print("Excel file path:", excel_path)  # Added print statement to check the file path
    try:
        df_existing = pd.read_excel(excel_path, engine='openpyxl')  # Specify the engine as 'openpyxl'
    except Exception as e:
        print("Error occurred while reading Excel file:", e)  # Print any errors that occur
        return

    # Generate a new DataFrame with the desired columns
    data = {
        'EDI_Customer': [EDI_Customer] * len(df_existing),  # Repeat the customer name for each row
        'EDI_Company': df_existing['VENDOR'],   # Placeholder for company name
        'EDI_DocType': df_existing['Transaction_Type'], # Placeholder for document type
        'EDI_TransType': [None] * len(df_existing),
        'EDI_PORef': df_existing['PO Number.'],  # Assuming "PO Number" is the exact field name
        'EDI_InvRef': df_existing['Invoice No'],  # Assuming "Invoice No" is the exact field name
        'EDI_Gross': df_existing['RC Amount'],  # Assuming "RC Amount" is the exact field name
        'EDI_Discount': [None] * len(df_existing),
        'EDI_EWT': df_existing['EWT'],  # Assuming "EWT" is the exact field name
        'EDI_Net': df_existing['NET AMOUNT'],  # Assuming "NET AMOUNT" is the exact field name
        'EDI_RARef': df_existing['Payment Ref No'],  # Assuming "Payment Ref No" is the exact field name
        'EDI_RADate': [None] * len(df_existing),  # Placeholder for payment date
        'EDI_RAAmt': df_existing['Cheque Amount']  # Assuming "Cheque Amount" is the exact field name
    }
    df_new = pd.DataFrame(data)

    # Save the DataFrame to a new Excel file
    outbound_dir = os.path.join(parent_dir, 'Inbound', 'Outbound', company_folder)
    if not os.path.exists(outbound_dir):
        os.makedirs(outbound_dir)
    new_excel_file = os.path.join(outbound_dir, f"opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    df_new.to_excel(new_excel_file, index=False)
    print(f"New Excel file generated and saved to: {new_excel_file}")

# Call the functions to execute the merging process
merge_excel_files_robd('ROBD')
merge_excel_files_robs('ROBS')

generate_inbound_outbound_excel('ROBD', company_names)
generate_inbound_outbound_excel('ROBS', company_names)
