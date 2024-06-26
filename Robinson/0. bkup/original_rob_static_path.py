import pandas as pd
from datetime import datetime
import os
import shutil

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

def generate_excel(company_folder, source_folder, target_folder):
    # Read the merged Excel file
    excel_file = os.path.join(source_folder, f"opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    df = pd.read_excel(excel_file)
    
    # Define the company names
    # company_names = {
    #     "ROBS": "ROBINSONS SUPERMARKET",
    #     "ROBD": "ROBINSONS DEPARTMENT STORE"
    # }
    
    # Define the mapping of fields
    field_mapping = {
        "EDI_Customer": lambda x: company_names.get(company_folder, ''),
        "EDI_Company": "VENDOR",
        "EDI_DocType": "Transaction_Type",
        "EDI_TransType": None,
        "EDI_PORef": "PO Number.",
        "EDI_InvRef": "Invoice No",
        "EDI_Gross": "RC Amount",
        "EDI_Discount": None,
        "EDI_EWT": "EWT",
        "EDI_Net": "NET AMOUNT",
        "EDI_RARef": "Payment Ref No",
        "EDI_RADate": "PAYMENT_DATE",
        "EDI_RAAmt": "Cheque Amount"
    }
    
    # Apply the field mapping
    for field, column in field_mapping.items():
        if column is not None:
            df[field] = df[column]
        else:
            df[field] = None
    
    # Create the target directory if it doesn't exist
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # Save the DataFrame to a new Excel file
    save_path = os.path.join(target_folder, f"{company_folder}_EDI_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    df.to_excel(save_path, index=False)
    print(f"Excel file generated for {company_folder}: {save_path}")

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

    # Generate Excel file with specific fields
    # generate_excel("ROBD", "C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/Merged/ROBD", "C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/Outbound/ROBD")

def merge_excel_files_robs():
    # Read Excel files for ROBS
    summary_file = pd.read_excel("C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBS/Outright Summary of Payments Day.xlsx", skiprows=13)
    advice_file = pd.read_excel("C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBS/Outright Payment Advice Day.xlsx", skiprows=15)
    
    # Print data types of columns for debugging
    print("Summary File Data Types for ROBS:")
    print(summary_file.dtypes)
    print("\nAdvice File Data Types for ROBS:")
    print(advice_file.dtypes)
    
    # Merge "Cheque Amount" from summary to advice file
    advice_file["Cheque Amount"] = summary_file["Cheque Amount"]
    
    # Save the merged DataFrame to a new Excel file
    save_path = f"C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/Merged/ROBS/opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    advice_file.to_excel(save_path, index=False)
    print(f"Merged ROBS files saved to: {save_path}")

    # Move original files to archive
    source_folder = "C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBS/"
    target_folder = f"C:/Users/User/Documents/Project/rax_converter/Robinson/Archive/excel/Original/ROBS/{datetime.now().strftime('%Y%m%d%H%M%S')}/"
    move_to_archive(source_folder, target_folder)
    print("Original files moved to archive.")

    # Generate Excel file with specific fields
    # generate_excel("ROBS", "C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/Merged/ROBS", "C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/Outbound/ROBS")

# Call the functions to execute the merging process
merge_excel_files_robd('ROBD')
# merge_excel_files_robs('ROBS')