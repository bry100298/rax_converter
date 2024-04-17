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

# Call the functions to execute the merging process
merge_excel_files_robd('ROBD')
merge_excel_files_robs('ROBS')