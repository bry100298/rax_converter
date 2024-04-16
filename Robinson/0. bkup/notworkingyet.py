import pandas as pd
from datetime import datetime
import os
import shutil

# Define parent directory
parent_dir = 'Robinson'

# Define paths
inbound_folder = os.path.join(parent_dir, 'Inbound')
outbound_folder = os.path.join(parent_dir, 'Outbound')
inbound_outbound_folder = os.path.join(parent_dir, 'Inbound', 'Outbound')
archive_folder = os.path.join(parent_dir, 'Archive')
error_folder = os.path.join(parent_dir, 'Error')

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

def merge_excel_files_robd(company_folder, summary_file_path, advice_file_path, merged_folder, archive_folder):
    # Read Excel files for ROBD
    summary_file = pd.read_excel(summary_file_path, skiprows=12)
    advice_file = pd.read_excel(advice_file_path, skiprows=14)
    
    # Clean column names
    clean_column_names(summary_file)
    clean_column_names(advice_file)
    
    # Print data types of columns for debugging
    print("Summary File Data Types for ROBD:")
    print(summary_file.dtypes)
    print("\nAdvice File Data Types for ROBD:")
    print(advice_file.dtypes)
    
    # Merge "Cheque Amount" from summary to advice file
    advice_file["Cheque Amount"] = summary_file["Cheque Amount"]
    
    # Save the merged DataFrame to a new Excel file
    save_path = os.path.join(merged_folder, f"opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    advice_file.to_excel(save_path, index=False)
    print(f"Merged {company_folder} files saved to: {save_path}")

    # Move original files to archive
    move_to_archive(company_folder, archive_folder)
    print("Original files moved to archive.")

def merge_excel_files_robs(summary_file_path, advice_file_path, merged_folder, archive_folder):
    # Read Excel files for ROBS
    summary_file = pd.read_excel(summary_file_path, skiprows=13)
    advice_file = pd.read_excel(advice_file_path, skiprows=15)
    
    # Print data types of columns for debugging
    print("Summary File Data Types for ROBS:")
    print(summary_file.dtypes)
    print("\nAdvice File Data Types for ROBS:")
    print(advice_file.dtypes)
    
    # Merge "Cheque Amount" from summary to advice file
    advice_file["Cheque Amount"] = summary_file["Cheque Amount"]
    
    # Save the merged DataFrame to a new Excel file
    save_path = os.path.join(merged_folder, f"opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    advice_file.to_excel(save_path, index=False)
    print(f"Merged ROBS files saved to: {save_path}")

    # Move original files to archive
    move_to_archive(summary_file_path, archive_folder)
    print("Original files moved to archive.")

# Call the function to execute the merging process
merge_excel_files_robd(os.path.join(inbound_folder, 'ROBD', 'Outright Summary of Payments Date.xls'), os.path.join(inbound_folder, 'ROBD', 'Outright Payment Advice Date.xls'), os.path.join(inbound_folder, 'Merged', 'ROBD'), os.path.join(archive_folder, 'excel', 'Original', 'ROBD'), archive_folder)

merge_excel_files_robs(os.path.join(inbound_folder, 'ROBS', 'Outright Summary of Payments Day.xlsx'), os.path.join(inbound_folder, 'ROBS', 'Outright Payment Advice Day.xlsx'), os.path.join(inbound_folder, 'Merged', 'ROBS'), os.path.join(archive_folder, 'excel', 'Original', 'ROBS'))
