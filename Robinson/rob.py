import pandas as pd
from datetime import datetime
import os
import shutil

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

def merge_excel_files_robd():
    # Read Excel files for ROBD
    summary_file = pd.read_excel("C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBD/Outright Summary of Payments Date.xls", skiprows=12)
    advice_file = pd.read_excel("C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBD/Outright Payment Advice Date.xls", skiprows=14)
    
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
    save_path = f"C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/Merged/ROBD/opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    advice_file.to_excel(save_path, index=False)
    print(f"Merged ROBD files saved to: {save_path}")

    # Move original files to archive
    source_folder = "C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBD/"
    target_folder = f"C:/Users/User/Documents/Project/rax_converter/Robinson/Archive/excel/Original/ROBD/{datetime.now().strftime('%Y%m%d%H%M%S')}/"
    move_to_archive(source_folder, target_folder)
    print("Original files moved to archive.")

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

# Call the functions to execute the merging process
merge_excel_files_robd()
merge_excel_files_robs()