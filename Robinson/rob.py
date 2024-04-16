import pandas as pd
from datetime import datetime

def clean_column_names(df):
    # Remove trailing white spaces from column names
    df.columns = df.columns.str.strip()

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

# Call the functions to execute the merging process
merge_excel_files_robd()
merge_excel_files_robs()