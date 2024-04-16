import pandas as pd
from datetime import datetime

def merge_excel_files_robd():
    # Read Excel files for ROBD
    summary_file = pd.read_excel("C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBD/Outright Summary of Payments Date.xls")
    advice_file = pd.read_excel("C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBD/Outright Payment Advice Date.xls")
    
    # Convert "Payment Ref No" column in advice file to object type
    advice_file["Payment Ref No"] = advice_file["Payment Ref No"].astype(str)
    
    # Print data types of columns for debugging
    print("Summary File Data Types for ROBD:")
    print(summary_file.dtypes)
    print("\nAdvice File Data Types for ROBD:")
    print(advice_file.dtypes)
    
    # Left join based on "VENDOR CODE" and "Payment Ref No"
    merged_df = pd.merge(advice_file, summary_file, on=["VENDOR CODE", "Payment Ref No"], how="left")
    
    # Rename "Cheque Amount" column
    merged_df.rename(columns={"Cheque Amount_y": "Cheque Amount"}, inplace=True)
    
    # Save the merged DataFrame to a new Excel file
    save_path = f"C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/Merged/ROBD/opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    merged_df.to_excel(save_path, index=False)
    print(f"Merged ROBD files saved to: {save_path}")

def merge_excel_files_robs():
    # Read Excel files for ROBS
    summary_file = pd.read_excel("C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBS/Outright Summary of Payments Day.xlsx")
    advice_file = pd.read_excel("C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/ROBS/Outright Payment Advice Day.xlsx")
    
    # Convert "Payment Ref No" column in advice file to object type
    advice_file["Payment Ref No"] = advice_file["Payment Ref No"].astype(str)
    
    # Print data types of columns for debugging
    print("Summary File Data Types for ROBS:")
    print(summary_file.dtypes)
    print("\nAdvice File Data Types for ROBS:")
    print(advice_file.dtypes)
    
    # Left join based on "Vendor Code" and "Payment Ref No"
    merged_df = pd.merge(advice_file, summary_file, on=["Vendor Code", "Payment Ref No"], how="left")
    
    # Rename "Cheque Amount" column
    merged_df.rename(columns={"Cheque Amount_y": "Cheque Amount"}, inplace=True)
    
    # Save the merged DataFrame to a new Excel file
    save_path = f"C:/Users/User/Documents/Project/rax_converter/Robinson/Inbound/Merged/ROBS/opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    merged_df.to_excel(save_path, index=False)
    print(f"Merged ROBS files saved to: {save_path}")

# Call the functions to execute the merging process
merge_excel_files_robd()
merge_excel_files_robs()
