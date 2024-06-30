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

    # Find the indices of rows containing "Company Confidential"
    confidential_indices = df.index[df.isin(["Company Confidential"]).any(axis=1)].tolist()

    # Delete the row before each occurrence of "Company Confidential"
    for idx in confidential_indices:
        if idx > 0:
            df.drop(idx - 1, inplace=True)

    # Now drop rows containing "Company Confidential"
    df.drop(df[df.isin(["Company Confidential"]).any(axis=1)].index, inplace=True)

    # Remove rows containing "Company Confidential" in any column
    # df.drop(df[df.isin(["Company Confidential"]).any(axis=1)].index, inplace=True)
    # Remove rows where all columns have no characters or only whitespace
    # df.drop(df[df.apply(lambda x: x.str.strip().eq('').all(), axis=1)].index, inplace=True)

def move_to_archive(source_folder, target_folder):
    # Create the target directory if it doesn't exist
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # Move files from source to target directory
    for root, dirs, files in os.walk(source_folder):
        for file in files:
            if file != ".gitkeep":
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
    summary_file = pd.read_excel(summary_file_path, skiprows=12, dtype=str)
    advice_file = pd.read_excel(advice_file_path, skiprows=14, dtype=str)
    
    # Clean column names
    clean_column_names(summary_file)
    clean_column_names(advice_file)
    
    # Print data types of columns for debugging
    print(f"Summary File Data Types for {company_code}:")
    print(summary_file.dtypes)
    print(f"\nAdvice File Data Types for {company_code}:")
    print(advice_file.dtypes)
    
    # Merge based on "Payment Ref No" column
    merged_df = pd.merge(advice_file, summary_file[['Payment Ref No', 'Cheque Amount', 'Payment Date']], on='Payment Ref No', how='left')

    # Merge "Cheque Amount" from summary to advice file
    # advice_file["Cheque Amount"] = summary_file["Cheque Amount"]

    # Update "Cheque Amount" column in advice file with values from summary file where applicable
    advice_file["Cheque Amount"] = merged_df["Cheque Amount"]

    advice_file["Payment Date"] = merged_df["Payment Date"]
    # advice_file["Payment Date"] = summary_file["Payment Date"]  # Add Payment Date column
    # advice_file["Payment Ref No"] = summary_file["Payment Ref No"]  # Add Payment Date column

    # Format 'Payment Ref No' column to have leading zeros and a fixed width of 10 characters
    # advice_file["Payment Ref No"] = advice_file["Payment Ref No"].astype(str).str.zfill(10)

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
    summary_file = pd.read_excel(summary_file_path, skiprows=13, dtype=str)
    advice_file = pd.read_excel(advice_file_path, skiprows=15, dtype=str)
    
    # Clean column names
    clean_column_names(summary_file)
    clean_column_names(advice_file)
    
    # Print data types of columns for debugging
    print(f"Summary File Data Types for {company_code}:")
    print(summary_file.dtypes)
    print(f"\nAdvice File Data Types for {company_code}:")
    print(advice_file.dtypes)
    

    # Merge based on "Payment Ref No" column
    merged_df = pd.merge(advice_file, summary_file[['Payment Ref No', 'Cheque Amount', 'Payment Date.']], on='Payment Ref No', how='left')

    # Merge "Cheque Amount" "Payment Date." and "Payment Ref No"
    # advice_file["Cheque Amount"] = summary_file["Cheque Amount"]

    # Update "Cheque Amount" column in advice file with values from summary file where applicable
    advice_file["Cheque Amount"] = merged_df["Cheque Amount"]
    
    # # Create a dictionary mapping Payment Ref No to Cheque Amount in summary_file
    # payment_ref_to_cheque_amount = summary_file.set_index('Payment Ref No')['Cheque Amount'].to_dict()

    # # Update "Cheque Amount" column in advice_file with values from summary file
    # advice_file['Cheque Amount'] = advice_file['Payment Ref No'].map(payment_ref_to_cheque_amount)


    advice_file["Payment Date."] = merged_df["Payment Date."]
    # advice_file["Payment Date."] = summary_file["Payment Date."]  # Add Payment Date column


    # advice_file["Payment Ref No"] = summary_file["Payment Ref No"]  # Add Payment Date column
    
    # Save the merged DataFrame to a new Excel file
    save_path = os.path.join(merged_folder, f"opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    advice_file.to_excel(save_path, index=False)
    print(f"Merged {company_code} files saved to: {save_path}")

    # Move original files to archive
    move_to_archive(inbound_folder, archive_folder)
    print("Original files moved to archive.")

def get_payment_date(row):
    if 'PAYMENT_DATE' in row.index and pd.notnull(row['PAYMENT_DATE']) and row['PAYMENT_DATE'].strip() != '':
        return row['PAYMENT_DATE']
    else:
        return row.get('Payment Date.', row.get('Payment Date'))
    
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
        # df_existing = pd.read_excel(excel_path, engine='openpyxl')  # Specify the engine as 'openpyxl'
        df_existing = pd.read_excel(excel_path, engine='openpyxl', dtype=str)  # Specify dtype as str to treat all columns as strings
    except Exception as e:
        print("Error occurred while reading Excel file:", e)  # Print any errors that occur
        return

    # Generate a new DataFrame with the desired columns
    data = {
        'EDI_Customer': [EDI_Customer] * len(df_existing),  # Repeat the customer name for each row
        'EDI_Company': df_existing['VENDOR'],   # Placeholder for company name
        'EDI_DocType': df_existing['Transaction_Type'], # Placeholder for document type
        # 'EDI_TransType': [None] * len(df_existing),
        # 'EDI_DocDescr': df_existing['VENDOR CODE'].astype(str) + '_' + df_existing['Document Description'].astype(str),
        # 'EDI_DocDescr': df_existing.get('VENDOR CODE', df_existing.get('Vendor Code')).astype(str) + '_' + df_existing['Document Description'].astype(str),
        'EDI_DocDescr': (
            df_existing.get('VENDOR CODE', df_existing.get('Vendor Code', '')).fillna('').astype(str).replace('^$', '', regex=True) + 
            '_' + 
            df_existing['Document Description'].fillna('').astype(str).replace('^$', '', regex=True)
        ).replace('^_$', '', regex=True).replace('^_', '', regex=True).replace('_$', '', regex=True),
        'EDI_PORef': df_existing['PO Number.'],  # Assuming "PO Number" is the exact field name
        'EDI_InvRef': df_existing['Invoice No'],  # Assuming "Invoice No" is the exact field name
        'EDI_Gross': df_existing['RC Amount'],  # Assuming "RC Amount" is the exact field name
        'EDI_Discount': [None] * len(df_existing),
        'EDI_EWT': df_existing['EWT'],  # Assuming "EWT" is the exact field name
        'EDI_Net': df_existing['NET AMOUNT'],  # Assuming "NET AMOUNT" is the exact field name
        'EDI_RARef': df_existing['Payment Ref No'],  # Assuming "Payment Ref No" is the exact field name
        # 'EDI_RARef': df_existing['Payment Ref No'].astype(int).astype(str).str.zfill(10),
        # 'EDI_RARef': df_existing['Payment Ref No'].fillna(0).astype(int).astype(str).str.zfill(10),
        # 'EDI_RADate': [None] * len(df_existing),  # Placeholder for payment date
        # 'EDI_RADate': df_existing['Payment Date'],
        # 'EDI_RADate': df_existing.get('Payment Date', df_existing.get('Payment Date.')),  # Use get() to handle both variations

        # 'EDI_RADate': df_existing.apply(get_payment_date, axis=1),
        # 'EDI_RADate': df_existing.get('Payment Date', df_existing['PAYMENT_DATE'], df_existing.get('Payment Date.')),

        
        'EDI_RADate': df_existing.apply(get_payment_date, axis=1),

        'EDI_RAAmt': df_existing['Cheque Amount']  # Assuming "Cheque Amount" is the exact field name
    }
    df_new = pd.DataFrame(data)

    # Save the DataFrame to a new Excel file
    outbound_dir = os.path.join(parent_dir, 'Inbound', 'Outbound', company_folder)
    if not os.path.exists(outbound_dir):
        os.makedirs(outbound_dir)
    # new_excel_file = os.path.join(outbound_dir, f"opadosopd_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    new_excel_file = os.path.join(outbound_dir, f"{company_folder}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    df_new.to_excel(new_excel_file, index=False)
    print(f"New Excel file generated and saved to: {new_excel_file}")

    archive_original_merged_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Original_Merged', company_folder)

    # Move original_merged files to archive
    move_to_archive(excel_dir, archive_original_merged_folder)
    print("Original Merged files moved to archive.")

    # shutil.move(excel_dir, os.path.join(archive_original_merged_folder, 'xlsx', company_folder))
    # print("Original Merged files moved to archive.")

    # outbound_dir_main = os.path.join(parent_dir, 'Outbound', company_folder)
    archive_excel_converted = os.path.join(parent_dir, 'Archive', 'excel', 'Converted', company_folder)
    # Copy Excel file to Archive excel Folder
    shutil.copy(new_excel_file, archive_excel_converted)

    outbound_dir_main = os.path.join(parent_dir, 'Outbound', company_folder)
    shutil.move(new_excel_file, outbound_dir_main)



# Call the functions to execute the merging process
merge_excel_files_robd('ROBD')
merge_excel_files_robs('ROBS')

generate_inbound_outbound_excel('ROBD', company_names)
generate_inbound_outbound_excel('ROBS', company_names)
