import os
import shutil
import pandas as pd
import time

# Variable dictionary mapping company folders to company names
company_names = {
    'ROBS': "ROBINSONS SUPERMARKET",
    'ROBD': "ROBINSONS DEPARTMENT STORE"
}

def merge_and_process_excel(company_folder, parent_dir, company_names):
    # Convert company folder name to lowercase
    company_folder_lower = company_folder.lower()
    
    # Construct paths to the Excel files
    summary_file_path = os.path.join(parent_dir, 'Inbound', company_folder_lower, 'Outright Summary of Payments Date.xlsx')
    advice_file_path = os.path.join(parent_dir, 'Inbound', company_folder_lower, 'Outright Payment Advice Date.xlsx')
    
    print(f"Summary file path: {summary_file_path}")
    print(f"Advice file path: {advice_file_path}")
    
    # Check if both Excel files exist
    if os.path.exists(summary_file_path) and os.path.exists(advice_file_path):
        print("Both Excel files exist.")
        # Read Excel files
        summary_df = pd.read_excel(summary_file_path)
        advice_df = pd.read_excel(advice_file_path)

        # Check if 'VENDOR CODE' column exists in both dataframes
        vendor_code_col_summary = summary_df.columns[summary_df.columns.str.upper() == 'VENDOR CODE'].tolist()
        vendor_code_col_advice = advice_df.columns[advice_df.columns.str.upper() == 'VENDOR CODE'].tolist()

        if vendor_code_col_summary and vendor_code_col_advice:
            # Merge the dataframes based on common columns
            merged_df = pd.merge(advice_df, summary_df[[vendor_code_col_summary[0], 'Payment Ref No', 'Cheque Amount']], 
                                on=[vendor_code_col_summary[0], 'Payment Ref No'], how='left')

            # Check if the merged DataFrame is not empty
            if not merged_df.empty:
                print("Merged DataFrame is not empty.")
                # Create a new DataFrame with the desired columns
                new_df = pd.DataFrame({
                    'EDI_Customer': company_names.get(company_folder_lower, company_folder),
                    'EDI_Company': merged_df['VENDOR'],
                    'EDI_DocType': merged_df['Document Type'] if 'Document Type' in merged_df.columns else merged_df.get('Transaction_Type', ''),
                    'EDI_TransType': merged_df.get('Transaction_Type', ''),
                    'EDI_PORef': merged_df.get('PO Number.', ''),
                    'EDI_InvRef': merged_df.get('Invoice No', ''),
                    'EDI_Gross': merged_df.get('RC Amount', ''),
                    'EDI_Discount': None,
                    'EDI_EWT': merged_df.get('EWT', ''),
                    'EDI_Net': merged_df.get('NET AMOUNT', ''),
                    'EDI_RARef': merged_df['Payment Ref No'],
                    'EDI_RADate': merged_df.get('PAYMENT_DATE', ''),
                    'EDI_RAAmt': merged_df['Cheque Amount']
                })

                # Create Excel file path for the new merged file
                merged_folder = os.path.join(parent_dir, 'Inbound', 'Merged', company_folder_lower)
                os.makedirs(merged_folder, exist_ok=True)
                merged_excel_file = os.path.join(merged_folder, f'opadosopd_{int(time.time())}.xlsx')

                # Write the new DataFrame to Excel
                new_df.to_excel(merged_excel_file, index=False)

                # Archive original Excel files
                archive_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Original', company_folder_lower, str(int(time.time())))
                os.makedirs(archive_folder, exist_ok=True)
                shutil.move(summary_file_path, archive_folder)
                shutil.move(advice_file_path, archive_folder)

                # Move merged Excel file to Outbound folder
                outbound_folder = os.path.join(parent_dir, 'Outbound', company_folder_lower)
                os.makedirs(outbound_folder, exist_ok=True)
                shutil.move(merged_excel_file, outbound_folder)

                # Archive merged Excel file
                archive_merged_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Original_Merged', company_folder_lower, str(int(time.time())))
                os.makedirs(archive_merged_folder, exist_ok=True)
                shutil.move(merged_excel_file, archive_merged_folder)

                # Create a copy of the merged Excel file in Converted folder
                converted_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Converted', company_folder_lower, str(int(time.time())))
                os.makedirs(converted_folder, exist_ok=True)
                shutil.copy(os.path.join(outbound_folder, os.path.basename(merged_excel_file)), converted_folder)
            else:
                print(f"Merged DataFrame is empty for {company_folder}")
        else:
            print(f"'VENDOR CODE' column is missing in one of the dataframes for {company_folder}")
    else:
        print(f"Excel files not found for {company_folder}")

def move_files_to_archive_and_outbound(parent_dir):
    # Move raw files from Inbound to Archive
    for company_folder in os.listdir(os.path.join(parent_dir, 'Inbound')):
        if os.path.isdir(os.path.join(parent_dir, 'Inbound', company_folder)):
            for file_name in os.listdir(os.path.join(parent_dir, 'Inbound', company_folder)):
                file_path = os.path.join(parent_dir, 'Inbound', company_folder, file_name)
                if os.path.isfile(file_path):
                    # Create directory structure in Archive/excel/Original
                    archive_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Original', company_folder, str(int(time.time())))
                    os.makedirs(archive_folder, exist_ok=True)
                    # Move file to archive folder
                    shutil.move(file_path, archive_folder)

    # Move merged files from Merged to Outbound
    for company_folder in os.listdir(os.path.join(parent_dir, 'Inbound', 'Merged')):
        if os.path.isdir(os.path.join(parent_dir, 'Inbound', 'Merged', company_folder)):
            for file_name in os.listdir(os.path.join(parent_dir, 'Inbound', 'Merged', company_folder)):
                file_path = os.path.join(parent_dir, 'Inbound', 'Merged', company_folder, file_name)
                if os.path.isfile(file_path) and file_name.endswith('.xlsx'):
                    # Create directory structure in Outbound
                    outbound_folder = os.path.join(parent_dir, 'Outbound', company_folder)
                    os.makedirs(outbound_folder, exist_ok=True)
                    # Move file to outbound folder if it doesn't already exist
                    destination_path = os.path.join(outbound_folder, file_name)
                    if not os.path.exists(destination_path):
                        shutil.move(file_path, outbound_folder)
                    else:
                        print(f"Excel file '{file_name}' already exists in destination. Skipping move operation.")

                    # Archive merged files
                    archive_merged_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Original_Merged', company_folder, str(int(time.time())))
                    os.makedirs(archive_merged_folder, exist_ok=True)
                    shutil.move(file_path, archive_merged_folder)

def main():
    parent_dir = 'Robinson'

    # Wait for 5 seconds before merging and processing Excel files
    time.sleep(5)

    # Iterate over company folders for merging and processing Excel files
    for company_folder in os.listdir(os.path.join(parent_dir, 'Inbound')):
        if os.path.isdir(os.path.join(parent_dir, 'Inbound', company_folder)):
            merge_and_process_excel(company_folder, parent_dir, company_names)

    # Move files to Archive and Outbound directories
    move_files_to_archive_and_outbound(parent_dir)

if __name__ == "__main__":
    main()
