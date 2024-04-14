import os
import shutil
import pandas as pd
import time

# Variable dictionary mapping company folders to company names
company_names = {
    'ROBS': "ROBINSONS SUPERMARKET",
    'ROBD': "ROBINSONS DEPARTMENT STORE"
}

def remove_quotes(value):
    # Function to remove quotes from a string
    return str(value).replace('"', '')

def merge_excel_files(company_folder, parent_dir):
    # Convert company folder name to lowercase
    company_folder_lower = company_folder.lower()
    
    # Construct paths to the Excel files based on company_folder
    merged_folder = os.path.join(parent_dir, 'Inbound', 'Merged', company_folder)
    summary_file_path = os.path.join(parent_dir, 'Inbound', company_folder, 'Outright Summary of Payments Date.xlsx')
    advice_file_path = os.path.join(parent_dir, 'Inbound', company_folder, 'Outright Payment Advice Date.xlsx')
    
    print(f"Summary file path: {summary_file_path}")
    print(f"Advice file path: {advice_file_path}")

    # Check if both Excel files exist
    if os.path.exists(summary_file_path) and os.path.exists(advice_file_path):
        print("Both Excel files exist.")
        try:
            # Read Excel files without header and handle exact values
            summary_df = pd.read_excel(summary_file_path, header=None)
            advice_df = pd.read_excel(advice_file_path, header=None)

            # Search for the row containing the fields by checking the first 3 rows
            for i in range(3):
                if "VENDOR CODE" in summary_df.iloc[i].values and "Payment Ref No" in summary_df.iloc[i].values:
                    summary_df = pd.read_excel(summary_file_path, header=i)
                    break
            else:
                raise ValueError("Fields not found in the first 3 rows of the Summary file.")

            for i in range(3):
                if "VENDOR CODE" in advice_df.iloc[i].values and "Payment Ref No" in advice_df.iloc[i].values:
                    advice_df = pd.read_excel(advice_file_path, header=i)
                    break
            else:
                raise ValueError("Fields not found in the first 3 rows of the Advice file.")

            # Merge the dataframes based on common columns
            merged_df = pd.merge(advice_df, summary_df[['VENDOR CODE', 'Payment Ref No']], 
                                on=['VENDOR CODE', 'Payment Ref No'], how='left')

            # Check if the merged DataFrame is not empty
            if not merged_df.empty:
                print("Merged DataFrame is not empty.")
                # Create a new DataFrame with the desired columns
                new_df = pd.DataFrame({
                    'EDI_Customer': company_names.get(company_folder_lower, company_folder),
                    'EDI_Company': merged_df['VENDOR'],
                    'EDI_DocType': merged_df['Transaction_Type'],
                    'EDI_TransType': None,
                    'EDI_PORef': merged_df['PO Number.'],
                    'EDI_InvRef': merged_df['Invoice No'],
                    'EDI_Gross': merged_df['RC Amount'],
                    'EDI_Discount': None,
                    'EDI_EWT': merged_df['EWT'],
                    'EDI_Net': merged_df['NET AMOUNT'],
                    'EDI_RARef': merged_df['Payment Ref No'],
                    'EDI_RADate': merged_df['PAYMENT_DATE'],
                    'EDI_RAAmt': merged_df['Cheque Amount']
                })

                # Create Excel file path for the new merged file
                merged_excel_file = os.path.join(merged_folder, f'opadosopd_{int(time.time())}.xlsx')

                # Write the new DataFrame to Excel
                new_df.to_excel(merged_excel_file, index=False)

                # Archive original Excel files
                archive_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Original', company_folder, str(int(time.time())))
                os.makedirs(archive_folder, exist_ok=True)
                print(f"Archive folder path: {archive_folder}")
                shutil.move(summary_file_path, archive_folder)
                shutil.move(advice_file_path, archive_folder)

                # Move merged Excel file to Outbound folder
                outbound_folder = os.path.join(parent_dir, 'Outbound', company_folder)
                os.makedirs(outbound_folder, exist_ok=True)
                shutil.move(merged_excel_file, outbound_folder)

                # Archive merged Excel file
                archive_merged_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Original_Merged', company_folder, str(int(time.time())))
                os.makedirs(archive_merged_folder, exist_ok=True)
                shutil.move(merged_excel_file, archive_merged_folder)

                # Create a copy of the merged Excel file in Converted folder
                converted_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Converted', company_folder, str(int(time.time())))
                os.makedirs(converted_folder, exist_ok=True)
                shutil.copy(os.path.join(outbound_folder, os.path.basename(merged_excel_file)), converted_folder)
            else:
                print(f"Merged DataFrame is empty for {company_folder}")
        except Exception as e:
            print(f"Error occurred while merging Excel files for {company_folder}: {e}")
    else:
        print(f"Excel files not found for {company_folder}")


def move_files_to_archive(parent_dir):
    # Move raw files from Inbound to Archive
    for company_folder in os.listdir(os.path.join(parent_dir, 'Inbound')):
        if os.path.isdir(os.path.join(parent_dir, 'Inbound', company_folder)):
            for file_name in os.listdir(os.path.join(parent_dir, 'Inbound', company_folder)):
                file_path = os.path.join(parent_dir, 'Inbound', company_folder, file_name)
                if os.path.isfile(file_path):
                    try:
                        # Create directory structure in Archive/excel/Original
                        archive_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Original', company_folder, str(int(time.time())))
                        os.makedirs(archive_folder, exist_ok=True)
                        # Move file to archive folder
                        shutil.move(file_path, archive_folder)
                    except Exception as e:
                        print(f"Error occurred while moving file to archive for {company_folder}: {e}")

def move_files_to_outbound(parent_dir):
    # Move merged files from Merged to Outbound
    for company_folder in os.listdir(os.path.join(parent_dir, 'Inbound', 'Merged')):
        if os.path.isdir(os.path.join(parent_dir, 'Inbound', 'Merged', company_folder)):
            for file_name in os.listdir(os.path.join(parent_dir, 'Inbound', 'Merged', company_folder)):
                file_path = os.path.join(parent_dir, 'Inbound', 'Merged', company_folder, file_name)
                if os.path.isfile(file_path) and file_name.endswith('.xlsx'):
                    try:
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

                        # Create a copy of the merged Excel file in Converted folder
                        converted_folder = os.path.join(parent_dir, 'Archive', 'excel', 'Converted', company_folder, str(int(time.time())))
                        os.makedirs(converted_folder, exist_ok=True)
                        shutil.copy(destination_path, converted_folder)
                    except Exception as e:
                        print(f"Error occurred while moving merged file to outbound for {company_folder}: {e}")

def main():
    parent_dir = 'Robinson'

    # Wait for 5 seconds before merging Excel files
    time.sleep(5)

    # Iterate over company folders for merging Excel files
    for company_folder in os.listdir(os.path.join(parent_dir, 'Inbound')):
        if os.path.isdir(os.path.join(parent_dir, 'Inbound', company_folder)):
            merge_excel_files(company_folder, parent_dir)

    # Move files to Archive directory
    move_files_to_archive(parent_dir)

    # Move merged files to Outbound directory
    move_files_to_outbound(parent_dir)

if __name__ == "__main__":
    main()
