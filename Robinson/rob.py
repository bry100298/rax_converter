import os
import shutil
import pandas as pd
import time

# Variable dictionary mapping company folders to company names
company_names = {
    'ROBS': "ROBINSONS SUPERMARKET",
    'ROBD': "ROBINSONS DEPARTMENT STORE"
}

# Function to merge and update Excel files
def merge_and_update_excel(inbound_dir, archive_dir):
    for company_folder in os.listdir(inbound_dir):
        if os.path.isdir(os.path.join(inbound_dir, company_folder)):
            summary_file = os.path.join(inbound_dir, company_folder, 'Outright Summary of Payments Date.xlsx')
            advice_file = os.path.join(inbound_dir, company_folder, 'Outright Payment Advice Date.xlsx')

            # Read Excel files into DataFrames
            df_summary = pd.read_excel(summary_file)
            df_advice = pd.read_excel(advice_file)

            # Merge DataFrames on common columns
            merged_df = pd.merge(df_summary, df_advice[['VENDOR CODE', 'Payment Ref No', 'Cheque Amount']], 
                                 on=['VENDOR CODE', 'Payment Ref No'], how='left')

            # Generate timestamp for filename
            timestamp = int(time.time())

            # Save merged DataFrame with additional column to a new Excel file
            merged_folder = os.path.join(inbound_dir, 'Merged', company_folder)
            os.makedirs(merged_folder, exist_ok=True)
            merged_file = os.path.join(merged_folder, f'opadosopd_{timestamp}.xlsx')
            merged_df.to_excel(merged_file, index=False)

            # Move raw Excel files to Archive/original
            archive_original_folder = os.path.join(archive_dir, 'excel', 'Original', company_folder, str(timestamp))
            os.makedirs(archive_original_folder, exist_ok=True)
            shutil.move(summary_file, archive_original_folder)
            shutil.move(advice_file, archive_original_folder)

            # Add additional columns and update Excel file in Outbound folder
            update_excel(merged_file, company_folder)

            # Move merged Excel file to Archive/original_merged
            archive_merged_folder = os.path.join(archive_dir, 'excel', 'Original_Merged', company_folder, str(timestamp))
            os.makedirs(archive_merged_folder, exist_ok=True)
            shutil.move(merged_file, archive_merged_folder)

            # Copy merged Excel file to Outbound folder
            outbound_folder = os.path.join(inbound_dir, 'Outbound', company_folder)
            os.makedirs(outbound_folder, exist_ok=True)
            shutil.copy(merged_file, outbound_folder)

            # Copy converted Excel file to Archive/Converted
            archive_converted_folder = os.path.join(archive_dir, 'excel', 'Converted', company_folder, str(timestamp))
            os.makedirs(archive_converted_folder, exist_ok=True)
            shutil.copy(merged_file, archive_converted_folder)

# Function to add additional columns and update Excel file
def update_excel(excel_file, company_folder):
    # Read Excel file into DataFrame
    df = pd.read_excel(excel_file)

    # Add additional columns based on instructions
    df['EDI_Customer'] = company_names[company_folder]
    df['EDI_Company'] = df['VENDOR']
    df['EDI_DocType'] = df['Transaction_Type']
    df['EDI_TransType'] = None
    df['EDI_PORef'] = df['PO Number']
    df['EDI_InvRef'] = df['Invoice No']
    df['EDI_Gross'] = df['RC Amount']
    df['EDI_Discount'] = None
    df['EDI_EWT'] = df['EWT']
    df['EDI_Net'] = df['NET AMOUNT']
    df['EDI_RARef'] = df['Payment Ref No']
    df['EDI_RADate'] = df['PAYMENT_DATE']
    df['EDI_RAAmt'] = df['Cheque Amount']

    # Save updated DataFrame to Excel
    df.to_excel(excel_file, index=False)

# Main function
def main():
    parent_dir = 'Robinson'

    inbound_dir = os.path.join(parent_dir, 'Inbound')
    archive_dir = 'Archive'

    # Merge and update Excel files for each company folder
    merge_and_update_excel(inbound_dir, archive_dir)

# Execute the main function
if __name__ == "__main__":
    main()
