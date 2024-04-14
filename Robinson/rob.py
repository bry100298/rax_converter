import os
import shutil
import pandas as pd
from datetime import datetime

# Variable dictionary mapping company folders to company names
company_names = {
    'PPCI': "PUREGOLD PRICE CLUB INC.",
    'AGRI': "AYAGOLD RETAILERS, INC."
}

# Function to merge Excel files
def merge_files(company_folder):
    # File paths
    advice_file = os.path.join('Inbound', company_folder, 'Outright Payment Advice Date.xls')
    summary_file = os.path.join('Inbound', company_folder, 'Outright Summary of Payments Date.xls')
    print(f"Advice file path: {advice_file}")
    print(f"Summary file path: {summary_file}")

    # Read Excel files
    advice_df = pd.read_excel(advice_file)
    summary_df = pd.read_excel(summary_file)

    # Merge and get Cheque Amount
    merged_df = pd.merge(advice_df, summary_df[['Payment Ref No', 'Cheque Amount']], on='Payment Ref No', how='left')
    merged_df.rename(columns={'Cheque Amount': 'EDI_RAAmt'}, inplace=True)

    # Move files to Archive/Original
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    os.makedirs(os.path.join('Archive', 'excel', 'Original', company_folder, timestamp), exist_ok=True)
    shutil.move(advice_file, os.path.join('Archive', 'excel', 'Original', company_folder, timestamp, 'Outright Payment Advice Date.xls'))
    shutil.move(summary_file, os.path.join('Archive', 'excel', 'Original', company_folder, timestamp, 'Outright Summary of Payments Date.xls'))

    # Write merged file to Merged directory
    os.makedirs(os.path.join('Inbound', 'Merged', company_folder), exist_ok=True)
    merged_file = os.path.join('Inbound', 'Merged', company_folder, f'opadosopd_{timestamp}.xlsx')
    merged_df.to_excel(merged_file, index=False)

    # Move merged file to Archive/Original_Merged
    os.makedirs(os.path.join('Archive', 'excel', 'Original_Merged', company_folder, timestamp), exist_ok=True)
    shutil.move(merged_file, os.path.join('Archive', 'excel', 'Original_Merged', company_folder, timestamp, f'opadosopd_{timestamp}.xlsx'))

    # Generate outbound file
    os.makedirs(os.path.join('Inbound', 'Outbound', company_folder), exist_ok=True)
    outbound_file = os.path.join('Inbound', 'Outbound', company_folder, f'opadosopd_{timestamp}.xlsx')
    merged_df.to_excel(outbound_file, index=False)

    # Create copy in Archive/Converted
    os.makedirs(os.path.join('Archive', 'excel', 'Converted', company_folder, timestamp), exist_ok=True)
    shutil.copy(outbound_file, os.path.join('Archive', 'excel', 'Converted', company_folder, timestamp, f'opadosopd_{timestamp}.xlsx'))

    # Move merged file to Outbound directory
    shutil.move(outbound_file, os.path.join('Inbound', 'Outbound', company_folder))

# Main function
def main():
    # Iterate through company folders
    company_folders = ['PPCI', 'AGRI']
    for company_folder in company_folders:
        merge_files(company_folder)

if __name__ == "__main__":
    main()
