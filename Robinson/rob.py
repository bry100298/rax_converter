import os
import shutil
import pandas as pd
import time

# Function to merge and update Excel files
def merge_and_update_excel(parent_dir):
    inbound_dir = os.path.join(parent_dir, 'Inbound')
    archive_dir = os.path.join(parent_dir, 'Archive')

    for company_folder in os.listdir(inbound_dir):
        company_dir = os.path.join(inbound_dir, company_folder)
        if os.path.isdir(company_dir):
            # Paths for original files
            summary_file = os.path.join(company_dir, 'Outright Summary of Payments Date.xlsx')
            advice_file = os.path.join(company_dir, 'Outright Payment Advice Date.xlsx')

            # Check if both files exist
            if os.path.exists(summary_file) and os.path.exists(advice_file):
                # Read Excel files into DataFrames
                df_summary = pd.read_excel(summary_file)
                df_advice = pd.read_excel(advice_file)

                # Merge DataFrames
                merged_df = pd.merge(df_summary, df_advice, how='outer')

                # Generate timestamp for filename
                timestamp = int(time.time())

                # Save merged DataFrame to a new Excel file
                merged_folder = os.path.join(inbound_dir, 'Merged', company_folder)
                os.makedirs(merged_folder, exist_ok=True)
                merged_file = os.path.join(merged_folder, f'opadosopd_{timestamp}.xlsx')
                merged_df.to_excel(merged_file, index=False)

                # Move original files to Archive/Original
                archive_original_folder = os.path.join(archive_dir, 'excel', 'Original', company_folder, str(timestamp))
                os.makedirs(archive_original_folder, exist_ok=True)
                shutil.move(summary_file, archive_original_folder)
                shutil.move(advice_file, archive_original_folder)

                # Add additional columns and update merged Excel file
                update_excel(merged_file, company_folder)

                # Move updated merged file to Outbound
                outbound_folder = os.path.join(inbound_dir, 'Outbound', company_folder)
                os.makedirs(outbound_folder, exist_ok=True)
                shutil.move(merged_file, outbound_folder)

                # Move previous version of updated merged file to Original_Merged in Archive
                archive_merged_folder = os.path.join(archive_dir, 'excel', 'Original_Merged', company_folder, str(timestamp))
                os.makedirs(archive_merged_folder, exist_ok=True)
                shutil.move(merged_file, archive_merged_folder)

                # Backup generated file to Converted in Archive
                archive_converted_folder = os.path.join(archive_dir, 'excel', 'Converted', company_folder, str(timestamp))
                os.makedirs(archive_converted_folder, exist_ok=True)
                shutil.copy(merged_file, archive_converted_folder)
            else:
                print(f"Error: Required files not found in {company_folder}")

# Function to add additional columns and update Excel file
def update_excel(excel_file, company_folder):
    # Read Excel file into DataFrame
    df = pd.read_excel(excel_file)

    # Add additional columns based on instructions
    # (Modify this part according to the specific columns and instructions)
    df['EDI_Customer'] = company_folder
    df['New_Column'] = df['Existing_Column']  # Example of adding a new column

    # Save updated DataFrame to Excel
    df.to_excel(excel_file, index=False)

# Main function
def main():
    parent_dir = 'Robinson'  # Replace 'Robinson' with the parent directory name
    merge_and_update_excel(parent_dir)

# Execute the main function
if __name__ == "__main__":
    main()
