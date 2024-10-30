import pandas as pd
import os

# Function to read the Excel file and get the values from Column A
def read_excel_column_a(file_path):
    # Read the Excel file (assumes the first sheet and column A contains the data)
    df = pd.read_excel(file_path, usecols="A", header=None)  # header=None if there is no header row
    return df[0].tolist()  # Convert the column to a list of values

# Function to get the folder names in the specified directory
def get_folder_names(directory_path):
    # List all folders in the given directory that start with 'small'
    folder_names = [folder for folder in os.listdir(directory_path) if os.path.isdir(os.path.join(directory_path, folder)) and folder.startswith('small')]
    return folder_names

# Function to compare the Excel column with the folder names (handling the 'small' prefix)
def find_missing_folders(excel_entries, folder_names):
    # Create the expected folder names by adding the 'small' prefix to each Excel entry
    expected_folders = set(f"tov{int(entry)}" for entry in excel_entries)
    
    # Convert folder_names to a set for comparison
    folder_set = set(folder_names)
    
    # Find the expected folders that are missing in the folder set
    missing_folders = expected_folders - folder_set
    
    # Return the missing folders as a sorted list
    return sorted(missing_folders)

# Main function to tie it all together
def main(excel_file_path, directory_path):
    # Step 1: Read column A of the Excel file
    excel_entries = read_excel_column_a(excel_file_path)
    
    # Step 2: Get all folder names in the specified directory
    folder_names = get_folder_names(directory_path)
    
    # Step 3: Compare and find the missing folders
    missing_folders = find_missing_folders(excel_entries, folder_names)
    
    if missing_folders:
        print("The following entries from the Excel file do not have corresponding folders (in ascending order):")
        for folder in missing_folders:
            print(folder)
    else:
        print("All entries in the Excel file have corresponding folders.")

# Example usage:
# Replace these paths with your actual Excel file path and folder path


excel_file_path = r'C:\GitHub\SF_BESS_Horsham/sim_list.xlsx'
directory_path = r'C:\GitHub\SF_BESS_Horsham\PSSE_sim\result_data\dynamic_smib\20241025-1801_S5255Iq1'

main(excel_file_path, directory_path)
