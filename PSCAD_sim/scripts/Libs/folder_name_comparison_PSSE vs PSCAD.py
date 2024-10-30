import os

# Function to get folder names in a specified directory
def get_folder_names(directory_path):
    # List all folders in the given directory
    folder_names = [folder for folder in os.listdir(directory_path) if os.path.isdir(os.path.join(directory_path, folder))]
    return set(folder_names)  # Convert to set for easy comparison

# Function to compare folders in two directories
def compare_folders(dir1, dir2):
    # Get folder names in both directories
    folders_in_dir1 = get_folder_names(dir1)
    folders_in_dir2 = get_folder_names(dir2)
    
    # Determine folders in dir1 not in dir2 and vice versa
    only_in_dir1 = folders_in_dir1 - folders_in_dir2
    only_in_dir2 = folders_in_dir2 - folders_in_dir1
    
    return sorted(only_in_dir1), sorted(only_in_dir2)

# Main function to print the comparison result
def main(dir1, dir2):
    only_in_dir1, only_in_dir2 = compare_folders(dir1, dir2)
    
    if only_in_dir1:
        print("Folders in Directory 1 but not in Directory 2:")
        for folder in only_in_dir1:
            print(folder)
    else:
        print("All folders in Directory 1 are present in Directory 2.")
    
    if only_in_dir2:
        print("\nFolders in Directory 2 but not in Directory 1:")
        for folder in only_in_dir2:
            print(folder)
    else:
        print("All folders in Directory 2 are present in Directory 1.")

# Example usage:
# Replace these paths with the actual paths of the directories you want to compare
dir1 = r'C:\GitHub\SF_BESS_Horsham\PSCAD_sim\result_data\dynamic_smib\20241025-1619_S5255Iq1'
dir2 = r'C:\GitHub\SF_BESS_Horsham\PSSE_sim\result_data\dynamic_smib\20241025-1801_S5255Iq1'

main(dir1, dir2)
