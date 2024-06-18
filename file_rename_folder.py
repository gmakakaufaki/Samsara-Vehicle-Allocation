import os
import shutil

# Function to rename a file and remove a part of the string in the filename
def rename_file_with_long_path(original_path, part_to_remove):
    # Enable long path support
    long_path_prefix = "\\\\?\\"
    if not original_path.startswith(long_path_prefix):
        original_path = long_path_prefix + original_path

    # Get the directory and the filename
    directory, filename = os.path.split(original_path)
    
    # Remove the specified part from the filename
    new_filename = filename.replace(part_to_remove, "")
    
    # Construct the new file path
    new_path = os.path.join(directory, new_filename)
    
    # Rename the file
    shutil.move(original_path, new_path)

# Function to rename all files in a directory
def rename_files_in_directory(directory_path, part_to_remove):
    # Enable long path support
    long_path_prefix = "\\\\?\\"
    if not directory_path.startswith(long_path_prefix):
        directory_path = long_path_prefix + directory_path
    
    # Iterate over all files in the directory
    for filename in os.listdir(directory_path):
        if part_to_remove in filename:
            original_file_path = os.path.join(directory_path, filename)
            rename_file_with_long_path(original_file_path, part_to_remove)

# Example usage
directory_path = r"C:\Users\gmakakaufaki\Sukut\Accounting - GL\01 General Ledger\01 SCLLC\Month-End Tasks\01 Vehicle Job Allocation\2024-05\Samsara Data"
part_to_remove = "Time on Site Report - "

rename_files_in_directory(directory_path, part_to_remove)
