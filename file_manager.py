import os

def rename_files(directory, string_var):
    # Check each file in the specified directory
    for filename in os.listdir(directory):
        if string_var in filename:
            # Construct the new filename by replacing the string_var with nothing
            new_filename = filename.replace(string_var, '')
            # Full path for the current file and the new file
            old_file = os.path.join(directory, filename)
            new_file = os.path.join(directory, new_filename)
            # Rename the file
            os.rename(old_file, new_file)
            print(f"Renamed '{filename}' to '{new_filename}'")
