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

def list_files(directory, extension=None):
    """
    Lists all files in a given directory.

    :param directory: Path to the directory.
    :param extension: Optional. If provided, filters files to only those with the given extension.
        ex: '.png' or '.xlsx'
    :return: A list of full file paths matching the criteria.
    """
    file_paths = []  # List to store the full paths of files.
    # Walk through all files and folders within the directory.
    for root, dirs, files in os.walk(directory):
        for file in files:
            if extension:
                # Check if the file ends with the specified extension
                if file.endswith(extension):
                    file_paths.append(os.path.join(root, file))
            else:
                file_paths.append(os.path.join(root, file))

    return file_paths