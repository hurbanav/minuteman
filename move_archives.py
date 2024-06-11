import os
import shutil
from datetime import datetime
import locale
import time

def filter_df(df, column_name, value):
    df = df[df[column_name] == value]
    return df

def organize_files(filepath):
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
    except Exception as e:
        print(f"Failed to set locale to Portuguese: {e}")
    
    directory_to_organize = filepath
    base_dir = os.path.join(directory_to_organize, 'Historico')
    
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
    
    def get_year_month_from_modification_date(file_path):
        modification_time = os.path.getmtime(file_path)
        date = datetime.fromtimestamp(modification_time)
        return date.strftime('%Y'), date.strftime('%B')
    
    doc_tracker = {}
    
    for file in os.listdir(directory_to_organize):
        file_path = os.path.join(directory_to_organize, file)
        if os.path.isdir(file_path):  # Skip directories
            continue
    
        year, month = get_year_month_from_modification_date(file_path)
        month_year_key = f"{year}_{month}"
    
        # Move the file to a temporary month directory (will rename later)
        temp_month_dir = os.path.join(base_dir, month_year_key)
        if not os.path.exists(temp_month_dir):
            os.makedirs(temp_month_dir)
        shutil.move(file_path, os.path.join(temp_month_dir, file))
    
        # Continue with naming convention logic only for non-zip files
        if file.lower().endswith('.zip'):
            continue  # Skip zip files for naming convention
    
        # Check if the file name follows the expected pattern (number_*)
        if '_' in file and file.split('_')[0].lstrip('0').isdigit():
            doc_number = int(file.split('_')[0].lstrip('0'))
            # Update doc_tracker with the smallest and largest document numbers
            if month_year_key in doc_tracker:
                doc_tracker[month_year_key]['min'] = min(doc_tracker[month_year_key]['min'], doc_number)
                doc_tracker[month_year_key]['max'] = max(doc_tracker[month_year_key]['max'], doc_number)
            else:
                doc_tracker[month_year_key] = {'min': doc_number, 'max': doc_number}
        else:
            print(f"Skipping file {file} for naming convention as it does not match the expected pattern.")
    
    # Rename month directories according to the first and last document (excluding zip files from naming logic)
    for month_year_key, doc_numbers in doc_tracker.items():
        year, month = month_year_key.split('_')
        # Only rename if 'min' and 'max' values have been set
        if 'min' in doc_numbers and 'max' in doc_numbers:
            new_dir_name = f"{month} ({doc_numbers['min']} - {doc_numbers['max']})"
            original_dir = os.path.join(base_dir, month_year_key)
            new_dir_path = os.path.join(base_dir, year, new_dir_name)
    
            # Create year directory if it doesn't exist
            if not os.path.exists(os.path.join(base_dir, year)):
                os.makedirs(os.path.join(base_dir, year))
    
            # Rename the directory
            os.rename(original_dir, new_dir_path)
    
    print("Files have been organized and folders have been renamed accordingly.")

def organize_files_by_status(filepath, df, status='Uploaded'):
    status = 'Uploaded'
    
    df = filter_df(df, 'Status', 'Uploaded')
    
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
    except Exception as e:
        print(f"Failed to set locale to Portuguese: {e}")
    
    directory_to_organize = filepath
    base_dir = os.path.join(directory_to_organize, 'Historico')
    
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
    
    def get_year_month_from_modification_date(file_path):
        modification_time = os.path.getmtime(file_path)
        date = datetime.fromtimestamp(modification_time)
        return date.strftime('%Y'), date.strftime('%B')
    
    doc_tracker = {}
    
    for file in os.listdir(directory_to_organize):
        # Check if the file is in the DataFrame before processing
        ###############################################DEFINE THE FILEPATH COLUMN NAME HERE###############################################
        def is_zip_or_xml(file_path):
            return file_path.lower().endswith(('.zip', '.xml'))
        
        if not is_zip_or_xml(file):
            continue
        def get_only_lot_numer(filename):
            if filename.lower().endswith('.xml'):
                number_part, _ = filename.split("_", 1)
                actual_number = int(number_part)
                actual_number_str = str(actual_number)
                return actual_number_str
            if filename.lower().endswith('.zip'):
                filename = filename.rsplit(".", 1)[0]
                number_part = filename.split("-")[-1].strip()
                actual_number = int(number_part)
                actual_number_str = str(actual_number)
                return actual_number_str
        
        if get_only_lot_numer(file) not in df['Numero_lote'].astype(str).str.strip().tolist():
            continue  # Skip files not listed in the DataFrame
        ##################################################################################################################################

        
        file_path = os.path.join(directory_to_organize, file)
        if os.path.isdir(file_path):  # Skip directories
            continue
    
        year, month = get_year_month_from_modification_date(file_path)
        month_year_key = f"{year}_{month}"
    
        temp_month_dir = os.path.join(base_dir, month_year_key)
        if not os.path.exists(temp_month_dir):
            os.makedirs(temp_month_dir)
        shutil.move(file_path, os.path.join(temp_month_dir, file))
    
        if file.lower().endswith('.zip'):
            continue  # Skip zip files for naming convention
    
        if '_' in file and file.split('_')[0].lstrip('0').isdigit():
            doc_number = int(file.split('_')[0].lstrip('0'))
            if month_year_key in doc_tracker:
                doc_tracker[month_year_key]['min'] = min(doc_tracker[month_year_key]['min'], doc_number)
                doc_tracker[month_year_key]['max'] = max(doc_tracker[month_year_key]['max'], doc_number)
            else:
                doc_tracker[month_year_key] = {'min': doc_number, 'max': doc_number}
        else:
            print(f"Skipping file {file} due to naming convention mismatch.")
    
        for month_year_key, doc_numbers in doc_tracker.items():
            year, month = month_year_key.split('_')
            if 'min' in doc_numbers and 'max' in doc_numbers:
                new_dir_name = f"{month} ({doc_numbers['min']} - {doc_numbers['max']})"
                original_dir = os.path.join(base_dir, month_year_key)
                new_dir_path = os.path.join(base_dir, year, new_dir_name)
        
                if not os.path.exists(os.path.join(base_dir, year)):
                    os.makedirs(os.path.join(base_dir, year))
        
                os.rename(original_dir, new_dir_path)
                
        # Loop through files in directory and delete if not in DataFrame
        for file in os.listdir(directory_to_organize):
            file_path = os.path.join(directory_to_organize, file)
    
            # Delete zip files not present in the DataFrame
            if file_path not in df['File Path'] and file.lower().endswith('.zip'):
                os.remove(file_path)
                print(f"Deleted file not listed in DataFrame: {file}")
        
        print("Files have been organized and folders have been renamed accordingly.")

def process_files(directory_to_organize, df):
    def robust_move(src, dst, max_retries=5, wait=2):
        for attempt in range(max_retries):
            try:
                shutil.move(src, dst)
                break
            except PermissionError as e:
                if attempt < max_retries - 1:
                    print(f"Attempt {attempt+1} failed with error: {e}. Retrying in {wait} seconds...")
                    time.sleep(wait)
                else:
                    print(f"Final attempt failed. Unable to move {src} to {dst}")
                    raise

    def set_locale_to_portuguese():
        try:
            locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        except Exception as e:
            print(f"Failed to set locale to Portuguese: {e}")

    def create_directory_if_not_exists(directory_path):
        if not os.path.exists(directory_path):
            os.makedirs(directory_path)

    def is_zip_or_xml(file_path):
        return file_path.lower().endswith(('.zip', '.xml'))

    def get_only_lot_number(filename):
        try:
            if filename.lower().endswith('.xml') or filename.lower().endswith('.zip'):
                number_part, _ = filename.split("_", 1) if '_' in filename else filename.rsplit(".", 1)[0].split("-")[-1], None
                if filename.lower().endswith('.xml'):
                    number_part = number_part[0]
                return str(int(number_part.strip()))
        except:
            filename = filename
        return None

    def get_year_month_from_modification_date(file_path):
        modification_time = os.path.getmtime(file_path)
        date = datetime.fromtimestamp(modification_time)
        return date.strftime('%Y'), date.strftime('%B')

    def update_document_tracker(doc_tracker, file, month_year_key):
        lot_number = get_only_lot_number(file)
        if lot_number:
            doc_number = int(lot_number)
            if month_year_key in doc_tracker:
                doc_tracker[month_year_key]['min'] = min(doc_tracker[month_year_key]['min'], doc_number)
                doc_tracker[month_year_key]['max'] = max(doc_tracker[month_year_key]['max'], doc_number)
            else:
                doc_tracker[month_year_key] = {'min': doc_number, 'max': doc_number}

    # Major modification: Merge the logic of considering existing files
    def process_existing_files(target_dir_path, doc_tracker, year, month):
        for existing_file in os.listdir(target_dir_path):
            if is_zip_or_xml(existing_file):
                update_document_tracker(doc_tracker, existing_file, f"{year}_{month}")

    def move_file_to_monthly_directory(file, base_dir):
        file_path = os.path.join(directory_to_organize, file)
        if os.path.isdir(file_path):
            return

        year, month = get_year_month_from_modification_date(file_path)
        target_dir_path = os.path.join(base_dir, year, month)

        if not os.path.exists(target_dir_path):
            create_directory_if_not_exists(target_dir_path)
        else:
            # New logic to include existing files in lot number assessment
            process_existing_files(target_dir_path, doc_tracker, year, month)

        robust_move(file_path, os.path.join(target_dir_path, file))
        return f"{year}_{month}"

    def rename_directories_according_to_doc_tracker(doc_tracker, base_dir):
        for month_year_key, doc_numbers in doc_tracker.items():
            year, month = month_year_key.split('_')
            new_dir_name = f"{month} ({doc_numbers['min']} - {doc_numbers['max']})"

            # Path adjustment to account for the new directory structure
            original_dir_path = os.path.join(base_dir, year, month)
            new_dir_path = os.path.join(base_dir, year, new_dir_name)

            if not os.path.exists(new_dir_path) and os.path.exists(original_dir_path):
                os.rename(original_dir_path, new_dir_path)

    set_locale_to_portuguese()
    base_dir = os.path.join(directory_to_organize, 'Historico')
    create_directory_if_not_exists(base_dir)

    doc_tracker = {}
    for file in os.listdir(directory_to_organize):
        if not is_zip_or_xml(file):
            continue
        month_year_key = move_file_to_monthly_directory(file, base_dir)
        if month_year_key:
            update_document_tracker(doc_tracker, file, month_year_key)

    rename_directories_according_to_doc_tracker(doc_tracker, base_dir)
        