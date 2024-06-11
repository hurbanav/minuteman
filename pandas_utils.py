import pandas as pd
import os


class PandasUtils:
    def __init__(self):
        pass

    @staticmethod
    def get_or_create_excel(file_path, sheet_name='Sheet1'):
        """
        Check if an Excel file exists. If it does, read and return a DataFrame.
        If not, create the Excel file and return an empty DataFrame.

        :param file_path: Path to the Excel file
        :param sheet_name: Name of the sheet to read or create
        :return: DataFrame
        """
        if os.path.exists(file_path):
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                print(f"Excel file '{file_path}' found. Loaded sheet '{sheet_name}' into DataFrame.")
            except Exception as e:
                print(f"Error reading the Excel file: {e}")
                df = pd.DataFrame()
        else:
            df = pd.DataFrame()
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Excel file '{file_path}' not found. Created new file with an empty sheet '{sheet_name}'.")

        return df
