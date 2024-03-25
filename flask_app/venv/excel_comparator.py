# Author @ Gabriele Malatesta

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Declaring class structure


class ExcelComparator:
    # First method: connection to open and initialize sheets
    def __init__(self, file1, file2, sheet_name):
        self.file1 = file1
        self.file2 = file2
        self.sheet_name = sheet_name

    # where the compare files arguments are hosted how data is read
    def compare_files(self):
        try:
            data_1 = pd.read_excel(
                self.file1, sheet_name=self.sheet_name, header=2)
        except ValueError as e:
            return f"Error in file1: {e}"
        try:
            data_2 = pd.read_excel(
                self.file2, sheet_name=self.sheet_name, header=2)
        except ValueError as e:
            return f"Error in file2: {e}"

        # Preparing to include column letters for mismatches
        wb = load_workbook(filename=self.file1, read_only=True)
        ws = wb[self.sheet_name]
        # setting range of column letters in order to print them later
        column_letters = {
            i+1: get_column_letter(i+1) for i in range(ws.max_column)}
        wb.close()

        mismatches = []  # hold mismatches

        for index, row_1 in data_1.iterrows():  # loop that iterates through rows
            if index >= len(data_2):
                break
            row_2 = data_2.iloc[index]
            for col_index, col_name in enumerate(data_1.columns, start=1):
                value_1 = row_1[col_name]
                value_2 = row_2[col_name]
                if pd.isna(value_1) and pd.isna(value_2):
                    continue
                if value_1 != value_2:  # creating not equals rule to append mismatches to a list
                    mismatches.append({
                        'Sheet': self.sheet_name,
                        # 'Column Name': col_name,
                        # Reflects good column letter
                        'Column Letter': column_letters[col_index],
                        'Row': index + 4,  # Adjusted for header +1-based indexing
                        'Value 1': value_1,
                        'Value 2': value_2
                    })
        return mismatches  # returens mismatches from the list
