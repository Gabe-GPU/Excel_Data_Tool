import pandas as pd
from openpyxl.utils import get_column_letter


class ExcelComparator:
    def __init__(self, file1, file2, sheet_name):
        self.file1 = file1
        self.file2 = file2
        self.sheet_name = sheet_name

    def compare_files(self):
        # Load the data from both files
        data_1 = pd.read_excel(
            self.file1, sheet_name=self.sheet_name, header=2)
        data_2 = pd.read_excel(
            self.file2, sheet_name=self.sheet_name, header=2)

        # Find mismatches
        mismatches = []
        for index, row_1 in data_1.iterrows():
            if index >= len(data_2):
                break  # Stop if there are no more rows in data_2
            row_2 = data_2.iloc[index]
            for col in data_1.columns:
                value_1 = row_1[col]
                value_2 = row_2[col]
                if pd.isna(value_1) and pd.isna(value_2):
                    continue  # Ignore NaN vs NaN
                if value_1 != value_2:
                    mismatches.append({
                        'Sheet': self.sheet_name,
                        'Column': col,
                        'Value 1': value_1,
                        'Value 2': value_2,
                        'Row': index + 3  # Adjust the row index to match Excel's 1-based indexing
                    })
        return mismatches
