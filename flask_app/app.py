from flask import Flask, render_template, request
import pandas as pd

class ExcelComparator:
    def __init__(self, file1, file2):
        self.file1 = file1
        self.file2 = file2

    def compare_files(self):
        try:
            data1 = pd.read_excel(self.file1)
            data2 = pd.read_excel(self.file2)
        except Exception as e:
            return f"Error reading XLSX file: {e}"

        # Ensure the same number of rows
        min_rows = min(len(data1), len(data2))
        data1 = data1.iloc[:min_rows]
        data2 = data2.iloc[:min_rows]

        # Compare each cell
        mismatches = []
        for i in range(min_rows):
            for j in range(min(len(data1.columns), len(data2.columns))):
                try:
                    if str(data1.iloc[i, j]).strip() != str(data2.iloc[i, j]).strip():
                        # If a mismatch is found beyond the first row, identify the field where the mismatch occurred
                        field = data1.columns[j]
                        mismatches.append((i+1, field, data1.iloc[i, j], data2.iloc[i, j]))
                except Exception as e:
                    # If an error occurs, identify the field from the first row
                    field = data1.columns[j]
                    mismatches.append((i+1, field, "Error", "Error"))

        return mismatches

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/compare', methods=['POST'])
def compare():
    file1 = request.files['file1']
    file2 = request.files['file2']
    comparator = ExcelComparator(file1, file2)
    mismatches = comparator.compare_files()
    if mismatches:
        return render_template('output.html', mismatches=mismatches)
    else:
        return "No mismatches found"

if __name__ == '__main__':
    app.run(debug=True)


