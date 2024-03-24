from flask import Flask, request, render_template
from excel_comparator import ExcelComparator

app = Flask(__name__)


@app.route('/', methods=['GET'])
def index():
    # Render the upload page
    return render_template('upload.html')


@app.route('/output', methods=['POST'])
def output():
    # Retrieve the uploaded files and sheet name
    file1 = request.files['file1']
    file2 = request.files['file2']
    sheet_name = request.form.get('sheet_name')

    # Create an instance of ExcelComparator and compare the files
    comparator = ExcelComparator(file1, file2, sheet_name)
    mismatches = comparator.compare_files()

    # Render the output template with the mismatches
    return render_template('output.html', mismatches=mismatches)


if __name__ == '__main__':
    app.run(debug=True)
