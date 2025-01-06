import os
import pandas as pd
from datetime import datetime
from flask import Flask, render_template_string, send_from_directory, request, jsonify

app = Flask(__name__)

# Folder where Excel files are stored
EXCEL_FILES_FOLDER = 'excel_files'

# Function to parse DOB in multiple formats
def parse_date(dob):
    formats = ["%d-%m-%Y", "%d/%m/%Y"]  # Include both hyphen and slash formats
    for fmt in formats:
        try:
            return datetime.strptime(dob, fmt).date()
        except ValueError:
            continue
    return None

# Function to clean up and normalize column names
def clean_column_names(df):
    # Strip whitespace, convert to lowercase, and replace problematic characters
    df.columns = df.columns.str.strip().str.lower()
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9 ]', '', regex=True)
    return df

# Function to find a person by Date of Birth
def find_person_by_dob(df, dob, filename):
    dob = parse_date(dob)
    if not dob:
        print(f"Invalid DOB format: {dob}")
        return None

    # Normalize column names
    df = clean_column_names(df)

    # Find matching columns (lenient matching)
    dob_column = next((col for col in df.columns if 'dob' in col), None)
    name_column = next((col for col in df.columns if 'name' in col), None)
    srno_column = next((col for col in df.columns if 'srno' in col), None)

    if not dob_column or not name_column or not srno_column:
        print(f"Required columns not found in file {filename}. Skipping file.")
        return None

    # Normalize the 'DOB' column
    df['dob'] = pd.to_datetime(df[dob_column], errors='coerce', dayfirst=True).dt.date
    print(f"Parsed Dates in file {filename}:\n{df['dob'].head()}")  # Debugging output

    matched_rows = df[df['dob'] == dob]
    print(f"Matching rows for DOB {dob} in file {filename}:\n{matched_rows}")  # Debugging output

    if not matched_rows.empty:
        results = []
        for _, row in matched_rows.iterrows():
            source_file = filename.replace('.xlsx', '').replace('.xls', '')
            results.append({
                "Name": row.get(name_column, 'N/A'),
                "SRNO": row.get(srno_column, 'N/A'),
                "SourceFile": f"{source_file}"
            })
        return results
    else:
        return None

# Function to read all Excel files from the directory
def read_excel_files():
    excel_files = []
    for filename in os.listdir(EXCEL_FILES_FOLDER):
        if filename.endswith(('.xls', '.xlsx')) and not filename.startswith('~$'):
            filepath = os.path.join(EXCEL_FILES_FOLDER, filename)
            try:
                print(f"Reading file: {filename}")  # Debugging output
                df = pd.read_excel(filepath)
                df = clean_column_names(df)  # Clean column names
                print(f"Columns in {filename}: {df.columns}")  # Debugging output
                # Ensure that the necessary columns are present
                if any('dob' in col for col in df.columns) and any('name' in col for col in df.columns) and any('srno' in col for col in df.columns):
                    excel_files.append((filename, df))
                else:
                    print(f"Missing required columns in {filename}. Skipping file.")
            except Exception as e:
                print(f"Error reading file {filename}: {e}")
    return excel_files

# Home route to render the HTML page from file (index.html in root directory)
@app.route('/')
def index():
    with open('index.html', 'r') as f:
        return render_template_string(f.read())

# Serve static files like CSS directly from root directory
@app.route('/static/<path:filename>')
def send_static(filename):
    return send_from_directory('.', filename)

# Search route to process the form and return results
@app.route('/search', methods=['POST'])
def search():
    data = request.get_json()
    dob = data.get('dob', '').strip()

    if not dob:
        return jsonify({"match": None, "error": "Please enter a valid Date of Birth."})

    excel_files = read_excel_files()
    for filename, df in excel_files:
        result = find_person_by_dob(df, dob, filename)
        if result:
            print(f"Match found in file {filename}: {result}")  # Debugging output
            return jsonify({"match": result[0]})
        else:
            print(f"No match found in file {filename}")  # Debugging output

    return jsonify({"match": None, "error": "No matching data found. Please check the DOB and try again."})

if __name__ == "__main__":
    app.run(debug=True)
