import os
import pandas as pd
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory, render_template

app = Flask(__name__, static_url_path='', static_folder='.')

# Folder where Excel files are stored
EXCEL_FILES_FOLDER = 'excel_files'

# Function to parse DOB in multiple formats
def parse_date(dob):
    formats = ["%d-%m-%Y", "%d/%m/%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(dob, fmt).date()
        except ValueError:
            continue
    return None

# Function to clean up and normalize column names
def clean_column_names(df):
    df.columns = df.columns.str.strip().str.lower()
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9 ]', '', regex=True)
    return df

# Function to find a person by Date of Birth
def find_person_by_dob(df, dob, filename):
    dob = parse_date(dob)
    if not dob:
        print(f"Invalid DOB format: {dob}")
        return None

    df = clean_column_names(df)
    dob_column = next((col for col in df.columns if 'dob' in col), None)
    name_column = next((col for col in df.columns if 'name' in col), None)
    srno_column = next((col for col in df.columns if 'srno' in col), None)

    if not dob_column or not name_column or not srno_column:
        print(f"Required columns not found in file {filename}. Skipping file.")
        return None

    df['dob'] = pd.to_datetime(df[dob_column], errors='coerce', dayfirst=True).dt.date
    matched_rows = df[df['dob'] == dob]

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
                df = pd.read_excel(filepath)
                df = clean_column_names(df)
                if any('dob' in col for col in df.columns) and any('name' in col for col in df.columns) and any('srno' in col for col in df.columns):
                    excel_files.append((filename, df))
            except Exception as e:
                print(f"Error reading file {filename}: {e}")
    return excel_files

# Serve the index.html file from the templates folder
@app.route('/')
def index():
    return render_template('index.html')  # Flask will look for 'index.html' in the 'templates' folder

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
            return jsonify({"match": result[0]})
    return jsonify({"match": None, "error": "No matching data found. Please check the DOB and try again."})

# Serve CSS file directly
@app.route('/styles.css')
def serve_css():
    return send_from_directory('.', 'styles.css')

if __name__ == "__main__":
    app.run(debug=True)
