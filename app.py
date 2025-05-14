from flask import Flask, request, send_file, render_template, redirect, url_for
import os
from werkzeug.utils import secure_filename
import zipfile
import io
import pandas as pd

from file_processor import process_file_with_columns


app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create folders if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Allowed file types
ALLOWED_EXTENSIONS = {'.csv', '.xls', '.xlsx'}

def allowed_file(filename):
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('welcome.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'GET':
        return render_template('upload.html')
        
    uploaded_files = request.files.getlist('files')
    saved_filenames = []

    # for file in uploaded_files:
    #     if file and allowed_file(file.filename):
    #         filename = secure_filename(file.filename)
    #         input_path = os.path.join(UPLOAD_FOLDER, filename)
    #         output_filename = f"{os.path.splitext(filename)[0]}_filtered.xlsx"
    #         output_path = os.path.join(OUTPUT_FOLDER, output_filename)
    #         file.save(input_path)
    #         process_file(input_path, output_path)
    #         filenames.append(output_filename)
    #     else:
    #         return f"Error: {file.filename} has an unsupported file type.", 400

    # return redirect(url_for('result', filenames=",".join(filenames)))
    for file in uploaded_files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            input_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(input_path)
            saved_filenames.append(filename)

    if not saved_filenames:
        return "No valid files uploaded.", 400

    # Show column selection using the first file
    first_file_path = os.path.join(UPLOAD_FOLDER, saved_filenames[0])
    ext = os.path.splitext(first_file_path)[1].lower()
    if ext == '.csv':
        df = pd.read_csv(first_file_path, nrows=1)
    else:
        df = pd.read_excel(first_file_path, engine='openpyxl', nrows=1)

    columns = df.columns.tolist()
    return render_template('select_columns.html', columns=columns, filenames=saved_filenames)
    
@app.route('/process-columns', methods=['POST'])
def process_columns():
    selected_columns = request.form.getlist('selected_columns')
    filenames = request.form.get('filenames').split(',')
    output_files = []

    for filename in filenames:
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        output_filename = f"{os.path.splitext(filename)[0]}_filtered.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        try:
            process_file_with_columns(input_path, output_path, selected_columns)
            output_files.append(output_filename)
        except Exception as e:
            print(f"Failed to process {filename}: {e}")

    return redirect(url_for('result', filenames=",".join(output_files)))

@app.route('/result')
def result():
    filenames = request.args.get('filenames', '').split(',')
    return render_template('result.html', filenames=filenames)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

@app.route('/download-all')
def download_all():
    filenames = request.args.get('filenames', '').split(',')
    zip_stream = io.BytesIO()

    with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in filenames:
            file_path = os.path.join(OUTPUT_FOLDER, fname)
            if os.path.exists(file_path):
                zf.write(file_path, arcname=fname)

    zip_stream.seek(0)
    return send_file(zip_stream, as_attachment=True, download_name='processed_files.zip')


if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0", port=10000)
