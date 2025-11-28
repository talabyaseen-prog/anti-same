from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import os
import zipfile
import tempfile
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Store for created folder structures (in a real app, you might use a database or persistent storage)
folder_storage = {}

@app.route('/')
def index():
    return send_file('index.html')

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    try:
        # Read the Excel file - specify engine explicitly based on file extension
        if file.filename.lower().endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl')
        elif file.filename.lower().endswith('.xls'):
            df = pd.read_excel(file, engine='xlrd')
        else:
            # Try to read without specifying engine if extension doesn't match
            df = pd.read_excel(file)
        
        # Get the second column (index 1), which is "Learner Name (Edexcel Online) "
        # Column names might have trailing spaces, so we check the actual columns
        column_b = df.columns[1]
        student_names = df[column_b].dropna().tolist()  # Get column B and remove NaN values
        
        return jsonify({
            'student_names': student_names,
            'count': len(student_names)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/create_folders', methods=['POST'])
def create_folders():
    data = request.json
    unit_title = data.get('unit_title')
    student_names = data.get('student_names', [])
    
    if not unit_title:
        return jsonify({'error': 'Unit title is required'}), 400
    
    # Create a temporary directory to hold the folder structure
    temp_dir = tempfile.mkdtemp()
    unit_folder_path = os.path.join(temp_dir, unit_title)
    os.makedirs(unit_folder_path, exist_ok=True)
    
    # Create a folder for each student
    for student_name in student_names:
        # Sanitize the student name to be a valid folder name
        safe_student_name = "".join(c for c in student_name if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
        if not safe_student_name:
            safe_student_name = "unnamed_student"  # fallback for empty names after sanitization
        
        student_folder_path = os.path.join(unit_folder_path, safe_student_name)
        os.makedirs(student_folder_path, exist_ok=True)
        
        # Create Learner Work and Assignment Files subfolders
        learner_work_path = os.path.join(student_folder_path, 'Learner Work')
        assignment_files_path = os.path.join(student_folder_path, 'Assignment Files')
        os.makedirs(learner_work_path, exist_ok=True)
        os.makedirs(assignment_files_path, exist_ok=True)
    
    # Store the path temporarily with a unique identifier
    import uuid
    folder_id = str(uuid.uuid4())
    zip_path = os.path.join(temp_dir, f"{unit_title}.zip")
    
    # Create a zip file of the folder structure
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(unit_folder_path):
            # Add directory to zip if it's empty
            if not files and not dirs:
                arcname = os.path.relpath(root, os.path.dirname(unit_folder_path))
                zipf.writestr(arcname + '/', '')  # Add empty directory
            
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, os.path.dirname(unit_folder_path))
                zipf.write(file_path, arcname)
    
    # Store the zip path in our temporary storage
    folder_storage[folder_id] = {
        'path': zip_path,
        'name': f"{unit_title}.zip"
    }
    
    # Return success message with the folder ID
    return jsonify({'success': True, 'folder_id': folder_id, 'message': f'لقد تم انشاء ملفات التقييم ل {len(student_names)} طالب'})

@app.route('/download_folders/<folder_id>', methods=['GET'])
def download_folders(folder_id):
    if folder_id not in folder_storage:
        return jsonify({'error': 'Folder not found'}), 404
    
    storage_info = folder_storage[folder_id]
    zip_path = storage_info['path']
    zip_name = storage_info['name']
    
    return send_file(zip_path, as_attachment=True, download_name=zip_name)

if __name__ == '__main__':
    app.run(debug=True, port=5000)