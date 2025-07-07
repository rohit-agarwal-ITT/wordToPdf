from flask import Blueprint, render_template, request, jsonify, send_file, current_app
import os
import uuid
from werkzeug.utils import secure_filename
from app.utils.word_processor import WordProcessor
from app.utils.pdf_generator import PDFGenerator
from docx2pdf import convert
import io
import subprocess
import zipfile
import tempfile
import shutil
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

main = Blueprint('main', __name__)

ALLOWED_EXTENSIONS = {'docx', 'doc', 'xlsx'}

# Global progress tracking
conversion_progress = {
    'status': 'idle',  # idle, converting, completed, error
    'current': 0,
    'total': 0,
    'message': '',
    'error': None
}

def reset_progress():
    global conversion_progress
    conversion_progress = {
        'status': 'idle',
        'current': 0,
        'total': 0,
        'message': '',
        'error': None
    }

def update_progress(current, total, message):
    global conversion_progress
    conversion_progress['current'] = current
    conversion_progress['total'] = total
    conversion_progress['message'] = message
    conversion_progress['status'] = 'converting'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_single_file(file_info):
    """Convert a single file using LibreOffice - optimized for parallel processing"""
    file_path, filename = file_info
    output_dir = os.path.dirname(file_path)
    soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
    
    try:
        # Convert using LibreOffice
        result = subprocess.run([
            soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, file_path
        ], check=True, capture_output=True, timeout=60)
        
        output_pdf = file_path.rsplit('.', 1)[0] + '.pdf'
        
        # Extract name from filename for PDF naming
        name_part = filename.rsplit('.', 1)[0]
        pdf_name = f"{name_part}-Appointment_letter.pdf"
        
        return (output_pdf, pdf_name, None)
    except subprocess.TimeoutExpired:
        return (None, None, f"Conversion timeout for {filename}")
    except Exception as e:
        return (None, None, f"Conversion failed for {filename}: {str(e)}")

@main.route('/')
def index():
    return render_template('index.html')

@main.route('/progress')
def get_progress():
    """Return current conversion progress"""
    global conversion_progress
    return jsonify(conversion_progress)

@main.route('/upload', methods=['POST'])
def upload_file():
    global conversion_progress
    
    if 'files[]' not in request.files:
        return jsonify({'error': 'No files provided'}), 400
    
    files = request.files.getlist('files[]')
    if not files or files[0].filename == '':
        return jsonify({'error': 'No files selected'}), 400
    
    # Reset progress for new conversion
    reset_progress()
    
    # Excel to Word to PDF batch logic
    if len(files) == 1 and files[0].filename.lower().endswith('.xlsx'):
        word_template = os.path.join('samples', 'sample_word_document.docx')
        temp_dir = tempfile.mkdtemp()
        output_dir = tempfile.mkdtemp()
        pdf_files = []
        errors = []
        try:
            excel_file = files[0]
            excel_path = os.path.join(temp_dir, secure_filename(excel_file.filename))
            excel_file.save(excel_path)
            df = pd.read_excel(excel_path)
            total_rows = len(df)
            update_progress(0, total_rows, 'Generating Word documents from Excel...')
            wp = WordProcessor()
            def generate_docx(row_tuple):
                i, row = row_tuple
                data = {str(col): str(row[col]) for col in df.columns}
                docx_name = f"{data.get('Name', 'Candidate')}_{i+1}.docx"
                docx_path = os.path.join(temp_dir, docx_name)
                wp.fill_placeholders(word_template, docx_path, data)
                return docx_path
            docx_files = []
            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = {executor.submit(generate_docx, (i, row)): i for i, row in df.iterrows()}
                for idx, future in enumerate(as_completed(futures)):
                    docx_path = future.result()
                    docx_files.append(docx_path)
                    update_progress(idx + 1, total_rows, f'Generated {idx + 1}/{total_rows} Word docs...')
            update_progress(total_rows, total_rows, 'Converting generated docs to PDF...')
            soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
            try:
                subprocess.run([
                    soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir
                ] + docx_files, check=True)
            except Exception as e:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = f'LibreOffice batch conversion failed: {e}'
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': f'LibreOffice batch conversion failed: {e}'}), 500
            update_progress(total_rows, total_rows, 'Collecting converted PDFs...')
            for docx_file in docx_files:
                base = os.path.splitext(os.path.basename(docx_file))[0]
                pdf_name = f"{base}-Appointment_letter.pdf"
                pdf_path = os.path.join(output_dir, base + '.pdf')
                if os.path.exists(pdf_path):
                    pdf_files.append((pdf_path, pdf_name))
                else:
                    errors.append(f'PDF not found for {base}')
            if errors:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = '; '.join(errors)
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': '; '.join(errors)}), 500
            update_progress(total_rows, total_rows, 'Creating ZIP file...')
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zip_file:
                for pdf_path, pdf_name in pdf_files:
                    with open(pdf_path, 'rb') as pdf_file:
                        zip_file.writestr(pdf_name, pdf_file.read())
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            zip_buffer.seek(0)
            conversion_progress['status'] = 'completed'
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name='Appointment_letters.zip',
                mimetype='application/zip'
            )
        except Exception as e:
            conversion_progress['status'] = 'error'
            conversion_progress['error'] = f'Error processing Excel: {e}'
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            return jsonify({'error': f'Error processing Excel: {e}'}), 500
    
    # Handle single file case
    if len(files) == 1:
        file = files[0]
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4().hex}_{filename}"
            input_path = os.path.join(current_app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(input_path)
            
            update_progress(0, 1, f'Converting {filename}...')
            
            # Convert single file
            result = convert_single_file((input_path, filename))
            output_pdf, pdf_name, error = result
            
            if error or output_pdf is None:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = error or 'Conversion failed'
                os.remove(input_path)
                return jsonify({'error': error or 'Conversion failed'}), 500
            
            update_progress(1, 1, 'Preparing download...')
            
            # Read PDF and clean up
            with open(output_pdf, 'rb') as f:
                pdf_data = f.read()
            os.remove(input_path)
            os.remove(output_pdf)
            
            conversion_progress['status'] = 'completed'
            
            return send_file(
                io.BytesIO(pdf_data),
                as_attachment=True,
                download_name=pdf_name,
                mimetype='application/pdf'
            )
        else:
            return jsonify({'error': 'Invalid file type. Only .docx, .doc, and .xlsx files are allowed.'}), 400
    
    # Handle multiple files case - BATCH CONVERSION
    else:
        temp_dir = tempfile.mkdtemp()
        output_dir = tempfile.mkdtemp()
        pdf_files = []
        errors = []
        
        try:
            total_files = len(files)
            update_progress(0, total_files, 'Preparing files for conversion...')
            
            # Save all files to temp_dir
            for i, file in enumerate(files):
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(temp_dir, filename)
                    file.save(file_path)
                    update_progress(i + 1, total_files, f'Prepared {i + 1}/{total_files} files...')
                else:
                    conversion_progress['status'] = 'error'
                    conversion_progress['error'] = f'Invalid file type for {file.filename}'
                    shutil.rmtree(temp_dir)
                    shutil.rmtree(output_dir)
                    return jsonify({'error': f'Invalid file type for {file.filename}. Only .docx, .doc, and .xlsx files are allowed.'}), 400
            
            update_progress(total_files, total_files, 'Converting files with LibreOffice...')
            
            # Batch convert all files in one soffice call
            soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
            docx_files = [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(('.docx', '.doc'))]
            
            if not docx_files:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = 'No valid DOCX/DOC files found'
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': 'No valid DOCX/DOC files found.'}), 400
            
            try:
                subprocess.run([
                    soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir
                ] + docx_files, check=True)
            except Exception as e:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = f'LibreOffice batch conversion failed: {e}'
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': f'LibreOffice batch conversion failed: {e}'}), 500
            
            update_progress(total_files, total_files, 'Collecting converted PDFs...')
            
            # Collect PDFs and rename for zipping
            for docx_file in docx_files:
                base = os.path.splitext(os.path.basename(docx_file))[0]
                pdf_name = f"{base}-Appointment_letter.pdf"
                pdf_path = os.path.join(output_dir, base + '.pdf')
                if os.path.exists(pdf_path):
                    pdf_files.append((pdf_path, pdf_name))
                else:
                    errors.append(f'PDF not found for {base}')
            
            if errors:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = '; '.join(errors)
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': '; '.join(errors)}), 500
            
            update_progress(total_files, total_files, 'Creating ZIP file...')
            
            # Zip all PDFs
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zip_file:
                for pdf_path, pdf_name in pdf_files:
                    with open(pdf_path, 'rb') as pdf_file:
                        zip_file.writestr(pdf_name, pdf_file.read())
            
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            zip_buffer.seek(0)
            
            conversion_progress['status'] = 'completed'
            
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name='Appointment_letters.zip',
                mimetype='application/zip'
            )
        except Exception as e:
            conversion_progress['status'] = 'error'
            conversion_progress['error'] = f'Error processing files: {e}'
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            return jsonify({'error': f'Error processing files: {e}'}), 500

@main.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(current_app.config['DOWNLOAD_FOLDER'], filename)
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': 'File not found'}), 404 