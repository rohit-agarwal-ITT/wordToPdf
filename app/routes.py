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
import platform
from app.utils.error_handler import ErrorHandler
from datetime import datetime

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

def format_date_field(value, field_name):
    """
    Format date fields (Date of Joining, Effective Date) to DD-MonthName-YYYY format.
    Handles input formats: '2024-07-01' (ISO) or '6/30/2025' (US format).
    
    Args:
        value: The date value from Excel (can be string or datetime object)
        field_name: The name of the field/column
    
    Returns:
        Formatted date string like '05-August-2025' or original value if not a date field or parsing fails
    """
    # Only format specific date fields
    date_fields = ['Date of Joining', 'Effective Date']
    if field_name not in date_fields:
        return str(value)
    
    # If value is already a datetime object (pandas sometimes reads dates as datetime)
    if isinstance(value, pd.Timestamp):
        try:
            # Check for NaT (Not a Time) - pandas null timestamp
            if pd.isna(value):
                return str(value)
            return value.strftime('%d-%B-%Y')
        except:
            return str(value)
    
    if isinstance(value, datetime):
        try:
            return value.strftime('%d-%B-%Y')
        except:
            return str(value)
    
    # Convert to string for parsing
    date_str = str(value).strip()
    if not date_str or date_str.lower() in ['nan', 'none', '']:
        return str(value)
    
    # Try parsing different date formats
    date_formats = [
        '%Y-%m-%d',      # 2024-07-01
        '%m/%d/%Y',      # 6/30/2025
        '%m-%d-%Y',      # 6-30-2025
        '%d/%m/%Y',      # 01/07/2024 (alternative)
        '%d-%m-%Y',      # 01-07-2024 (alternative)
        '%Y/%m/%d',      # 2024/07/01
    ]
    
    parsed_date = None
    for fmt in date_formats:
        try:
            parsed_date = datetime.strptime(date_str, fmt)
            break
        except ValueError:
            continue
    
    if parsed_date:
        # Format as DD-MonthName-YYYY (e.g., 05-August-2025)
        return parsed_date.strftime('%d-%B-%Y')
    
    # If parsing failed, return original value
    return str(value)

def convert_single_file(file_info):
    """Convert a single file using LibreOffice - optimized for parallel processing"""
    file_path, filename = file_info
    output_dir = os.path.dirname(file_path)
    
    # Replace all hardcoded soffice_path assignments with platform-aware logic
    if platform.system() == "Windows":
        soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
    else:
        soffice_path = 'soffice'
    
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
        return jsonify({'error': 'No files were provided. Please select at least one Word or Excel file to upload.'}), 400
    
    files = request.files.getlist('files[]')
    if not files or files[0].filename == '' or files[0].filename is None:
        return jsonify({'error': 'No files selected. Please choose a file to upload.'}), 400
    
    # Reset progress for new conversion
    reset_progress()
    
    # Excel to Word to PDF batch logic
    if len(files) == 1 and files[0].filename and files[0].filename.lower().endswith('.xlsx'):
        word_template = os.path.join('samples', 'sample_document_for_placeholder.docx')
        temp_dir = tempfile.mkdtemp()
        output_dir = tempfile.mkdtemp()
        pdf_files = []
        errors = []
        try:
            excel_file = files[0]
            if not excel_file.filename:
                return jsonify({'error': 'Uploaded Excel file has no filename. Please re-upload.'}), 400
            excel_path = os.path.join(temp_dir, secure_filename(excel_file.filename or 'uploaded.xlsx'))
            excel_file.save(excel_path)
            try:
                df = pd.read_excel(excel_path)
            except Exception as e:
                return jsonify({'error': f'Failed to read Excel file. Please check your file format.'}), 400
            if df is None or df.empty:
                return jsonify({'error': 'The uploaded Excel file is empty or invalid. Please provide a valid file with data.'}), 400
            total_rows = len(df)
            update_progress(0, total_rows, 'Generating Word documents from Excel. Please wait...')
            
            def generate_docx(row_tuple):
                i, row = row_tuple
                # Process data with date formatting for specific fields
                data = {}
                for col in df.columns:
                    col_str = str(col)
                    value = row[col]
                    # Format date fields appropriately
                    formatted_value = format_date_field(value, col_str)
                    data[col_str] = formatted_value
                docx_name = f"{data.get('Name', 'Candidate')}_{i+1}.docx"
                docx_path = os.path.join(temp_dir, docx_name)
                wp = WordProcessor()  # Create a new instance per row/thread
                wp.fill_placeholders(word_template, docx_path, data)
                return docx_path
            docx_files = []
            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = {executor.submit(generate_docx, (i, row)): i for i, row in df.iterrows()}
                for idx, future in enumerate(as_completed(futures)):
                    docx_path = future.result()
                    docx_files.append(docx_path)
                    update_progress(idx + 1, total_rows, f'Generated {idx + 1} of {total_rows} Word documents. Please wait...')
            update_progress(total_rows, total_rows, 'Converting generated documents to PDF. This may take a moment...')
            
            # Replace all hardcoded soffice_path assignments with platform-aware logic
            if platform.system() == "Windows":
                soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
            else:
                soffice_path = 'soffice'
            try:
                # Validate all file paths before passing to subprocess
                validated_files = []
                for docx_file in docx_files:
                    if os.path.exists(docx_file) and os.path.isfile(docx_file):
                        # Ensure file is within temp directory
                        real_temp_path = os.path.realpath(temp_dir)
                        real_file_path = os.path.realpath(docx_file)
                        if real_file_path.startswith(real_temp_path):
                            validated_files.append(docx_file)
                
                if not validated_files:
                    raise Exception("No valid files found for conversion")
                
                subprocess.run([
                    soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir
                ] + validated_files, check=True)
            except Exception as e:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = f'Failed to convert Word documents to PDF. Please ensure LibreOffice is installed.'
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': conversion_progress['error']}), 500
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
                conversion_progress['error'] = 'Some PDFs could not be generated: ' + '; '.join(errors)
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': conversion_progress['error']}), 500
            update_progress(total_rows, total_rows, 'Creating ZIP file. Almost done!')
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zip_file:
                for pdf_path, pdf_name in pdf_files:
                    with open(pdf_path, 'rb') as pdf_file:
                        zip_file.writestr(pdf_name, pdf_file.read())
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            zip_buffer.seek(0)
            conversion_progress['status'] = 'completed'
            update_progress(total_rows, total_rows, 'Conversion complete! Your files are ready for download.')
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name='Appointment_letters.zip',
                mimetype='application/zip'
            )
        except Exception as e:
            conversion_progress['status'] = 'error'
            conversion_progress['error'] = f'An unexpected error occurred while processing your Excel file. Please try again.'
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            return jsonify({'error': conversion_progress['error']}), 500
    
    # Handle single file case
    if len(files) == 1:
        file = files[0]
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename or 'uploaded.docx')
            unique_filename = f"{uuid.uuid4().hex}_{filename}"
            input_path = os.path.join(current_app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(input_path)
            
            update_progress(0, 1, f'Converting {filename}. Please wait...')
            
            # Convert single file
            result = convert_single_file((input_path, filename))
            output_pdf, pdf_name, error = result
            
            if error or output_pdf is None:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = f'Failed to convert file: {error or "Unknown error"}. Please check your file and try again.'
                os.remove(input_path)
                return jsonify({'error': conversion_progress['error']}), 500
            
            update_progress(1, 1, 'Preparing download. Almost done!')
            
            # Read PDF and clean up
            with open(output_pdf, 'rb') as f:
                pdf_data = f.read()
            os.remove(input_path)
            os.remove(output_pdf)
            
            conversion_progress['status'] = 'completed'
            update_progress(1, 1, 'Conversion complete! Your file is ready for download.')
            
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
                    filename = secure_filename(file.filename or f'file_{i}.docx')
                    file_path = os.path.join(temp_dir, filename)
                    file.save(file_path)
                    update_progress(i + 1, total_files, f'Prepared {i + 1} of {total_files} files. Please wait...')
                else:
                    conversion_progress['status'] = 'error'
                    conversion_progress['error'] = f'Invalid file type for {file.filename}. Only .docx, .doc, and .xlsx files are allowed.'
                    shutil.rmtree(temp_dir)
                    shutil.rmtree(output_dir)
                    return jsonify({'error': conversion_progress['error']}), 400
            
            update_progress(total_files, total_files, 'Converting files to PDF. This may take a moment...')
            
            # Batch convert all files in one soffice call
            
            docx_files = [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(('.docx', '.doc'))]
            
            if not docx_files:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = 'No valid DOCX/DOC files found. Please upload valid Word documents.'
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': conversion_progress['error']}), 400
            
            # Replace all hardcoded soffice_path assignments with platform-aware logic
            if platform.system() == "Windows":
                soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
            else:
                soffice_path = 'soffice'
            try:
                # Validate all file paths before passing to subprocess
                validated_files = []
                for docx_file in docx_files:
                    if os.path.exists(docx_file) and os.path.isfile(docx_file):
                        # Ensure file is within temp directory
                        real_temp_path = os.path.realpath(temp_dir)
                        real_file_path = os.path.realpath(docx_file)
                        if real_file_path.startswith(real_temp_path):
                            validated_files.append(docx_file)
                
                if not validated_files:
                    raise Exception("No valid files found for conversion")
                
                subprocess.run([
                    soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir
                ] + validated_files, check=True)
            except Exception as e:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = f'Failed to convert files to PDF. Please ensure LibreOffice is installed.'
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': conversion_progress['error']}), 500
            
            update_progress(total_files, total_files, 'Collecting converted PDFs...')
            
            # Collect PDFs and rename for zipping, updating progress as each PDF is found (Excel case)
            pdfs_found = 0
            for idx, docx_file in enumerate(docx_files):
                base = os.path.splitext(os.path.basename(docx_file))[0]
                pdf_name = f"{base}-Appointment_letter.pdf"
                pdf_path = os.path.join(output_dir, base + '.pdf')
                if os.path.exists(pdf_path):
                    pdf_files.append((pdf_path, pdf_name))
                    pdfs_found += 1
                    update_progress(pdfs_found, total_files, f'Converted {pdfs_found} of {total_files} files to PDF...')
                else:
                    errors.append(f'PDF not found for {base}')
            
            if errors:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = 'Some PDFs could not be generated: ' + '; '.join(errors)
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': conversion_progress['error']}), 500
            
            # After all PDFs are collected and before sending the response (Excel case)
            conversion_progress['status'] = 'completed'
            update_progress(total_files, total_files, 'Conversion complete! Your files are ready for download.')
            
            # Zip all PDFs
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zip_file:
                for pdf_path, pdf_name in pdf_files:
                    with open(pdf_path, 'rb') as pdf_file:
                        zip_file.writestr(pdf_name, pdf_file.read())
            
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            zip_buffer.seek(0)
            
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name='Appointment_letters.zip',
                mimetype='application/zip'
            )
        except Exception as e:
            conversion_progress['status'] = 'error'
            conversion_progress['error'] = f'An unexpected error occurred while processing your files. Please try again.'
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            return jsonify({'error': conversion_progress['error']}), 500

@main.route('/download/<filename>')
def download_file(filename):
    try:
        # Validate filename to prevent path traversal
        if not filename or '..' in filename or '/' in filename or '\\' in filename:
            return jsonify({'error': 'Invalid filename'}), 400
        
        # Sanitize filename
        safe_filename = secure_filename(filename)
        if not safe_filename:
            return jsonify({'error': 'Invalid filename'}), 400
        
        file_path = os.path.join(current_app.config['DOWNLOAD_FOLDER'], safe_filename)
        
        # Additional path validation
        real_download_path = os.path.realpath(current_app.config['DOWNLOAD_FOLDER'])
        real_file_path = os.path.realpath(file_path)
        
        if not real_file_path.startswith(real_download_path):
            return jsonify({'error': 'Access denied'}), 403
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': 'File access error'}), 500 

def register_error_handlers(app):
    """Register error handlers with the Flask app"""
    
    @app.errorhandler(413)
    def too_large(e):
        """Handle file too large error"""
        return ErrorHandler.create_error_response(
            Exception("File too large"), 
            context='upload'
        )

    @app.errorhandler(500)
    def internal_error(e):
        """Handle internal server errors"""
        return ErrorHandler.create_error_response(
            Exception("Internal server error"), 
            context='system'
        )

    @app.errorhandler(404)
    def not_found(e):
        """Handle 404 errors"""
        return ErrorHandler.create_error_response(
            Exception("Page not found"), 
            context='system'
        ) 