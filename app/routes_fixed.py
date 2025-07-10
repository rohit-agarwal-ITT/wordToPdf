from flask import Blueprint, render_template, request, jsonify, send_file, current_app
import os
import uuid
import logging
from werkzeug.utils import secure_filename
from app.utils.word_processor import WordProcessor
from app.utils.pdf_generator import PDFGenerator
from app.utils.validators import FileValidator
from app.utils.error_handler import ErrorHandler
from app.utils.conversion_manager import ConversionManager
import io
import subprocess
import zipfile
import tempfile
import shutil
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import platform
import time

main = Blueprint('main', __name__)

ALLOWED_EXTENSIONS = {'docx', 'doc', 'xlsx'}

# Initialize conversion manager
conversion_manager = ConversionManager()

# Setup logging
logger = logging.getLogger(__name__)

def reset_progress():
    """Reset conversion progress"""
    conversion_manager.reset_progress()

def update_progress(current, total, message):
    """Update conversion progress"""
    conversion_manager.update_progress(current, total, message)

def allowed_file(filename):
    """Check if file has allowed extension"""
    if not filename:
        return False
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_single_file(file_info):
    """Convert a single file using LibreOffice - optimized for parallel processing"""
    file_path, filename = file_info
    output_dir = os.path.dirname(file_path)
    
    # Use conversion manager for better error handling
    output_pdf, pdf_name, error = conversion_manager.convert_single_file(file_path, filename, output_dir)
    return (output_pdf, pdf_name, error)

@main.route('/')
def index():
    """Main page"""
    return render_template('index.html')

@main.route('/progress')
def get_progress():
    """Return current conversion progress"""
    return jsonify(conversion_manager.conversion_progress)

@main.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and conversion"""
    
    # Validate file upload
    if 'files[]' not in request.files:
        return jsonify({'error': 'No files provided'}), 400
    
    files = request.files.getlist('files[]')
    if not files or not files[0] or not files[0].filename:
        return jsonify({'error': 'No files selected'}), 400
    
    # Comprehensive file validation
    is_valid, error_msg, valid_files = FileValidator.validate_file_upload(files)
    if not is_valid:
        return jsonify({'error': error_msg}), 400
    
    # Reset progress for new conversion
    reset_progress()
    
    # Check system requirements
    requirements_ok, requirement_errors = conversion_manager.validate_conversion_requirements()
    if not requirements_ok:
        return jsonify({'error': f"System requirements not met: {'; '.join(requirement_errors)}"}), 500
    
    # Excel to Word to PDF batch logic
    if len(valid_files) == 1 and valid_files[0].filename and valid_files[0].filename.lower().endswith('.xlsx'):
        return _handle_excel_conversion(valid_files[0])
    
    # Handle single file case
    if len(valid_files) == 1:
        return _handle_single_file_conversion(valid_files[0])
    
    # Handle multiple files case - BATCH CONVERSION
    return _handle_batch_conversion(valid_files)

def _handle_excel_conversion(excel_file):
    """Handle Excel file conversion with template filling"""
    temp_dir = tempfile.mkdtemp()
    output_dir = tempfile.mkdtemp()
    
    try:
        # Validate template file
        word_template = os.path.join('samples', 'sample_document_for_placeholder.docx')
        template_ok, template_error = FileValidator.validate_template_file(word_template)
        if not template_ok:
            conversion_manager.conversion_progress['status'] = 'error'
            conversion_manager.conversion_progress['error'] = template_error
            return jsonify({'error': template_error}), 500
        
        # Save Excel file
        excel_filename = FileValidator.sanitize_filename(excel_file.filename)
        excel_path = os.path.join(temp_dir, excel_filename)
        excel_file.save(excel_path)
        
        # Validate Excel structure
        df_ok, df_error, df = FileValidator.validate_excel_structure(excel_path)
        if not df_ok:
            conversion_manager.conversion_progress['status'] = 'error'
            conversion_manager.conversion_progress['error'] = df_error
            return jsonify({'error': df_error}), 500
        
        total_rows = len(df)
        update_progress(0, total_rows, 'Generating Word documents from Excel...')
        
        # Generate Word documents
        docx_files = []
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(_generate_docx_from_excel_row, i, row, word_template, temp_dir): i 
                      for i, row in df.iterrows()}
            
            for idx, future in enumerate(as_completed(futures)):
                try:
                    docx_path = future.result(timeout=30)
                    docx_files.append(docx_path)
                    update_progress(idx + 1, total_rows, f'Generated {idx + 1}/{total_rows} Word docs...')
                except Exception as e:
                    ErrorHandler.log_error(e, "excel_to_docx_generation")
                    conversion_manager.conversion_progress['status'] = 'error'
                    conversion_manager.conversion_progress['error'] = f"Error generating Word document: {str(e)}"
                    return jsonify({'error': f"Error generating Word document: {str(e)}"}), 500
        
        update_progress(total_rows, total_rows, 'Converting generated docs to PDF...')
        
        # Convert to PDF
        successful_conversions, errors = conversion_manager.convert_batch_files(
            [(docx_path, os.path.basename(docx_path)) for docx_path in docx_files], 
            output_dir
        )
        
        if errors:
            conversion_manager.conversion_progress['status'] = 'error'
            conversion_manager.conversion_progress['error'] = '; '.join(errors)
            return jsonify({'error': '; '.join(errors)}), 500
        
        update_progress(total_rows, total_rows, 'Creating ZIP file...')
        
        # Create ZIP file
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zip_file:
            for pdf_path, pdf_name in successful_conversions:
                with open(pdf_path, 'rb') as pdf_file:
                    zip_file.writestr(pdf_name, pdf_file.read())
        
        zip_buffer.seek(0)
        conversion_manager.conversion_progress['status'] = 'completed'
        
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name='Appointment_letters.zip',
            mimetype='application/zip'
        )
        
    except Exception as e:
        error_dict = ErrorHandler.handle_conversion_error(e, [temp_dir, output_dir], "Excel processing failed")
        return jsonify(error_dict), 500

def _generate_docx_from_excel_row(i, row, template_path, output_dir):
    """Generate Word document from Excel row data"""
    data = {str(col): str(row[col]) for col in row.index}
    docx_name = f"{data.get('Name', 'Candidate')}_{i+1}.docx"
    docx_path = os.path.join(output_dir, docx_name)
    
    wp = WordProcessor()
    wp.fill_placeholders(template_path, docx_path, data)
    return docx_path

def _handle_single_file_conversion(file):
    """Handle single file conversion"""
    if not file or not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only .docx, .doc, and .xlsx files are allowed.'}), 400
    
    filename = FileValidator.sanitize_filename(file.filename)
    unique_filename = f"{uuid.uuid4().hex}_{filename}"
    input_path = os.path.join(current_app.config['UPLOAD_FOLDER'], unique_filename)
    
    try:
        file.save(input_path)
        update_progress(0, 1, f'Converting {filename}...')
        
        # Convert single file
        result = convert_single_file((input_path, filename))
        output_pdf, pdf_name, error = result
        
        if error or output_pdf is None:
            conversion_manager.conversion_progress['status'] = 'error'
            conversion_manager.conversion_progress['error'] = error or 'Conversion failed'
            os.remove(input_path)
            return jsonify({'error': error or 'Conversion failed'}), 500
        
        update_progress(1, 1, 'Preparing download...')
        
        # Read PDF and clean up
        with open(output_pdf, 'rb') as f:
            pdf_data = f.read()
        os.remove(input_path)
        os.remove(output_pdf)
        
        conversion_manager.conversion_progress['status'] = 'completed'
        
        return send_file(
            io.BytesIO(pdf_data),
            as_attachment=True,
            download_name=pdf_name,
            mimetype='application/pdf'
        )
        
    except Exception as e:
        error_dict = ErrorHandler.handle_file_processing_error(e, input_path, "Single file conversion failed")
        return jsonify(error_dict), 500

def _handle_batch_conversion(files):
    """Handle batch file conversion"""
    temp_dir = tempfile.mkdtemp()
    output_dir = tempfile.mkdtemp()
    
    try:
        total_files = len(files)
        update_progress(0, total_files, 'Preparing files for conversion...')
        
        # Save all files to temp_dir
        file_paths = []
        for i, file in enumerate(files):
            if file and file.filename and allowed_file(file.filename):
                filename = FileValidator.sanitize_filename(file.filename)
                file_path = os.path.join(temp_dir, filename)
                file.save(file_path)
                file_paths.append((file_path, filename))
                update_progress(i + 1, total_files, f'Prepared {i + 1}/{total_files} files...')
            else:
                conversion_manager.conversion_progress['status'] = 'error'
                conversion_manager.conversion_progress['error'] = f'Invalid file type for {file.filename if file else "unknown"}'
                return jsonify({'error': f'Invalid file type for {file.filename if file else "unknown"}. Only .docx, .doc, and .xlsx files are allowed.'}), 400
        
        if not file_paths:
            conversion_manager.conversion_progress['status'] = 'error'
            conversion_manager.conversion_progress['error'] = 'No valid files found'
            return jsonify({'error': 'No valid files found.'}), 400
        
        update_progress(total_files, total_files, 'Converting files with LibreOffice...')
        
        # Convert files
        successful_conversions, errors = conversion_manager.convert_batch_files(file_paths, output_dir)
        
        if errors:
            conversion_manager.conversion_progress['status'] = 'error'
            conversion_manager.conversion_progress['error'] = '; '.join(errors)
            return jsonify({'error': '; '.join(errors)}), 500
        
        update_progress(total_files, total_files, 'Creating ZIP file...')
        
        # Create ZIP file
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zip_file:
            for pdf_path, pdf_name in successful_conversions:
                with open(pdf_path, 'rb') as pdf_file:
                    zip_file.writestr(pdf_name, pdf_file.read())
        
        zip_buffer.seek(0)
        conversion_manager.conversion_progress['status'] = 'completed'
        
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name='Appointment_letters.zip',
            mimetype='application/zip'
        )
        
    except Exception as e:
        error_dict = ErrorHandler.handle_conversion_error(e, [temp_dir, output_dir], "Batch conversion failed")
        return jsonify(error_dict), 500

@main.route('/download/<filename>')
def download_file(filename):
    """Download file from download folder"""
    try:
        file_path = os.path.join(current_app.config['DOWNLOAD_FOLDER'], filename)
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        return jsonify({'error': 'File not found'}), 404

@main.route('/health')
def health_check():
    """Health check endpoint"""
    try:
        # Check system requirements
        requirements_ok, requirement_errors = conversion_manager.validate_conversion_requirements()
        
        return jsonify({
            'status': 'healthy' if requirements_ok else 'unhealthy',
            'requirements_ok': requirements_ok,
            'requirement_errors': requirement_errors,
            'system_info': ErrorHandler.get_system_info()
        })
    except Exception as e:
        return jsonify({
            'status': 'unhealthy',
            'error': str(e)
        }), 500 