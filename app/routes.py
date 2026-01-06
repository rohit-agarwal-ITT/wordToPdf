from flask import Blueprint, render_template, request, jsonify, send_file, current_app
import os
import uuid
import re
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
import threading
import time

main = Blueprint('main', __name__)

ALLOWED_EXTENSIONS = {'xlsx'}  # Only Excel files allowed

# Thread safety: Lock for conversion_progress dictionary
conversion_progress_lock = threading.Lock()

# Concurrent conversion limits: Semaphore to limit simultaneous conversions
# Default: 2 concurrent conversions (configurable via environment variable)
MAX_CONCURRENT_CONVERSIONS = int(os.environ.get('MAX_CONCURRENT_CONVERSIONS', '2'))
conversion_semaphore = threading.Semaphore(MAX_CONCURRENT_CONVERSIONS)

# Track semaphore acquisition time for debugging
_semaphore_acquisition_time = {}
_semaphore_lock = threading.Lock()

# Global progress tracking
conversion_progress = {
    'status': 'idle',  # idle, converting, completed, error
    'current': 0,
    'total': 0,
    'message': '',
    'error': None,
    'percentage': 0,
    'eta_seconds': None,
    'start_time': None,
    'elapsed_time': 0,
    'files': [],  # List of file statuses
    'display_total': 0,  # Actual number of files/records to display to user
    'display_current': 0,  # Current progress mapped to display_total scale
    'conversion_id': None  # Unique ID for each conversion to help frontend detect new conversions
}

def reset_progress():
    """Reset progress tracking - thread-safe
    This MUST be called at the start of each new conversion to clear previous state.
    """
    global conversion_progress
    with conversion_progress_lock:
        # Explicitly reset all fields to ensure no stale data persists
        conversion_progress = {
            'status': 'idle',
            'current': 0,
            'total': 0,
            'message': 'Initializing...',
            'error': None,
            'percentage': 0,
            'eta_seconds': None,
            'start_time': None,
            'elapsed_time': 0,
            'files': [],
            'display_total': 0,  # CRITICAL: Reset to 0 so new conversion sets correct value
            'display_current': 0,  # CRITICAL: Reset to 0 so new conversion starts fresh
            'conversion_id': None  # Reset conversion ID
        }

def set_progress_status(status, error=None, eta_seconds=None):
    """Set conversion progress status - thread-safe helper"""
    global conversion_progress
    with conversion_progress_lock:
        conversion_progress['status'] = status
        if error is not None:
            conversion_progress['error'] = error
        if eta_seconds is not None:
            conversion_progress['eta_seconds'] = eta_seconds

def update_progress(current, total, message, current_file=None, file_status=None, display_total=None):
    """Update conversion progress - thread-safe"""
    global conversion_progress
    from time import time
    
    with conversion_progress_lock:
        # Only update start_time if it's None (new conversion) or if conversion_id changed
        # This ensures start_time is reset for new conversions
        if conversion_progress['start_time'] is None:
            conversion_progress['start_time'] = time()
        
        conversion_progress['current'] = current
        conversion_progress['total'] = total
        conversion_progress['message'] = message
        conversion_progress['status'] = 'converting'
        # Ensure conversion_id is preserved during updates (don't overwrite it)
        # conversion_id should only be set at the start of a new conversion
        
        # Set display_total (actual number of files/records to show to user)
        # If explicitly provided, always use it (this allows resetting from previous conversion)
        if display_total is not None:
            conversion_progress['display_total'] = display_total
        elif 'display_total' not in conversion_progress or conversion_progress['display_total'] == 0:
            # Default to total if not set, but only if it's not already set to a non-zero value
            # This prevents overwriting a valid display_total with total (which might be total_steps)
            conversion_progress['display_total'] = total
        # If display_total is already set to a non-zero value and display_total parameter is None,
        # keep the existing value (don't overwrite it) - BUT this should only happen during a single conversion,
        # not across conversions since reset_progress() sets it to 0
        
        # Calculate percentage based on total steps (internal tracking)
        if total > 0:
            conversion_progress['percentage'] = min(100, int((current / total) * 100))
        else:
            conversion_progress['percentage'] = 0
        
        # Calculate ETA based on display_total (user-facing count)
        elapsed = time() - conversion_progress['start_time']
        conversion_progress['elapsed_time'] = int(elapsed)
        
        display_total_val = conversion_progress.get('display_total', total)
        # Calculate current display progress (for ETA calculation and frontend display)
        if display_total_val > 0 and total > 0:
            # Map internal progress to display progress
            # Ensure display_current never decreases (monotonic increase)
            calculated_display_current = min(display_total_val, int((current / total) * display_total_val))
            existing_display_current = conversion_progress.get('display_current', 0)
            # Only update if the new value is greater than or equal to existing (prevent decreases)
            display_current = max(existing_display_current, calculated_display_current)
            # Store display_current for frontend to use directly
            conversion_progress['display_current'] = display_current
            
            # Calculate ETA only when we have meaningful progress
            # Require at least 3 files processed for more accurate ETA calculation
            # This prevents incorrect ETA calculations early in the process
            if display_current > 0 and display_total_val > display_current and elapsed > 0:
                remaining_files = display_total_val - display_current
                
                # Use different calculation strategies based on sample size
                if display_current >= 3:
                    # With 3+ samples, use direct average (more accurate)
                    avg_time_per_file = elapsed / display_current
                elif display_current >= 2:
                    # With 2 samples, apply 1.5x multiplier to be more conservative
                    avg_time_per_file = (elapsed / display_current) * 1.5
                else:
                    # With only 1 sample, apply 2x multiplier and use minimum 2 seconds per file
                    # This prevents unrealistic ETAs from a single slow file
                    avg_time_per_file = max(2, (elapsed / display_current) * 2)
                
                # Calculate ETA
                estimated_eta = avg_time_per_file * remaining_files
                
                # Cap ETA at 2 hours (7200 seconds) to prevent unrealistic estimates
                # This handles edge cases where early files are much slower than average
                conversion_progress['eta_seconds'] = min(int(estimated_eta), 7200)
            elif display_current >= display_total_val:
                # All items completed - check if we're still processing (ZIP creation, etc.)
                # If status is still 'converting', show a small ETA for final processing
                if conversion_progress.get('status') == 'converting':
                    # Estimate 5-10 seconds for ZIP creation and finalization
                    conversion_progress['eta_seconds'] = 5
                else:
                    # Fully completed
                    conversion_progress['eta_seconds'] = 0
            else:
                conversion_progress['eta_seconds'] = None
        else:
            conversion_progress['display_current'] = 0
            conversion_progress['eta_seconds'] = None
        
        # Update file status
        if current_file and file_status:
            # Find or create file entry
            file_found = False
            for file_entry in conversion_progress['files']:
                if file_entry.get('name') == current_file:
                    file_entry['status'] = file_status
                    file_entry['progress'] = current
                    file_found = True
                    break
            
            if not file_found:
                conversion_progress['files'].append({
                    'name': current_file,
                    'status': file_status,
                    'progress': current
                })

def allowed_file(filename):
    # Only allow Excel files (.xlsx)
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'xlsx'

def get_template_path(location_value, designation_value=None):
    """
    Get the appropriate template path based on designation and location.
    
    Args:
        location_value: The location value from Excel (e.g., 'Jaipur', 'Bangalore')
        designation_value: The designation value from Excel (e.g., 'Trainee', 'Software Engineer', 'Junior Software Engineer')
    
    Returns:
        Path to the appropriate template file
    """
    # Normalize designation value (case-insensitive, strip whitespace)
    if designation_value:
        designation = str(designation_value).strip().lower()
        
        # Check if designation is Trainee (case-insensitive)
        if 'trainee' in designation:
            # Trainee template is not location-specific
            template_name = 'sample_document_for_trainee_placeholder.docx'
            template_path = os.path.join('samples', template_name)
            
            # Fallback to Jaipur template if trainee template doesn't exist
            if not os.path.exists(template_path):
                template_path = os.path.join('samples', 'sample_document_for_placeholder_jaipur.docx')
            
            return template_path
    
    # For non-Trainee designations (Software Engineer, Junior Software Engineer, etc.)
    # Use location-based templates
    # Normalize location value (case-insensitive, strip whitespace)
    if location_value:
        location = str(location_value).strip().lower()
        
        # Check for Bangalore location (case-insensitive)
        if 'bangalore' in location or 'bengaluru' in location:
            template_name = 'sample_document_for_placeholder_bangalore.docx'
        else:
            # Default to Jaipur template
            template_name = 'sample_document_for_placeholder_jaipur.docx'
    else:
        # Default to Jaipur template if location is empty/None
        template_name = 'sample_document_for_placeholder_jaipur.docx'
    
    template_path = os.path.join('samples', template_name)
    
    # Fallback to Jaipur template if location-specific template doesn't exist
    if not os.path.exists(template_path):
        template_path = os.path.join('samples', 'sample_document_for_placeholder_jaipur.docx')
    
    return template_path

def format_date_field(value, field_name):
    """
    Format date fields (Date of Joining, Effective Date, Date) to DD-MonthName-YYYY format.
    Handles input formats: '2024-07-01' (ISO) or '6/30/2025' (US format).
    
    Args:
        value: The date value from Excel (can be string or datetime object)
        field_name: The name of the field/column
    
    Returns:
        Formatted date string like '05-August-2025' or original value if not a date field or parsing fails
    """
    # Only format specific date fields (case-insensitive matching)
    date_fields = ['Date of Joining', 'Effective Date', 'Date']
    field_name_normalized = str(field_name).strip()
    # Check if field_name matches any date field (case-insensitive)
    is_date_field = any(field_name_normalized.lower() == df.lower() for df in date_fields)
    if not is_date_field:
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
    """Return current conversion progress - thread-safe"""
    global conversion_progress
    try:
        with conversion_progress_lock:
            # Create a copy to avoid holding lock during JSON serialization
            progress_copy = conversion_progress.copy()
            # Deep copy the files list to avoid race conditions
            progress_copy['files'] = conversion_progress['files'].copy()
        return jsonify(progress_copy)
    except Exception as e:
        # Ensure we always return JSON, even on errors
        current_app.logger.error(f'Error getting progress: {e}', exc_info=True)
        return jsonify({
            'status': 'error',
            'error': 'Failed to retrieve progress information',
            'current': 0,
            'total': 0,
            'percentage': 0,
            'message': 'Error retrieving progress',
            'files': []
        }), 500

@main.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and conversion with thread safety and concurrency limits"""
    global conversion_progress
    
    # Acquire semaphore to limit concurrent conversions
    # If limit is reached, return 503 Service Unavailable
    semaphore_acquired = False
    request_id = None
    try:
        if not conversion_semaphore.acquire(blocking=False):
            # Check if semaphore might be stuck (held for more than 10 minutes)
            with _semaphore_lock:
                current_time = time.time()
                stuck_requests = [req_id for req_id, acquire_time in _semaphore_acquisition_time.items() 
                                if current_time - acquire_time > 600]  # 10 minutes
                if stuck_requests:
                    current_app.logger.warning(f'Detected potentially stuck semaphore acquisitions: {stuck_requests}')
                    # Don't auto-release, but log for monitoring
            
            return jsonify({
                'error': 'Server is busy processing other conversions. Please try again in a moment.'
            }), 503
        
        semaphore_acquired = True
        # Track semaphore acquisition
        request_id = str(uuid.uuid4())
        with _semaphore_lock:
            _semaphore_acquisition_time[request_id] = time.time()
    except Exception as e:
        current_app.logger.error(f'Error acquiring semaphore: {e}', exc_info=True)
        return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 500
    
    # Initialize variables for cleanup
    temp_dir = None
    output_dir = None
    conversion_thread = None
    monitor_thread = None
    
    try:
        # Reset progress for new conversion immediately and set initial status
        # This MUST happen first to clear any previous conversion state
        # Do this atomically to prevent frontend from seeing stale data
        new_conversion_id = str(uuid.uuid4())  # Generate unique ID for this conversion
        with conversion_progress_lock:
            # Inline reset_progress() logic to avoid deadlock (reset_progress also tries to acquire the same lock)
            # Explicitly reset all fields to ensure no stale data persists
            conversion_progress = {
                'status': 'converting',  # Set to converting immediately
                'current': 0,
                'total': 0,
                'message': 'Initializing new conversion...',
                'error': None,
                'percentage': 0,
                'eta_seconds': None,
                'start_time': None,
                'elapsed_time': 0,
                'files': [],  # Clear old file entries
                'display_total': 0,  # CRITICAL: Reset to 0 so new conversion sets correct value
                'display_current': 0,  # CRITICAL: Reset to 0 so new conversion starts fresh
                'conversion_id': new_conversion_id  # Set unique ID to help frontend detect new conversion
            }
        
        if 'files[]' not in request.files:
            return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 400
        
        files = request.files.getlist('files[]')
        if not files or files[0].filename == '' or files[0].filename is None:
            return jsonify({'error': 'No files selected. Please choose an Excel file (.xlsx) to upload.'}), 400
        
        # Validate that only Excel files are uploaded
        for file in files:
            if file and file.filename:
                if not allowed_file(file.filename):
                    return jsonify({'error': f'Invalid file type: {file.filename}. Only Excel files (.xlsx) are allowed.'}), 400
        
        # Excel to Word to PDF batch logic
        if len(files) == 1 and files[0].filename and files[0].filename.lower().endswith('.xlsx'):
            # Create temp directories for Excel conversion
            try:
                temp_dir = tempfile.mkdtemp()
                output_dir = tempfile.mkdtemp()
            except Exception as e:
                current_app.logger.error(f'Error creating temp directories: {e}', exc_info=True)
                return jsonify({'error': 'An error occurred while setting up conversion. Please try again.'}), 500
            
            pdf_files = []
            errors = []
            
            try:  # Main try block for Excel conversion
                excel_file = files[0]
                if not excel_file.filename:
                    return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 400
                
                # Get Excel filename for dynamic ZIP name - use same name with .zip extension
                excel_filename = secure_filename(excel_file.filename)
                excel_base_name = os.path.splitext(excel_filename)[0]  # Remove .xlsx extension
                zip_filename = f"{excel_base_name}.zip"
                excel_path = os.path.join(temp_dir, secure_filename(excel_file.filename or 'uploaded.xlsx'))
                
                # Save Excel file - this should be quick
                # IMPORTANT: Ensure we're saving the NEW file, not reusing old data
                try:
                    # Reset file pointer to beginning in case it was read before
                    excel_file.seek(0)
                    excel_file.save(excel_path)
                    current_app.logger.info(f'Saved Excel file: {excel_filename} to {excel_path}, conversion_id: {new_conversion_id}')
                except Exception as e:
                    current_app.logger.error(f'Error saving Excel file: {e}', exc_info=True)
                    return jsonify({'error': 'An error occurred while saving the file. Please try again.'}), 400
                
                # Read Excel file - this should also be quick for validation
                # IMPORTANT: Read from the file we just saved, not any cached data
                try:
                    df = pd.read_excel(excel_path)
                    current_app.logger.info(f'Read Excel file: {excel_filename}, rows: {len(df) if df is not None else 0}, conversion_id: {new_conversion_id}')
                except Exception as e:
                    current_app.logger.error(f'Error reading Excel file: {e}', exc_info=True)
                    return jsonify({'error': 'An error occurred while reading the Excel file. Please ensure the file is valid.'}), 400
                
                if df is None or df.empty:
                    return jsonify({'error': 'The Excel file appears to be empty. Please check your file.'}), 400
                
                # Read Excel file and get row count FIRST, then set progress with correct count
                total_rows = len(df)
                # Total steps: DOCX generation (50%) + PDF conversion (50%)
                total_steps = total_rows * 2
                
                # CRITICAL: Set display_total immediately after reading Excel, before any progress updates
                # This ensures frontend sees the correct count from the start and can start polling
                # The frontend should be able to poll /progress immediately after this
                # Do this atomically to ensure frontend sees the update
                with conversion_progress_lock:
                    # ALWAYS use the new_conversion_id that was generated at the start of this request
                    # Don't preserve any old conversion_id - this is a new conversion
                    conversion_progress['display_total'] = total_rows
                    conversion_progress['display_current'] = 0  # Reset display_current as well
                    conversion_progress['status'] = 'converting'  # Ensure status is set
                    conversion_progress['message'] = f'Processing {total_rows} records from Excel file...'
                    conversion_progress['files'] = []  # Clear any old file entries - start fresh
                    conversion_progress['current'] = 0
                    conversion_progress['total'] = total_steps
                    conversion_progress['percentage'] = 0
                    conversion_progress['conversion_id'] = new_conversion_id  # ALWAYS use the new ID for this conversion
                    conversion_progress['start_time'] = time.time()  # Reset start time for new conversion
                    conversion_progress['elapsed_time'] = 0  # Reset elapsed time
                    conversion_progress['eta_seconds'] = None  # Reset ETA
                    current_app.logger.info(f'Set progress for Excel conversion: {total_rows} rows, conversion_id: {new_conversion_id}')
                
                # Now update progress with the correct display_total
                # This allows frontend to immediately see progress and start polling
                # IMPORTANT: This happens BEFORE any long-running operations
                update_progress(0, total_steps, 'Preparing appointment letters for PDF generation...', display_total=total_rows)
                
                def generate_docx(row_tuple):
                    i, row = row_tuple
                    try:
                        # Process data with date formatting for specific fields
                        data = {}
                        for col in df.columns:
                            col_str = str(col)
                            value = row[col]
                            # Handle NaN/None values from pandas
                            if pd.isna(value):
                                value = ''
                            # Format date fields appropriately
                            formatted_value = format_date_field(value, col_str)
                            data[col_str] = formatted_value
                        
                        # Get location value to determine which template to use
                        # Primary: Look for "Place of Joining" column (exact match, case-insensitive)
                        location_value = None
                        for col in df.columns:
                            col_str = str(col).strip()
                            # Check for exact match with "Place of Joining" (case-insensitive)
                            if col_str.lower() == 'place of joining':
                                location_value = data.get(col_str)
                                break
                        
                        # Fallback: Try common location column names (case-insensitive)
                        if location_value is None:
                            location_column_names = ['Location', 'location', 'LOCATION', 'City', 'city', 'CITY', 'Location Name', 'location name', 'Place of Joining', 'place of joining']
                            for loc_col in location_column_names:
                                if loc_col in data:
                                    location_value = data[loc_col]
                                    break
                        
                        # If still not found, try to find any column containing 'location', 'city', or 'place'
                        if location_value is None:
                            for col in df.columns:
                                col_lower = str(col).lower()
                                if 'location' in col_lower or 'city' in col_lower or ('place' in col_lower and 'joining' in col_lower):
                                    location_value = data.get(str(col))
                                    break
                        
                        # Get designation value to determine which template to use
                        # Primary: Look for "Designation" column (exact match, case-insensitive)
                        designation_value = None
                        for col in df.columns:
                            col_str = str(col).strip()
                            # Check for exact match with "Designation" (case-insensitive)
                            if col_str.lower() == 'designation':
                                designation_value = data.get(col_str)
                                break
                        
                        # Fallback: Try common designation column names (case-insensitive)
                        if designation_value is None:
                            designation_column_names = ['Designation', 'designation', 'DESIGNATION', 'Role', 'role', 'ROLE', 'Job Title', 'job title']
                            for desig_col in designation_column_names:
                                if desig_col in data:
                                    designation_value = data[desig_col]
                                    break
                        
                        # If still not found, try to find any column containing 'designation' or 'role'
                        if designation_value is None:
                            for col in df.columns:
                                col_lower = str(col).lower()
                                if 'designation' in col_lower or ('role' in col_lower and 'title' not in col_lower):
                                    designation_value = data.get(str(col))
                                    break
                        
                        # Get the appropriate template based on designation and location
                        word_template = get_template_path(location_value, designation_value)
                        
                        docx_name = f"{data.get('Name', 'Candidate')}_{i+1}.docx"
                        docx_path = os.path.join(temp_dir, docx_name)
                        wp = WordProcessor()  # Create a new instance per row/thread
                        wp.fill_placeholders(word_template, docx_path, data)
                        return docx_path
                    except Exception as e:
                        # Log the actual error for debugging
                        import traceback
                        error_msg = f"Error generating DOCX for row {i+1}: {str(e)}\n{traceback.format_exc()}"
                        print(error_msg)  # Print to console/log
                        raise Exception(f"Error processing row {i+1}: {str(e)}")
                
                docx_files = []
                with ThreadPoolExecutor(max_workers=4) as executor:
                    futures = {executor.submit(generate_docx, (i, row)): i for i, row in df.iterrows()}
                    for idx, future in enumerate(as_completed(futures)):
                        try:
                            docx_path = future.result()
                            docx_files.append(docx_path)
                        except Exception as e:
                            # Log the error and continue with other rows
                            import traceback
                            error_msg = f"Error in future result: {str(e)}\n{traceback.format_exc()}"
                            print(error_msg)
                            # Re-raise to stop processing
                            raise Exception(f"Error generating appointment letter: {str(e)}")
                        # Progress for DOCX generation: 0 to 50% - but show as PDF preparation
                        # Map progress to show as if we're creating PDFs directly
                        # Update progress more frequently to show row-by-row progress
                        rows_processed = idx + 1
                        progress_pct = rows_processed / total_rows * 0.3  # First 30% is preparation
                        current_progress = int(total_steps * progress_pct)
                        # Update progress - update_progress will calculate display_current correctly
                        # But we need to ensure it reflects the actual row count
                        update_progress(current_progress, total_steps, 
                                      f'Preparing appointment letters... ({rows_processed}/{total_rows} records)', 
                                      display_total=total_rows)
                        
                        # Explicitly update display_current to match rows processed for better accuracy
                        # This ensures frontend sees the correct row-by-row progress
                        with conversion_progress_lock:
                            # Ensure display_current matches rows processed (for row-by-row visibility)
                            conversion_progress['display_current'] = max(
                                conversion_progress.get('display_current', 0),
                                rows_processed
                            )
                
                # DOCX generation complete, now starting PDF conversion (50% of progress)
                # Show as if we're starting PDF creation
                update_progress(total_rows, total_steps, 'Generating PDFs. This may take a moment...', display_total=total_rows)
                
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
                    
                    # Start conversion in background and monitor progress
                    conversion_complete = threading.Event()
                    conversion_error = [None]
                    
                    def monitor_pdf_conversion_excel():
                        """Monitor output directory for PDF files appearing"""
                        expected_pdfs = {os.path.splitext(os.path.basename(f))[0] + '.pdf': os.path.basename(f) for f in validated_files}
                        pdfs_found = set()
                        start_time = time.time()
                        max_wait_time = 300  # 5 minutes max
                        
                        while not conversion_complete.is_set() and (time.time() - start_time) < max_wait_time:
                            # Check for new PDFs
                            if os.path.exists(output_dir):
                                existing_pdfs = set(f for f in os.listdir(output_dir) if f.endswith('.pdf'))
                                new_pdfs = existing_pdfs - pdfs_found
                                
                                for pdf_file in new_pdfs:
                                    if pdf_file in expected_pdfs:
                                        pdfs_found.add(pdf_file)
                                        # Update progress: 30% to 90% based on PDFs found (show as PDF creation)
                                        progress_pct = 0.3 + (len(pdfs_found) / len(expected_pdfs)) * 0.6
                                        current_progress = int(total_steps * progress_pct)
                                        update_progress(current_progress, total_steps, 
                                                      'Creating PDFs...', 
                                                      display_total=total_rows)
                                
                                # If all PDFs are found, we're done
                                if len(pdfs_found) == len(expected_pdfs):
                                    break
                            
                            time.sleep(0.5)  # Check every 500ms
                    
                    def run_conversion_excel():
                        """Run the actual conversion"""
                        try:
                            subprocess.run([
                                soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir
                            ] + validated_files, check=True, timeout=300)
                            conversion_complete.set()
                        except Exception as e:
                            conversion_error[0] = e
                            conversion_complete.set()
                    
                    # Start conversion and monitoring in separate threads
                    conversion_thread = threading.Thread(target=run_conversion_excel, daemon=True)
                    monitor_thread = threading.Thread(target=monitor_pdf_conversion_excel, daemon=True)
                    
                    conversion_thread.start()
                    monitor_thread.start()
                    
                    # Wait for conversion to complete
                    conversion_thread.join(timeout=300)
                    conversion_complete.set()
                    monitor_thread.join(timeout=5)
                    
                    if conversion_error[0]:
                        raise conversion_error[0]
                
                except Exception as e:
                    current_app.logger.error(f'PDF conversion error: {e}', exc_info=True)
                    raise Exception(f"PDF conversion failed: {str(e)}")
                
                # Collect PDFs and update progress (90% to 100%)
                pdfs_collected = 0
                for idx, docx_file in enumerate(docx_files):
                    base = os.path.splitext(os.path.basename(docx_file))[0]
                    # Extract name by removing the _number suffix (e.g., "John Doe_1" -> "John Doe")
                    name_match = re.match(r'^(.+?)_\d+$', base)
                    if name_match:
                        name = name_match.group(1)
                    else:
                        name = base  # Fallback if pattern doesn't match
                    pdf_name = f"Appointment letter and Training Agreement- {name}.pdf"
                    pdf_path = os.path.join(output_dir, base + '.pdf')
                    if os.path.exists(pdf_path):
                        pdf_files.append((pdf_path, pdf_name))
                        pdfs_collected += 1
                        # Progress for PDF collection: 90% to 100%
                        progress_pct = 0.9 + (pdfs_collected / total_rows) * 0.1
                        current_progress = int(total_steps * progress_pct)
                        update_progress(current_progress, total_steps, 
                                      'Finalizing PDFs...', 
                                      display_total=total_rows)
                    else:
                        errors.append(f'PDF not found for {base}')
                
                if errors:
                    set_progress_status('error', error='An error occurred during conversion. Please try again.')
                    if temp_dir and os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir)
                    if output_dir and os.path.exists(output_dir):
                        shutil.rmtree(output_dir)
                    with conversion_progress_lock:
                        error_msg = conversion_progress.get('error', 'An error occurred during conversion. Please try again.')
                    return jsonify({'error': error_msg}), 500
                
                # All PDFs collected, creating ZIP
                # Keep status as 'converting' during ZIP creation so ETA still shows
                zip_start_time = time.time()
                update_progress(total_steps, total_steps, 'All PDFs created! Creating ZIP package...', display_total=total_rows)
                # Set a small ETA for ZIP creation - estimate based on number of files
                estimated_zip_time = min(10, max(3, len(pdf_files) * 0.1))  # 0.1s per file, min 3s, max 10s
                set_progress_status('converting', eta_seconds=int(estimated_zip_time))
                # Use lower compression for faster ZIP creation (compresslevel=1 is much faster than 6)
                # For large files, use a temporary file instead of BytesIO to avoid memory issues
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as zip_file:
                    for idx, (pdf_path, pdf_name) in enumerate(pdf_files):
                        # Read file directly - zipfile handles large files efficiently
                        # Lower compression (compresslevel=1) makes this much faster
                        with open(pdf_path, 'rb') as pdf_file:
                            zip_file.writestr(pdf_name, pdf_file.read())
                        # Update progress during ZIP creation for large files
                        if (idx + 1) % 10 == 0 or idx == len(pdf_files) - 1:
                            update_progress(total_steps, total_steps, 
                                          f'Creating ZIP package... ({idx + 1}/{len(pdf_files)} files)', 
                                          display_total=total_rows)
                zip_creation_time = time.time() - zip_start_time
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if output_dir and os.path.exists(output_dir):
                    shutil.rmtree(output_dir)
                zip_buffer.seek(0)
                # Update progress one more time before marking as completed
                update_progress(total_steps, total_steps, 'Successfully created all PDF appointment letters! Download starting...', display_total=total_rows)
                set_progress_status('completed', eta_seconds=0)
                # Send file with explicit timeout and chunk size for better performance
                return send_file(
                    zip_buffer,
                    as_attachment=True,
                    download_name=zip_filename,
                    mimetype='application/zip',
                    max_age=0  # Prevent caching
                )
            
            except Exception as e:  # Main exception handler for Excel conversion
                current_app.logger.error(f'Excel conversion error: {e}', exc_info=True)
                set_progress_status('error', error='An error occurred during conversion. Please try again.')
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir)
                    except Exception:
                        pass
                if output_dir and os.path.exists(output_dir):
                    try:
                        shutil.rmtree(output_dir)
                    except Exception:
                        pass
                with conversion_progress_lock:
                    error_msg = conversion_progress.get('error', 'An error occurred during conversion. Please try again.')
                return jsonify({'error': error_msg}), 500
        
        # Handle single file case
        if len(files) == 1:
            file = files[0]
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename or 'uploaded.docx')
                unique_filename = f"{uuid.uuid4().hex}_{filename}"
                input_path = os.path.join(current_app.config['UPLOAD_FOLDER'], unique_filename)
                file.save(input_path)
                
                # For single file: 0-50% preparation, 50-100% PDF conversion
                with conversion_progress_lock:
                    conversion_progress['display_total'] = 1
                update_progress(0, 2, 'Preparing file for PDF conversion...', filename, 'processing', display_total=1)
                update_progress(1, 2, 'Creating PDF...', filename, 'processing', display_total=1)
                
                # Convert single file
                result = convert_single_file((input_path, filename))
                output_pdf, pdf_name, error = result
                
                if error or output_pdf is None:
                    set_progress_status('error', error='An error occurred during conversion. Please try again.')
                    update_progress(1, 2, 'Error converting file', filename, 'error')
                    os.remove(input_path)
                    with conversion_progress_lock:
                        error_msg = conversion_progress['error']
                    return jsonify({'error': error_msg}), 500
                
                update_progress(2, 2, 'PDF created successfully! Preparing download...', filename, 'completed', display_total=1)
                
                # Read PDF and clean up
                with open(output_pdf, 'rb') as f:
                    pdf_data = f.read()
                os.remove(input_path)
                os.remove(output_pdf)
                
                set_progress_status('completed')
                update_progress(2, 2, 'PDF ready! Download starting...', display_total=1)
                
                return send_file(
                    io.BytesIO(pdf_data),
                    as_attachment=True,
                    download_name=pdf_name,
                    mimetype='application/pdf'
                )
            else:
                return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 400
        
        # Handle multiple files case - BATCH CONVERSION
        else:
            temp_dir = tempfile.mkdtemp()
            output_dir = tempfile.mkdtemp()
            pdf_files = []
            errors = []
        
        try:
            total_files = len(files)
            with conversion_progress_lock:
                conversion_progress['display_total'] = total_files
            update_progress(0, total_files, 'Preparing files for PDF conversion...', display_total=total_files)
            
            # Save all files to temp_dir
            for i, file in enumerate(files):
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename or f'file_{i}.docx')
                    file_path = os.path.join(temp_dir, filename)
                    file.save(file_path)
                    # Show as PDF preparation, not file preparation
                    progress_pct = (i + 1) / total_files * 0.2  # First 20% is file prep
                    current_progress = int(total_files * progress_pct)
                    update_progress(current_progress, total_files, 
                                  'Preparing files for PDF conversion...', 
                                  filename, 'processing', display_total=total_files)
                else:
                    set_progress_status('error', error='An error occurred during conversion. Please try again.')
                    shutil.rmtree(temp_dir)
                    shutil.rmtree(output_dir)
                    with conversion_progress_lock:
                        error_msg = conversion_progress['error']
                    return jsonify({'error': error_msg}), 400
            
            # Progress: 0-20% file prep, 20-90% PDF conversion, 90-100% collection
            update_progress(int(total_files * 0.2), total_files, 'Creating PDFs. This may take a moment...', display_total=total_files)
            
            # Batch convert all files in one soffice call
            
            docx_files = [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(('.docx', '.doc'))]
            
            if not docx_files:
                set_progress_status('error', error='An error occurred during conversion. Please try again.')
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                with conversion_progress_lock:
                    error_msg = conversion_progress['error']
                return jsonify({'error': error_msg}), 400
            
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
                
                # Update progress to show PDF conversion is starting
                update_progress(int(total_files * 0.25), total_files, f'Generating PDFs. Please wait...', display_total=total_files)
                
                # Start conversion in background and monitor progress
                conversion_complete = threading.Event()
                conversion_error = [None]
                
                def monitor_pdf_conversion():
                    """Monitor output directory for PDF files appearing"""
                    expected_pdfs = {os.path.splitext(os.path.basename(f))[0] + '.pdf': os.path.basename(f) for f in validated_files}
                    pdfs_found = set()
                    start_time = time.time()
                    max_wait_time = 300  # 5 minutes max
                    
                    while not conversion_complete.is_set() and (time.time() - start_time) < max_wait_time:
                        # Check for new PDFs
                        if os.path.exists(output_dir):
                            existing_pdfs = set(f for f in os.listdir(output_dir) if f.endswith('.pdf'))
                            new_pdfs = existing_pdfs - pdfs_found
                            
                            for pdf_file in new_pdfs:
                                if pdf_file in expected_pdfs:
                                    pdfs_found.add(pdf_file)
                                    original_filename = expected_pdfs[pdf_file]
                                    # Update progress: 25% to 85% based on PDFs found
                                    progress_pct = 0.25 + (len(pdfs_found) / len(expected_pdfs)) * 0.6
                                    current_progress = int(total_files * progress_pct)
                                    update_progress(current_progress, total_files, 
                                                  f'Creating PDF {len(pdfs_found)} of {len(expected_pdfs)}...', 
                                                  original_filename, 'processing', display_total=total_files)
                            
                            # If all PDFs are found, we're done
                            if len(pdfs_found) == len(expected_pdfs):
                                break
                        
                        time.sleep(0.5)  # Check every 500ms
                
                def run_conversion():
                    """Run the actual conversion"""
                    try:
                        subprocess.run([
                            soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir
                        ] + validated_files, check=True, timeout=300)
                        conversion_complete.set()
                    except Exception as e:
                        conversion_error[0] = e
                        conversion_complete.set()
                
                # Start conversion and monitoring in separate threads
                conversion_thread = threading.Thread(target=run_conversion, daemon=True)
                monitor_thread = threading.Thread(target=monitor_pdf_conversion, daemon=True)
                
                conversion_thread.start()
                monitor_thread.start()
                
                # Wait for conversion to complete
                conversion_thread.join(timeout=300)
                conversion_complete.set()
                monitor_thread.join(timeout=5)
                
                if conversion_error[0]:
                    raise conversion_error[0]
                
                # PDF conversion complete
                update_progress(int(total_files * 0.85), total_files, 'PDFs created! Finalizing files...', display_total=total_files)
            except Exception as e:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = 'An error occurred during conversion. Please try again.'
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': conversion_progress['error']}), 500
            
            # Collect PDFs and rename for zipping, updating progress as each PDF is found
            pdfs_found = 0
            for idx, docx_file in enumerate(docx_files):
                base = os.path.splitext(os.path.basename(docx_file))[0]
                # Extract name by removing the _number suffix (e.g., "John Doe_1" -> "John Doe")
                name_match = re.match(r'^(.+?)_\d+$', base)
                if name_match:
                    name = name_match.group(1)
                else:
                    name = base  # Fallback if pattern doesn't match
                pdf_name = f"Appointment letter and Training Agreement- {name}.pdf"
                pdf_path = os.path.join(output_dir, base + '.pdf')
                filename = os.path.basename(docx_file)
                if os.path.exists(pdf_path):
                    pdf_files.append((pdf_path, pdf_name))
                    pdfs_found += 1
                    # Progress from 85% to 100% as PDFs are collected
                    progress_value = int(total_files * 0.85) + int((pdfs_found / total_files) * total_files * 0.15)
                    update_progress(progress_value, total_files, 
                                  'Finalizing PDFs...', 
                                  filename, 'completed', display_total=total_files)
                else:
                    errors.append(f'PDF not found for {base}')
                    update_progress(int(total_files * 0.85) + pdfs_found, total_files, 
                                  'Error creating PDF...', 
                                  filename, 'error', display_total=total_files)
            
            if errors:
                conversion_progress['status'] = 'error'
                conversion_progress['error'] = 'An error occurred during conversion. Please try again.'
                shutil.rmtree(temp_dir)
                shutil.rmtree(output_dir)
                return jsonify({'error': conversion_progress['error']}), 500
            
            # After all PDFs are collected and before sending the response
            # Zip all PDFs
            # Keep status as 'converting' during ZIP creation so ETA still shows
            zip_start_time = time.time()
            update_progress(total_files, total_files, 'All PDFs created! Creating ZIP package...', display_total=total_files)
            # Set a small ETA for ZIP creation - estimate based on number of files
            estimated_zip_time = min(10, max(3, len(pdf_files) * 0.1))  # 0.1s per file, min 3s, max 10s
            set_progress_status('converting', eta_seconds=int(estimated_zip_time))
            # Use lower compression for faster ZIP creation (compresslevel=1 is much faster than 6)
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as zip_file:
                for idx, (pdf_path, pdf_name) in enumerate(pdf_files):
                    # Read file directly - zipfile handles large files efficiently
                    # Lower compression (compresslevel=1) makes this much faster
                    with open(pdf_path, 'rb') as pdf_file:
                        zip_file.writestr(pdf_name, pdf_file.read())
                    # Update progress during ZIP creation for large files
                    if (idx + 1) % 10 == 0 or idx == len(pdf_files) - 1:
                        update_progress(total_files, total_files, 
                                      f'Creating ZIP package... ({idx + 1}/{len(pdf_files)} files)', 
                                      display_total=total_files)
            
            zip_creation_time = time.time() - zip_start_time
            shutil.rmtree(temp_dir)
            shutil.rmtree(output_dir)
            zip_buffer.seek(0)
            
            # Update progress one more time before marking as completed
            update_progress(total_files, total_files, 'Successfully created all PDFs! Download starting...', display_total=total_files)
            set_progress_status('completed', eta_seconds=0)
            
            # Generate dynamic ZIP filename from first file (should be Excel file) - use same name with .zip extension
            if files and files[0] and files[0].filename:
                first_filename = secure_filename(files[0].filename)
                first_base_name = os.path.splitext(first_filename)[0]  # Remove extension
                zip_filename = f"{first_base_name}.zip"
            else:
                zip_filename = 'Appointment_letters.zip'
            
            # Send file with explicit timeout and chunk size for better performance
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name=zip_filename,
                mimetype='application/zip',
                max_age=0  # Prevent caching
            )
        except Exception as e:
            current_app.logger.error(f'Upload error: {e}', exc_info=True)
            set_progress_status('error', error='An error occurred during conversion. Please try again.')
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir)
                except Exception as cleanup_error:
                    current_app.logger.error(f'Error cleaning up temp_dir: {cleanup_error}')
            if output_dir and os.path.exists(output_dir):
                try:
                    shutil.rmtree(output_dir)
                except Exception as cleanup_error:
                    current_app.logger.error(f'Error cleaning up output_dir: {cleanup_error}')
            with conversion_progress_lock:
                error_msg = conversion_progress.get('error', 'An error occurred during conversion. Please try again.')
            return jsonify({'error': error_msg}), 500
    finally:
        # Always release semaphore, even if an exception occurs
        if semaphore_acquired:
            try:
                conversion_semaphore.release()
                # Remove from tracking
                if request_id:
                    with _semaphore_lock:
                        _semaphore_acquisition_time.pop(request_id, None)
                else:
                    # Fallback: remove oldest entry if request_id not available
                    with _semaphore_lock:
                        if _semaphore_acquisition_time:
                            oldest_key = min(_semaphore_acquisition_time.keys(), 
                                           key=lambda k: _semaphore_acquisition_time[k])
                            _semaphore_acquisition_time.pop(oldest_key, None)
            except Exception as e:
                current_app.logger.error(f'Error releasing semaphore: {e}', exc_info=True)
        
        # Clean up any remaining threads
        if conversion_thread and conversion_thread.is_alive():
            current_app.logger.warning('Conversion thread still alive, attempting to join...')
            conversion_thread.join(timeout=1)
        
        if monitor_thread and monitor_thread.is_alive():
            current_app.logger.warning('Monitor thread still alive, attempting to join...')
            monitor_thread.join(timeout=1)
        
        # Final cleanup of temp directories if they still exist
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                current_app.logger.error(f'Final cleanup error for temp_dir: {e}')
        
        if output_dir and os.path.exists(output_dir):
            try:
                shutil.rmtree(output_dir)
            except Exception as e:
                current_app.logger.error(f'Final cleanup error for output_dir: {e}')

@main.route('/download/<filename>')
def download_file(filename):
    try:
        # Validate filename to prevent path traversal
        if not filename or '..' in filename or '/' in filename or '\\' in filename:
            return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 400
        
        # Sanitize filename
        safe_filename = secure_filename(filename)
        if not safe_filename:
            return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 400
        
        file_path = os.path.join(current_app.config['DOWNLOAD_FOLDER'], safe_filename)
        
        # Additional path validation
        real_download_path = os.path.realpath(current_app.config['DOWNLOAD_FOLDER'])
        real_file_path = os.path.realpath(file_path)
        
        if not real_file_path.startswith(real_download_path):
            return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 403
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 404
        
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 500 

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