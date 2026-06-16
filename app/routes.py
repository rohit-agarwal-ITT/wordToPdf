from flask import Blueprint, render_template, request, jsonify, send_file, current_app
import os
import uuid
import re
from werkzeug.utils import secure_filename
from app.utils.word_processor import WordProcessor, OrdinalDateValue
from app.utils.validators import FileValidator
from app.utils.excel_helpers import (
    ConversionSummary,
    build_pdf_filename,
    count_eligible_rows,
    count_skipped_training_rows,
    find_column as _find_column_name,
    get_emp_code_from_row,
    is_completed_status,
    is_trainee_designation,
    is_training_workbook as is_training_excel,
    sanitize_person_name,
    validate_excel_upload_files,
    validate_templates_exist,
    validate_workbook_columns,
)
from app.template_config import (
    BANGALORE_TEMPLATE_NAME,
    JAIPUR_TEMPLATE_NAME,
    SAMPLE_FILES,
    SAMPLES_DIR,
    TRAINEE_TEMPLATE_NAME,
    TRAINING_TEMPLATE_NAME,
    sample_path,
)
import io
import zipfile
import tempfile
import shutil
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from app.utils.libreoffice_helper import convert_docx_files_to_pdf
from app.utils.error_handler import ErrorHandler
from datetime import datetime
import threading
import time

main = Blueprint('main', __name__)

ALLOWED_EXTENSIONS = {'xlsx'}  # Only Excel files allowed

# Progress tracking per conversion (supports concurrent users)
conversion_progress_store = {}
conversion_progress_lock = threading.Lock()
MAX_STORED_CONVERSIONS = 100
_rate_limit_lock = threading.Lock()
_rate_limit_buckets = {}

# Active progress mirror for the current in-flight conversion in this worker
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
    'conversion_id': None,
    'cancel_requested': False,
    'summary': [],
}

MAX_CONCURRENT_CONVERSIONS = int(os.environ.get('MAX_CONCURRENT_CONVERSIONS', '2'))
conversion_semaphore = threading.Semaphore(MAX_CONCURRENT_CONVERSIONS)
_semaphore_acquisition_time = {}
_semaphore_lock = threading.Lock()
RATE_LIMIT_REQUESTS = int(os.environ.get('RATE_LIMIT_REQUESTS', '30'))
RATE_LIMIT_WINDOW_SECONDS = int(os.environ.get('RATE_LIMIT_WINDOW_SECONDS', '300'))


def _prune_progress_store():
    if len(conversion_progress_store) <= MAX_STORED_CONVERSIONS:
        return
    oldest_ids = sorted(
        conversion_progress_store.keys(),
        key=lambda cid: conversion_progress_store[cid].get('start_time') or 0,
    )
    for cid in oldest_ids[: len(conversion_progress_store) - MAX_STORED_CONVERSIONS]:
        conversion_progress_store.pop(cid, None)


def _bind_progress_state(state):
    global conversion_progress
    conversion_progress = state


def _create_progress_state(conversion_id):
    state = {
        'status': 'converting',
        'current': 0,
        'total': 0,
        'message': 'Initializing new conversion...',
        'error': None,
        'percentage': 0,
        'eta_seconds': None,
        'start_time': time.time(),
        'elapsed_time': 0,
        'files': [],
        'display_total': 0,
        'display_current': 0,
        'conversion_id': conversion_id,
        'cancel_requested': False,
        'summary': [],
    }
    with conversion_progress_lock:
        conversion_progress_store[conversion_id] = state
        _prune_progress_store()
    _bind_progress_state(state)
    return state


def _parse_conversion_id(value):
    """Use client-provided UUID when valid; otherwise generate a new id."""
    if value:
        try:
            return str(uuid.UUID(str(value)))
        except (ValueError, AttributeError):
            pass
    return str(uuid.uuid4())


def _is_cancelled():
    with conversion_progress_lock:
        return bool(conversion_progress.get('cancel_requested'))


def _check_rate_limit():
    client_ip = request.remote_addr or 'unknown'
    now = time.time()
    with _rate_limit_lock:
        bucket = _rate_limit_buckets.setdefault(client_ip, [])
        bucket[:] = [ts for ts in bucket if now - ts < RATE_LIMIT_WINDOW_SECONDS]
        if len(bucket) >= RATE_LIMIT_REQUESTS:
            return False
        bucket.append(now)
    return True

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
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'xlsx'


def enrich_gender_placeholders(data, gender_value):
    """Add salutation placeholders for training letters based on Gender."""
    gender = str(gender_value).strip().lower() if gender_value is not None and not pd.isna(gender_value) else ''
    if gender == 'male':
        data['Mr_Ms'] = 'Mr.'
        data['Mr_Mrs'] = 'Mr.'
        data['his_her'] = 'his'
        data['he_she'] = 'he'
        data['him_her'] = 'him'
    elif gender == 'female':
        data['Mr_Ms'] = 'Ms.'
        data['Mr_Mrs'] = 'Mrs.'
        data['his_her'] = 'her'
        data['he_she'] = 'she'
        data['him_her'] = 'her'
    else:
        data['Mr_Ms'] = ''
        data['Mr_Mrs'] = ''
        data['his_her'] = ''
        data['he_she'] = ''
        data['him_her'] = ''


def get_training_template_path():
    template_path = sample_path(TRAINING_TEMPLATE_NAME)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Training template not found: {template_path}")
    return template_path


def get_appointment_letter_type(designation_value):
    return 'trainee' if is_trainee_designation(designation_value) else 'employment'


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
        
        if is_trainee_designation(designation_value):
            template_path = sample_path(TRAINEE_TEMPLATE_NAME)

            if not os.path.exists(template_path):
                template_path = sample_path(JAIPUR_TEMPLATE_NAME)

            return template_path
    
    # For non-Trainee designations (Software Engineer, Junior Software Engineer, etc.)
    # Use location-based templates
    # Normalize location value (case-insensitive, strip whitespace)
    if location_value:
        location = str(location_value).strip().lower()
        
        # Check for Bangalore location (case-insensitive)
        if 'bangalore' in location or 'bengaluru' in location:
            template_name = BANGALORE_TEMPLATE_NAME
        else:
            # Default to Jaipur template
            template_name = JAIPUR_TEMPLATE_NAME
    else:
        # Default to Jaipur template if location is empty/None
        template_name = JAIPUR_TEMPLATE_NAME
    
    template_path = sample_path(template_name)

    if not os.path.exists(template_path):
        template_path = sample_path(JAIPUR_TEMPLATE_NAME)

    return template_path

def _ordinal_suffix(day):
    """Return st/nd/rd/th for a day of month."""
    if 11 <= day % 100 <= 13:
        return 'th'
    return {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')

def _parse_date_value(value):
    """Parse Excel/string date values into a datetime, or None if parsing fails."""
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        return value.to_pydatetime()

    if isinstance(value, datetime):
        return value

    date_str = str(value).strip()
    if not date_str or date_str.lower() in ['nan', 'none', '']:
        return None

    date_formats = [
        '%Y-%m-%d',
        '%m/%d/%Y',
        '%m-%d-%Y',
        '%d/%m/%Y',
        '%d-%m-%Y',
        '%Y/%m/%d',
        '%d-%B-%Y',
        '%d %B %Y',
    ]

    for fmt in date_formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue

    return None

def _format_ordinal_date(parsed_date):
    """Format as 4th February' 26 (ordinal suffix applied as superscript in Word)."""
    day = parsed_date.day
    month = parsed_date.strftime('%B')
    year = parsed_date.strftime('%y')
    return OrdinalDateValue(day, _ordinal_suffix(day), f" {month}' {year}")

def format_date_field(value, field_name):
    """
    Format date fields for Word templates as 4th February' 26
    (ordinal suffix rendered as superscript in the PDF).
    """
    field_name_normalized = str(field_name).strip().lower()
    ordinal_date_fields = {
        'date',
        'start_date',
        'end_date',
        'start date',
        'end date',
        'date of joining',
        'effective date',
    }

    if field_name_normalized not in ordinal_date_fields:
        return str(value)

    parsed_date = _parse_date_value(value)
    if not parsed_date:
        return str(value)

    return _format_ordinal_date(parsed_date)

def _generate_docx_from_row(i, row, df, temp_dir, file_prefix=''):
    """Generate one filled Word document from an Excel row."""
    data = {}
    for col in df.columns:
        col_str = str(col)
        value = row[col]
        if pd.isna(value):
            value = ''
        data[col_str] = format_date_field(value, col_str)

    status_col = _find_column_name(df.columns, 'Status')
    letter_type = 'employment'

    if status_col is not None:
        status_value = data.get(status_col, '')
        if not is_completed_status(status_value):
            return None
        letter_type = 'training'
        gender_col = _find_column_name(df.columns, 'Gender')
        gender_value = data.get(gender_col, '') if gender_col else ''
        enrich_gender_placeholders(data, gender_value)
        word_template = get_training_template_path()
    else:
        location_value = None
        for col in df.columns:
            col_str = str(col).strip()
            if col_str.lower() == 'place of joining':
                location_value = data.get(col_str)
                break

        if location_value is None:
            location_column_names = [
                'Location', 'location', 'LOCATION', 'City', 'city', 'CITY',
                'Location Name', 'location name', 'Place of Joining', 'place of joining',
            ]
            for loc_col in location_column_names:
                if loc_col in data:
                    location_value = data[loc_col]
                    break

        if location_value is None:
            for col in df.columns:
                col_lower = str(col).lower()
                if 'location' in col_lower or 'city' in col_lower or ('place' in col_lower and 'joining' in col_lower):
                    location_value = data.get(str(col))
                    break

        designation_value = None
        for col in df.columns:
            col_str = str(col).strip()
            if col_str.lower() == 'designation':
                designation_value = data.get(col_str)
                break

        if designation_value is None:
            designation_column_names = [
                'Designation', 'designation', 'DESIGNATION', 'Role', 'role', 'ROLE',
                'Job Title', 'job title',
            ]
            for desig_col in designation_column_names:
                if desig_col in data:
                    designation_value = data[desig_col]
                    break

        if designation_value is None:
            for col in df.columns:
                col_lower = str(col).lower()
                if 'designation' in col_lower or ('role' in col_lower and 'title' not in col_lower):
                    designation_value = data.get(str(col))
                    break

        word_template = get_template_path(location_value, designation_value)
        letter_type = get_appointment_letter_type(designation_value)

    name_part = sanitize_person_name(data.get('Name', 'Candidate'))
    emp_code = get_emp_code_from_row(row, df.columns)
    safe_docx_name = FileValidator.sanitize_filename(name_part)
    if file_prefix:
        docx_name = f"{file_prefix}_{safe_docx_name}_{i + 1}.docx"
    else:
        docx_name = f"{safe_docx_name}_{i + 1}.docx"
    docx_path = os.path.join(temp_dir, docx_name)
    WordProcessor().fill_placeholders(word_template, docx_path, data)
    return (docx_path, letter_type, name_part, i, emp_code)

@main.route('/')
def index():
    return render_template('index.html')


@main.route('/health')
def health_check():
    libreoffice_ok, libreoffice_error = FileValidator.validate_libreoffice_installation()
    templates_ok, templates_error = validate_templates_exist()
    is_ready = libreoffice_ok and templates_ok
    return jsonify({
        'status': 'ok' if is_ready else 'degraded',
        'libreoffice': {'ok': libreoffice_ok, 'message': libreoffice_error},
        'templates': {'ok': templates_ok, 'message': templates_error},
    }), 200 if is_ready else 503


@main.route('/samples/<path:filename>')
def download_sample(filename):
    if '..' in filename or filename.startswith(('/', '\\')):
        return jsonify({'error': 'Sample file not found.'}), 404
    safe_name = os.path.basename(filename)
    if safe_name not in SAMPLE_FILES:
        return jsonify({'error': 'Sample file not found.'}), 404
    file_path = sample_path(safe_name)
    if not os.path.exists(file_path):
        return jsonify({'error': 'Sample file not found.'}), 404
    return send_file(file_path, as_attachment=True, download_name=safe_name)


@main.route('/cancel', methods=['POST'])
def cancel_conversion():
    payload = request.get_json(silent=True) or {}
    conversion_id = payload.get('conversion_id')
    if not conversion_id:
        return jsonify({'error': 'conversion_id is required.'}), 400
    with conversion_progress_lock:
        state = conversion_progress_store.get(conversion_id)
        if not state:
            return jsonify({'error': 'Conversion not found or already completed.'}), 404
        state['cancel_requested'] = True
        state['message'] = 'Cancellation requested...'
    return jsonify({'status': 'cancelling', 'conversion_id': conversion_id})


@main.route('/progress')
def get_progress():
    """Return conversion progress for a specific conversion_id."""
    conversion_id = request.args.get('conversion_id')
    try:
        with conversion_progress_lock:
            if conversion_id and conversion_id in conversion_progress_store:
                progress_copy = conversion_progress_store[conversion_id].copy()
            else:
                progress_copy = conversion_progress.copy()
            progress_copy['files'] = list(progress_copy.get('files', []))
            progress_copy['summary'] = list(progress_copy.get('summary', []))
        return jsonify(progress_copy)
    except Exception as e:
        current_app.logger.error(f'Error getting progress: {e}', exc_info=True)
        return jsonify({
            'status': 'error',
            'error': 'Failed to retrieve progress information',
            'current': 0,
            'total': 0,
            'percentage': 0,
            'message': 'Error retrieving progress',
            'files': [],
            'conversion_id': conversion_id,
        }), 500

@main.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and conversion with thread safety and concurrency limits"""
    global conversion_progress
    
    # Record request start time for timeout tracking
    request_start_time = time.time()
    request_timeout = 600  # 10 minutes max request time
    
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
        if not _check_rate_limit():
            return jsonify({
                'error': 'Too many requests. Please wait a few minutes and try again.'
            }), 429

        new_conversion_id = _parse_conversion_id(request.form.get('conversion_id'))
        _create_progress_state(new_conversion_id)

        libreoffice_ok, libreoffice_error = FileValidator.validate_libreoffice_installation()
        if not libreoffice_ok:
            set_progress_status('error', error=libreoffice_error)
            return jsonify({'error': libreoffice_error}), 500

        templates_ok, templates_error = validate_templates_exist()
        if not templates_ok:
            set_progress_status('error', error=templates_error)
            return jsonify({'error': templates_error}), 500

        if 'files[]' not in request.files:
            return jsonify({'error': 'No files were uploaded. Please select at least one Excel file.'}), 400

        files = request.files.getlist('files[]')
        upload_ok, upload_error, excel_files = validate_excel_upload_files(files)
        if not upload_ok:
            set_progress_status('error', error=upload_error)
            return jsonify({'error': upload_error}), 400

        # Excel to Word to PDF batch logic (supports one or more .xlsx files)
        if excel_files:
            try:
                temp_dir = tempfile.mkdtemp()
                output_dir = tempfile.mkdtemp()
            except Exception as e:
                current_app.logger.error(f'Error creating temp directories: {e}', exc_info=True)
                return jsonify({'error': 'An error occurred while setting up conversion. Please try again.'}), 500

            pdf_files = []
            errors = []
            summary = ConversionSummary()
            used_zip_names = set()

            try:
                workbooks = []
                for excel_file in excel_files:
                    excel_filename = secure_filename(excel_file.filename)
                    excel_path = os.path.join(temp_dir, f"{len(workbooks)}_{excel_filename}")
                    try:
                        excel_file.seek(0)
                        excel_file.save(excel_path)
                    except Exception as e:
                        summary.add_error(f'{excel_filename}: could not be saved ({e})')
                        continue

                    try:
                        df = pd.read_excel(excel_path)
                    except Exception as e:
                        summary.add_error(f'{excel_filename}: could not be read ({e})')
                        continue

                    if df is None or df.empty:
                        summary.add_error(f'{excel_filename}: file is empty')
                        continue

                    if len(df) > 1000:
                        summary.add_error(f'{excel_filename}: has more than 1000 rows')
                        continue

                    columns_ok, columns_error = validate_workbook_columns(df)
                    if not columns_ok:
                        summary.add_error(f'{excel_filename}: {columns_error}')
                        continue

                    eligible_rows = count_eligible_rows(df)
                    skipped_rows = count_skipped_training_rows(df)
                    if eligible_rows == 0:
                        reason = (
                            'No records with Status "Completed"'
                            if is_training_excel(df.columns)
                            else 'No records to process'
                        )
                        summary.add_error(f'{excel_filename}: {reason}')
                        continue

                    file_prefix = os.path.splitext(excel_filename)[0] if len(excel_files) > 1 else ''
                    workbooks.append((excel_filename, df, file_prefix, skipped_rows))

                if not workbooks:
                    error_message = 'No valid records found in the uploaded Excel file(s). See summary for details.'
                    set_progress_status('error', error=error_message)
                    return jsonify({'error': error_message, 'summary': summary.to_text()}), 400

                total_rows = sum(count_eligible_rows(df) for _, df, _, _ in workbooks)
                total_steps = total_rows * 2
                multiple_workbooks = len(workbooks) > 1

                if len(workbooks) == 1:
                    zip_filename = f"{os.path.splitext(workbooks[0][0])[0]}.zip"
                else:
                    zip_filename = 'Appointment_letters.zip'

                with conversion_progress_lock:
                    conversion_progress['display_total'] = total_rows
                    conversion_progress['display_current'] = 0
                    conversion_progress['status'] = 'converting'
                    conversion_progress['message'] = (
                        f'Processing {total_rows} records from {len(workbooks)} Excel file(s)...'
                    )
                    conversion_progress['files'] = []
                    conversion_progress['current'] = 0
                    conversion_progress['total'] = total_steps
                    conversion_progress['percentage'] = 0
                    conversion_progress['conversion_id'] = new_conversion_id
                    conversion_progress['start_time'] = time.time()
                    conversion_progress['elapsed_time'] = 0
                    conversion_progress['eta_seconds'] = None
                    conversion_progress['summary'] = [
                        f'{name}: {count_eligible_rows(df)} letter(s)'
                        + (f', {skipped} skipped' if skipped else '')
                        for name, df, _, skipped in workbooks
                    ]

                update_progress(
                    0, total_steps,
                    'Preparing appointment letters for PDF generation...',
                    display_total=total_rows
                )

                all_docx_files = []
                for workbook_index, (excel_filename, df, file_prefix, skipped_rows) in enumerate(workbooks, start=1):
                    if _is_cancelled():
                        raise Exception('Conversion cancelled by user.')

                    if skipped_rows:
                        for row_idx, row in df.iterrows():
                            status_col = _find_column_name(df.columns, 'Status')
                            if status_col and not is_completed_status(row.get(status_col)):
                                name = sanitize_person_name(row.get('Name', f'Row {row_idx + 1}'))
                                summary.add_skipped(
                                    f'{excel_filename} row {row_idx + 1} ({name}): Status not Completed'
                                )

                    workbook_label = (
                        f' ({workbook_index}/{len(workbooks)}: {excel_filename})'
                        if multiple_workbooks else ''
                    )

                    def generate_docx(row_tuple, current_df=df, prefix=file_prefix, source_name=excel_filename):
                        if _is_cancelled():
                            return None
                        i, row = row_tuple
                        try:
                            return _generate_docx_from_row(i, row, current_df, temp_dir, prefix)
                        except Exception as e:
                            raise Exception(f"Error processing row {i + 1} in {source_name}: {str(e)}")

                    with ThreadPoolExecutor(max_workers=4) as executor:
                        futures = {
                            executor.submit(generate_docx, (i, row)): i for i, row in df.iterrows()
                        }
                        try:
                            for future in as_completed(futures):
                                if _is_cancelled():
                                    for pending in futures:
                                        pending.cancel()
                                    raise Exception('Conversion cancelled by user.')
                                result = future.result()
                                if result is not None:
                                    all_docx_files.append((*result, file_prefix))
                                    if result[1] == 'training':
                                        gender_col = _find_column_name(df.columns, 'Gender')
                                        row_index = result[3]
                                        gender_val = df.at[row_index, gender_col] if gender_col else ''
                                        if not str(gender_val).strip() or str(gender_val).strip().lower() not in ('male', 'female'):
                                            summary.add_warning(
                                                f'{excel_filename} row {row_index + 1}: missing or invalid Gender'
                                            )

                                rows_processed = len(all_docx_files)
                                progress_pct = rows_processed / total_rows * 0.3 if total_rows else 0
                                current_progress = int(total_steps * progress_pct)
                                update_progress(
                                    current_progress, total_steps,
                                    f'Preparing appointment letters... ({rows_processed}/{total_rows} records){workbook_label}',
                                    display_total=total_rows
                                )
                                with conversion_progress_lock:
                                    conversion_progress['display_current'] = max(
                                        conversion_progress.get('display_current', 0),
                                        rows_processed
                                    )
                        finally:
                            if _is_cancelled():
                                for pending in futures:
                                    pending.cancel()
                                executor.shutdown(wait=False, cancel_futures=True)

                if _is_cancelled():
                    raise Exception('Conversion cancelled by user.')

                update_progress(
                    total_rows, total_steps,
                    'Generating PDFs. This may take a moment...',
                    display_total=total_rows
                )

                try:
                    validated_files = []
                    for docx_file, _letter_type, _name, _row_idx, _emp_code, _file_prefix in all_docx_files:
                        if os.path.exists(docx_file) and os.path.isfile(docx_file):
                            real_temp_path = os.path.realpath(temp_dir)
                            real_file_path = os.path.realpath(docx_file)
                            if real_file_path.startswith(real_temp_path):
                                validated_files.append(docx_file)

                    if not validated_files:
                        raise Exception("No valid files found for conversion")

                    conversion_complete = threading.Event()
                    conversion_error = [None]

                    def monitor_pdf_conversion_excel():
                        expected_pdfs = {
                            os.path.splitext(os.path.basename(f))[0] + '.pdf': os.path.basename(f)
                            for f in validated_files
                        }
                        pdfs_found = set()
                        start_time = time.time()
                        max_wait_time = 300

                        while not conversion_complete.is_set() and (time.time() - start_time) < max_wait_time:
                            if _is_cancelled():
                                conversion_error[0] = Exception('Conversion cancelled by user.')
                                conversion_complete.set()
                                break
                            if os.path.exists(output_dir):
                                existing_pdfs = set(f for f in os.listdir(output_dir) if f.endswith('.pdf'))
                                new_pdfs = existing_pdfs - pdfs_found

                                for pdf_file in new_pdfs:
                                    if pdf_file in expected_pdfs:
                                        pdfs_found.add(pdf_file)
                                        progress_pct = 0.3 + (len(pdfs_found) / len(expected_pdfs)) * 0.6
                                        current_progress = int(total_steps * progress_pct)
                                        update_progress(
                                            current_progress, total_steps,
                                            'Creating PDFs...',
                                            display_total=total_rows
                                        )

                                if len(pdfs_found) == len(expected_pdfs):
                                    break

                            time.sleep(0.5)

                    def run_conversion_excel():
                        try:
                            if _is_cancelled():
                                raise Exception('Conversion cancelled by user.')
                            convert_docx_files_to_pdf(
                                validated_files,
                                output_dir,
                                timeout=300,
                                should_cancel=_is_cancelled,
                            )
                            conversion_complete.set()
                        except Exception as e:
                            conversion_error[0] = e
                            conversion_complete.set()

                    conversion_thread = threading.Thread(target=run_conversion_excel, daemon=True)
                    monitor_thread = threading.Thread(target=monitor_pdf_conversion_excel, daemon=True)
                    conversion_thread.start()
                    monitor_thread.start()
                    conversion_thread.join(timeout=300)
                    conversion_complete.set()
                    monitor_thread.join(timeout=5)

                    if conversion_error[0]:
                        raise conversion_error[0]

                except Exception as e:
                    current_app.logger.error(f'PDF conversion error: {e}', exc_info=True)
                    raise Exception(f"PDF conversion failed: {str(e)}")

                pdfs_collected = 0
                for docx_file, letter_type, name, row_idx, emp_code, file_prefix in all_docx_files:
                    base = os.path.splitext(os.path.basename(docx_file))[0]
                    pdf_name = build_pdf_filename(letter_type, name, used_zip_names)
                    pdf_path = os.path.join(output_dir, base + '.pdf')
                    zip_entry = f"{file_prefix}/{pdf_name}" if file_prefix else pdf_name

                    if os.path.exists(pdf_path):
                        pdf_files.append((pdf_path, zip_entry))
                        pdfs_collected += 1
                        progress_pct = 0.9 + (pdfs_collected / total_rows) * 0.1
                        current_progress = int(total_steps * progress_pct)
                        update_progress(
                            current_progress, total_steps,
                            'Finalizing PDFs...',
                            display_total=total_rows
                        )
                    else:
                        msg = f'PDF not found for {base}'
                        errors.append(msg)
                        summary.add_error(msg)

                if not pdf_files:
                    error_message = 'No PDFs were generated. ' + (
                        errors[0] if errors else 'Please verify LibreOffice is installed and templates are valid.'
                    )
                    set_progress_status('error', error=error_message)
                    if temp_dir and os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir)
                    if output_dir and os.path.exists(output_dir):
                        shutil.rmtree(output_dir)
                    return jsonify({'error': error_message, 'summary': summary.to_text()}), 500

                zip_start_time = time.time()
                update_progress(
                    total_steps, total_steps,
                    'All PDFs created! Creating ZIP package...',
                    display_total=total_rows
                )
                estimated_zip_time = min(10, max(3, len(pdf_files) * 0.1))
                set_progress_status('converting', eta_seconds=int(estimated_zip_time))
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as zip_file:
                    if summary.has_skipped_or_warnings():
                        zip_file.writestr('summary.txt', summary.to_text())
                    for idx, (pdf_path, zip_entry_name) in enumerate(pdf_files):
                        with open(pdf_path, 'rb') as pdf_file:
                            zip_file.writestr(zip_entry_name, pdf_file.read())
                        if (idx + 1) % 10 == 0 or idx == len(pdf_files) - 1:
                            update_progress(
                                total_steps, total_steps,
                                f'Creating ZIP package... ({idx + 1}/{len(pdf_files)} files)',
                                display_total=total_rows
                            )
                zip_creation_time = time.time() - zip_start_time
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                if output_dir and os.path.exists(output_dir):
                    shutil.rmtree(output_dir)
                zip_buffer.seek(0)
                update_progress(
                    total_steps, total_steps,
                    'Successfully created all PDF appointment letters! Download starting...',
                    display_total=total_rows
                )
                set_progress_status('completed', eta_seconds=0)
                if time.time() - request_start_time > request_timeout:
                    current_app.logger.error(f'Request timeout before sending file: {request_id}')
                    set_progress_status('error', error='Request timeout. Please try again.')
                    return jsonify({
                        'error': 'Request timeout. The conversion took too long. Please try again with a smaller file.'
                    }), 504

                try:
                    response = send_file(
                        zip_buffer,
                        as_attachment=True,
                        download_name=zip_filename,
                        mimetype='application/zip',
                        max_age=0,
                        conditional=True
                    )
                    response.headers['X-Conversion-Id'] = new_conversion_id
                    if summary.has_skipped_or_warnings():
                        response.headers['X-Has-Summary'] = 'true'
                    return response
                except Exception as send_error:
                    current_app.logger.error(f'Error sending file: {send_error}', exc_info=True)
                    set_progress_status('error', error='Error sending file. Please try again.')
                    return jsonify({'error': 'Error sending file. Please try again.'}), 500

            except Exception as e:
                current_app.logger.error(f'Excel conversion error: {e}', exc_info=True)
                error_message = str(e) or 'An error occurred during conversion. Please try again.'
                is_cancelled = 'cancelled by user' in error_message.lower()
                set_progress_status('error', error=error_message)
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
                status_code = 499 if is_cancelled else 500
                return jsonify({'error': error_message}), status_code

        return jsonify({'error': 'Only Excel files (.xlsx) are supported.'}), 400

    except Exception as e:
        current_app.logger.error(f'Upload error: {e}', exc_info=True)
        return jsonify({'error': 'An error occurred during conversion. Please try again.'}), 500
    finally:
        if semaphore_acquired:
            conversion_semaphore.release()
            if request_id:
                with _semaphore_lock:
                    _semaphore_acquisition_time.pop(request_id, None)

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