"""Helpers for Excel validation, naming, and conversion summaries."""
import os
import re
from typing import List, Optional, Set, Tuple

import pandas as pd

from app.utils.validators import FileValidator

MAX_EXCEL_FILES = 20
MAX_ROWS_PER_SHEET = 1000

TRAINING_REQUIRED_COLUMNS = ['Name', 'Status', 'Gender']
APPOINTMENT_REQUIRED_COLUMNS = ['Name', 'Designation']
TRAINING_OPTIONAL_COLUMNS = ['EmpCode', 'Date', 'Start_date', 'End_date']
APPOINTMENT_OPTIONAL_COLUMNS = ['Place of Joining', 'Email', 'Contact', 'Date of Joining']


def find_column(columns, target_name: str) -> Optional[str]:
    target = target_name.strip().lower()
    for col in columns:
        if str(col).strip().lower() == target:
            return str(col)
    return None


def is_training_workbook(columns) -> bool:
    return find_column(columns, 'Status') is not None


def is_trainee_designation(designation_value) -> bool:
    if designation_value is None or (isinstance(designation_value, float) and pd.isna(designation_value)):
        return False
    return re.search(r'\btrainee\b', str(designation_value).strip(), re.IGNORECASE) is not None


def sanitize_person_name(name) -> str:
    """Keep the person's name readable; only strip characters invalid in filenames."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return 'Candidate'
    cleaned = str(name).strip()
    cleaned = re.sub(r'[<>:"/\\|?*]', '_', cleaned)
    return cleaned or 'Candidate'


def get_emp_code_from_row(row, columns) -> str:
    for col_name in ('EmpCode', 'Employee Code', 'Emp Code', 'Employee ID'):
        col = find_column(columns, col_name)
        if col:
            value = row.get(col, '')
            if value is not None and not (isinstance(value, float) and pd.isna(value)):
                text = str(value).strip()
                if text:
                    return FileValidator.sanitize_filename(text)
    return ''


def get_required_columns_for_workbook(columns) -> List[str]:
    if is_training_workbook(columns):
        return list(TRAINING_REQUIRED_COLUMNS)
    return list(APPOINTMENT_REQUIRED_COLUMNS)


def validate_workbook_columns(df: pd.DataFrame) -> Tuple[bool, str]:
    required = get_required_columns_for_workbook(df.columns)
    missing = [col for col in required if find_column(df.columns, col) is None]
    if missing:
        sheet_type = 'training completion' if is_training_workbook(df.columns) else 'appointment'
        return False, (
            f'Missing required column(s) for {sheet_type} letters: {", ".join(missing)}. '
            f'Expected columns like: {", ".join(required)}.'
        )
    return True, ''


def count_skipped_training_rows(df: pd.DataFrame) -> int:
    status_col = find_column(df.columns, 'Status')
    if status_col is None:
        return 0
    return int(df[status_col].apply(
        lambda v: str(v).strip().lower() != 'completed' if not pd.isna(v) else True
    ).sum())


def count_eligible_rows(df: pd.DataFrame) -> int:
    if is_training_workbook(df.columns):
        status_col = find_column(df.columns, 'Status')
        return int(df[status_col].apply(
            lambda v: str(v).strip().lower() == 'completed' if not pd.isna(v) else False
        ).sum())
    return len(df)


def is_completed_status(status_value) -> bool:
    if status_value is None or pd.isna(status_value):
        return False
    return str(status_value).strip().lower() == 'completed'


def build_pdf_filename(
    letter_type: str,
    name: str,
    used_names: Optional[Set[str]] = None,
) -> str:
    """Build PDF filename: Appointment/Training letter - {Name}.pdf"""
    safe_name = sanitize_person_name(name)

    if letter_type == 'training':
        base = f'Training letter- {safe_name}'
    elif letter_type == 'trainee':
        base = f'Appointment Letter and Training Agreement - {safe_name}'
    else:
        base = f'Appointment Letter and Employment Agreement - {safe_name}'

    filename = f'{base}.pdf'
    if used_names is not None:
        counter = 2
        while filename in used_names:
            filename = f'{base} ({counter}).pdf'
            counter += 1
        used_names.add(filename)
    return filename


class ConversionSummary:
    """Tracks skipped rows, warnings, and errors only."""

    def __init__(self):
        self.skipped: List[str] = []
        self.errors: List[str] = []
        self.warnings: List[str] = []

    def add_skipped(self, message: str):
        self.skipped.append(message)

    def add_error(self, message: str):
        self.errors.append(message)

    def add_warning(self, message: str):
        self.warnings.append(message)

    def has_issues(self) -> bool:
        return bool(self.skipped or self.warnings or self.errors)

    def has_skipped_or_warnings(self) -> bool:
        return bool(self.skipped or self.warnings)

    def to_text(self) -> str:
        if not self.has_skipped_or_warnings() and not self.errors:
            return ''

        lines = []
        if self.skipped:
            lines.append('Skipped rows:')
            for item in self.skipped:
                lines.append(f'  - {item}')
            lines.append('')

        if self.warnings:
            lines.append('Warnings:')
            for item in self.warnings:
                lines.append(f'  - {item}')
            lines.append('')

        if self.errors:
            lines.append('Errors:')
            for item in self.errors:
                lines.append(f'  - {item}')
            lines.append('')

        return '\n'.join(lines).strip() + '\n'


def validate_excel_upload_files(files) -> Tuple[bool, str, List]:
    """Validate uploaded Excel files (type, count, total size)."""
    if not files:
        return False, 'No files selected. Please choose at least one Excel file (.xlsx).', []

    valid = [f for f in files if f and f.filename and str(f.filename).lower().endswith('.xlsx')]
    if not valid:
        return False, 'Only Excel files (.xlsx) are allowed.', []

    if len(valid) > MAX_EXCEL_FILES:
        return False, f'You can upload up to {MAX_EXCEL_FILES} Excel files at a time.', []

    total_size = 0
    for file in valid:
        file.seek(0, 2)
        size = file.tell()
        file.seek(0)
        if size > FileValidator.MAX_FILE_SIZE:
            return False, (
                f'File "{file.filename}" is too large. '
                f'Maximum size is {FileValidator.MAX_FILE_SIZE // (1024 * 1024)}MB per file.'
            ), []
        total_size += size

    if total_size > FileValidator.MAX_TOTAL_SIZE:
        return False, (
            f'Total upload size ({total_size // (1024 * 1024)}MB) exceeds the '
            f'{FileValidator.MAX_TOTAL_SIZE // (1024 * 1024)}MB limit.'
        ), []

    return True, '', valid


def validate_templates_exist() -> Tuple[bool, str]:
    from app.template_config import ALL_TEMPLATE_NAMES, sample_path

    missing = [name for name in ALL_TEMPLATE_NAMES if not os.path.exists(sample_path(name))]
    if missing:
        return False, f'Missing template file(s) in samples folder: {", ".join(missing)}'
    return True, ''
