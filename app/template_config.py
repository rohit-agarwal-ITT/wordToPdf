"""Word template filenames and sample Excel references."""
import os

_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
SAMPLES_DIR = os.path.join(_PROJECT_ROOT, 'samples')

TRAINEE_TEMPLATE_NAME = 'Appointment Letter and Training Agreement.docx'
JAIPUR_TEMPLATE_NAME = 'Appointment Letter and Employment Agreement - Jaipur.docx'
BANGALORE_TEMPLATE_NAME = 'Appointment Letter and Employment Agreement - Bangalore.docx'
TRAINING_TEMPLATE_NAME = 'Training letter.docx'

SAMPLE_EXCEL_EMPLOYMENT = 'Appointment Letter and Employment Agreement - JaipurBangalore.xlsx'
SAMPLE_EXCEL_TRAINEE = 'Appointment Letter and Training Agreement.xlsx'
SAMPLE_EXCEL_TRAINING = 'Training letter.xlsx'

ALL_TEMPLATE_NAMES = [
    TRAINEE_TEMPLATE_NAME,
    JAIPUR_TEMPLATE_NAME,
    BANGALORE_TEMPLATE_NAME,
    TRAINING_TEMPLATE_NAME,
]

SAMPLE_FILES = [
    SAMPLE_EXCEL_EMPLOYMENT,
    SAMPLE_EXCEL_TRAINEE,
    SAMPLE_EXCEL_TRAINING,
]


def sample_path(filename: str) -> str:
    return os.path.join(SAMPLES_DIR, filename)
