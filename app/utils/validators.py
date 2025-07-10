import os
import pandas as pd
from typing import List, Dict, Tuple, Optional
from werkzeug.datastructures import FileStorage
import logging
import mimetypes

logger = logging.getLogger(__name__)

class FileValidator:
    """Comprehensive file validation for security and functionality"""
    
    # File size limits (in bytes)
    MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB per file
    MAX_TOTAL_SIZE = 200 * 1024 * 1024  # 200MB total
    
    # Allowed MIME types
    ALLOWED_MIME_TYPES = {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',  # .docx
        'application/msword',  # .doc
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',  # .xlsx
        'application/vnd.ms-excel',  # .xls
    }
    
    # Allowed file extensions
    ALLOWED_EXTENSIONS = {'docx', 'doc', 'xlsx', 'xls'}
    
    @staticmethod
    def validate_file_upload(files: List[FileStorage]) -> Tuple[bool, str, List[FileStorage]]:
        """
        Comprehensive file upload validation
        
        Returns:
            Tuple[bool, str, List[FileStorage]]: (is_valid, error_message, valid_files)
        """
        if not files:
            return False, "No files provided", []
        
        valid_files = []
        total_size = 0
        
        for file in files:
            # Check if file object is valid
            if not file or not hasattr(file, 'filename'):
                continue
                
            # Check filename
            if not file.filename or file.filename.strip() == '':
                continue
            
            # Check file extension
            if not FileValidator._has_valid_extension(file.filename):
                return False, f"Invalid file type: {file.filename}. Only .docx, .doc, .xlsx, .xls files are allowed.", []
            
            # Check file size
            file.seek(0, 2)  # Seek to end
            file_size = file.tell()
            file.seek(0)  # Reset to beginning
            
            if file_size > FileValidator.MAX_FILE_SIZE:
                return False, f"File {file.filename} is too large. Maximum size is {FileValidator.MAX_FILE_SIZE // (1024*1024)}MB.", []
            
            total_size += file_size
            
            # Validate MIME type using mimetypes module
            try:
                # Get MIME type from file extension
                mime_type, _ = mimetypes.guess_type(file.filename or '')
                file.seek(0)  # Reset to beginning
                
                if mime_type and mime_type not in FileValidator.ALLOWED_MIME_TYPES:
                    return False, f"Invalid file type detected: {file.filename}. Please upload a valid Word or Excel file.", []
                    
            except Exception as e:
                logger.warning(f"Could not determine MIME type for {file.filename}: {e}")
                # Continue with extension-based validation as fallback
        
        # Check total size
        if total_size > FileValidator.MAX_TOTAL_SIZE:
            return False, f"Total file size ({total_size // (1024*1024)}MB) exceeds limit ({FileValidator.MAX_TOTAL_SIZE // (1024*1024)}MB).", []
        
        valid_files = [f for f in files if f and f.filename and FileValidator._has_valid_extension(f.filename)]
        
        if not valid_files:
            return False, "No valid files found.", []
        
        return True, "", valid_files
    
    @staticmethod
    def _has_valid_extension(filename: str) -> bool:
        """Check if file has valid extension"""
        if not filename or '.' not in filename:
            return False
        return filename.rsplit('.', 1)[1].lower() in FileValidator.ALLOWED_EXTENSIONS
    
    @staticmethod
    def validate_excel_structure(file_path: str, required_columns: Optional[List[str]] = None) -> Tuple[bool, str, Optional[pd.DataFrame]]:
        """
        Validate Excel file structure and required columns
        
        Args:
            file_path: Path to Excel file
            required_columns: List of required column names
            
        Returns:
            Tuple[bool, str, Optional[pd.DataFrame]]: (is_valid, error_message, dataframe)
        """
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                return False, "Excel file not found", None
            
            # Check file size
            file_size = os.path.getsize(file_path)
            if file_size > FileValidator.MAX_FILE_SIZE:
                return False, f"Excel file is too large ({file_size // (1024*1024)}MB). Maximum size is {FileValidator.MAX_FILE_SIZE // (1024*1024)}MB.", None
            
            # Read Excel file
            try:
                df = pd.read_excel(file_path)
            except Exception as e:
                return False, f"Error reading Excel file: {str(e)}", None
            
            # Check if dataframe is empty
            if df.empty:
                return False, "Excel file is empty", None
            
            # Check if dataframe has too many rows
            if len(df) > 1000:  # Reasonable limit
                return False, f"Excel file has too many rows ({len(df)}). Maximum allowed is 1000.", None
            
            # Check required columns if specified
            if required_columns:
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    return False, f"Missing required columns: {', '.join(missing_columns)}", None
            
            return True, "", df
            
        except Exception as e:
            logger.error(f"Error validating Excel file {file_path}: {e}")
            return False, f"Error validating Excel file: {str(e)}", None
    
    @staticmethod
    def validate_template_file(template_path: str) -> Tuple[bool, str]:
        """
        Validate template file exists and is accessible
        
        Returns:
            Tuple[bool, str]: (is_valid, error_message)
        """
        if not template_path:
            return False, "Template path is required"
        
        if not os.path.exists(template_path):
            return False, f"Template file not found: {template_path}"
        
        if not os.path.isfile(template_path):
            return False, f"Template path is not a file: {template_path}"
        
        # Check if file is readable
        try:
            with open(template_path, 'rb') as f:
                f.read(1024)  # Try to read first 1KB
        except Exception as e:
            return False, f"Cannot read template file: {str(e)}"
        
        return True, ""
    
    @staticmethod
    def validate_libreoffice_installation() -> Tuple[bool, str]:
        """
        Check if LibreOffice is installed and accessible
        
        Returns:
            Tuple[bool, str]: (is_installed, error_message)
        """
        import platform
        import subprocess
        
        if platform.system() == "Windows":
            soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
            if not os.path.exists(soffice_path):
                return False, "LibreOffice not found. Please install LibreOffice to convert documents."
        else:
            soffice_path = 'soffice'
        
        try:
            # Test if LibreOffice can be executed
            result = subprocess.run([soffice_path, '--version'], 
                                  capture_output=True, timeout=10)
            if result.returncode != 0:
                return False, f"LibreOffice test failed: {result.stderr.decode()}"
            return True, ""
        except subprocess.TimeoutExpired:
            return False, "LibreOffice test timed out"
        except FileNotFoundError:
            return False, "LibreOffice not found in PATH"
        except Exception as e:
            return False, f"Error testing LibreOffice: {str(e)}"
    
    @staticmethod
    def sanitize_filename(filename: str) -> str:
        """
        Sanitize filename to prevent path traversal and other security issues
        
        Args:
            filename: Original filename
            
        Returns:
            str: Sanitized filename
        """
        import re
        from pathlib import Path
        
        # Remove path separators and other dangerous characters
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
        
        # Remove leading/trailing spaces and dots
        sanitized = sanitized.strip('. ')
        
        # Limit length
        if len(sanitized) > 100:
            name, ext = os.path.splitext(sanitized)
            sanitized = name[:100-len(ext)] + ext
        
        # Ensure it's not empty
        if not sanitized:
            sanitized = "file"
        
        return sanitized
    
    @staticmethod
    def validate_output_directory(output_dir: str) -> Tuple[bool, str]:
        """
        Validate output directory is writable
        
        Returns:
            Tuple[bool, str]: (is_valid, error_message)
        """
        if not output_dir:
            return False, "Output directory is required"
        
        try:
            # Create directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)
            
            # Test if directory is writable
            test_file = os.path.join(output_dir, '.test_write')
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
            
            return True, ""
        except Exception as e:
            return False, f"Cannot write to output directory: {str(e)}" 