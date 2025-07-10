import os
import mimetypes
import hashlib
import tempfile
from typing import Tuple, List, Optional
import logging

logger = logging.getLogger(__name__)

class FileSecurity:
    """Comprehensive file security and validation utilities"""
    
    ALLOWED_EXTENSIONS = {'.docx', '.doc', '.xlsx'}
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
    MAX_FILES_PER_REQUEST = 100
    
    @staticmethod
    def validate_file_upload(files: List, max_files: Optional[int] = None) -> Tuple[bool, str, List]:
        """
        Comprehensive file upload validation
        
        Returns:
            Tuple[bool, str, List]: (is_valid, error_message, valid_files)
        """
        if not files:
            return False, "No files provided", []
        
        if max_files is None:
            max_files = FileSecurity.MAX_FILES_PER_REQUEST
            
        if len(files) > max_files:
            return False, f"Too many files. Maximum {max_files} files allowed.", []
        
        valid_files = []
        
        for file in files:
            if not file or not file.filename:
                continue
                
            # Check file extension
            is_valid, error_msg = FileSecurity.validate_file_extension(file.filename)
            if not is_valid:
                return False, error_msg, []
            
            # Check file size
            file.seek(0, 2)  # Seek to end
            file_size = file.tell()
            file.seek(0)  # Reset to beginning
            
            if file_size > FileSecurity.MAX_FILE_SIZE:
                return False, f"File {file.filename} is too large. Maximum size is 100MB.", []
            
            # Check MIME type
            is_valid, error_msg = FileSecurity.validate_mime_type(file)
            if not is_valid:
                return False, error_msg, []
            
            # Sanitize filename
            safe_filename = FileSecurity.sanitize_filename(file.filename)
            if not safe_filename:
                return False, f"Invalid filename: {file.filename}", []
            
            valid_files.append(file)
        
        if not valid_files:
            return False, "No valid files found", []
        
        return True, "", valid_files
    
    @staticmethod
    def validate_file_extension(filename: str) -> Tuple[bool, str]:
        """Validate file extension"""
        if not filename or '.' not in filename:
            return False, "Invalid filename format"
        
        extension = filename.lower().rsplit('.', 1)[1]
        if extension not in {ext[1:] for ext in FileSecurity.ALLOWED_EXTENSIONS}:
            return False, f"File type .{extension} not allowed. Supported types: {', '.join(FileSecurity.ALLOWED_EXTENSIONS)}"
        
        return True, ""
    
    @staticmethod
    def validate_mime_type(file) -> Tuple[bool, str]:
        """Validate MIME type of uploaded file"""
        try:
            # Read first few bytes to detect MIME type
            file.seek(0)
            header = file.read(512)
            file.seek(0)
            
            # Check for common Office file signatures
            if header.startswith(b'PK\x03\x04'):  # ZIP-based formats (DOCX, XLSX)
                return True, ""
            elif header.startswith(b'\xd0\xcf\x11\xe0'):  # OLE format (DOC, XLS)
                return True, ""
            else:
                return False, "File does not appear to be a valid Office document"
                
        except Exception as e:
            logger.error(f"Error validating MIME type: {e}")
            return False, "Unable to validate file type"
    
    @staticmethod
    def sanitize_filename(filename: str) -> Optional[str]:
        """Sanitize filename for security"""
        if not filename:
            return None
        
        # Remove path traversal attempts
        filename = os.path.basename(filename)
        
        # Remove or replace dangerous characters
        dangerous_chars = ['<', '>', ':', '"', '|', '?', '*', '\\', '/']
        for char in dangerous_chars:
            filename = filename.replace(char, '_')
        
        # Limit length
        if len(filename) > 255:
            name, ext = os.path.splitext(filename)
            filename = name[:255-len(ext)] + ext
        
        return filename if filename.strip() else None
    
    @staticmethod
    def create_secure_temp_file(prefix: str = "upload_", suffix: str = "") -> Tuple[str, str]:
        """
        Create a secure temporary file with proper permissions
        
        Returns:
            Tuple[str, str]: (file_path, file_descriptor)
        """
        try:
            fd, path = tempfile.mkstemp(prefix=prefix, suffix=suffix)
            # Set restrictive permissions
            os.chmod(path, 0o600)
            return path, str(fd)
        except Exception as e:
            logger.error(f"Error creating secure temp file: {e}")
            raise
    
    @staticmethod
    def calculate_file_hash(file_path: str, algorithm: str = 'sha256') -> str:
        """Calculate file hash for integrity checking"""
        try:
            hash_obj = hashlib.new(algorithm)
            with open(file_path, 'rb') as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_obj.update(chunk)
            return hash_obj.hexdigest()
        except Exception as e:
            logger.error(f"Error calculating file hash: {e}")
            return ""
    
    @staticmethod
    def cleanup_temp_files(file_paths: List[str]) -> None:
        """Safely cleanup temporary files"""
        for file_path in file_paths:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logger.info(f"Cleaned up temp file: {file_path}")
            except Exception as e:
                logger.error(f"Error cleaning up temp file {file_path}: {e}")
    
    @staticmethod
    def validate_directory_path(path: str, base_dir: str) -> bool:
        """Validate that a path is within the base directory (path traversal protection)"""
        try:
            real_path = os.path.realpath(path)
            real_base = os.path.realpath(base_dir)
            return real_path.startswith(real_base)
        except Exception:
            return False 