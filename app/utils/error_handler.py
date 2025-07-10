import logging
import traceback
import tempfile
import shutil
from typing import Dict, Any, Optional
from flask import current_app
import os

logger = logging.getLogger(__name__)

class ErrorHandler:
    """Comprehensive error handling and logging for the application"""
    
    @staticmethod
    def setup_logging():
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('app.log'),
                logging.StreamHandler()
            ]
        )
    
    @staticmethod
    def log_error(error: Exception, context: str = "", extra_data: Optional[Dict[str, Any]] = None):
        """Log error with context and extra data"""
        error_msg = f"Error in {context}: {str(error)}"
        if extra_data:
            error_msg += f" | Extra data: {extra_data}"
        
        logger.error(error_msg)
        logger.error(f"Traceback: {traceback.format_exc()}")
    
    @staticmethod
    def cleanup_temp_files(*temp_dirs):
        """Clean up temporary directories safely"""
        for temp_dir in temp_dirs:
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir)
                    logger.info(f"Cleaned up temp directory: {temp_dir}")
                except Exception as e:
                    logger.error(f"Failed to clean up temp directory {temp_dir}: {e}")
    
    @staticmethod
    def handle_conversion_error(error: Exception, temp_dirs: list, user_message: str = "Conversion failed") -> Dict[str, Any]:
        """Handle conversion errors with cleanup and logging"""
        ErrorHandler.log_error(error, "conversion", {"temp_dirs": temp_dirs})
        ErrorHandler.cleanup_temp_files(*temp_dirs)
        
        return {
            "error": user_message,
            "details": str(error) if current_app.debug else "Internal server error"
        }
    
    @staticmethod
    def handle_file_processing_error(error: Exception, file_path: str, user_message: str = "File processing failed") -> Dict[str, Any]:
        """Handle file processing errors"""
        ErrorHandler.log_error(error, "file_processing", {"file_path": file_path})
        
        # Clean up the problematic file
        if file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
                logger.info(f"Cleaned up problematic file: {file_path}")
            except Exception as e:
                logger.error(f"Failed to clean up file {file_path}: {e}")
        
        return {
            "error": user_message,
            "details": str(error) if current_app.debug else "Internal server error"
        }
    
    @staticmethod
    def handle_validation_error(error: Exception, context: str = "validation") -> Dict[str, Any]:
        """Handle validation errors"""
        ErrorHandler.log_error(error, context)
        
        return {
            "error": "Validation failed",
            "details": str(error)
        }
    
    @staticmethod
    def handle_system_error(error: Exception, context: str = "system") -> Dict[str, Any]:
        """Handle system-level errors"""
        ErrorHandler.log_error(error, context)
        
        return {
            "error": "System error occurred",
            "details": str(error) if current_app.debug else "Internal server error"
        }
    
    @staticmethod
    def validate_system_requirements() -> Dict[str, Any]:
        """Validate system requirements and dependencies"""
        errors = []
        warnings = []
        
        # Check LibreOffice installation
        from app.utils.validators import FileValidator
        libreoffice_ok, libreoffice_error = FileValidator.validate_libreoffice_installation()
        if not libreoffice_ok:
            errors.append(f"LibreOffice: {libreoffice_error}")
        
        # Check upload directory
        upload_dir = current_app.config.get('UPLOAD_FOLDER')
        if upload_dir:
            upload_ok, upload_error = FileValidator.validate_output_directory(upload_dir)
            if not upload_ok:
                errors.append(f"Upload directory: {upload_error}")
        
        # Check download directory
        download_dir = current_app.config.get('DOWNLOAD_FOLDER')
        if download_dir:
            download_ok, download_error = FileValidator.validate_output_directory(download_dir)
            if not download_ok:
                errors.append(f"Download directory: {download_error}")
        
        # Check template file
        template_path = os.path.join('samples', 'sample_document_for_placeholder.docx')
        template_ok, template_error = FileValidator.validate_template_file(template_path)
        if not template_ok:
            warnings.append(f"Template file: {template_error}")
        
        return {
            "errors": errors,
            "warnings": warnings,
            "is_ready": len(errors) == 0
        }
    
    @staticmethod
    def create_error_response(error_dict: Dict[str, Any], status_code: int = 500) -> tuple:
        """Create standardized error response"""
        from flask import jsonify
        
        response = jsonify(error_dict)
        response.status_code = status_code
        return response
    
    @staticmethod
    def handle_timeout_error(operation: str, timeout_seconds: int) -> Dict[str, Any]:
        """Handle timeout errors"""
        error_msg = f"{operation} timed out after {timeout_seconds} seconds"
        logger.error(error_msg)
        
        return {
            "error": "Operation timed out",
            "details": f"The {operation} took too long to complete. Please try with smaller files or fewer files."
        }
    
    @staticmethod
    def handle_memory_error(operation: str) -> Dict[str, Any]:
        """Handle memory-related errors"""
        error_msg = f"Memory error during {operation}"
        logger.error(error_msg)
        
        return {
            "error": "Memory limit exceeded",
            "details": "The operation requires more memory than available. Please try with smaller files or fewer files."
        }
    
    @staticmethod
    def handle_disk_space_error(operation: str, required_space: int) -> Dict[str, Any]:
        """Handle disk space errors"""
        error_msg = f"Insufficient disk space for {operation}. Required: {required_space} bytes"
        logger.error(error_msg)
        
        return {
            "error": "Insufficient disk space",
            "details": f"The operation requires {required_space // (1024*1024)}MB of disk space. Please free up space and try again."
        }
    
    @staticmethod
    def check_disk_space(path: str, required_bytes: int) -> bool:
        """Check if there's enough disk space"""
        try:
            stat = os.statvfs(path)
            free_bytes = stat.f_frsize * stat.f_bavail
            return free_bytes >= required_bytes
        except Exception as e:
            logger.warning(f"Could not check disk space: {e}")
            return True  # Assume OK if we can't check
    
    @staticmethod
    def get_system_info() -> Dict[str, Any]:
        """Get system information for debugging"""
        import platform
        import psutil
        
        try:
            return {
                "platform": platform.system(),
                "python_version": platform.python_version(),
                "cpu_count": psutil.cpu_count(),
                "memory_total": psutil.virtual_memory().total,
                "memory_available": psutil.virtual_memory().available,
                "disk_usage": psutil.disk_usage('/').percent if os.path.exists('/') else None
            }
        except Exception as e:
            logger.error(f"Error getting system info: {e}")
            return {"error": str(e)} 