import os
import json
from typing import Any, Dict, Optional
from pathlib import Path

class ConfigManager:
    """Centralized configuration management for the application"""
    
    def __init__(self):
        self.config = {}
        self.load_config()
    
    def load_config(self):
        """Load configuration from environment variables and config files"""
        # Default configuration
        self.config = {
            # Application settings
            'APP_NAME': 'Word to PDF Converter',
            'APP_VERSION': '1.0.0',
            'DEBUG': False,
            'SECRET_KEY': os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production'),
            
            # File upload settings
            'MAX_FILE_SIZE': 100 * 1024 * 1024,  # 100MB
            'MAX_FILES_PER_REQUEST': 100,
            'ALLOWED_EXTENSIONS': {'.docx', '.doc', '.xlsx'},
            'UPLOAD_FOLDER': 'static/uploads',
            'DOWNLOAD_FOLDER': 'static/downloads',
            
            # Conversion settings
            'LIBREOFFICE_PATH': self._get_libreoffice_path(),
            'CONVERSION_TIMEOUT': 300,  # 5 minutes
            'BATCH_SIZE': 10,  # Number of files to process in parallel
            
            # Performance settings
            'ENABLE_PERFORMANCE_MONITORING': True,
            'CLEANUP_INTERVAL': 3600,  # 1 hour
            'MAX_LOG_SIZE': 10 * 1024 * 1024,  # 10MB
            
            # Security settings
            'ENABLE_FILE_VALIDATION': True,
            'ENABLE_MIME_CHECK': True,
            'ENABLE_PATH_TRAVERSAL_PROTECTION': True,
            
            # Feature flags
            'ENABLE_BATCH_PROCESSING': True,
            'ENABLE_EXCEL_TO_PDF': True,
            'ENABLE_PROGRESS_TRACKING': True,
            'ENABLE_DOWNLOAD_COMPRESSION': True,
            
            # UI settings
            'ENABLE_DRAG_DROP': True,
            'ENABLE_FILE_PREVIEW': False,
            'MAX_PREVIEW_SIZE': 1024 * 1024,  # 1MB
        }
        
        # Override with environment variables
        self._load_from_env()
        
        # Override with config file if exists
        self._load_from_file()
    
    def _get_libreoffice_path(self) -> str:
        """Get LibreOffice path based on platform"""
        import platform
        if platform.system() == "Windows":
            return r'C:\Program Files\LibreOffice\program\soffice.exe'
        else:
            return 'soffice'
    
    def _load_from_env(self):
        """Load configuration from environment variables"""
        env_mappings = {
            'DEBUG': ('DEBUG', bool),
            'MAX_FILE_SIZE': ('MAX_FILE_SIZE', int),
            'MAX_FILES_PER_REQUEST': ('MAX_FILES_PER_REQUEST', int),
            'CONVERSION_TIMEOUT': ('CONVERSION_TIMEOUT', int),
            'BATCH_SIZE': ('BATCH_SIZE', int),
            'ENABLE_PERFORMANCE_MONITORING': ('ENABLE_PERFORMANCE_MONITORING', bool),
            'ENABLE_FILE_VALIDATION': ('ENABLE_FILE_VALIDATION', bool),
            'ENABLE_BATCH_PROCESSING': ('ENABLE_BATCH_PROCESSING', bool),
            'ENABLE_EXCEL_TO_PDF': ('ENABLE_EXCEL_TO_PDF', bool),
            'ENABLE_PROGRESS_TRACKING': ('ENABLE_PROGRESS_TRACKING', bool),
        }
        
        for env_var, (config_key, value_type) in env_mappings.items():
            env_value = os.environ.get(env_var)
            if env_value is not None:
                try:
                    if value_type == bool:
                        self.config[config_key] = env_value.lower() in ('true', '1', 'yes')
                    else:
                        self.config[config_key] = value_type(env_value)
                except (ValueError, TypeError):
                    pass  # Keep default value
    
    def _load_from_file(self):
        """Load configuration from config file"""
        config_file = Path('config.json')
        if config_file.exists():
            try:
                with open(config_file, 'r') as f:
                    file_config = json.load(f)
                    self.config.update(file_config)
            except (json.JSONDecodeError, IOError):
                pass  # Keep default configuration
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value"""
        return self.config.get(key, default)
    
    def set(self, key: str, value: Any):
        """Set configuration value"""
        self.config[key] = value
    
    def get_all(self) -> Dict[str, Any]:
        """Get all configuration values"""
        return self.config.copy()
    
    def is_enabled(self, feature: str) -> bool:
        """Check if a feature is enabled"""
        return self.config.get(f'ENABLE_{feature.upper()}', False)
    
    def get_file_settings(self) -> Dict[str, Any]:
        """Get file-related settings"""
        return {
            'max_file_size': self.get('MAX_FILE_SIZE'),
            'max_files_per_request': self.get('MAX_FILES_PER_REQUEST'),
            'allowed_extensions': list(self.get('ALLOWED_EXTENSIONS')),
            'upload_folder': self.get('UPLOAD_FOLDER'),
            'download_folder': self.get('DOWNLOAD_FOLDER'),
        }
    
    def get_conversion_settings(self) -> Dict[str, Any]:
        """Get conversion-related settings"""
        return {
            'libreoffice_path': self.get('LIBREOFFICE_PATH'),
            'conversion_timeout': self.get('CONVERSION_TIMEOUT'),
            'batch_size': self.get('BATCH_SIZE'),
        }
    
    def get_security_settings(self) -> Dict[str, Any]:
        """Get security-related settings"""
        return {
            'enable_file_validation': self.get('ENABLE_FILE_VALIDATION'),
            'enable_mime_check': self.get('ENABLE_MIME_CHECK'),
            'enable_path_traversal_protection': self.get('ENABLE_PATH_TRAVERSAL_PROTECTION'),
        }
    
    def get_feature_flags(self) -> Dict[str, bool]:
        """Get all feature flags"""
        return {
            'batch_processing': self.is_enabled('BATCH_PROCESSING'),
            'excel_to_pdf': self.is_enabled('EXCEL_TO_PDF'),
            'progress_tracking': self.is_enabled('PROGRESS_TRACKING'),
            'download_compression': self.is_enabled('DOWNLOAD_COMPRESSION'),
            'drag_drop': self.is_enabled('DRAG_DROP'),
            'file_preview': self.is_enabled('FILE_PREVIEW'),
        }
    
    def validate_config(self) -> tuple[bool, list[str]]:
        """Validate configuration settings"""
        errors = []
        
        # Check required directories
        upload_folder = Path(self.get('UPLOAD_FOLDER'))
        download_folder = Path(self.get('DOWNLOAD_FOLDER'))
        
        if not upload_folder.exists():
            errors.append(f"Upload folder does not exist: {upload_folder}")
        
        if not download_folder.exists():
            errors.append(f"Download folder does not exist: {download_folder}")
        
        # Check LibreOffice path
        libreoffice_path = self.get('LIBREOFFICE_PATH')
        if libreoffice_path and not os.path.exists(libreoffice_path):
            errors.append(f"LibreOffice not found at: {libreoffice_path}")
        
        # Validate numeric values
        if self.get('MAX_FILE_SIZE', 0) <= 0:
            errors.append("MAX_FILE_SIZE must be positive")
        
        if self.get('MAX_FILES_PER_REQUEST', 0) <= 0:
            errors.append("MAX_FILES_PER_REQUEST must be positive")
        
        if self.get('CONVERSION_TIMEOUT', 0) <= 0:
            errors.append("CONVERSION_TIMEOUT must be positive")
        
        return len(errors) == 0, errors

# Global configuration instance
config_manager = ConfigManager() 