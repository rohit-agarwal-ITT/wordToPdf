import os
import subprocess
import tempfile
import shutil
import threading
import time
from typing import List, Dict, Tuple, Optional, Callable, Any
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
import logging
import platform
from app.utils.validators import FileValidator
from app.utils.error_handler import ErrorHandler

logger = logging.getLogger(__name__)

class ConversionManager:
    """Manages document conversions with proper error handling and resource management"""
    
    # Timeout settings
    SINGLE_FILE_TIMEOUT = 120  # 2 minutes
    BATCH_TIMEOUT = 600  # 10 minutes
    SUBPROCESS_TIMEOUT = 60  # 1 minute
    
    # Resource limits
    MAX_CONCURRENT_CONVERSIONS = 4
    MAX_MEMORY_USAGE = 500 * 1024 * 1024  # 500MB
    
    def __init__(self):
        self.conversion_progress = {
            'status': 'idle',
            'current': 0,
            'total': 0,
            'message': '',
            'error': None,
            'start_time': None,
            'estimated_time': None
        }
        self._progress_callback = None
        self._stop_conversion = False
    
    def set_progress_callback(self, callback: Callable):
        """Set callback function for progress updates"""
        self._progress_callback = callback
    
    def update_progress(self, current: int, total: int, message: str):
        """Update conversion progress"""
        self.conversion_progress.update({
            'current': current,
            'total': total,
            'message': message,
            'status': 'converting'
        })
        
        if self._progress_callback:
            self._progress_callback(self.conversion_progress)
    
    def reset_progress(self):
        """Reset progress tracking"""
        self.conversion_progress = {
            'status': 'idle',
            'current': 0,
            'total': 0,
            'message': '',
            'error': None,
            'start_time': None,
            'estimated_time': None
        }
        self._stop_conversion = False
    
    def stop_conversion(self):
        """Stop ongoing conversion"""
        self._stop_conversion = True
        self.conversion_progress['status'] = 'stopped'
    
    def convert_single_file(self, file_path: str, filename: str, output_dir: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """
        Convert a single file to PDF
        
        Returns:
            Tuple[output_path, pdf_name, error_message]
        """
        try:
            # Validate input
            if not os.path.exists(file_path):
                return None, None, f"Input file not found: {file_path}"
            
            # Check disk space
            file_size = os.path.getsize(file_path)
            required_space = file_size * 3  # Estimate for conversion
            if not ErrorHandler.check_disk_space(output_dir, required_space):
                return None, None, "Insufficient disk space for conversion"
            
            # Get LibreOffice path
            soffice_path = self._get_libreoffice_path()
            if not soffice_path:
                return None, None, "LibreOffice not found. Please install LibreOffice."
            
            # Convert file
            output_pdf = file_path.rsplit('.', 1)[0] + '.pdf'
            name_part = filename.rsplit('.', 1)[0]
            pdf_name = f"{name_part}-Appointment_letter.pdf"
            
            # Run conversion with timeout
            try:
                result = subprocess.run([
                    soffice_path, '--headless', '--convert-to', 'pdf', 
                    '--outdir', output_dir, file_path
                ], check=True, capture_output=True, timeout=self.SUBPROCESS_TIMEOUT)
                
                # Check if output file was created
                if not os.path.exists(output_pdf):
                    return None, None, f"Conversion failed: Output file not created"
                
                return output_pdf, pdf_name, None
                
            except subprocess.TimeoutExpired:
                return None, None, f"Conversion timeout for {filename}"
            except subprocess.CalledProcessError as e:
                return None, None, f"Conversion failed for {filename}: {e.stderr.decode()}"
                
        except Exception as e:
            ErrorHandler.log_error(e, "single_file_conversion", {"file_path": file_path})
            return None, None, f"Conversion failed for {filename}: {str(e)}"
    
    def convert_batch_files(self, file_paths: List[Tuple[str, str]], output_dir: str) -> Tuple[List[Tuple[str, str]], List[str]]:
        """
        Convert multiple files to PDF using parallel processing
        
        Returns:
            Tuple[successful_conversions, errors]
        """
        successful_conversions = []
        errors = []
        
        try:
            # Validate LibreOffice installation
            soffice_path = self._get_libreoffice_path()
            if not soffice_path:
                return [], ["LibreOffice not found. Please install LibreOffice."]
            
            # Check disk space for batch conversion
            total_size = sum(os.path.getsize(path) for path, _ in file_paths)
            required_space = total_size * 3  # Estimate for conversion
            if not ErrorHandler.check_disk_space(output_dir, required_space):
                return [], ["Insufficient disk space for batch conversion"]
            
            # Convert files in parallel
            with ThreadPoolExecutor(max_workers=self.MAX_CONCURRENT_CONVERSIONS) as executor:
                # Submit all conversion tasks
                future_to_file = {
                    executor.submit(self._convert_file_with_progress, file_path, filename, output_dir): (file_path, filename)
                    for file_path, filename in file_paths
                }
                
                # Process completed conversions
                for i, future in enumerate(as_completed(future_to_file, timeout=self.BATCH_TIMEOUT)):
                    if self._stop_conversion:
                        break
                    
                    file_path, filename = future_to_file[future]
                    
                    try:
                        output_pdf, pdf_name, error = future.result(timeout=self.SUBPROCESS_TIMEOUT)
                        
                        if error:
                            errors.append(f"{filename}: {error}")
                        else:
                            successful_conversions.append((output_pdf, pdf_name))
                        
                        # Update progress
                        self.update_progress(i + 1, len(file_paths), f"Converted {i + 1}/{len(file_paths)} files...")
                        
                    except TimeoutError:
                        errors.append(f"{filename}: Conversion timeout")
                    except Exception as e:
                        errors.append(f"{filename}: {str(e)}")
            
            return successful_conversions, errors
            
        except Exception as e:
            ErrorHandler.log_error(e, "batch_conversion", {"file_count": len(file_paths)})
            return [], [f"Batch conversion failed: {str(e)}"]
    
    def _convert_file_with_progress(self, file_path: str, filename: str, output_dir: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """Convert a single file with progress tracking"""
        if self._stop_conversion:
            return None, None, "Conversion stopped by user"
        
        return self.convert_single_file(file_path, filename, output_dir)
    
    def _get_libreoffice_path(self) -> Optional[str]:
        """Get LibreOffice executable path"""
        if platform.system() == "Windows":
            soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
            if os.path.exists(soffice_path):
                return soffice_path
        else:
            # Try to find soffice in PATH
            try:
                result = subprocess.run(['which', 'soffice'], capture_output=True, text=True)
                if result.returncode == 0:
                    return 'soffice'
            except Exception:
                pass
        
        return None
    
    def validate_conversion_requirements(self) -> Tuple[bool, List[str]]:
        """Validate all requirements for conversion"""
        errors = []
        
        # Check LibreOffice
        soffice_path = self._get_libreoffice_path()
        if not soffice_path:
            errors.append("LibreOffice not found. Please install LibreOffice.")
        
        # Check system resources
        try:
            import psutil
            memory = psutil.virtual_memory()
            if memory.available < self.MAX_MEMORY_USAGE:
                errors.append(f"Insufficient memory. Available: {memory.available // (1024*1024)}MB, Required: {self.MAX_MEMORY_USAGE // (1024*1024)}MB")
        except ImportError:
            logger.warning("psutil not available, skipping memory check")
        
        return len(errors) == 0, errors
    
    def cleanup_temp_files(self, *temp_dirs):
        """Clean up temporary files"""
        ErrorHandler.cleanup_temp_files(*temp_dirs)
    
    def get_conversion_stats(self) -> Dict[str, Any]:
        """Get conversion statistics"""
        if self.conversion_progress['start_time']:
            elapsed = time.time() - self.conversion_progress['start_time']
            if self.conversion_progress['current'] > 0:
                rate = self.conversion_progress['current'] / elapsed
                remaining = (self.conversion_progress['total'] - self.conversion_progress['current']) / rate if rate > 0 else 0
            else:
                remaining = 0
        else:
            elapsed = 0
            remaining = 0
        
        return {
            'elapsed_time': elapsed,
            'estimated_remaining': remaining,
            'progress_percentage': (self.conversion_progress['current'] / self.conversion_progress['total'] * 100) if self.conversion_progress['total'] > 0 else 0
        } 