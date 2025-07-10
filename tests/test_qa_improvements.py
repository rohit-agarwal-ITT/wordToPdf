#!/usr/bin/env python3
"""
Comprehensive QA Test Suite for Word to PDF Converter
Tests all security, validation, error handling, and edge cases
"""

import unittest
import tempfile
import os
import shutil
import io
from unittest.mock import Mock, patch, MagicMock
from werkzeug.datastructures import FileStorage
import pandas as pd
import time

# Import the modules we want to test
from app.utils.validators import FileValidator
from app.utils.error_handler import ErrorHandler
from app.utils.conversion_manager import ConversionManager

class TestFileValidator(unittest.TestCase):
    """Test file validation functionality"""
    
    def setUp(self):
        """Set up test environment"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Clean up test environment"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_validate_file_upload_no_files(self):
        """Test validation with no files"""
        is_valid, error_msg, valid_files = FileValidator.validate_file_upload([])
        self.assertFalse(is_valid)
        self.assertIn("No files provided", error_msg)
        self.assertEqual(len(valid_files), 0)
    
    def test_validate_file_upload_invalid_extension(self):
        """Test validation with invalid file extension"""
        # Create mock file with invalid extension
        mock_file = Mock(spec=FileStorage)
        mock_file.filename = "test.txt"
        mock_file.seek = Mock()
        mock_file.tell = Mock(return_value=1024)
        mock_file.read = Mock(return_value=b"test content")
        
        is_valid, error_msg, valid_files = FileValidator.validate_file_upload([mock_file])
        self.assertFalse(is_valid)
        self.assertIn("Invalid file type", error_msg)
        self.assertEqual(len(valid_files), 0)
    
    def test_validate_file_upload_file_too_large(self):
        """Test validation with file too large"""
        # Create mock file that's too large
        mock_file = Mock(spec=FileStorage)
        mock_file.filename = "test.docx"
        mock_file.seek = Mock()
        mock_file.tell = Mock(return_value=FileValidator.MAX_FILE_SIZE + 1024)
        mock_file.read = Mock(return_value=b"test content")
        
        is_valid, error_msg, valid_files = FileValidator.validate_file_upload([mock_file])
        self.assertFalse(is_valid)
        self.assertIn("too large", error_msg)
        self.assertEqual(len(valid_files), 0)
    
    def test_validate_file_upload_valid_file(self):
        """Test validation with valid file"""
        # Create mock file with valid extension
        mock_file = Mock(spec=FileStorage)
        mock_file.filename = "test.docx"
        mock_file.seek = Mock()
        mock_file.tell = Mock(return_value=1024)
        mock_file.read = Mock(return_value=b"test content")
        
        # Test with valid file (mimetypes will be used instead of magic)
        is_valid, error_msg, valid_files = FileValidator.validate_file_upload([mock_file])
        self.assertTrue(is_valid)
        self.assertEqual(error_msg, "")
        self.assertEqual(len(valid_files), 1)
    
    def test_sanitize_filename(self):
        """Test filename sanitization"""
        # Test dangerous characters
        dangerous_filename = "../../../etc/passwd"
        sanitized = FileValidator.sanitize_filename(dangerous_filename)
        self.assertNotIn("..", sanitized)
        self.assertNotIn("/", sanitized)
        
        # Test normal filename
        normal_filename = "test document.docx"
        sanitized = FileValidator.sanitize_filename(normal_filename)
        self.assertEqual(sanitized, "test document.docx")
        
        # Test empty filename
        sanitized = FileValidator.sanitize_filename("")
        self.assertEqual(sanitized, "file")
    
    def test_validate_excel_structure(self):
        """Test Excel file structure validation"""
        # Create a test Excel file
        test_data = {'Name': ['John', 'Jane'], 'Age': [30, 25]}
        df = pd.DataFrame(test_data)
        excel_path = os.path.join(self.temp_dir, "test.xlsx")
        df.to_excel(excel_path, index=False)
        
        # Test valid Excel file
        is_valid, error_msg, df_result = FileValidator.validate_excel_structure(excel_path)
        self.assertTrue(is_valid)
        self.assertEqual(error_msg, "")
        self.assertIsNotNone(df_result)
        if df_result is not None:
            self.assertEqual(len(df_result), 2)
        
        # Test non-existent file
        is_valid, error_msg, df_result = FileValidator.validate_excel_structure("nonexistent.xlsx")
        self.assertFalse(is_valid)
        self.assertIn("not found", error_msg)
        self.assertIsNone(df_result)
    
    def test_validate_template_file(self):
        """Test template file validation"""
        # Test non-existent template
        is_valid, error_msg = FileValidator.validate_template_file("nonexistent.docx")
        self.assertFalse(is_valid)
        self.assertIn("not found", error_msg)
        
        # Test with valid file (create a dummy file)
        test_file = os.path.join(self.temp_dir, "test.docx")
        with open(test_file, 'w') as f:
            f.write("test content")
        
        is_valid, error_msg = FileValidator.validate_template_file(test_file)
        self.assertTrue(is_valid)
        self.assertEqual(error_msg, "")

class TestErrorHandler(unittest.TestCase):
    """Test error handling functionality"""
    
    def setUp(self):
        """Set up test environment"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Clean up test environment"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_log_error(self):
        """Test error logging"""
        test_error = ValueError("Test error")
        ErrorHandler.log_error(test_error, "test_context", {"test": "data"})
        # This should not raise an exception
    
    def test_cleanup_temp_files(self):
        """Test temporary file cleanup"""
        # Create some test files
        test_file1 = os.path.join(self.temp_dir, "test1.txt")
        test_file2 = os.path.join(self.temp_dir, "test2.txt")
        
        with open(test_file1, 'w') as f:
            f.write("test1")
        with open(test_file2, 'w') as f:
            f.write("test2")
        
        # Test cleanup
        ErrorHandler.cleanup_temp_files(self.temp_dir)
        
        # Verify files are cleaned up
        self.assertFalse(os.path.exists(test_file1))
        self.assertFalse(os.path.exists(test_file2))
    
    def test_handle_conversion_error(self):
        """Test conversion error handling"""
        test_error = ValueError("Conversion failed")
        temp_dirs = [self.temp_dir]
        
        result = ErrorHandler.handle_conversion_error(test_error, temp_dirs, "Test conversion failed")
        
        self.assertIn("error", result)
        self.assertIn("Test conversion failed", result["error"])
    
    def test_validate_system_requirements(self):
        """Test system requirements validation"""
        result = ErrorHandler.validate_system_requirements()
        
        self.assertIn("errors", result)
        self.assertIn("warnings", result)
        self.assertIn("is_ready", result)
        self.assertIsInstance(result["errors"], list)
        self.assertIsInstance(result["warnings"], list)
        self.assertIsInstance(result["is_ready"], bool)

class TestConversionManager(unittest.TestCase):
    """Test conversion manager functionality"""
    
    def setUp(self):
        """Set up test environment"""
        self.conversion_manager = ConversionManager()
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Clean up test environment"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_reset_progress(self):
        """Test progress reset"""
        self.conversion_manager.conversion_progress['status'] = 'converting'
        self.conversion_manager.conversion_progress['current'] = 5
        
        self.conversion_manager.reset_progress()
        
        self.assertEqual(self.conversion_manager.conversion_progress['status'], 'idle')
        self.assertEqual(self.conversion_manager.conversion_progress['current'], 0)
    
    def test_update_progress(self):
        """Test progress update"""
        self.conversion_manager.update_progress(3, 10, "Converting...")
        
        self.assertEqual(self.conversion_manager.conversion_progress['current'], 3)
        self.assertEqual(self.conversion_manager.conversion_progress['total'], 10)
        self.assertEqual(self.conversion_manager.conversion_progress['message'], "Converting...")
        self.assertEqual(self.conversion_manager.conversion_progress['status'], 'converting')
    
    def test_stop_conversion(self):
        """Test conversion stop"""
        self.conversion_manager.conversion_progress['status'] = 'converting'
        
        self.conversion_manager.stop_conversion()
        
        self.assertEqual(self.conversion_manager.conversion_progress['status'], 'stopped')
        self.assertTrue(self.conversion_manager._stop_conversion)
    
    def test_validate_conversion_requirements(self):
        """Test conversion requirements validation"""
        is_ready, errors = self.conversion_manager.validate_conversion_requirements()
        
        self.assertIsInstance(is_ready, bool)
        self.assertIsInstance(errors, list)
    
    def test_get_conversion_stats(self):
        """Test conversion statistics"""
        # Set up some progress
        self.conversion_manager.conversion_progress['start_time'] = time.time() - 10
        self.conversion_manager.conversion_progress['current'] = 5
        self.conversion_manager.conversion_progress['total'] = 10
        
        stats = self.conversion_manager.get_conversion_stats()
        
        self.assertIn('elapsed_time', stats)
        self.assertIn('estimated_remaining', stats)
        self.assertIn('progress_percentage', stats)
        self.assertIsInstance(stats['elapsed_time'], (int, float))
        self.assertIsInstance(stats['estimated_remaining'], (int, float))
        self.assertIsInstance(stats['progress_percentage'], (int, float))

class TestSecurityEdgeCases(unittest.TestCase):
    """Test security edge cases and vulnerabilities"""
    
    def test_path_traversal_prevention(self):
        """Test prevention of path traversal attacks"""
        malicious_filenames = [
            "../../../etc/passwd",
            "..\\..\\..\\windows\\system32\\config\\sam",
            "....//....//....//etc/passwd",
            "..%2F..%2F..%2Fetc%2Fpasswd"
        ]
        
        for filename in malicious_filenames:
            sanitized = FileValidator.sanitize_filename(filename)
            self.assertNotIn("..", sanitized)
            self.assertNotIn("/", sanitized)
            self.assertNotIn("\\", sanitized)
    
    def test_file_size_limits(self):
        """Test file size limit enforcement"""
        # Test with file size exceeding limit
        mock_file = Mock(spec=FileStorage)
        mock_file.filename = "test.docx"
        mock_file.seek = Mock()
        mock_file.tell = Mock(return_value=FileValidator.MAX_FILE_SIZE + 1024)
        mock_file.read = Mock(return_value=b"test content")
        
        is_valid, error_msg, valid_files = FileValidator.validate_file_upload([mock_file])
        self.assertFalse(is_valid)
        self.assertIn("too large", error_msg)
    
    def test_mime_type_validation(self):
        """Test MIME type validation"""
        # Test with invalid MIME type
        mock_file = Mock(spec=FileStorage)
        mock_file.filename = "test.txt"  # Invalid extension
        mock_file.seek = Mock()
        mock_file.tell = Mock(return_value=1024)
        mock_file.read = Mock(return_value=b"test content")
        
        is_valid, error_msg, valid_files = FileValidator.validate_file_upload([mock_file])
        self.assertFalse(is_valid)
        self.assertIn("Invalid file type", error_msg)

class TestErrorRecovery(unittest.TestCase):
    """Test error recovery and cleanup"""
    
    def setUp(self):
        """Set up test environment"""
        self.temp_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Clean up test environment"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_cleanup_on_error(self):
        """Test cleanup when errors occur"""
        # Create some test files
        test_file = os.path.join(self.temp_dir, "test.txt")
        with open(test_file, 'w') as f:
            f.write("test content")
        
        # Simulate error and cleanup
        try:
            raise ValueError("Test error")
        except Exception as e:
            ErrorHandler.cleanup_temp_files(self.temp_dir)
        
        # Verify cleanup occurred
        self.assertFalse(os.path.exists(test_file))
    
    def test_memory_error_handling(self):
        """Test memory error handling"""
        error_dict = ErrorHandler.handle_memory_error("file conversion")
        
        self.assertIn("error", error_dict)
        self.assertIn("Memory limit exceeded", error_dict["error"])
    
    def test_timeout_error_handling(self):
        """Test timeout error handling"""
        error_dict = ErrorHandler.handle_timeout_error("file conversion", 60)
        
        self.assertIn("error", error_dict)
        self.assertIn("timed out", error_dict["error"])

if __name__ == '__main__':
    # Run the tests
    unittest.main(verbosity=2) 