#!/usr/bin/env python3
"""
Test script for Word to PDF Converter
This script tests the basic functionality of the application.
"""

import os
import sys
import socket
import time

def test_app():
    """Test the application functionality"""
    
    print("🧪 Testing Word to PDF Converter")
    print("=" * 40)
    
    # Test 1: Check if app is running
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(5)
        result = sock.connect_ex(('localhost', 5000))
        sock.close()
        
        if result == 0:
            print("✅ Application is running on http://localhost:5000")
        else:
            print("❌ Application is not running. Please start it with:")
            print("   python run.py")
            return False
    except Exception as e:
        print(f"❌ Error checking application: {e}")
        return False
    
    # Test 2: Check if port is accessible
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(5)
        result = sock.connect_ex(('localhost', 5000))
        sock.close()
        
        if result == 0:
            print("✅ Port 5000 is accessible")
        else:
            print("⚠️  Port 5000 is not accessible")
    except Exception as e:
        print(f"❌ Error testing port accessibility: {e}")
    
    # Test 3: Check required directories
    required_dirs = ['logs', 'app/static/uploads', 'app/static/downloads']
    for directory in required_dirs:
        if os.path.exists(directory):
            print(f"✅ Directory exists: {directory}")
        else:
            print(f"❌ Missing directory: {directory}")
    
    # Test 4: Check sample files
    sample_files = [
        'samples/sample_document_for_placeholder.docx',
        'samples/sample_excel_sheet.xlsx'
    ]
    for file_path in sample_files:
        if os.path.exists(file_path):
            print(f"✅ Sample file exists: {file_path}")
        else:
            print(f"⚠️  Sample file missing: {file_path}")
    
    print("\n🎉 Testing complete!")
    print("\nTo use the application:")
    print("1. Open http://localhost:5000 in your browser")
    print("2. Upload Word or Excel files")
    print("3. Click 'Convert to PDF'")
    print("4. Download your converted files")
    
    return True

if __name__ == '__main__':
    test_app() 