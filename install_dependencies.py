#!/usr/bin/env python3
"""
Simple dependency installer for Word to PDF Converter
"""

import subprocess
import sys
import os

def install_requirements():
    """Install required packages"""
    print("Installing required packages...")
    
    try:
        # Upgrade pip first
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        print("✅ pip upgraded successfully")
        
        # Install requirements
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ All packages installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install packages: {e}")
        return False

def test_imports():
    """Test if all required modules can be imported"""
    print("\nTesting imports...")
    
    required_modules = [
        'flask',
        'docx',
        'reportlab',
        'PIL',
        'pandas',
        'openpyxl',
        'psutil'
    ]
    
    failed_imports = []
    
    for module in required_modules:
        try:
            __import__(module)
            print(f"✅ {module}")
        except ImportError:
            print(f"❌ {module}")
            failed_imports.append(module)
    
    if failed_imports:
        print(f"\n❌ Failed to import: {', '.join(failed_imports)}")
        return False
    
    print("✅ All modules imported successfully")
    return True

def check_system():
    """Check system requirements"""
    print("\nChecking system requirements...")
    
    # Check Python version
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 7):
        print(f"❌ Python 3.7+ required, found {version.major}.{version.minor}.{version.micro}")
        return False
    else:
        print(f"✅ Python {version.major}.{version.minor}.{version.micro}")
    
    # Check if directories exist
    required_dirs = ['app', 'samples']
    for dir_name in required_dirs:
        if os.path.exists(dir_name):
            print(f"✅ {dir_name} directory exists")
        else:
            print(f"⚠️  {dir_name} directory not found")
    
    return True

def main():
    """Main installation function"""
    print("Word to PDF Converter - Dependency Installer")
    print("=" * 50)
    
    # Check system
    if not check_system():
        print("\n❌ System requirements not met")
        sys.exit(1)
    
    # Install requirements
    if not install_requirements():
        print("\n❌ Failed to install requirements")
        sys.exit(1)
    
    # Test imports
    if not test_imports():
        print("\n❌ Some modules failed to import")
        sys.exit(1)
    
    print("\n🎉 Installation completed successfully!")
    print("\nYou can now run the application with:")
    print("  python run.py")
    print("\nOr test it with:")
    print("  python -m pytest tests/")

if __name__ == "__main__":
    main() 