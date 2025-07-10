#!/usr/bin/env python3
"""
Environment setup script for Word to PDF Converter
This script helps configure the environment for both development and production.
"""

import os
import sys
import secrets

def setup_environment():
    """Setup environment variables for the application"""
    
    print("🔧 Setting up Word to PDF Converter Environment")
    print("=" * 50)
    
    # Check if SECRET_KEY is already set
    if os.environ.get('SECRET_KEY'):
        print("✅ SECRET_KEY is already set")
    else:
        # Generate a secure secret key
        secret_key = secrets.token_hex(32)
        print(f"🔑 Generated new SECRET_KEY: {secret_key[:16]}...")
        
        # Set environment variable for current session
        os.environ['SECRET_KEY'] = secret_key
        
        # Create .env file for future sessions
        env_file = '.env'
        with open(env_file, 'w') as f:
            f.write(f"SECRET_KEY={secret_key}\n")
            f.write("# Add other environment variables here\n")
            f.write("# FLASK_ENV=development\n")
            f.write("# FLASK_DEBUG=1\n")
        
        print(f"📝 Created {env_file} file with environment variables")
        print("💡 For production, set SECRET_KEY environment variable securely")
    
    # Check Python version
    python_version = sys.version_info
    if python_version.major < 3 or (python_version.major == 3 and python_version.minor < 7):
        print("⚠️  Warning: Python 3.7+ is recommended")
    else:
        print(f"✅ Python version {python_version.major}.{python_version.minor} is compatible")
    
    # Check required directories
    required_dirs = ['logs', 'app/static/uploads', 'app/static/downloads']
    for directory in required_dirs:
        if not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
            print(f"📁 Created directory: {directory}")
        else:
            print(f"✅ Directory exists: {directory}")
    
    print("\n🚀 Environment setup complete!")
    print("\nTo run the application:")
    print("  Development: python run.py")
    print("  Production:  python wsgi.py")
    print("\nFor production deployment:")
    print("  - Set SECRET_KEY environment variable")
    print("  - Use a proper WSGI server (gunicorn, uwsgi)")
    print("  - Configure reverse proxy (nginx, apache)")

if __name__ == '__main__':
    setup_environment() 