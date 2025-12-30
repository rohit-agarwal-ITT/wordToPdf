from flask import Flask, jsonify
import os
import logging
from logging.handlers import RotatingFileHandler
import tempfile
import shutil

def create_app():
    app = Flask(__name__)
    
    # Enhanced Configuration
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'your-secret-key-here-change-in-production')
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
    app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0  # Prevent file caching
    # Increase timeout for large file operations (if using a WSGI server like gunicorn, set timeout there)
    app.config['PERMANENT_SESSION_LIFETIME'] = 1800  # 30 minutes
    
    # Handle static folder safely
    static_folder = app.static_folder or os.path.join(app.root_path, 'static')
    app.config['UPLOAD_FOLDER'] = os.path.join(static_folder, 'uploads')
    app.config['DOWNLOAD_FOLDER'] = os.path.join(static_folder, 'downloads')
    app.config['TEMP_FOLDER'] = tempfile.mkdtemp(prefix='wordtopdf_')
    
    # Ensure directories exist
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
    
    # Setup logging
    if not app.debug:
        if not os.path.exists('logs'):
            os.mkdir('logs')
        file_handler = RotatingFileHandler('logs/wordtopdf.log', maxBytes=10240, backupCount=10)
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
        ))
        file_handler.setLevel(logging.INFO)
        app.logger.addHandler(file_handler)
        app.logger.setLevel(logging.INFO)
        app.logger.info('Word to PDF Converter startup')
    
    # Global error handlers
    @app.errorhandler(413)
    def too_large(e):
        return jsonify({'error': 'File too large. Maximum size is 100MB.'}), 413
    
    @app.errorhandler(404)
    def not_found(e):
        return jsonify({'error': 'Page not found.'}), 404
    
    @app.errorhandler(500)
    def internal_error(e):
        app.logger.error(f'Server Error: {e}')
        return jsonify({'error': 'Internal server error. Please try again.'}), 500
    
    @app.errorhandler(Exception)
    def handle_exception(e):
        app.logger.error(f'Unhandled Exception: {e}')
        return jsonify({'error': 'An unexpected error occurred. Please try again.'}), 500
    
    # Cleanup function for temp files
    def cleanup_temp_files():
        try:
            if os.path.exists(app.config['TEMP_FOLDER']):
                shutil.rmtree(app.config['TEMP_FOLDER'])
        except Exception as e:
            app.logger.error(f'Error cleaning up temp files: {e}')
    
    # Register cleanup on app context teardown
    @app.teardown_appcontext
    def cleanup(error):
        cleanup_temp_files()
    
    # Register blueprints
    from app.routes import main
    app.register_blueprint(main)
    
    return app 