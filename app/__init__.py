from flask import Flask, jsonify, request
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
    
    # Global error handlers - ensure all errors return JSON
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
    
    @app.errorhandler(502)
    def bad_gateway(e):
        app.logger.error(f'Bad Gateway Error: {e}')
        return jsonify({'error': 'Server temporarily unavailable. Please try again in a moment.'}), 502
    
    @app.errorhandler(503)
    def service_unavailable(e):
        app.logger.error(f'Service Unavailable: {e}')
        return jsonify({'error': 'Service temporarily unavailable. Please try again in a moment.'}), 503
    
    @app.errorhandler(504)
    def gateway_timeout(e):
        app.logger.error(f'Gateway Timeout: {e}')
        return jsonify({'error': 'Request timed out. The conversion may be taking longer than expected. Please try again with smaller files.'}), 504
    
    @app.errorhandler(Exception)
    def handle_exception(e):
        app.logger.error(f'Unhandled Exception: {e}', exc_info=True)
        return jsonify({'error': 'An unexpected error occurred. Please try again.'}), 500
    
    # Middleware to ensure JSON responses for API routes
    @app.after_request
    def ensure_json_response(response):
        """Ensure API routes return JSON, even on errors"""
        # Only apply to routes that should return JSON (not static files or templates)
        if request.path.startswith('/upload') or request.path.startswith('/progress'):
            # If response is not already JSON and is an error, convert it
            if response.status_code >= 400 and not response.is_json:
                try:
                    # Try to get the response data
                    data = response.get_data(as_text=True)
                    # If it's HTML or plain text, convert to JSON
                    if data and (data.strip().startswith('<') or not data.strip().startswith('{')):
                        return jsonify({
                            'error': f'Server error: {response.status_code}',
                            'details': data[:200] if len(data) > 200 else data
                        }), response.status_code
                except Exception:
                    pass
        return response
    
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