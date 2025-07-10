# Word & Excel to PDF Converter

A modern, secure, and high-performance web application for converting Word and Excel files to PDF format. Built with Flask and featuring comprehensive error handling, security measures, and user-friendly interface.

## ✨ Features

### 🔄 Conversion Capabilities
- **Single File Conversion**: Convert individual Word (.docx, .doc) files to PDF
- **Batch Processing**: Convert multiple files simultaneously with progress tracking
- **Excel to PDF**: Upload Excel files to generate personalized PDFs for each row
- **Template Support**: Use Word templates with placeholders for dynamic content

### 🛡️ Security & Validation
- **File Type Validation**: Comprehensive MIME type and extension checking
- **Path Traversal Protection**: Secure file handling and sanitization
- **Size Limits**: Configurable file size limits (default: 100MB)
- **Input Sanitization**: Automatic filename sanitization and validation
- **Secure Temp Files**: Proper file permissions and cleanup

### 📊 Performance & Monitoring
- **Performance Tracking**: Monitor conversion times and resource usage
- **System Health Checks**: Automatic validation of system resources
- **Progress Tracking**: Real-time progress updates with estimated completion times
- **Resource Management**: Memory and disk space monitoring

### 🎨 User Experience
- **Modern UI**: Beautiful, responsive design with drag-and-drop support
- **Accessibility**: Full keyboard navigation and screen reader support
- **Real-time Feedback**: Progress bars, status updates, and error messages
- **File Preview**: Visual file list with size and status information
- **Statistics Dashboard**: File count, total size, and estimated processing time

### ⚙️ Configuration & Management
- **Environment-based Config**: Support for environment variables and config files
- **Feature Flags**: Enable/disable features via configuration
- **Logging**: Comprehensive logging with rotation and size limits
- **Error Handling**: Graceful error handling with user-friendly messages

## 🚀 Quick Start

### Prerequisites
- Python 3.8+
- LibreOffice (for PDF conversion)
- 500MB+ free memory
- 1GB+ free disk space

### Installation

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd wordToPdf
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install LibreOffice** (if not already installed)
   - **Windows**: Download from [LibreOffice.org](https://www.libreoffice.org/)
   - **macOS**: `brew install libreoffice`
   - **Linux**: `sudo apt-get install libreoffice`

4. **Run the application**
   ```bash
   python run.py
   ```

5. **Access the application**
   Open your browser and navigate to `http://localhost:5000`

## 📁 Project Structure

```
wordToPdf/
├── app/
│   ├── __init__.py              # Flask app initialization
│   ├── routes.py                # Main application routes
│   ├── templates/
│   │   └── index.html          # Enhanced UI with accessibility
│   ├── static/                 # Static files (CSS, JS, uploads)
│   └── utils/
│       ├── file_security.py    # File validation and security
│       ├── performance_monitor.py # Performance tracking
│       ├── config_manager.py   # Configuration management
│       ├── word_processor.py   # Word document processing
│       ├── pdf_generator.py    # PDF generation utilities
│       ├── validators.py       # File validation utilities
│       ├── error_handler.py    # Error handling utilities
│       └── conversion_manager.py # Conversion orchestration
├── samples/                    # Sample files for testing
├── logs/                       # Application logs
├── requirements.txt            # Python dependencies
├── run.py                     # Application entry point
└── README.md                  # This file
```

## ⚙️ Configuration

### Environment Variables
```bash
# Application settings
SECRET_KEY=your-secret-key-here
DEBUG=false

# File upload settings
MAX_FILE_SIZE=104857600  # 100MB in bytes
MAX_FILES_PER_REQUEST=100
CONVERSION_TIMEOUT=300   # 5 minutes

# Performance settings
ENABLE_PERFORMANCE_MONITORING=true
BATCH_SIZE=10

# Feature flags
ENABLE_BATCH_PROCESSING=true
ENABLE_EXCEL_TO_PDF=true
ENABLE_PROGRESS_TRACKING=true
```

### Configuration File
Create a `config.json` file in the root directory:
```json
{
  "MAX_FILE_SIZE": 104857600,
  "ENABLE_PERFORMANCE_MONITORING": true,
  "LIBREOFFICE_PATH": "/usr/bin/soffice"
}
```

## 🔧 Advanced Features

### Performance Monitoring
The application includes comprehensive performance monitoring:
- Conversion time tracking
- Memory usage monitoring
- System resource validation
- Performance metrics collection

### Security Features
- File type validation using MIME signatures
- Path traversal protection
- Secure temporary file handling
- Input sanitization and validation

### Error Handling
- Comprehensive error messages
- Graceful degradation
- User-friendly error reporting
- Automatic cleanup on errors

### Accessibility
- Full keyboard navigation
- Screen reader support
- ARIA labels and descriptions
- High contrast mode support

## 📊 Usage Examples

### Single File Conversion
1. Upload a Word document (.docx or .doc)
2. Click "Convert to PDF"
3. Download the converted PDF

### Batch Processing
1. Upload multiple Word files
2. View file list with statistics
3. Click "Convert to PDF"
4. Download ZIP file containing all PDFs

### Excel to PDF Processing
1. Upload an Excel file (.xlsx)
2. The system will generate a PDF for each row
3. Download ZIP file with all generated PDFs

## 🐛 Troubleshooting

### Common Issues

**LibreOffice not found**
- Ensure LibreOffice is installed and accessible
- Check the path in configuration
- Verify installation on your platform

**File upload errors**
- Check file size limits
- Verify file type is supported
- Ensure sufficient disk space

**Conversion timeouts**
- Increase `CONVERSION_TIMEOUT` in configuration
- Check system resources
- Reduce batch size for large files

**Memory errors**
- Increase available memory
- Reduce batch size
- Close other applications

### Logs
Check the `logs/wordtopdf.log` file for detailed error information and performance metrics.

## 🔒 Security Considerations

- All uploaded files are validated for type and content
- Temporary files are automatically cleaned up
- File paths are sanitized to prevent traversal attacks
- Secure file permissions are enforced
- Input validation prevents malicious uploads

## 📈 Performance Tips

- Use SSD storage for better I/O performance
- Ensure adequate RAM (2GB+ recommended)
- Configure appropriate batch sizes for your system
- Monitor system resources during conversion
- Use network storage for large file processing

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For issues and questions:
1. Check the troubleshooting section
2. Review the logs for error details
3. Open an issue with detailed information
4. Include system specifications and error messages

---

**Built with ❤️ using Flask, Python, and modern web technologies** 