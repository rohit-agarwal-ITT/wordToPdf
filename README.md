# Word to PDF Converter

A Flask-based web application that converts Word documents (.docx and .doc) to PDF format. Users can upload single or multiple Word documents and download the converted PDF files.

## Features

- **Multiple File Upload**: Upload single or multiple Word documents at once
- **Drag & Drop Interface**: Modern, responsive web interface with drag & drop functionality
- **Batch Processing**: Convert multiple files simultaneously
- **Format Preservation**: Maintains text formatting, tables, and basic styling
- **Secure File Handling**: Unique file naming and secure file storage
- **Real-time Progress**: Visual feedback during conversion process

## Prerequisites

- Python 3.11 or higher
- Windows OS (for docx2pdf library compatibility)

## Installation

1. **Clone or download the project**:
   ```bash
   git clone <repository-url>
   cd wordToPdf
   ```

2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Start the application**:
   ```bash
   python run.py
   ```

2. **Open your web browser** and navigate to:
   ```
   http://localhost:5000
   ```

3. **Upload Word documents**:
   - Drag and drop Word files (.docx or .doc) onto the upload area
   - Or click the upload area to browse and select files
   - You can select multiple files at once

4. **Convert to PDF**:
   - Click the "Convert to PDF" button
   - Wait for the conversion process to complete
   - Download the converted PDF files

## Project Structure

```
wordToPdf/
├── app/
│   ├── __init__.py          # Flask application factory
│   ├── routes.py            # Main routes and endpoints
│   ├── utils/
│   │   ├── __init__.py
│   │   ├── word_processor.py # Word document processing
│   │   └── pdf_generator.py  # PDF generation
│   ├── static/
│   │   ├── uploads/         # Temporary uploaded files
│   │   └── downloads/       # Generated PDF files
│   └── templates/
│       └── index.html       # Main web interface
├── samples/                 # Sample documents
├── requirements.txt         # Python dependencies
├── run.py                  # Application entry point
└── README.md               # This file
```

## Technical Details

### Dependencies

- **Flask**: Web framework
- **python-docx**: Word document processing
- **reportlab**: PDF generation
- **docx2pdf**: Direct Word to PDF conversion
- **Pillow**: Image processing
- **Werkzeug**: File handling utilities

### Supported File Formats

- **Input**: Microsoft Word documents (.docx, .doc)
- **Output**: PDF documents (.pdf)

### File Size Limits

- Maximum file size: 16MB per file
- Multiple files can be processed simultaneously

## API Endpoints

- `GET /`: Main application interface
- `POST /upload`: File upload and conversion endpoint
- `GET /download/<filename>`: Download converted PDF files

## Error Handling

The application includes comprehensive error handling for:
- Invalid file types
- File size limits
- Conversion errors
- File not found errors

## Security Features

- Secure filename handling
- File type validation
- Unique file naming to prevent conflicts
- Temporary file cleanup

## Troubleshooting

### Common Issues

1. **"python-docx2pdf not found" error**:
   - This is expected - the application uses alternative libraries for conversion

2. **Conversion fails**:
   - Ensure the Word document is not corrupted
   - Check that the file is a valid .docx or .doc format
   - Verify the file is not password-protected

3. **Port already in use**:
   - Change the port in `run.py` or stop other applications using port 5000

### Performance Tips

- For large documents, the conversion may take longer
- The application processes files sequentially for stability
- Consider breaking very large documents into smaller parts

## Development

To run in development mode with auto-reload:
```bash
python run.py
```

The application will be available at `http://localhost:5000` with debug mode enabled.

## License

This project is open source and available under the MIT License.

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve the application. 