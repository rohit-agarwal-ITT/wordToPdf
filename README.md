# Word to PDF Converter

A Flask web application that converts Word documents (DOCX) to PDF format with high fidelity to the original formatting.

## Features

- **Batch Conversion**: Convert up to 100 files at once
- **High Fidelity**: Preserves formatting, images, and layout using LibreOffice
- **Modern UI**: Beautiful, responsive interface with progress tracking
- **Fast Processing**: Optimized batch conversion with estimated time display
- **ZIP Download**: Multiple files are automatically zipped for easy download

## Prerequisites

- Python 3.8+
- LibreOffice (for PDF conversion)
- Git

## Installation

1. **Clone the repository**:
   ```bash
   git clone <your-repo-url>
   cd wordToPdf
   ```

2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Install LibreOffice** (if not already installed):
   - **Windows**: Download from [LibreOffice.org](https://www.libreoffice.org/download/download/)
   - **macOS**: `brew install libreoffice`
   - **Ubuntu/Debian**: `sudo apt-get install libreoffice`

4. **Set environment variables**:
   ```bash
   # Windows
   set SECRET_KEY=your-secret-key-here
   
   # macOS/Linux
   export SECRET_KEY=your-secret-key-here
   ```

## Local Development

Run the application locally:

```bash
python run.py
```

Or using Flask:

```bash
flask run --host=0.0.0.0 --port=5000
```

The app will be available at `http://localhost:5000`

## Deployment Options

### Option 1: Heroku (Recommended for beginners)

1. **Install Heroku CLI** and login:
   ```bash
   heroku login
   ```

2. **Create Heroku app**:
   ```bash
   heroku create your-app-name
   ```

3. **Set environment variables**:
   ```bash
   heroku config:set SECRET_KEY=your-secret-key-here
   ```

4. **Deploy**:
   ```bash
   git add .
   git commit -m "Deploy to Heroku"
   git push heroku main
   ```

5. **Open the app**:
   ```bash
   heroku open
   ```

### Option 2: Railway

1. **Connect your GitHub repository** to Railway
2. **Set environment variables** in Railway dashboard:
   - `SECRET_KEY`: Your secret key
3. **Deploy automatically** from your repository

### Option 3: Render

1. **Connect your GitHub repository** to Render
2. **Create a new Web Service**
3. **Configure**:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn wsgi:app`
4. **Set environment variables**:
   - `SECRET_KEY`: Your secret key

### Option 4: DigitalOcean App Platform

1. **Connect your GitHub repository** to DigitalOcean
2. **Create a new app**
3. **Configure**:
   - Source: Your repository
   - Build Command: `pip install -r requirements.txt`
   - Run Command: `gunicorn wsgi:app`
4. **Set environment variables**

### Option 5: VPS (Ubuntu/Debian)

1. **SSH into your VPS**:
   ```bash
   ssh user@your-server-ip
   ```

2. **Install dependencies**:
   ```bash
   sudo apt update
   sudo apt install python3 python3-pip nginx libreoffice
   ```

3. **Clone and setup**:
   ```bash
   git clone <your-repo-url>
   cd wordToPdf
   pip3 install -r requirements.txt
   ```

4. **Create systemd service**:
   ```bash
   sudo nano /etc/systemd/system/wordtopdf.service
   ```

   Add this content:
   ```ini
   [Unit]
   Description=Word to PDF Converter
   After=network.target

   [Service]
   User=www-data
   WorkingDirectory=/path/to/your/wordToPdf
   Environment="PATH=/path/to/your/wordToPdf/venv/bin"
   ExecStart=/path/to/your/wordToPdf/venv/bin/gunicorn wsgi:app
   Restart=always

   [Install]
   WantedBy=multi-user.target
   ```

5. **Start the service**:
   ```bash
   sudo systemctl start wordtopdf
   sudo systemctl enable wordtopdf
   ```

6. **Configure Nginx** (optional for domain):
   ```bash
   sudo nano /etc/nginx/sites-available/wordtopdf
   ```

   Add this content:
   ```nginx
   server {
       listen 80;
       server_name your-domain.com;

       location / {
           proxy_pass http://127.0.0.1:8000;
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
       }
   }
   ```

   Enable the site:
   ```bash
   sudo ln -s /etc/nginx/sites-available/wordtopdf /etc/nginx/sites-enabled/
   sudo nginx -t
   sudo systemctl reload nginx
   ```

## Environment Variables

- `SECRET_KEY`: Flask secret key (required for production)
- `FLASK_ENV`: Set to `production` for production deployment

## File Structure

```
wordToPdf/
├── app/
│   ├── __init__.py          # Flask app factory
│   ├── routes.py            # Main routes
│   ├── static/              # Static files (CSS, JS, uploads, downloads)
│   ├── templates/           # HTML templates
│   └── utils/               # Utility functions
├── samples/                 # Sample files
├── requirements.txt         # Python dependencies
├── wsgi.py                 # WSGI entry point
├── Procfile               # Heroku configuration
├── runtime.txt            # Python version
└── README.md              # This file
```

## Usage

1. **Upload Files**: Drag and drop or click to select Word documents
2. **Convert**: Click "Convert to PDF" button
3. **Download**: Single files download as PDF, multiple files as ZIP

## Limitations

- Maximum 100 files per upload
- Maximum 100MB total upload size
- Requires LibreOffice for conversion
- Files are temporarily stored and cleaned up automatically

## Troubleshooting

### Common Issues

1. **LibreOffice not found**: Ensure LibreOffice is installed and in PATH
2. **Permission errors**: Check file permissions on upload/download folders
3. **Memory issues**: Reduce batch size for large files

### Logs

- **Heroku**: `heroku logs --tail`
- **Railway**: Check logs in dashboard
- **VPS**: `sudo journalctl -u wordtopdf -f`

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

This project is licensed under the MIT License. 