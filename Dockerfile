 # Use official Python image
FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV FLASK_APP=wsgi:app

# Install LibreOffice and other system dependencies
RUN apt-get update && \
    apt-get install -y libreoffice libreoffice-writer libreoffice-calc libreoffice-impress && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Create necessary directories and set permissions
RUN mkdir -p app/static/uploads app/static/downloads && \
    chmod -R 755 app/static/uploads app/static/downloads

# Expose port (App Runner expects 8080)
EXPOSE 8080

# Run the app with Gunicorn (production-ready WSGI server)
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "2", "--timeout", "120", "wsgi:app"]