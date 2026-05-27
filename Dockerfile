# Use an official Python runtime
FROM python:3.11-slim

WORKDIR /app

# Install LibreOffice (for pixel-perfect DOCX→PDF conversion) and fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-core \
    fontconfig \
    fonts-dejavu \
    fonts-liberation \
    fonts-liberation2 \
    fonts-noto \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Create output directories
RUN mkdir -p documents generated

# Expose the port
EXPOSE 5000

ENV PORT=5000

# Run with Gunicorn (increased timeout for LibreOffice conversions)
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "120", "app:app"]
