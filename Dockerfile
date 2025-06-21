FROM python:3.13.15

# Install Tesseract and its dependencies
RUN apt-get update && \
    apt-get install -y tesseract-ocr libgl1 && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set up app directory
WORKDIR /app
COPY . .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Start the Flask app via Gunicorn
CMD ["gunicorn", "wsgi:app"]
