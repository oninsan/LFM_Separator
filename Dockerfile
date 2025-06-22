FROM python:3.13.5

# Set up app directory
WORKDIR /app
COPY . .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Start the Flask app via Gunicorn
CMD ["gunicorn", "wsgi:app"]
