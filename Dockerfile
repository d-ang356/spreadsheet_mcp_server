FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    && rm -rf /var/lib/apt/lists/*

RUN apt-get update && apt-get install -y \
    libxml2 libxslt1.1 \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy server code
COPY spreadsheet_server.py .

# Create directories for spreadsheets
RUN mkdir -p /spreadsheets /imports

# Set environment variables
ENV PYTHONUNBUFFERED=1

# Run the server
CMD ["python", "spreadsheet_server.py"]