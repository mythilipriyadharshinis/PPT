FROM python:3.10-slim

# Install system dependencies
RUN apt-get update && \
    apt-get install -y curl gnupg2 unixodbc-dev && \
    curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add - && \
    curl https://packages.microsoft.com/config/debian/10/prod.list > /etc/apt/sources.list.d/mssql-release.list && \
    apt-get update && \
    ACCEPT_EULA=Y apt-get install -y msodbcsql17 && \
    apt-get clean

# Set working directory
# WORKDIR /app

# Copy your code
COPY . .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Start the app (edit if you're using Flask, FastAPI, etc.)
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
