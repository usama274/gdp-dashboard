# Dockerfile for gdp-dashboard Streamlit app
FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive
WORKDIR /app

# Install system deps
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    git \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copy app
COPY . /app

# Expose port
EXPOSE 8502

ENV STREAMLIT_SERVER_PORT=8502
ENV STREAMLIT_SERVER_HEADLESS=true
ENV PYTHONUNBUFFERED=1

CMD ["streamlit", "run", "app.py", "--server.port", "8502", "--server.address", "0.0.0.0", "--server.enableCORS", "false"]
