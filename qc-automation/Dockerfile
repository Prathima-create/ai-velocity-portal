FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app
COPY app.py .
COPY .streamlit/ .streamlit/
RUN mkdir -p data

# Expose Streamlit port
EXPOSE 8501

# Health check
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# Run Streamlit
ENV QC_S3_BUCKET=fincom-qc-data
ENV QC_S3_PREFIX=General_Apr_2026/
ENV QC_S3_PREV_PREFIX=General_Mar_2026/

CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0", "--server.headless=true"]
