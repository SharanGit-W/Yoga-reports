# 1. Use a lightweight Python environment
FROM python:3.11-slim

# 2. Set the working directory inside the container
WORKDIR /app

# 3. Copy your requirements file first
COPY requirements.txt .

# 4. Install the Python libraries
RUN pip install --no-cache-dir -r requirements.txt

# 5. Copy your actual application code
COPY app.py .

# 6. Expose port 8080 (Cloud Run requires this specific port)
EXPOSE 8080

# 7. Start the Streamlit server on the correct port
CMD ["streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0", "--server.enableCORS=false"]
