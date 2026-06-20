FROM python:3.11-slim

# Prevent Python from writing .pyc files and force stdout/stderr to be unbuffered
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Install dependencies
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy all source files
COPY . ./

# Set the default timezone
ENV TZ=Asia/Ho_Chi_Minh

ENTRYPOINT ["python", "src/pipeline.py"]
