FROM python:3.11-slim
WORKDIR /app
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV RENDER=true
EXPOSE 3000
CMD ["python", "backend/main.py"]
