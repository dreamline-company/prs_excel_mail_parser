FROM python:3.11-slim

WORKDIR /app

# Копируем requirements и устанавливаем зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем код приложения
COPY main.py .
COPY .env .env 2>/dev/null || true

# Создаём директории для результатов
RUN mkdir -p processed results

# Запускаем приложение
CMD ["python", "main.py"]
