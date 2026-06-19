FROM python:3.11-slim

WORKDIR /app

# Устанавливаем CA сертификаты и утилиты для обработки файлов
RUN apt-get update && apt-get install -y \
    ca-certificates \
    dos2unix \
    curl \
    && update-ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Копируем requirements и устанавливаем зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем код приложения
COPY main.py .
COPY .env .env 2>/dev/null || true

# Конвертируем .env в Unix формат (убираем CRLF если есть)
RUN if [ -f .env ]; then dos2unix .env; fi

# Создаём директории для результатов
RUN mkdir -p processed results

# Запускаем приложение
CMD ["python", "main.py"]
