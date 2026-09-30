FROM python:3.11-slim

WORKDIR /app

# Устанавливаем CA сертификаты и утилиты
RUN apt-get update && apt-get install -y \
    ca-certificates \
    curl \
    && update-ca-certificates --fresh \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Корневой сертификат шлюза USIAG: HTTPS в интернет из сети сервера идёт через
# него (тот же, что в образе бэкенда). pip берёт системное хранилище.
COPY usia-ca.crt /usr/local/share/ca-certificates/usia-ca.crt
RUN update-ca-certificates
ENV PIP_CERT=/etc/ssl/certs/ca-certificates.crt

# Копируем requirements и устанавливаем зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем код приложения
COPY main.py summary_parser.py upload_summaries.py ./
# .env в образ не кладём (в нём пароль): docker compose передаёт его
# переменные при запуске (env_file).

# Создаём директорию для обработанных файлов
RUN mkdir -p inbox processed failed

# Запускаем приложение
CMD ["python", "main.py"]
