# PRS Parser - Gmail IMAP Bot

Бот для автоматической загрузки и парсинга Excel файлов с работ из Gmail, с отправкой результатов в REST API.

## Быстрый старт на Linux

### 1. Клонирование репозитория
```bash
cd /opt
git clone <your-repo-url> prs-parser
cd prs-parser
```

### 2. Настройка конфигурации
```bash
cp .env.example .env
# Отредактируйте .env и добавьте реальные значения:
nano .env
```

### 3. Запуск с Docker Compose

```bash
# Сборка образа и запуск контейнера
docker-compose up -d

# Просмотр логов
docker-compose logs -f prs-parser

# Остановка
docker-compose down
```

## Требования

- Docker & Docker Compose
- Gmail аккаунт с включённой 2FA
- App Password (16 символов, не обычный пароль)

## Диагностика проблем

### SSL Error на Linux: "socket error: EOF"

**Причина:** Отсутствуют CA сертификаты в контейнере.

**Решение:**
1. Пересоберите образ (Dockerfile содержит `update-ca-certificates`)
   ```bash
   docker-compose up -d --build
   ```

2. Проверьте логи:
   ```bash
   docker-compose logs prs-parser | grep ПОЧТА
   ```

3. Если ошибка сохраняется, проверьте на хосте:
   ```bash
   docker exec prs-parser python -c "import ssl; print(ssl.get_default_verify_paths())"
   ```

### App Password ошибка

Gmail не принимает обычные пароли через IMAP. Требуется:
1. Включить 2FA: https://myaccount.google.com/security
2. Создать App Password: https://myaccount.google.com/apppasswords
3. Использовать 16-символьный App Password в .env (без пробелов)

### Контейнер не начинает работу

```bash
# Проверить логи
docker-compose logs prs-parser

# Проверить что .env загружен
docker exec prs-parser env | grep PARSER

# Тестовое подключение к Gmail
docker exec prs-parser python -c "
import imaplib
import os
email = os.getenv('PARSER_EMAIL')
password = os.getenv('PARSER_PASSWORD')
try:
    imap = imaplib.IMAP4_SSL('imap.gmail.com', 993, timeout=10)
    imap.login(email, password)
    print('✓ OK')
    imap.close()
except Exception as e:
    print(f'✗ Error: {e}')
"
```

## Архитектура

- **main.py** - Основной скрипт
  - Подключается к Gmail IMAP (SSL на порту 993)
  - Скачивает .xlsx/.xls вложения
  - Парсит работы из Excel
  - Отправляет JSON результаты на REST API

- **Dockerfile** - Контейнизация
  - Python 3.11-slim
  - Установка CA сертификатов для SSL
  - Конвертация .env в Unix формат (LF вместо CRLF)

- **docker-compose.yml** - Оркестрация
  - Автоматический перезапуск контейнера
  - Подключение .env файла

## Формат данных

Парсёр отправляет на API массив объектов, каждый содержит:
```json
{
  "start_date": "2024-01-15",
  "brigade_number": 123,
  "well_name": "Скважина №5",
  "pump_type": "ЦБ-160",
  "car": "АВ1234АК",
  "device_number": "0234",
  "shift_details": ["Текст работы 1", "Текст работы 2"]
}
```

## Примечания

- Бот проверяет почту каждые 300 секунд (5 минут)
- Уже обработанные письма автоматически помечаются как прочитанные и удаляются
- Ошибки парсинга логируются без остановки работы бота
- На первом запуске контейнер может занять время при установке зависимостей

## API Endpoint

Парсёр отправляет данные на:
```
POST http://108.181.186.12:8023/prs-analytics/api/repairs/v1/summaries/parsed
Content-Type: application/json

{
  "summaries": [...]
}
```
