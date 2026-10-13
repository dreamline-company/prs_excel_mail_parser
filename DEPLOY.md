# PRS Parser — почтовый бот сводок ПРС (IMAP)

Бот забирает Excel-сводки ПРС из почтового ящика и отправляет результаты в REST API.
Рабочий ящик — `emg_prs-svodki@emg.kmg.kz` на корпоративном Exchange `mail.emg.kmg.kz`
(10.32.10.62 / 10.32.10.63 — та же сеть, что и сервер 10.32.10.90; Gmail с сервера недоступен).

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
- Ящик Exchange `emg_prs-svodki@emg.kmg.kz` и его пароль (IMAP SSL 993 на `mail.emg.kmg.kz`)
- Для Gmail (`IMAP_HOST=imap.gmail.com`): 2FA и App Password

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

### Ошибка входа (IMAP: LOGIN failed / AUTHENTICATE failed)

- Exchange: проверьте пароль; если не принимает адрес ящика, задайте в `.env`
  `PARSER_LOGIN=EMG\emg_prs-svodki` (домен\пользователь; имя домена уточнить у админов).
- Gmail: обычные пароли через IMAP не принимаются — нужен App Password при 2FA
  (https://myaccount.google.com/apppasswords).

### SSL: CERTIFICATE_VERIFY_FAILED

Сертификат Exchange самоподписанный (CN=emg-ex-mbx01): нужен `IMAP_SSL_VERIFY=false`
(в docker-compose это значение по умолчанию).

### Контейнер не начинает работу

```bash
# Проверить логи
docker-compose logs prs-parser

# Проверить что .env загружен
docker exec prs-parser env | grep PARSER

# Тестовое подключение к ящику (те же настройки, что у бота)
docker exec prs-parser python -c "
import main
imap = main.connect_to_mail()
print('✓ OK' if imap else '✗ см. сообщение выше')
imap and imap.logout()
"
```

## Архитектура

- **main.py** - Основной скрипт
  - Подключается к ящику по IMAP SSL (`IMAP_HOST:IMAP_PORT`, по умолчанию Exchange `mail.emg.kmg.kz:993`)
  - Письма адресуются по UID; письмо удаляется только после сохранения всех его вложений
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
