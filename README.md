# PRS Excel Mail Parser

Email-бот для парсинга Excel файлов из Gmail и отправки в REST API.

## Требования

- Docker и Docker Compose установлены
- Файл `.env` с переменными окружения

## Переменные окружения (.env)

```
PARSER_EMAIL=<ваш Gmail адрес>
PARSER_PASSWORD=<16-символьный App Password>
SUMMARIES_API_URL=http://localhost:8023/prs-analytics/api/repairs/v1/summaries/parsed
```

## Быстрый запуск

### С Docker Compose (рекомендуется)

```bash
# Сборка и запуск
docker-compose up -d

# Просмотр логов
docker-compose logs -f

# Остановка
docker-compose down
```

### С Docker напрямую

```bash
# Сборка образа
docker build -t prs-parser .

# Запуск контейнера
docker run -d \
  --name prs-parser \
  --env-file .env \
  prs-parser

# Просмотр логов
docker logs -f prs-parser

# Остановка
docker stop prs-parser
docker rm prs-parser
```

## Структура проекта

```
.
├── main.py              # Почтовый бот (Gmail -> inbox/ -> API)
├── summary_parser.py    # Разбор xlsx-сводки в записи API
├── upload_summaries.py  # CLI-загрузка xlsx без почты
├── requirements.txt     # Python зависимости
├── Dockerfile          # Конфигурация Docker
├── docker-compose.yml  # Docker Compose конфигурация
├── .env                # Переменные окружения (не коммитить!)
├── .env.example        # Пример конфигурации
└── processed/          # Обработанные Excel файлы
```

## Загрузка без почты

`upload_summaries.py` разбирает те же xlsx и отправляет их в бэкенд напрямую —
удобно грузить архив сводок или проверять файл, не дожидаясь письма:

```bash
pip install -r requirements.txt
# все xlsx из папки (адрес API — из .env SUMMARIES_API_URL или --api-url)
python upload_summaries.py ~/Documents/prs
# только разобрать и посмотреть статистику, JSON сложить в ./parsed
python upload_summaries.py "Сводка ПРС за сентябрь.xlsx" --dry-run --out ./parsed
# после успешной загрузки переносить файлы в processed/
python upload_summaries.py ./inbox --move-to ./processed
```

Разбор книги — `summary_parser.py` (общий для бота и CLI): лист = сутки,
дата записи берётся из названия листа; секции по НГДУ; имя скважины собирается
как `<КОД_МЕСТОРОЖДЕНИЯ>_<NNNN>` (`BLG_0251`, `ZPV_403R`, `DSR_37/9`) из колонок
C и D; номер бригады — из «Бригада №N» (колонки C или B), колонка A — порядковый
номер, не бригада. Статистика разбора печатается по каждому файлу: неизвестные
месторождения и нераспознанные скважины стоит проверить глазами.

## Функционал

- ✅ Подключение к Gmail по IMAP SSL
- ✅ Скачивание .xlsx/.xls вложений
- ✅ Парсинг данных в JSON структуру
- ✅ Отправка на REST API
- ✅ Автоматическое удаление обработанных писем
- ✅ Фильтрация мусорных записей (Резерв, Ремонт и т.д.)
- ✅ Распознавание номеров машин (с фильтром телефонов)
- ✅ Нормализация Cyrillic → Latin в номерах

## Примечания

- Приложение работает постоянно в цикле, проверяя почту каждые `MAIL_CHECK_INTERVAL` секунд (300)
- Вложения складываются в `inbox/`, после загрузки уходят в `processed/`, при ошибке — в `failed/`
- Адрес бэкенда — `SUMMARIES_API_URL` в `.env`
