# PRS Excel Mail Parser

Email-бот для парсинга Excel файлов из Gmail и отправки в REST API.

## Требования

- Docker и Docker Compose установлены
- Файл `.env` с переменными окружения

## Переменные окружения (.env)

```
PARSER_EMAIL=<ваш Gmail адрес>
PARSER_PASSWORD=<16-символьный App Password>
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
  -v $(pwd)/results:/app/results \
  -v $(pwd)/processed:/app/processed \
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
├── main.py              # Основной скрипт
├── requirements.txt     # Python зависимости
├── Dockerfile          # Конфигурация Docker
├── docker-compose.yml  # Docker Compose конфигурация
├── .env                # Переменные окружения (не коммитить!)
├── .dockerignore       # Файлы исключаемые из образа
└── results/            # Выходные JSON файлы
    └── YYYY-MM/        # Организованы по месяцам
```

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

- Приложение работает постоянно в цикле, проверяя почту каждые 300 секунд
- Обработанные файлы перемещаются в папку `processed`
- Результаты сохраняются в `results/YYYY-MM/DD.MM.YY.json`
