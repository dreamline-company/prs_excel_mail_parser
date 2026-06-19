#!/bin/bash

# Скрипт для быстрого развёртывания PRS Parser на Linux

set -e

echo "=== PRS Parser Deploy Script ==="
echo ""

# Проверяем Docker
if ! command -v docker &> /dev/null; then
    echo "✗ Docker не установлен. Установите Docker: https://docs.docker.com/install/"
    exit 1
fi

if ! command -v docker-compose &> /dev/null; then
    echo "✗ Docker Compose не установлен. Установите Docker Compose: https://docs.docker.com/compose/install/"
    exit 1
fi

echo "✓ Docker найден"
echo "✓ Docker Compose найден"
echo ""

# Проверяем .env
if [ ! -f .env ]; then
    echo "✗ .env файл не найден!"
    echo "Создаю .env из .env.example..."
    cp .env.example .env
    echo ""
    echo "⚠️  ВНИМАНИЕ: Отредактируйте .env и добавьте App Password:"
    echo "    nano .env"
    echo ""
    exit 1
fi

# Проверяем что .env содержит правильные переменные
if ! grep -q "PARSER_EMAIL" .env; then
    echo "✗ PARSER_EMAIL не найден в .env"
    exit 1
fi

if ! grep -q "PARSER_PASSWORD" .env; then
    echo "✗ PARSER_PASSWORD не найден в .env"
    exit 1
fi

echo "✓ .env файл настроен"
echo ""

# Останавливаем предыдущий контейнер если запущен
echo "Проверяю предыдущие экземпляры..."
docker-compose down 2>/dev/null || true

# Собираем и запускаем
echo ""
echo "Собираю Docker образ..."
docker-compose build

echo ""
echo "Запускаю контейнер..."
docker-compose up -d

echo ""
echo "=== ✓ Развёртывание завершено ==="
echo ""
echo "Просмотр логов:"
echo "  docker-compose logs -f prs-parser"
echo ""
echo "Остановка:"
echo "  docker-compose down"
echo ""
echo "Статус:"
echo "  docker-compose ps"
