"""Почтовый бот: забирает xlsx-сводки ПРС из Gmail и грузит их в бэкенд.

Разбор книги — ``summary_parser``, отправка — ``upload_summaries.send_records``.
Загрузить файлы без почты: ``python upload_summaries.py <папка или файлы>``.

Переменные окружения (.env):
    PARSER_EMAIL, PARSER_PASSWORD — Gmail и App Password;
    SUMMARIES_API_URL — адрес POST .../repairs/v1/summaries/parsed
        (по умолчанию локальный бэкенд на сервере);
    MAIL_CHECK_INTERVAL — период проверки почты, сек (300).
"""

import email
import imaplib
import os
import shutil
import socket
import ssl
import time
from email.header import decode_header
from pathlib import Path

import certifi
from dotenv import load_dotenv

from summary_parser import parse_workbook
from upload_summaries import DEFAULT_API_URL, format_result, send_records

load_dotenv()

EMAIL = os.getenv("PARSER_EMAIL")
PASSWORD = os.getenv("PARSER_PASSWORD")
API_URL = os.getenv("SUMMARIES_API_URL", DEFAULT_API_URL)
MAIL_CHECK_INTERVAL = int(os.getenv("MAIL_CHECK_INTERVAL", "300"))
# false — не проверять сертификат IMAP: в сети сервера FortiGate подменяет
# сертификат imap.gmail.com своим. Принимается любой сертификат, решение
# владельца (30.09.2026); по умолчанию проверка включена.
IMAP_SSL_VERIFY = os.getenv("IMAP_SSL_VERIFY", "true").lower() != "false"
GMAIL_IMAP_SERVER = "imap.gmail.com"
GMAIL_IMAP_PORT = 993
script_dir = Path(__file__).resolve().parent
inbox_dir = script_dir / "inbox"
processed_dir = script_dir / "processed"
failed_dir = script_dir / "failed"

if not EMAIL or not PASSWORD:
    print("✗ КРИТИЧЕСКАЯ ОШИБКА: Проверьте файл .env")
    print(f"  PARSER_EMAIL: {'✓ загружен' if EMAIL else '✗ ОТСУТСТВУЕТ'}")
    print(f"  PARSER_PASSWORD: {'✓ загружен' if PASSWORD else '✗ ОТСУТСТВУЕТ'}")
    raise SystemExit(1)


def connect_to_gmail():
    """Подключение к Gmail по IMAP SSL с явным контекстом сертификатов"""
    # Проверяем переменные окружения
    if not EMAIL:
        print("[ПОЧТА] ✗ КРИТИЧЕСКАЯ ОШИБКА: PARSER_EMAIL не установлен в .env!")
        return None
    if not PASSWORD:
        print("[ПОЧТА] ✗ КРИТИЧЕСКАЯ ОШИБКА: PARSER_PASSWORD не установлен в .env!")
        return None
    
    try:
        print(f"[ПОЧТА] Подключаюсь к {GMAIL_IMAP_SERVER}:{GMAIL_IMAP_PORT}...")
        print(f"[ПОЧТА] Email: {EMAIL}")
        
        # Создаём SSL контекст с явным путём к сертификатам
        ssl_context = ssl.create_default_context(cafile=certifi.where())
        ssl_context.check_hostname = IMAP_SSL_VERIFY
        ssl_context.verify_mode = ssl.CERT_REQUIRED if IMAP_SSL_VERIFY else ssl.CERT_NONE
        if not IMAP_SSL_VERIFY:
            print("[ПОЧТА] Проверка сертификата IMAP отключена (IMAP_SSL_VERIFY=false)")
        
        imap = imaplib.IMAP4_SSL(GMAIL_IMAP_SERVER, GMAIL_IMAP_PORT, ssl_context=ssl_context, timeout=10)
        print(f"[ПОЧТА] Авторизация пользователя: {EMAIL}...")
        imap.login(EMAIL, PASSWORD)
        print("[ПОЧТА] ✓ Успешно авторизован в Gmail")
        return imap
    except imaplib.IMAP4.error as e:
        print(f"[ПОЧТА] ✗ Ошибка IMAP: {e}")
        print("[ПОЧТА] Возможные причины:")
        print("  1. App Password неправильный (используй пароль приложения, не основной)")
        print("  2. 2FA (двухфакторная аутентификация) не включена на аккаунте")
        print("  3. Email неправильный в .env файле")
        print(f"[ПОЧТА] Email в конфиге: {EMAIL}")
        return None
    except ssl.SSLError as e:
        print(f"[ПОЧТА] ✗ SSL ошибка: {e}")
        print(f"[ПОЧТА] Используются сертификаты из: {certifi.where()}")
        return None
    except socket.timeout:
        print("[ПОЧТА] ✗ Timeout - сервер Gmail не отвечает. Проверьте интернет соединение")
        return None
    except socket.error as e:
        print(f"[ПОЧТА] ✗ Socket error: {e}")
        print("[ПОЧТА] На Linux это часто означает отсутствие CA сертификатов!")
        print("[ПОЧТА] Решение для Docker:")
        print("  RUN apt-get update && apt-get install -y ca-certificates && update-ca-certificates")
        return None
    except Exception as e:
        print(f"[ПОЧТА] ✗ Ошибка подключения ({type(e).__name__}): {e}")
        import traceback
        traceback.print_exc()
        return None

def get_gmail_attachments():
    """Скачивает .xlsx/.xls вложения из непрочитанных писем Gmail"""
    imap = connect_to_gmail()
    if not imap:
        return []

    result = []
    try:
        imap.select("INBOX")
        status, messages = imap.search(None, "UNSEEN")
        message_ids = messages[0].split()
        print(f"[ПОЧТА] Непрочитанных писем: {len(message_ids)}")

        for msg_id in message_ids:
            try:
                status, msg_data = imap.fetch(msg_id, "(RFC822)")
                msg = email.message_from_bytes(msg_data[0][1])

                has_xlsx = False
                for part in msg.walk():
                    if part.get_content_disposition() != "attachment":
                        continue
                    raw_filename = part.get_filename()
                    if not raw_filename:
                        continue
                    decoded = decode_header(raw_filename)
                    filename_bytes, charset = decoded[0]
                    if isinstance(filename_bytes, bytes):
                        filename = filename_bytes.decode(charset or "utf-8", errors="replace")
                    else:
                        filename = filename_bytes

                    if not filename.lower().endswith((".xlsx", ".xls")):
                        continue

                    has_xlsx = True
                    file_data = part.get_payload(decode=True)
                    print(f"[ПОЧТА] Найдено вложение: {filename}")
                    result.append({
                        "filename": filename,
                        "data": file_data,
                        "msg_id": msg_id
                    })

                # Письма без xlsx помечаем как прочитанные (чтобы не мешали)
                # Письма с xlsx НЕ трогаем — пометим/удалим после успешного сохранения
                if not has_xlsx:
                    imap.store(msg_id, "+FLAGS", "\\Seen")

            except Exception as e:
                print(f"[ПОЧТА] Ошибка обработки письма {msg_id}: {e}")

        imap.close()
        imap.logout()
    except Exception as e:
        print(f"[ПОЧТА] Ошибка: {e}")
        try:
            imap.close()
            imap.logout()
        except Exception:
            pass

    return result

def delete_gmail_message(msg_id):
    """Удаляет письмо из Gmail после успешного скачивания вложения"""
    imap = connect_to_gmail()
    if not imap:
        return False
    try:
        imap.select("INBOX")
        imap.store(msg_id, "+FLAGS", "\\Deleted")
        imap.expunge()
        imap.close()
        imap.logout()
        print(f"[ПОЧТА] Письмо {msg_id.decode()} удалено")
        return True
    except Exception as e:
        print(f"[ПОЧТА] Ошибка удаления письма {msg_id}: {e}")
        try:
            imap.close()
            imap.logout()
        except Exception:
            pass
        return False


def fetch_mail_to_inbox() -> None:
    """Скачать вложения из непрочитанных писем в inbox/ и удалить письма."""
    attachments = get_gmail_attachments()
    msg_status: dict[bytes, dict[str, int]] = {}
    for att in attachments:
        mid = att["msg_id"]
        msg_status.setdefault(mid, {"total": 0, "saved": 0})
        msg_status[mid]["total"] += 1
        try:
            base, ext = os.path.splitext(att["filename"])
            target = inbox_dir / f"{base}_{mid.decode()}{ext}"
            target.write_bytes(att["data"])
            msg_status[mid]["saved"] += 1
            print(f"[ПОЧТА] Сохранён: {target.name}")
        except Exception as e:  # noqa: BLE001
            print(f"[ПОЧТА] Ошибка сохранения файла {att['filename']}: {e}")
    for mid, counts in msg_status.items():
        if counts["saved"] == counts["total"]:
            delete_gmail_message(mid)
        else:
            print(f"[ПОЧТА] Письмо {mid.decode()} оставлено (ошибка сохранения, повтор при следующей проверке)")


def process_inbox() -> None:
    """Разобрать и отправить все xlsx из inbox/; успешные — в processed/, ошибочные — в failed/."""
    files = sorted(p for p in inbox_dir.iterdir() if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$"))
    for path in files:
        print(f"\n[ПАРСЕР] {path.name}")
        try:
            result = parse_workbook(path)
        except Exception as e:  # noqa: BLE001
            print(f"[ПАРСЕР] Ошибка разбора: {type(e).__name__}: {e}")
            shutil.move(str(path), str(failed_dir / path.name))
            continue
        print(result.report.summary())
        if not result.records:
            print("[API] нет данных для отправки")
            shutil.move(str(path), str(processed_dir / path.name))
            continue
        ok, body = send_records(result.records, API_URL)
        print(f"[API] {'OK' if ok else 'ОШИБКА'}: {format_result(body)}")
        shutil.move(str(path), str((processed_dir if ok else failed_dir) / path.name))


if __name__ == "__main__":
    for d in (inbox_dir, processed_dir, failed_dir):
        d.mkdir(exist_ok=True)
    print("\nРобот успешно запущен!")
    print(f"Папка для работы: {script_dir}")
    print(f"API: {API_URL}")
    print("Перехожу в режим ожидания писем...")
    last_mail_check = 0.0
    while True:
        if time.time() - last_mail_check > MAIL_CHECK_INTERVAL:
            last_mail_check = time.time()
            print("[СИСТЕМА] Проверка почты Gmail...")
            try:
                fetch_mail_to_inbox()
            except Exception as e:  # noqa: BLE001
                print(f"[ПОЧТА] Сбой проверки почты: {type(e).__name__}: {e}")
        try:
            process_inbox()
        except Exception as e:  # noqa: BLE001
            print(f"[ПАРСЕР] Сбой обработки inbox: {type(e).__name__}: {e}")
        time.sleep(15)
