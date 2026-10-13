"""Почтовый бот: забирает xlsx-сводки ПРС из почтового ящика (IMAP) и грузит их в бэкенд.

Разбор книги — ``summary_parser``, отправка — ``upload_summaries.send_records``.
Загрузить файлы без почты: ``python upload_summaries.py <папка или файлы>``.

Переменные окружения (.env):
    PARSER_EMAIL, PARSER_PASSWORD — адрес ящика и пароль;
    PARSER_LOGIN — логин IMAP, если отличается от адреса (например,
        ``DOMAIN\\user`` для Exchange); по умолчанию — PARSER_EMAIL;
    IMAP_HOST, IMAP_PORT — сервер IMAP SSL (imap.gmail.com:993; корпоративный
        Exchange — mail.emg.kmg.kz:993);
    IMAP_FOLDER — папка со сводками (INBOX);
    IMAP_SSL_VERIFY — false, если сертификат сервера не проверяется
        (Exchange: самоподписанный, Gmail в сети сервера: подменён FortiGate);
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
LOGIN = os.getenv("PARSER_LOGIN") or EMAIL
API_URL = os.getenv("SUMMARIES_API_URL", DEFAULT_API_URL)
MAIL_CHECK_INTERVAL = int(os.getenv("MAIL_CHECK_INTERVAL", "300"))
IMAP_HOST = os.getenv("IMAP_HOST", "imap.gmail.com")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
IMAP_FOLDER = os.getenv("IMAP_FOLDER", "INBOX")
# false — не проверять сертификат IMAP: у корпоративного Exchange он
# самоподписанный (CN=emg-ex-mbx01), у Gmail в сети сервера его подменяет
# FortiGate. Принимается любой сертификат, решение владельца (30.09.2026);
# по умолчанию проверка включена.
IMAP_SSL_VERIFY = os.getenv("IMAP_SSL_VERIFY", "true").lower() != "false"
script_dir = Path(__file__).resolve().parent
inbox_dir = script_dir / "inbox"
processed_dir = script_dir / "processed"
failed_dir = script_dir / "failed"

if not EMAIL or not PASSWORD:
    print("✗ КРИТИЧЕСКАЯ ОШИБКА: Проверьте файл .env")
    print(f"  PARSER_EMAIL: {'✓ загружен' if EMAIL else '✗ ОТСУТСТВУЕТ'}")
    print(f"  PARSER_PASSWORD: {'✓ загружен' if PASSWORD else '✗ ОТСУТСТВУЕТ'}")
    raise SystemExit(1)


def _ssl_context() -> ssl.SSLContext:
    context = ssl.create_default_context(cafile=certifi.where())
    context.check_hostname = IMAP_SSL_VERIFY
    context.verify_mode = ssl.CERT_REQUIRED if IMAP_SSL_VERIFY else ssl.CERT_NONE
    return context


def connect_to_mail():
    """Подключение к почтовому серверу по IMAP SSL; None — не удалось."""
    try:
        print(f"[ПОЧТА] Подключаюсь к {IMAP_HOST}:{IMAP_PORT} (ящик {EMAIL})...")
        if not IMAP_SSL_VERIFY:
            print("[ПОЧТА] Проверка сертификата IMAP отключена (IMAP_SSL_VERIFY=false)")
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=_ssl_context(), timeout=30)
        imap.login(LOGIN, PASSWORD)
        print(f"[ПОЧТА] ✓ Авторизован как {LOGIN}")
        return imap
    except imaplib.IMAP4.error as e:
        print(f"[ПОЧТА] ✗ Ошибка IMAP (логин {LOGIN}): {e}")
        print("[ПОЧТА] Проверьте PARSER_PASSWORD; для Exchange попробуйте")
        print("  PARSER_LOGIN в виде DOMAIN\\user вместо адреса ящика.")
        if IMAP_HOST == "imap.gmail.com":
            print("  Gmail принимает только App Password (нужна 2FA).")
        return None
    except ssl.SSLError as e:
        print(f"[ПОЧТА] ✗ SSL ошибка: {e}")
        print("[ПОЧТА] Сертификат сервера не проходит проверку — IMAP_SSL_VERIFY=false")
        return None
    except (TimeoutError, socket.timeout):
        print(f"[ПОЧТА] ✗ Timeout — {IMAP_HOST}:{IMAP_PORT} не отвечает")
        return None
    except OSError as e:
        print(f"[ПОЧТА] ✗ Сетевая ошибка {IMAP_HOST}:{IMAP_PORT}: {e}")
        return None
    except Exception as e:  # noqa: BLE001
        print(f"[ПОЧТА] ✗ Ошибка подключения ({type(e).__name__}): {e}")
        import traceback
        traceback.print_exc()
        return None


def _close(imap) -> None:
    for step in (imap.close, imap.logout):
        try:
            step()
        except Exception:  # noqa: BLE001
            pass


def _decode_filename(raw_filename: str) -> str:
    filename_bytes, charset = decode_header(raw_filename)[0]
    if isinstance(filename_bytes, bytes):
        return filename_bytes.decode(charset or "utf-8", errors="replace")
    return filename_bytes


def get_mail_attachments():
    """Скачивает .xlsx/.xls вложения из непрочитанных писем.

    Письма адресуются по UID, а не по порядковому номеру: номер сдвигается,
    как только из папки удаляют письмо, и удаление по старым номерам
    задело бы чужие письма. Письма без xlsx помечаются прочитанными; письма
    с xlsx не трогаются — их удалит ``delete_messages`` после сохранения.
    """
    imap = connect_to_mail()
    if not imap:
        return []

    result = []
    try:
        imap.select(IMAP_FOLDER)
        status, data = imap.uid("SEARCH", None, "UNSEEN")
        uids = data[0].split() if status == "OK" and data and data[0] else []
        print(f"[ПОЧТА] Непрочитанных писем в {IMAP_FOLDER}: {len(uids)}")

        for uid in uids:
            try:
                # BODY.PEEK — не ставить \Seen при скачивании: письмо с
                # вложением остаётся непрочитанным до удаления.
                status, msg_data = imap.uid("FETCH", uid, "(BODY.PEEK[])")
                raw = next(
                    (part[1] for part in msg_data if isinstance(part, tuple)),
                    None,
                )
                if status != "OK" or raw is None:
                    print(f"[ПОЧТА] Письмо UID {uid.decode()} не скачалось: {status}")
                    continue
                msg = email.message_from_bytes(raw)

                has_xlsx = False
                for part in msg.walk():
                    if part.get_content_disposition() != "attachment":
                        continue
                    raw_filename = part.get_filename()
                    if not raw_filename:
                        continue
                    filename = _decode_filename(raw_filename)
                    if not filename.lower().endswith((".xlsx", ".xls")):
                        continue

                    has_xlsx = True
                    print(f"[ПОЧТА] Найдено вложение: {filename}")
                    result.append({
                        "filename": filename,
                        "data": part.get_payload(decode=True),
                        "msg_id": uid,
                    })

                # Письма без xlsx помечаем как прочитанные (чтобы не мешали)
                if not has_xlsx:
                    imap.uid("STORE", uid, "+FLAGS", "(\\Seen)")

            except Exception as e:  # noqa: BLE001
                print(f"[ПОЧТА] Ошибка обработки письма UID {uid}: {e}")
    except Exception as e:  # noqa: BLE001
        print(f"[ПОЧТА] Ошибка: {e}")
    finally:
        _close(imap)

    return result


def delete_messages(uids) -> bool:
    """Удалить письма по UID одним подключением (после сохранения вложений)."""
    if not uids:
        return True
    imap = connect_to_mail()
    if not imap:
        return False
    try:
        imap.select(IMAP_FOLDER)
        for uid in uids:
            imap.uid("STORE", uid, "+FLAGS", "(\\Deleted)")
        imap.expunge()
        print(f"[ПОЧТА] Удалено писем: {len(uids)} (UID {', '.join(u.decode() for u in uids)})")
        return True
    except Exception as e:  # noqa: BLE001
        print(f"[ПОЧТА] Ошибка удаления писем {uids}: {e}")
        return False
    finally:
        _close(imap)


def fetch_mail_to_inbox() -> None:
    """Скачать вложения из непрочитанных писем в inbox/ и удалить письма."""
    attachments = get_mail_attachments()
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
    done = [mid for mid, counts in msg_status.items() if counts["saved"] == counts["total"]]
    for mid, counts in msg_status.items():
        if counts["saved"] != counts["total"]:
            print(f"[ПОЧТА] Письмо UID {mid.decode()} оставлено (ошибка сохранения, повтор при следующей проверке)")
    delete_messages(done)


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
    print(f"Почта: {EMAIL} @ {IMAP_HOST}:{IMAP_PORT}/{IMAP_FOLDER}")
    print(f"API: {API_URL}")
    print("Перехожу в режим ожидания писем...")
    last_mail_check = 0.0
    while True:
        if time.time() - last_mail_check > MAIL_CHECK_INTERVAL:
            last_mail_check = time.time()
            print(f"[СИСТЕМА] Проверка почты {EMAIL} ({IMAP_HOST})...")
            try:
                fetch_mail_to_inbox()
            except Exception as e:  # noqa: BLE001
                print(f"[ПОЧТА] Сбой проверки почты: {type(e).__name__}: {e}")
        try:
            process_inbox()
        except Exception as e:  # noqa: BLE001
            print(f"[ПАРСЕР] Сбой обработки inbox: {type(e).__name__}: {e}")
        time.sleep(15)
