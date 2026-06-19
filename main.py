import requests
import time
import json
import os
import re
import openpyxl
import shutil
import datetime
import imaplib
import email
import socket
from email.header import decode_header
from dotenv import load_dotenv
# Глушим предупреждения о подменных сертификатах (защита от SSL ошибок)
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

load_dotenv()
# ==========================================
# 1. НАСТРОЙКИ ПОЧТЫ И ПУТЕЙ
# ==========================================
EMAIL = os.getenv("PARSER_EMAIL")
PASSWORD = os.getenv("PARSER_PASSWORD")
API_URL = "http://108.181.186.12:8023/prs-analytics/api/repairs/v1/summaries/parsed"
GMAIL_IMAP_SERVER = "imap.gmail.com"
GMAIL_IMAP_PORT = 993
script_dir = os.path.dirname(os.path.abspath(__file__))

# Проверяем что Email и Password загружены
if not EMAIL or not PASSWORD:
    print("✗ КРИТИЧЕСКАЯ ОШИБКА: Проверьте файл .env")
    print(f"  PARSER_EMAIL: {'✓ загружен' if EMAIL else '✗ ОТСУТСТВУЕТ'}")
    print(f"  PARSER_PASSWORD: {'✓ загружен' if PASSWORD else '✗ ОТСУТСТВУЕТ'}")
    print("\nВсе параметры должны быть указаны в .env файле!")
    exit(1)

# ==========================================
# 2. СЛОВАРИ ДЛЯ ПАРСЕРА
# ==========================================
CAR_PLATE_MAP = {
    "024BH": "024BH06", "250AKD": "250AMD", "564AZ06": "564AY06",
    "E247AHD": "AHD247E", "153AAS": "153AS06", "153AC": "153AS06", "259ALD06": "ALD259E",
    "480AZD": "AZD480E", "730AU": "730AU06"
}

FIELD_CODE_MAP = {
    "акингень": "AKG", "акинген": "AKG",
    "аккудук": "AKD", "аккудык": "AKD",
    "актюбе": "ATB", "актобе": "ATB",
    "алтыколь": "ATK", "алтыкуль": "ATK",
    "б. жоламанов": "JLM", "б жоламанов": "JLM", "б.жоламанов": "JLM",
    "ботахан": "BTN",
    "в. макат": "VMT", "в макат": "VMT", "в.макат": "VMT",
    "восточный молдабек": "VMB", "ш. молдабек": "VMB", "ш молдабек": "VMB", "ш.молдабек": "VMB",
    "гран": "GRN",
    "досмухамбетовское": "DMB", "досмухамбетиовскоее": "DMB",
    "доссор": "DSR",
    "жанаталап": "ZHT",
    "забурунье": "ZBN",
    "западная прорва": "ZPV", "з. прорва": "ZPV", "з прорва": "ZPV", "з.прорва": "ZPV", "зап.прорва": "ZPV",
    "карасор западный": "ZKS",
    "каратон": "KRT",
    "карсак": "KRK",
    "кисимбай": "KSB", "кисымбай": "KSB",
    "косшагил": "KSG", "косшагыл": "KSG",
    "кошкар": "KKR",
    "кульсары": "KLR",
    "новобогатинск ю-в": "UVN", "новобогатинск юв": "UVN", "ювн": "UVN",
    "с. балгимбаев": "BLG", "с балгимбаев": "BLG", "с.балгимбаев": "BLG",
    "с. жолдыбай": "SJB", "с жолдыбай": "SJB", "с.жолдыбай": "SJB",
    "с. нуржанов": "NRG", "с нуржанов": "NRG", "с.нуржанов": "NRG",
    "северный котыртас": "SKS", "с.қотыртас": "SKS", "с қотыртас": "SKS", "с.котыртас": "SKS", "с котыртас": "SKS", "котыртас": "SKS",
    "терен-узек": "TNU", "терен узек": "TNU",
    "уаз": "UAZ", "с. уаз": "UAZ", "с уаз": "UAZ", "с.уаз": "UAZ",
    "уаз восточный": "UZV", "ш.уаз": "UZV",
    "уаз северный": "UZS",
    "ю-в камышитовое": "UVK", "ювк": "UVK",
    "ю-з камышитовое": "UZK", "юзк": "UZK", "южный забурунный купол": "UZK",
    "жайықмұнайгаз": "JMG", "жайыкмунайгаз": "JMG", "жмг": "JMG", "jmg": "JMG",
    "жылыоймұнайгаз": "ZhylMG", "жылыоймунайгаз": "ZhylMG", "жылмг": "ZhylMG", "жылымг": "ZhylMG", "zhylmg": "ZhylMG", "zmg": "ZhylMG",
    "доссормұнайгаз": "DMG", "доссормунайгаз": "DMG", "дмг": "DMG", "dmg": "DMG",
    "қайнармұнайгаз": "KMG", "кайнармунайгаз": "KMG", "кмг": "KMG", "kmg": "KMG"
}

FIELD_NAME_MAP = {
    "AKG": "Акинген", "AKD": "Аккудык", "ATB": "Актобе", "ATK": "Алтыкуль",
    "JLM": "Б. Жоламанов", "BTN": "Ботахан", "VMT": "В. Макат", "VMB": "Ш. Молдабек",
    "GRN": "Гран", "DMB": "Досмухамбетовское", "DSR": "Доссор", "ZHT": "Жанаталап",
    "ZBN": "Забурунье", "ZPV": "Западная Прорва", "ZKS": "Карасор Западный",
    "KRT": "Каратон", "KRK": "Карсак", "KSB": "Кисымбай", "KSG": "Косшагыл",
    "KKR": "Кошкар", "KLR": "Кульсары", "UVN": "Новобогатинск Ю-В",
    "BLG": "С. Балгимбаев", "SJB": "С. Жолдыбай", "NRG": "С. Нуржанов",
    "SKS": "Северный Котыртас", "TNU": "Терен-Узек", "UAZ": "С. Уаз",
    "UZV": "Уаз Восточный", "UZS": "Уаз Северный", "UVK": "Ю-В Камышитовое",
    "UZK": "Ю-З Камышитовое",
    "JMG": "НГДУ Жайыкмунайгаз",
    "ZhylMG": "НГДУ Жылыоймунайгаз",
    "DMG": "НГДУ Доссормунайгаз",
    "KMG": "НГДУ Кайнармунайгаз"
}

def format_date(raw_date):
    # 1. Если это уже объект даты (обработка openpyxl)
    if isinstance(raw_date, datetime.datetime):
        return raw_date.strftime("%d.%m.%Y")
    
    # 2. Если это число (стандартный формат Excel)
    if isinstance(raw_date, (int, float)):
        # Excel считает дни от 30 декабря 1899 года
        dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(raw_date))
        return dt.strftime("%d.%m.%Y")
        
    # 3. Если это строка — применяем вашу логику с регуляркой
    s = str(raw_date).strip().replace("\n", "").replace(" ", "")
    match = re.search(r'(\d{2})[.-](\d{2})[.-](\d{2,4})', s)
    
    if match:
        day, month, year = match.groups()
        if len(year) == 2:
            year = "20" + year
        elif len(year) > 4:
            year = year[-4:]
        return f"{day}.{month}.{year}"
        
    # 4. Fallback: если ничего не подошло
    return datetime.datetime.now().strftime("%d.%m.%Y")

def send_to_api(data):
    try:
        # Ключевое изменение: оборачиваем список в словарь с ключом 'summaries'
        payload = {"summaries": data}
        
        r = requests.post(
            API_URL,
            json=payload, # Теперь данные отправляются в формате {"summaries": [...]}
            timeout=60
        )

        print(f"[API] статус: {r.status_code}")
        # Если статус не 200, мы увидим подробную ошибку от сервера
        if r.status_code != 200:
            print(f"[API] подробности ошибки: {r.text}")

        return r.status_code == 200

    except Exception as e:
        print(f"[API] ошибка отправки: {e}")
        return False


# ==========================================
# 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==========================================
def normalize_plate(plate):
    p = str(plate).upper()
    cyr_to_lat = {
        'А': 'A', 'В': 'B', 'С': 'C', 'Е': 'E', 'Н': 'H',
        'К': 'K', 'М': 'M', 'О': 'O', 'Р': 'P', 'Т': 'T',
        'Х': 'X', 'У': 'Y'
    }
    for cyr, lat in cyr_to_lat.items():
        p = p.replace(cyr, lat)
    return p


def get_field_info(name):
    clean_name = str(name).lower().strip()
    if clean_name in FIELD_CODE_MAP:
        code = FIELD_CODE_MAP[clean_name]
        full_name = FIELD_NAME_MAP.get(code, name)
        return code, full_name
    return "UNKNOWN", name

def clean_spaces(text):
    if not text: return ""
    return re.sub(r'[ \t]+', ' ', str(text)).strip()

def parse_shift_text(raw_text):
    text_str = str(raw_text or "").strip()
    if not text_str or text_str.lower() in ("none", "нет", "-", "."):
        return []
    
    text_str = text_str.replace("::", ":")
    raw_lines = text_str.split('\n')
    temp_parts = []
    
    for line in raw_lines:
        parts_time = re.split(r'\s*(?=\d{2}[.:]\d{2}\s*-\s*\d{2}[.:]\d{2})', line)
        temp_parts.extend(parts_time)

    lines_list = []
    for part in temp_parts:
        cleaned = clean_spaces(part)
        cleaned = re.sub(r'^[Вв]р\s*\.?\s*', '', cleaned)
        cleaned = re.sub(r'\s*[Вв]р\s*\.?\s*$', '', cleaned)
        cleaned = clean_spaces(cleaned)
        
        if cleaned:
            cleaned = cleaned.replace('"', "'")
            if not cleaned.endswith('.'):
                cleaned += '.'
            lines_list.append(cleaned)
    return lines_list

def generate_clean_filename(original_name, sheet_name):
    date_match = re.search(r'(\d{2})[-.](\d{2})[-.](\d{2,4})', sheet_name)
    if date_match:
        day, month, year = date_match.groups()
        if len(year) == 4:
            year = year[-2:]
        return f"{day}.{month}.{year}"
        
    date_match = re.search(r'(\d{2})[-._\s](\d{2})[-._\s](\d{2,4})', original_name)
    if date_match:
        day, month, year = date_match.groups()
        if len(year) == 4:
            year = year[-2:]
        return f"{day}.{month}.{year}"
    
    clean_name = os.path.splitext(original_name)[0]
    clean_name = re.sub(r'[^a-zA-Z0-9]', '_', clean_name)
    clean_name = re.sub(r'_+', '_', clean_name).strip('_')
    return clean_name.lower()


def make_key(item):
    return (
        item.get("start_date"),
        item.get("well_name"),
        item.get("shift_type_number"),
        item.get("brigade_number"),
        item.get("car"),
        tuple(item.get("shift_details", []))
    )

def connect_to_gmail():
    """Подключение к Gmail по IMAP SSL"""
    try:
        print(f"[ПОЧТА] Подключаюсь к {GMAIL_IMAP_SERVER}:{GMAIL_IMAP_PORT}...")
        imap = imaplib.IMAP4_SSL(GMAIL_IMAP_SERVER, GMAIL_IMAP_PORT, timeout=10)
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
    except socket.timeout:
        print("[ПОЧТА] ✗ Timeout - сервер Gmail не отвечает. Проверьте интернет соединение")
        return None
    except Exception as e:
        print(f"[ПОЧТА] ✗ Ошибка подключения: {type(e).__name__}: {e}")
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

# ==========================================
# 4. ЗАПУСК И ГЛАВНЫЙ ЦИКЛ
# ==========================================
print(f"\nРобот успешно запущен!")
print(f"Папка для работы: {script_dir}")
print("Перехожу в режим ожидания писем...")
last_mail_check = 0
while True:
    # ---------------------------------------------------------
    # ШАГ 1: ПРОВЕРКА ПОЧТЫ И СКАЧИВАНИЕ НОВЫХ ФАЙЛОВ (Gmail IMAP)
    # ---------------------------------------------------------
    if time.time() - last_mail_check > 300:
        last_mail_check = time.time()
        print("[СИСТЕМА] Проверка почты Gmail...")

        attachments = get_gmail_attachments()

        # msg_id -> {"total": N, "saved": N} — отслеживаем успех по каждому письму
        msg_status = {}
        for att in attachments:
            mid = att["msg_id"]
            if mid not in msg_status:
                msg_status[mid] = {"total": 0, "saved": 0}
            msg_status[mid]["total"] += 1

            try:
                base, ext = os.path.splitext(att["filename"])
                unique_name = f"{base}_{att['msg_id'].decode()}{ext}"
                file_path = os.path.join(script_dir, unique_name)

                with open(file_path, "wb") as f:
                    f.write(att["data"])

                msg_status[mid]["saved"] += 1
                print(f"[ПОЧТА] Сохранён: {unique_name}")
            except Exception as e:
                print(f"[ПОЧТА] Ошибка сохранения файла {att['filename']}: {e}")

        # Удаляем письма где все вложения сохранены успешно
        for mid, counts in msg_status.items():
            if counts["saved"] == counts["total"]:
                delete_gmail_message(mid)
            else:
                print(f"[ПОЧТА] Письмо {mid.decode()} оставлено (ошибка сохранения, повтор при следующей проверке)")


    # ---------------------------------------------------------
    # ШАГ 2: ПАРСИНГ ВСЕХ EXCEL ФАЙЛОВ В ПАПКЕ
    # ---------------------------------------------------------
    processed_dir = os.path.join(script_dir, "processed")
    if not os.path.exists(processed_dir):
        os.makedirs(processed_dir)

    excel_files = [f for f in os.listdir(script_dir) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]

    for target_file in excel_files:
        excel_path = os.path.join(script_dir, target_file)
        print(f"\n[ПАРСЕР] Начинаю обработку нового файла: '{target_file}'")

        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            first_sheet_title = wb.sheetnames[0]
            result_json = []

            for sheet in wb.worksheets:
                print(f"  -> Обрабатываем лист: {sheet.title}")

                target_month = None
                target_year = None

                MONTH_WORDS = {
                    "январ": "01", "феврал": "02", "март": "03", "апрел": "04",
                    "май": "05", "мая": "05", "июн": "06", "июл": "07",
                    "август": "08", "сентябр": "09", "октябр": "10",
                    "ноябр": "11", "декабр": "12"
                }

                target_file_lower = target_file.lower()

                for word, m_num in MONTH_WORDS.items():
                    if word in target_file_lower:
                        target_month = m_num
                        break

                year_match = re.search(r'(202\d)', target_file_lower)
                if year_match:
                    target_year = year_match.group(1)

                if not target_month:
                    date_match = re.search(r'\d{2}[-.](\d{2})(?:[-.](\d{2,4}))?', sheet.title)
                    if date_match:
                        target_month = date_match.group(1)
                        if date_match.group(2) and not target_year:
                            target_year = date_match.group(2)
                            if len(target_year) == 2:
                                target_year = "20" + target_year

                if not target_year:
                    target_year = str(datetime.datetime.now().year)

                last_master_and_car = None
                last_field_and_brigade = None
                last_well = None
                last_start_date = None
                last_end_date = None

                for row in range(4, sheet.max_row + 1):
                    raw_b = sheet[f"B{row}"].value
                    raw_c = sheet[f"C{row}"].value
                    raw_d = sheet[f"D{row}"].value
                    raw_n = sheet[f"N{row}"].value
                    raw_o = sheet[f"O{row}"].value

                    if (raw_b and str(raw_b).strip()) or (raw_d and str(raw_d).strip()):
                        last_end_date = None

                    if raw_b is not None and str(raw_b).strip():
                        last_master_and_car = raw_b
                    if raw_c is not None and str(raw_c).strip():
                        last_field_and_brigade = raw_c
                    if raw_d is not None and str(raw_d).strip():
                        last_well = raw_d
                    if raw_n is not None and str(raw_n).strip():
                        last_start_date = raw_n
                    if raw_o is not None and str(raw_o).strip():
                        last_end_date = raw_o

                    shift_1_details = sheet[f"P{row}"].value
                    shift_2_details = sheet[f"Q{row}"].value

                    s1_str = str(shift_1_details or "").strip()
                    s2_str = str(shift_2_details or "").strip()

                    if s1_str in ("", "None", "-", "нет", ".") and s2_str in ("", "None", "-", "нет", "."):
                        continue

                    master_and_car = last_master_and_car
                    field_and_brigade = last_field_and_brigade
                    well = last_well
                    start_date = last_start_date
                    end_date = last_end_date

                    clean_well_str = clean_spaces(str(well or "").replace('\n', ' '))

                    # Стоп-слова (русские + казахские)
                    stop_words = [
                        # Русские
                        "ожидани", "списан", "резерв", "ремонт", "база", "демалыс", "гараж",
                        "не полный", "работает с подъем",
                        # Казахские статусные фразы
                        "кезекте", "сақтауда", "толық емес", "жиналыста",
                        "түнгі кезек", "оқуда болды",
                    ]

                    # Регулярка наличия временного диапазона (08:00-09:00 или 08.00-09.00)
                    _HAS_TIME = re.compile(r'\d{2}[:.]\d{2}\s*[-–]\s*\d{2}[:.]\d{2}')

                    def _is_trash(raw):
                        if not raw or raw.strip().lower() in ("none", "-", "нет", "."):
                            return True
                        # Нет ни одного временного отрезка → статусная заметка, не работа
                        if not _HAS_TIME.search(raw):
                            return True
                        # Содержит известное стоп-слово
                        return any(w in raw.lower() for w in stop_words)

                    is_s1_trash = _is_trash(s1_str)
                    is_s2_trash = _is_trash(s2_str)

                    if is_s1_trash and is_s2_trash:
                        continue

                    raw_end_date_str = str(end_date or "").strip()
                    second_well_name = ""

                    if raw_end_date_str:
                        match_well = re.search(r'(?i)скв[^0-9]*(\d{3,4})', raw_end_date_str)
                        if match_well:
                            w = match_well.group(1)
                            if not (len(w) == 4 and 2000 <= int(w) <= 2099):
                                second_well_name = w

                    raw_b_str = str(master_and_car or "").strip()
                    search_str = re.sub(r'(?i)kz\s*', ' ', raw_b_str)
                    
                    potential_plates = []
                    plate_pattern = r'\b([A-Za-zА-Яа-я]?\s*\d{3}\s*[A-Za-zА-Яа-я]{2,3}\s*(?:/|-)?\s*\d{0,3})\b'
                    
                    for match in re.finditer(plate_pattern, search_str):
                        raw_plate = match.group(1)
                        start_idx = match.start()
                        preceding_text = search_str[max(0, start_idx-35):start_idx].upper()
                        
                        is_priority = any(marker in preceding_text for marker in ["АПРС", "ПТП"])
                        
                        clean_p = re.sub(r'[\s/|-]', '', raw_plate)
                        potential_plates.append((clean_p, is_priority, raw_plate))
                        
                    plate_clean = ""
                    if potential_plates:
                        potential_plates.sort(key=lambda x: x[1], reverse=True)
                        plate_clean = potential_plates[0][0]
                    else:
                        for line in search_str.splitlines():
                            if '№' in line:
                                plate_clean = line.replace('№', '').replace(' ', '').strip()
                                potential_plates.append((plate_clean, False, line))
                                break
                    
                    if plate_clean:
                        plate_clean = normalize_plate(plate_clean)
                        if plate_clean in CAR_PLATE_MAP:
                            plate_clean = CAR_PLATE_MAP[plate_clean]
                    
                    b_lines = [str(x).strip() for x in raw_b_str.splitlines() if str(x).strip()]
                    first_line = b_lines[0] if b_lines else ""
                    clean_name_str = first_line
                    
                    for _, _, raw_p in potential_plates:
                        clean_name_str = clean_name_str.replace(raw_p, ' ')
                        
                    clean_name_str = re.sub(r'(?:\+7|8)[\s\-]?\(?[7]\d{2}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}', ' ', clean_name_str)
                    clean_name_str = re.sub(r'(?i)(АПРС|ПТП|Көтергіш|Подъемник|Сваб|УМК|АЦ|УПСТ)[\s-]*\d*', ' ', clean_name_str)
                    clean_name_str = re.sub(r'\b\d{4}\s*[гж]\.?(\s*в\.?)?', ' ', clean_name_str) 
                    clean_name_str = re.sub(r'(?i)kz\s*', ' ', clean_name_str)
                    clean_name_str = clean_name_str.replace('№', '')
                    clean_name_str = clean_spaces(clean_name_str)
                    
                    fio_part = ""
                    name_parts = clean_name_str.split()
                    
                    ignore_titles = ['мастер', 'ст.мастер', 'ст.', 'ст', 'смена']
                    filtered_parts = [p for p in name_parts if p.lower() not in ignore_titles]
                    
                    if len(filtered_parts) >= 2:
                        last_name = filtered_parts[0].capitalize()
                        first_initial = filtered_parts[1][0].upper()
                        fio_part = f"{last_name} {first_initial}."
                    elif len(filtered_parts) == 1:
                        fio_part = filtered_parts[0].capitalize()
                        
                    if fio_part and plate_clean:
                        car_field_value = f"{fio_part} {plate_clean}"
                    elif plate_clean:
                        car_field_value = plate_clean
                    elif fio_part:
                        car_field_value = fio_part
                    else:
                        car_field_value = "Не указана"

                    clean_c = str(field_and_brigade or "").strip()
                    cell_lines = [x.strip() for x in clean_c.splitlines() if x.strip()]

                    brigade_num_final = None
                    device_number = None
                    if cell_lines:
                        for line in cell_lines:
                            if "бригада" in line.lower():
                                # Ищем число ≤ 999 (реальный номер бригады)
                                for m in re.finditer(r'\d+', line):
                                    val = int(m.group(0))
                                    if val <= 999:
                                        brigade_num_final = val
                                        break

                        # Извлекаем номер агрегата из ячейки колонки C
                        # Номер агрегата — 4+ цифр (15679, 11983...), номер бригады — 1-3 цифры
                        for line in cell_lines:
                            if "бригада" in line.lower():
                                continue
                            m = re.search(r'№\s*(\d{4,})', line)
                            if m:
                                device_number = m.group(1)
                                break

                    well_clean = clean_well_str.upper()
                    pump_type = "Не указан"

                    tech_words = ["ШГН", "ЭБС", "ФОНТАН"]
                    for t in tech_words:
                        if t in well_clean:
                            pump_type = t

                    well_digit_match = re.search(r'\d+', well_clean)

                    if well_digit_match:
                        num_part = int(well_digit_match.group(0))
                        well_name_final = f"WL_{num_part:04d}"
                    else:
                        well_name_final = "WL_UNKNOWN"

                    if hasattr(start_date, 'strftime'):
                        date_str = start_date.strftime("%d.%m.%Y")
                    else:
                        date_str = str(start_date)

                    shift_1_list = parse_shift_text(shift_1_details)
                    shift_2_list = parse_shift_text(shift_2_details)

                    if shift_1_list:
                        record = {
                            "start_date": format_date(start_date),
                            "brigade_number": int(brigade_num_final) if brigade_num_final else 0,
                            "well_name": well_name_final,
                            "second_well_name": second_well_name,
                            "pump_type": pump_type,
                            "shift_type_number": 1,
                            "car": car_field_value,
                            "device_number": device_number if device_number else "",
                            "shift_details": shift_1_list
                        }
                        result_json.append(record)

                    if shift_2_list:
                        record = {
                            "start_date": format_date(start_date),
                            "brigade_number": int(brigade_num_final) if brigade_num_final else 0,
                            "well_name": well_name_final,
                            "second_well_name": second_well_name,
                            "pump_type": pump_type,
                            "shift_type_number": 2,
                            "car": car_field_value,
                            "device_number": device_number if device_number else "",
                            "shift_details": shift_2_list
                        }
                        result_json.append(record)

            # =========================
            # JSON САВИНГ (ВАЖНО)
            # =========================

            date_now = datetime.datetime.now().strftime("%Y-%m-%d")
            year_month = datetime.datetime.now().strftime("%Y-%m")

            output_dir = os.path.join(script_dir, "results", year_month)
            os.makedirs(output_dir, exist_ok=True)

            output_filename = f"{date_now}.json"
            output_path = os.path.join(output_dir, output_filename)

            def make_key(item):
                return (
                    item.get("start_date"),
                    item.get("well_name"),
                    item.get("shift_type_number"),
                    item.get("brigade_number"),
                    item.get("car")
                )

            existing_data = []

            if os.path.exists(output_path):
                try:
                    with open(output_path, "r", encoding="utf-8") as f:
                        existing_data = json.load(f)
                except:
                    existing_data = []

            existing_keys = set(str(x) for x in existing_data)

            for item in result_json:
                if str(item) not in existing_keys:
                    existing_data.append(item)
                    existing_keys.add(str(item))

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(existing_data, f, ensure_ascii=False, indent=2)

            print(f"[ПАРСЕР] Готово: {len(result_json)} записей")

            if result_json:
                print("[API] отправка данных...")

                success = send_to_api(result_json)

                if success:
                    print("[API] успешно отправлено")
                else:
                    print("[API] ошибка отправки")
            else:
                print("[API] нет данных для отправки")

            shutil.move(excel_path, os.path.join(processed_dir, target_file))
            print(f"[ПАРСЕР] Перемещён в processed")

        except Exception as e:
            print(f"[ПАРСЕР] Ошибка: {e}")

    time.sleep(15)
