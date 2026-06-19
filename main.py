import requests
import time
import json
import os
import re
import openpyxl
import hashlib
import datetime

# ==========================================
# 1. НАСТРОЙКИ ПОЧТЫ И ПУТЕЙ
# ==========================================
# Берутся из файла .env через окружение Docker
EMAIL = os.environ.get("PARSER_EMAIL", "dlc_prs@web-library.net")
PASSWORD = os.environ.get("PARSER_PASSWORD")

if not PASSWORD:
    print("✗ КРИТИЧЕСКАЯ ОШИБКА: В файле .env не задан PARSER_PASSWORD!")
    print("Работа скрипта остановлена.")
    exit()

API_URL = "https://api.mail.tm"
script_dir = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(script_dir, "processed_history.json")

# ==========================================
# 2. СЛОВАРИ ДЛЯ ПАРСЕРА
# ==========================================
CAR_PLATE_MAP = {
    "024BH": "024BH06",
    "250AKD": "250AMD",
    "564AZ06": "564AY06",
    "E247AHD": "AHD247E",
    "153AAS": "153AS06",
    "259ALD06": "ALD259E",
    "480AZD": "AZD480E",
    "730AU": "730AU06"
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

def get_file_hash(file_path):
    hasher = hashlib.md5()
    with open(file_path, 'rb') as f:
        buf = f.read()
        hasher.update(buf)
    return hasher.hexdigest()

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

def get_fresh_token(email, password):
    print("Попытка авторизации на сервере почты...")
    payload = {"address": email, "password": password}
    try:
        response = requests.post(f"{API_URL}/token", json=payload, timeout=15)
        if response.status_code == 200:
            print("✓ Авторизация успешна! Токен получен.")
            return response.json().get('token')
        else:
            print(f"✗ Ошибка авторизации: {response.status_code}. Проверьте пароль!")
            return None
    except Exception as e:
        print(f"✗ Ошибка при получении токена: {e}")
        return None

# ==========================================
# 4. ЗАПУСК И ГЛАВНЫЙ ЦИКЛ
# ==========================================
token = get_fresh_token(EMAIL, PASSWORD)
if not token:
    print("Внимание: Не удалось получить токен сразу. Будем пытаться в цикле...")

headers = {"Authorization": f"Bearer {token}"} if token else {}

print(f"\nРобот успешно запущен!")
print(f"Папка для работы: {script_dir}")
print("Перехожу в режим ожидания писем...")

while True:
    # ---------------------------------------------------------
    # ШАГ 1: ПРОВЕРКА ПОЧТЫ И СКАЧИВАНИЕ НОВЫХ ФАЙЛОВ
    # ---------------------------------------------------------
    if not headers.get("Authorization"):
        token = get_fresh_token(EMAIL, PASSWORD)
        if token:
            headers = {"Authorization": f"Bearer {token}"}
            
    if headers.get("Authorization"):
        try:
            response = requests.get(f"{API_URL}/messages", headers=headers, timeout=15)
            
            if response.status_code == 200:
                try:
                    data = response.json()
                except ValueError:
                    time.sleep(15)
                    continue
                    
                messages = data.get('hydra:member', [])
                
                for msg in messages:
                    msg_id = msg['id']
                    full_msg_response = requests.get(f"{API_URL}/messages/{msg_id}", headers=headers, timeout=15)
                    
                    if full_msg_response.status_code == 200:
                        full_msg = full_msg_response.json()
                        attachments = full_msg.get('attachments', [])
                        
                        should_delete = True  
                        
                        for att in attachments:
                            filename = att['filename']
                            if filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls'):
                                print(f"\n[ПОЧТА] Найдена таблица: {filename}. Начинаю скачивание...")
                                
                                download_url = f"{API_URL}{att['downloadUrl']}"
                                try:
                                    file_response = requests.get(download_url, headers=headers, timeout=30)
                                    if file_response.status_code == 200:
                                        file_path = os.path.join(script_dir, filename)
                                        with open(file_path, "wb") as f:
                                            f.write(file_response.content)
                                        print(f"[ПОЧТА] Успех! Файл сохранен: {filename}")
                                    else:
                                        print(f"[ПОЧТА] Ошибка скачивания: {file_response.status_code}. Письмо не будет удалено.")
                                        should_delete = False
                                except requests.exceptions.RequestException as e:
                                    print(f"[ПОЧТА] Срыв при скачивании файла: {e}. Письмо не будет удалено.")
                                    should_delete = False
                                
                        if should_delete:
                            try:
                                delete_response = requests.delete(f"{API_URL}/messages/{msg_id}", headers=headers, timeout=15)
                                if delete_response.status_code == 204:
                                    print("[ПОЧТА] Письмо успешно удалено из ящика.")
                                else:
                                    print(f"[ПОЧТА] Ошибка при удалении письма: {delete_response.status_code}")
                            except requests.exceptions.RequestException as e:
                                print(f"[ПОЧТА] Ошибка при удалении письма (таймаут): {e}")
                        else:
                            print(f"[ПОЧТА] Письмо {msg_id} оставлено в ящике для повторной попытки.")
                            
            elif response.status_code == 401:
                print("\n[ПОЧТА] Токен истек! Пытаюсь получить новый...")
                token = get_fresh_token(EMAIL, PASSWORD)
                if token:
                    headers = {"Authorization": f"Bearer {token}"}
                else:
                    headers = {}
                    
        except requests.exceptions.RequestException as e:
            print(f"\n[СЕТЬ] Ошибка подключения к почтовому серверу: {type(e).__name__} (возможно, сервер mail.tm недоступен).")

    # ---------------------------------------------------------
    # ШАГ 2: ПАРСИНГ ВСЕХ EXCEL ФАЙЛОВ В ПАПКЕ
    # ---------------------------------------------------------
    processed_history = {}
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            try:
                processed_history = json.load(f)
            except json.JSONDecodeError:
                processed_history = {}

    excel_files = [f for f in os.listdir(script_dir) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]

    for target_file in excel_files:
        excel_path = os.path.join(script_dir, target_file)
        current_hash = get_file_hash(excel_path)
        
        if target_file in processed_history and processed_history[target_file] == current_hash:
            continue 
            
        print(f"\n[ПАРСЕР] Начинаю обработку нового файла: '{target_file}'")
        
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            first_sheet_title = wb.sheetnames[0]
            result_json = []
            
            for sheet in wb.worksheets:
                print(f"  -> Обрабатываем лист: {sheet.title}")

                # --- УМНАЯ ЗАЩИТА ОТ ОПЕЧАТОК ДИСПЕТЧЕРОВ ---
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
                # ---------------------------------------------------------------

                last_master_and_car = None
                last_field_and_brigade = None
                last_well = None
                last_start_date = None

                for row in range(4, sheet.max_row + 1):
                    raw_b = sheet[f"B{row}"].value  
                    raw_c = sheet[f"C{row}"].value
                    raw_d = sheet[f"D{row}"].value
                    raw_n = sheet[f"N{row}"].value
                    
                    if raw_b is not None and str(raw_b).strip(): last_master_and_car = raw_b
                    if raw_c is not None and str(raw_c).strip(): last_field_and_brigade = raw_c
                    if raw_d is not None and str(raw_d).strip(): last_well = raw_d
                    if raw_n is not None and str(raw_n).strip(): last_start_date = raw_n
                    
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

                    clean_well_str = clean_spaces(str(well or "").replace('\n', ' '))
                    stop_words = ["ожидани", "списан", "резерв", "ремонт", "база", "демалыс", "гараж"]
                    
                    is_s1_trash = any(word in s1_str.lower() for word in stop_words) if s1_str else True
                    is_s2_trash = any(word in s2_str.lower() for word in stop_words) if s2_str else True
                    
                    if clean_well_str in ("", "None") and is_s1_trash and is_s2_trash:
                        continue

                    # ==========================================
                    # Парсинг столбца B (Мастер и машина) 
                    # ==========================================
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

                    # ==========================================
                    # Парсинг столбца C 
                    # ==========================================
                    clean_c = str(field_and_brigade or "").strip()
                    cell_lines = [str(x).strip() for x in clean_c.splitlines() if str(x).strip()]

                    brigade_name = "Не указана"
                    field_code = "UNKNOWN"
                    device_number = None 

                    if cell_lines:
                        found_brigade = False
                        for line in cell_lines:
                            if any(x in line.lower() for x in ["бригада", "бр.", "бр№", "бр-", "отряд"]):
                                brigade_name = clean_spaces(line)
                                found_brigade = True
                                break
                        if not found_brigade:
                            brigade_name = clean_spaces(cell_lines[-1]) if cell_lines else "Не указана"

                        for i, line in enumerate(cell_lines):
                            if "дэл" in line.lower():
                                match_num = re.search(r'№\s*(\d+)', line)
                                if match_num:
                                    device_number = match_num.group(1)
                                elif i + 1 < len(cell_lines) and "№" in cell_lines[i+1] and "бригада" not in cell_lines[i+1].lower():
                                    match_next = re.search(r'№\s*(\d+)', cell_lines[i+1])
                                    if match_next:
                                        device_number = match_next.group(1)
                                break

                        raw_field = cell_lines[0].strip()
                        if "ДЭЛ" in raw_field:
                            raw_field = raw_field.split("ДЭЛ")[0].strip()
                        
                        if not raw_field or any(x in raw_field.lower() for x in ["бригада", "бр.", "бр№", "бр-"]):
                            field_code = "UNKNOWN"
                        else:
                            field_code, _ = get_field_info(raw_field)

                    brigade_num_final = None
                    brigade_match = re.search(r'\d+', brigade_name)
                    if brigade_match:
                        brigade_num_final = int(brigade_match.group(0))

                    well_clean = clean_well_str.upper()
                    pump_type = "Не указан"
                    tech_words = ["ШГН", "ЭБС", "ФОНТАН", "Б/Д", "ПОГЛ", "ТҚ", "ТК", "ЛШПН", "ЛШ"]
                    
                    found_pumps = []
                    for tech_word in tech_words:
                        if tech_word in well_clean:
                            found_pumps.append(tech_word)
                            well_clean = well_clean.replace(tech_word, "")
                            
                    if found_pumps:
                        pump_type = ", ".join(found_pumps)
                        
                    well_clean = well_clean.strip()

                    well_digit_match = re.search(r'\d+', well_clean)
                    if well_digit_match:
                        num_part = int(well_digit_match.group(0))
                        suffix_match = re.search(r'\d+\s*([A-ZА-ЯЁ]{1,2})', well_clean)
                        suffix = suffix_match.group(1) if suffix_match else ""
                        well_name_final = f"{field_code}_{num_part:04d}{suffix}"
                    else:
                        well_name_final = f"{field_code}_{well_clean.replace(' ', '_')}" if well_clean else f"{field_code}_UNKNOWN"

                    # ==========================================
                    # Парсинг даты 
                    # ==========================================
                    if hasattr(start_date, 'strftime'):
                        orig_date_str = start_date.strftime("%d.%m.%Y")
                    else:
                        raw_date_str = str(start_date or "")
                        date_match = re.search(r'(\d{2})[./-](\d{2})[./-](\d{2,4})', raw_date_str)
                        if date_match:
                            d_day = date_match.group(1)
                            d_mon = date_match.group(2)
                            d_year = date_match.group(3)
                            if len(d_year) == 2:
                                d_year = "20" + d_year
                            orig_date_str = f"{d_day}.{d_mon}.{d_year}"
                        else:
                            orig_date_str = clean_spaces(raw_date_str).replace('\n', ' ')

                    date_str = orig_date_str
                    if target_month and target_year and re.match(r'\d{2}\.\d{2}\.\d{4}', orig_date_str):
                        d_day = orig_date_str.split('.')[0]
                        date_str = f"{d_day}.{target_month}.{target_year}"

                    shift_1_list = parse_shift_text(shift_1_details)
                    shift_2_list = parse_shift_text(shift_2_details)

                    if shift_1_list and not is_s1_trash:
                        result_json.append({
                            "start_date": date_str,
                            "brigade_number": brigade_num_final,
                            "well_name": well_name_final,
                            "pump_type": pump_type,
                            "shift_type_number": 1,
                            "car": car_field_value,       
                            "device_number": device_number, 
                            "shift_details": shift_1_list
                        })

                    if shift_2_list and not is_s2_trash:
                        result_json.append({
                            "start_date": date_str,
                            "brigade_number": brigade_num_final,
                            "well_name": well_name_final,
                            "pump_type": pump_type,
                            "shift_type_number": 2,
                            "car": car_field_value,       
                            "device_number": device_number, 
                            "shift_details": shift_2_list
                        })

            clean_base_name = generate_clean_filename(target_file, first_sheet_title)
            output_filename = f"{clean_base_name}.json"
            output_path = os.path.join(script_dir, output_filename)

            counter = 1
            while os.path.exists(output_path):
                output_filename = f"{clean_base_name}_{counter}.json"
                output_path = os.path.join(script_dir, output_filename)

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(result_json, f, ensure_ascii=False, indent=2)

            processed_history[target_file] = current_hash
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(processed_history, f, ensure_ascii=False, indent=2)

            print(f"[ПАРСЕР] Готово! Создан файл: {output_filename} (Смен: {len(result_json)})")
            
        except Exception as e:
            print(f"[ПАРСЕР] Ошибка при обработке файла {target_file}: {e}")

    time.sleep(15)