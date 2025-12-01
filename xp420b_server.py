from flask import Flask, request, jsonify
from datetime import datetime
import sys
import subprocess
import os
import json
from pathlib import Path
import copy

CONFIG_FILE = Path(__file__).with_name("printer_config.json")

DEFAULT_CONFIG = {
    "printer_name": "Xprinter XP-420B",

    "label": {
        "width_mm": 30,
        "height_mm": 20,
        "dots_per_mm": 8,
        "gap_mm": 2,
    },

    "features": {
        "show_underline": True,
        "show_datetime": True,
    },

    "line": {
        "thickness": 5,
        "y_from_bottom": 35,
    },

    "datetime": {
        "mul_x": 7,
        "mul_y": 7,
        "y_from_bottom": 20,
        "x": 73,
        "shift_by_slot": {
            "1": 20,
            "2": 8,
            "3": 0,
            "4": 0,
        },
    },

    "slots": {
        "1": {"mul": 45, "x": 34, "y": 40},
        "2": {"mul": 40, "x": 24, "y": 40},
          "3": {"mul": 34, "x": 14, "y": 43},
        "4": {"mul": 28, "x": 14, "y": 50},
    },

    "fallback": {
        "mul": 28,
        "x": 14,
        "y": 40,
    },
}


def _deep_update(base: dict, updates: dict):
    """Рекурсивное объединение словарей (для частичных конфигов)."""
    for k, v in updates.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_update(base[k], v)
        else:
            base[k] = v


def load_config() -> dict:
    cfg = copy.deepcopy(DEFAULT_CONFIG)

    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                _deep_update(cfg, data)
            else:
                print("[CONFIG] printer_config.json имеет неверный формат, использую дефолтные")
        except Exception as e:
            print("[CONFIG] Ошибка чтения printer_config.json:", e)
    else:
        print("[CONFIG] Файл настроек не найден, создаю с настройками по умолчанию")
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("[CONFIG] Не удалось создать printer_config.json:", e)

    return cfg


CONFIG = load_config()


# --- АВТООБНОВЛЕНИЕ ИЗ GITHUB ---
VERSION = "2025-12-01-1"

# ссылка на raw-версию файла xp420b_server.py в GitHub
# ⚠️ Обязательно замени USERNAME, REPO и ветку (main/master) на свои.
AUTOUPDATE_URL = (
    "https://raw.githubusercontent.com/voronovmaksim57-dotcom/Print/refs/heads/main/xp420b_server.py"
)

AUTOUPDATE_ENABLED = True   # можно выключить, если что
AUTOUPDATE_BACKUP_SUFFIX = ".bak"


REQUIRED_MODULES = [
    ("flask", "flask"),
    ("win32print", "pywin32"),
]


def ensure_dependencies():
    for module, package in REQUIRED_MODULES:
        try:
            __import__(module)
        except ImportError:
            print(f"[AUTO-INSTALL] Пакет '{module}' не найден. Устанавливаем '{package}'...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print(f"[OK] Установлено: {package}")
            except Exception as e:
                print(f"[ERROR] Не удалось установить {package}: {e}")
                print("Продолжаем выполнение, но сервер может не работать!")


ensure_dependencies()


def extract_version(code: str):
    """
    Ищет строку вида VERSION = "...." и возвращает её значение.
    Если VERSION не найден — возвращает None.
    """
    import re as _re
    m = _re.search(r'^VERSION\s*=\s*["\'](.+?)["\']', code, _re.MULTILINE)
    return m.group(1) if m else None


def check_and_update_from_github():
    """
    Проверяет raw-файл на GitHub.
    Если VERSION в GitHub отличается от локальной —
    скачивает файл, делает .bak и перезапускает скрипт.
    """
    if not AUTOUPDATE_ENABLED:
        return

    import urllib.request

    script_path = os.path.abspath(__file__)

    try:
        print("[AUTOUPDATE] Проверяю обновление с GitHub...", flush=True)
        with urllib.request.urlopen(AUTOUPDATE_URL, timeout=5) as resp:
            if resp.status != 200:
                print(f"[AUTOUPDATE] GitHub ответил статусом {resp.status}, пропускаю.", flush=True)
                return
            new_code = resp.read().decode("utf-8", errors="replace")
    except Exception as e:
        print(f"[AUTOUPDATE] Не удалось скачать файл с GitHub: {e}", flush=True)
        return

    # Читаем текущий файл
    try:
        with open(script_path, "r", encoding="utf-8") as f:
            current_code = f.read()
    except Exception as e:
        print(f"[AUTOUPDATE] Не удалось прочитать текущий файл: {e}", flush=True)
        return

    # Достаём версии
    remote_version = extract_version(new_code)
    local_version = extract_version(current_code)

    print(f"[AUTOUPDATE] Local VERSION={local_version}, Remote VERSION={remote_version}", flush=True)

    # Если в удалённом файле нет VERSION — лучше ничего не трогать
    if remote_version is None:
        print("[AUTOUPDATE] В удалённом файле нет VERSION, пропускаю.", flush=True)
        return

    # Если версии совпадают — обновление не нужно
    if remote_version == local_version:
        print("[AUTOUPDATE] Уже актуальная версия, обновление не требуется.", flush=True)
        return

    # sanity-check, что это действительно наш файл
    if "def build_tspl" not in new_code or "app.run" not in new_code:
        print("[AUTOUPDATE] Загруженный файл не похож на xp420b_server.py, пропускаю.", flush=True)
        return

    # Делаем бэкап
    backup_path = script_path + AUTOUPDATE_BACKUP_SUFFIX
    try:
        with open(backup_path, "w", encoding="utf-8") as f:
            f.write(current_code)
        print(f"[AUTOUPDATE] Резервная копия сохранена: {backup_path}", flush=True)
    except Exception as e:
        print(f"[AUTOUPDATE] Не удалось сохранить бэкап: {e}", flush=True)
        return

    # Перезаписываем текущий файл новым кодом
    try:
        with open(script_path, "w", encoding="utf-8") as f:
            f.write(new_code)
        print("[AUTOUPDATE] Файл обновлён (VERSION изменился), запускаю новую версию и выхожу.", flush=True)
    except Exception as e:
        print(f"[AUTOUPDATE] Ошибка записи нового файла: {e}", flush=True)
        return

    # Стартуем новый процесс и выходим из текущего
    try:
        subprocess.Popen([sys.executable, script_path])
    except Exception as e:
        print(f"[AUTOUPDATE] Не удалось перезапустить скрипт: {e}", flush=True)
    finally:
        sys.exit(0)


import win32print
import re

# Все настройки берём из CONFIG
PRINTER_NAME = CONFIG["printer_name"]

LABEL_WIDTH_MM = CONFIG["label"]["width_mm"]
LABEL_HEIGHT_MM = CONFIG["label"]["height_mm"]
DOTS_PER_MM = CONFIG["label"]["dots_per_mm"]
GAP_MM = CONFIG["label"]["gap_mm"]

SHOW_UNDERLINE   = CONFIG["features"]["show_underline"]
SHOW_DATETIME    = CONFIG["features"]["show_datetime"]
PRINT_ONLY_LEFT  = CONFIG["features"].get("print_only_left", False)

LINE_THICKNESS     = CONFIG["line"]["thickness"]
LINE_Y_FROM_BOTTOM = CONFIG["line"]["y_from_bottom"]

DATETIME_MUL_X        = CONFIG["datetime"]["mul_x"]
DATETIME_MUL_Y        = CONFIG["datetime"]["mul_y"]
DATETIME_Y_FROM_BOTTOM = CONFIG["datetime"]["y_from_bottom"]
DATETIME_X             = CONFIG["datetime"]["x"]
DATETIME_SHIFT         = CONFIG["datetime"]["shift_by_slot"]
DATETIME_SINGLE_SHIFT = CONFIG["datetime"].get("single_shift", {})

SLOT_CONFIG        = CONFIG["slots"]
SINGLE_SLOT_CONFIG = CONFIG.get("single_slots", {})
FALLBACK_CONFIG    = CONFIG["fallback"]


app = Flask(__name__)

@app.after_request
def add_cors_headers(response):
    # Разрешаем запросы с любых страниц + private network
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
    response.headers["Access-Control-Allow-Private-Network"] = "true"
    return response


@app.route("/print", methods=["OPTIONS"])
def print_options():
    # Ответ на preflight-запрос
    return "", 200


TASK_STARTUP_NAME = "xp420b_server_startup.bat"


def install_startup():
    """
    Создаёт файл автозапуска в папке Startup текущего пользователя,
    который тихо запускает xp420b_server.py через pythonw.exe.
    """
    # Путь к папке Startup текущего пользователя
    startup_dir = os.path.join(
        os.environ.get("APPDATA", ""),
        r"Microsoft\Windows\Start Menu\Programs\Startup"
    )

    if not os.path.isdir(startup_dir):
        print("[XP420B] Не удалось найти папку Startup:", startup_dir)
        return

    # Путь к самому скрипту
    script_path = os.path.abspath(__file__)

    # ТВОЙ путь к pythonw.exe (как ты написал)
    pythonw_path = r"C:\Users\USER\AppData\Local\Programs\Python\Python314\pythonw.exe"

    # Файл, который положим в Startup
    bat_path = os.path.join(startup_dir, TASK_STARTUP_NAME)

    # Содержимое батника: одна строка запуска
    cmd_line = f'"{pythonw_path}" "{script_path}"\r\n'

    try:
        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(cmd_line)
        print("[XP420B] Файл автозапуска создан:")
        print("  ", bat_path)
        print("[XP420B] При следующем входе в Windows сервер запустится автоматически.")
    except Exception as e:
        print("[XP420B] Ошибка при создании файла автозапуска:", e)


def build_tspl(label: str) -> str:
    font = "0"

    width_dots = LABEL_WIDTH_MM * DOTS_PER_MM
    height_dots = LABEL_HEIGHT_MM * DOTS_PER_MM

    # Пытаемся разобрать формат "числа-числа"
    m = re.match(r"^(\d+)-(\d+)$", label)
    if m:
        left, right = m.groups()
        left_len = len(left)
        right_len = len(right)

        # ===== СЛОТЫ РАЗМЕТКИ (как раньше) =====
        if left_len == 1 and right_len == 1:
            slot = "1"
        elif (left_len == 1 and right_len >= 2) or (left_len == 2 and right_len == 1):
            slot = "2"
        elif (left_len == 2 and right_len >= 2) or (left_len == 3 and right_len == 1):
            slot = "3"
        else:
            slot = "4"

        # ===== Выбираем, что печатать и какие размеры =====
        if PRINT_ONLY_LEFT:
            # Печатаем только левую часть (до тире)
            text_to_print = left
            # Берём отдельный слот по длине левой части (1 / 2 / 3 цифры)
            single_key = str(left_len)
            cfg = SINGLE_SLOT_CONFIG.get(
                single_key,
                SLOT_CONFIG.get(slot, SLOT_CONFIG["4"])  # запасной вариант — старый слот
            )
        else:
            # Старое поведение — печатаем "1-1", "12-34" целиком
            text_to_print = label
            cfg = SLOT_CONFIG.get(slot, SLOT_CONFIG["4"])

        mul = cfg["mul"]
        x = cfg["x"]
        y = cfg["y"]

        # Индивидуальное поднятие для каждого слота, если включена дата
        # === Коррекция y (смещение вверх под дату) ===
        if SHOW_DATETIME:
            if PRINT_ONLY_LEFT:
                # отдельный shift для режима "печатаем только слева"
                shift = DATETIME_SINGLE_SHIFT.get(str(left_len), 0)
            else:
                # стандартный shift по slot
                shift = DATETIME_SHIFT.get(slot, 0)

            y = max(0, y - shift)

        print(
            f"TSPL(label='{label}', left_len={left_len}, right_len={right_len}, "
            f"slot={slot}, mul={mul}, x={x}, y={y}, print='{text_to_print}')",
            flush=True,
        )

        lines = []
        lines.append(f"SIZE {LABEL_WIDTH_MM} mm,{LABEL_HEIGHT_MM} mm")
        lines.append(f"GAP {GAP_MM} mm,0")
        lines.append("DIRECTION 1")
        lines.append("CLS")

        # 1) ЧИСЛО (или "1-1", или только левая часть)
        lines.append(
            f'TEXT {x},{y},"{font}",0,{mul},{mul},"{text_to_print}"'
        )

        # 2) ЛИНИЯ НИЖЕ ЧИСЛА
        if SHOW_UNDERLINE:
            line_y = height_dots - LINE_Y_FROM_BOTTOM
            lines.append(
                f"BAR 0,{line_y},{width_dots},{LINE_THICKNESS}"
            )

        # 3) ДАТА + ВРЕМЯ ЕЩЁ НИЖЕ
        if SHOW_DATETIME:
            now_str = datetime.now().strftime("%d.%m %H:%M")
            date_y = height_dots - DATETIME_Y_FROM_BOTTOM
            lines.append(
                f'TEXT {DATETIME_X},{date_y},"{font}",0,{DATETIME_MUL_X},{DATETIME_MUL_Y},"{now_str}"'
            )

        lines.append("PRINT 1,1")
        return "\r\n".join(lines) + "\r\n"

    # ===== Если формат не "числа-числа" — запасной вариант =====
    mul = FALLBACK_CONFIG["mul"]
    x = FALLBACK_CONFIG["x"]
    y = FALLBACK_CONFIG["y"]

    print(f"TSPL(label='{label}' DEFAULT, mul={mul}, x={x}, y={y})", flush=True)

    lines = []
    lines.append(f"SIZE {LABEL_WIDTH_MM} mm,{LABEL_HEIGHT_MM} mm")
    lines.append(f"GAP {GAP_MM} mm,0")
    lines.append("DIRECTION 1")
    lines.append("CLS")

    # 1) текст как есть
    lines.append(
        f'TEXT {x},{y},"{font}",0,{mul},{mul},"{label}"'
    )

    # 2) линия
    if SHOW_UNDERLINE:
        line_y = height_dots - LINE_Y_FROM_BOTTOM
        lines.append(
            f"BAR 0,{line_y},{width_dots},{LINE_THICKNESS}"
        )

    # 3) дата/время
    if SHOW_DATETIME:
        now_str = datetime.now().strftime("%d.%m %H:%M")
        date_y = height_dots - DATETIME_Y_FROM_BOTTOM
        lines.append(
            f'TEXT {DATETIME_X},{date_y},"{font}",0,{DATETIME_MUL_X},{DATETIME_MUL_Y},"{now_str}"'
        )

    lines.append("PRINT 1,1")
    return "\r\n".join(lines) + "\r\n"

def send_raw_to_printer(tspl_cmd: str):
    h = win32print.OpenPrinter(PRINTER_NAME)
    try:
        win32print.StartDocPrinter(h, 1, ("XP-420B Label", None, "RAW"))
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, tspl_cmd.encode("ascii"))
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)
        

def normalize(code: str):
    # Убираем различия CRLF/LF и лишние пробелы
    return code.replace("\r\n", "\n").replace("\r", "\n").strip()


@app.post("/print")
def print_label():
    data = request.get_json(force=True)
    label = (data.get("label") or "").strip()

    print(f"REQUEST /print label={label}", flush=True)

    if not re.match(r"^\d+-\d+$", label):
        print("BAD FORMAT", flush=True)
        return jsonify({"error": "bad label format"}), 400

    tspl = build_tspl(label)
    try:
        send_raw_to_printer(tspl)
        print("PRINT OK", flush=True)
    except Exception as e:
        print("PRINT ERROR:", e, flush=True)
        return jsonify({"error": str(e)}), 500

    return jsonify({"status": "ok", "printed": label})


TASK_NAME = "XP420B_LabelServer"

STARTUP_VBS_NAME = "xp420b_server_startup.vbs"  # имя файла в автозагрузке


def install_startup_vbs():
    """
    Создаёт .vbs в папке Startup текущего пользователя,
    который тихо запускает xp420b_server.py через pythonw.exe.
    """
    # Папка автозагрузки текущего пользователя
    startup_dir = os.path.join(
        os.environ.get("APPDATA", ""),
        r"Microsoft\Windows\Start Menu\Programs\Startup"
    )

    if not os.path.isdir(startup_dir):
        print("[XP420B] Не удалось найти папку Startup:", startup_dir)
        return

    # Полный путь к текущему скрипту
    script_path = os.path.abspath(__file__)

    # Путь к pythonw.exe (как ты писал)
    pythonw_path = r"C:\Users\USER\AppData\Local\Programs\Python\Python314\pythonw.exe"

    # Полный путь к .vbs в автозагрузке
    vbs_path = os.path.join(startup_dir, STARTUP_VBS_NAME)

    # Содержимое .vbs — запуск pythonw без окна
    vbs_content = f'''Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """{pythonw_path}"" ""{script_path}""", 0, False
'''

    try:
        with open(vbs_path, "w", encoding="utf-8") as f:
            f.write(vbs_content)
        print("[XP420B] Файл автозапуска (.vbs) создан:")
        print("  ", vbs_path)
        print("[XP420B] При следующем входе в Windows сервер запустится автоматически и БЕЗ окна.")
    except Exception as e:
        print("[XP420B] Ошибка при создании .vbs в автозагрузке:", e)


if __name__ == "__main__":
    # 1) спец-режим: установка автозапуска через .vbs
    if "--install-startup" in sys.argv:
        install_startup_vbs()
        sys.exit(0)

    # 2) обычный режим: сначала пытаемся обновиться с GitHub
    check_and_update_from_github()

    # 3) если обновления не было или не удалось -- просто запускаем сервер
    print("[XP420B] Запуск Flask-сервера на 127.0.0.1:9123")
    app.run(host="127.0.0.1", port=9123)
