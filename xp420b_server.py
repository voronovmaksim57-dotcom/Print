from flask import Flask, request, jsonify
from datetime import datetime
import sys
import subprocess
import os

# --- АВТООБНОВЛЕНИЕ ИЗ GITHUB ---
VERSION = "2025-11-28-1"

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

# Включаем/выключаем доп. элементы
SHOW_UNDERLINE = True   # показывать линию под числом
SHOW_DATETIME  = True   # печатать дату и время

# Насколько поднимать основное число, если печатаем дату/время
DATETIME_SHIFT = {
    "1": 20,
    "2": 8,
    "3": 0,
    "4": 0
}    # точек вверх; 0 = не двигать

# Параметры линии
LINE_THICKNESS = 5       # толщина линии (было 2)

# Параметры даты/времени
DATETIME_MUL_X = 7       # масштаб по X (1..3)
DATETIME_MUL_Y = 7       # масштаб по Y

# Координаты линии и даты (относительно низа этикетки)
LINE_Y_FROM_BOTTOM     = 35   # расстояние от нижнего края до линии
DATETIME_Y_FROM_BOTTOM = 20   # расстояние от нижнего края до текста даты
DATETIME_X             = 73   # смещение даты слева

# ⚠️ Впиши ТОЧНОЕ имя принтера из "Устройства и принтеры"
PRINTER_NAME = "Xprinter XP-420B"

app = Flask(__name__)

LABEL_WIDTH_MM = 30
LABEL_HEIGHT_MM = 20
DOTS_PER_MM = 8  # 203 dpi ~ 8 точек/мм


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

        # ===== СЛОТЫ РАЗМЕТКИ =====
        # slot1: 1 цифра слева, 1 цифра справа (например 1-1, 2-3)
        # slot2: (1 слева, 2 справа) ИЛИ (2 слева, 1 справа) — сюда пойдёт 6-10
        # slot3: 2 слева, 2+ справа — сюда пойдёт 44-10, 22-10 и т.п.
        # slot4: всё остальное (3+ слева, длинные коды и т.п.)
        if left_len == 1 and right_len == 1:
            slot = "1"
        elif (left_len == 1 and right_len >= 2) or (left_len == 2 and right_len == 1):
            slot = "2"
        elif (left_len == 2 and right_len >= 2) or (left_len == 3 and right_len == 1):
            slot = "3"
        else:
            slot = "4"

        # ===== БАЗОВЫЕ НАСТРОЙКИ (КАК ТЫ ЗАДАЛ РАНЕЕ) =====
        BASE_SLOT_CONFIG = {
            "1": {"mul": 45, "x": 34, "y": 40},
            "2": {"mul": 40, "x": 24, "y": 40},
            "3": {"mul": 34, "x": 14, "y": 43},
            "4": {"mul": 28, "x": 14, "y": 50},
        }

        cfg = BASE_SLOT_CONFIG.get(slot, BASE_SLOT_CONFIG["4"])
        mul = cfg["mul"]
        x = cfg["x"]
        y = cfg["y"]

        # Индивидуальное поднятие для каждого слота
        if SHOW_DATETIME:
            shift = DATETIME_SHIFT.get(slot, 8)
            y = max(0, y - shift)

        print(
            f"TSPL(label='{label}', left_len={left_len}, right_len={right_len}, "
            f"slot={slot}, mul={mul}, x={x}, y={y})",
            flush=True
        )

        lines = []
        lines.append(f"SIZE {LABEL_WIDTH_MM} mm,{LABEL_HEIGHT_MM} mm")
        lines.append("GAP 2 mm,0")
        lines.append("DIRECTION 1")
        lines.append("CLS")

        # 1) ЧИСЛО (основной код — сверху)
        lines.append(
            f'TEXT {x},{y},"{font}",0,{mul},{mul},"{label}"'
        )

        # 2) ЛИНИЯ НИЖЕ ЧИСЛА
        if SHOW_UNDERLINE:
            line_y = height_dots - LINE_Y_FROM_BOTTOM
            # Прямоугольник шириной на всю этикетку, высотой = LINE_THICKNESS
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
    # Здесь тоже уважаем старые координаты и такое же поведение смещения
    mul = 28
    x = 14
    y = 40

    print(f"TSPL(label='{label}' DEFAULT, mul={mul}, x={x}, y={y})", flush=True)

    lines = []
    lines.append(f"SIZE {LABEL_WIDTH_MM} mm,{LABEL_HEIGHT_MM} mm")
    lines.append("GAP 2 mm,0")
    lines.append("DIRECTION 1")
    lines.append("CLS")

    # 1) текст (или любой произвольный)
    lines.append(
        f'TEXT {x},{y},"{"0"}",0,{mul},{mul},"{label}"'
    )

    # 2) линия
    if SHOW_UNDERLINE:
        line_y = height_dots - LINE_Y_FROM_BOTTOM
        # Прямоугольник шириной на всю этикетку, высотой = LINE_THICKNESS
        lines.append(
            f"BAR 0,{line_y},{width_dots},{LINE_THICKNESS}"
        )

    # 3) дата/время
    if SHOW_DATETIME:
        now_str = datetime.now().strftime("%d.%m %H:%M")
        date_y = height_dots - DATETIME_Y_FROM_BOTTOM
        lines.append(
            f'TEXT {DATETIME_X},{date_y},"{"0"}",0,{DATETIME_MUL_X},{DATETIME_MUL_Y},"{now_str}"'
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

