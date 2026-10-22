import os

import re

import json

import math

import time

import hashlib

import threading

import traceback

import subprocess

import logging

import tkinter as tk

from tkinter import ttk, filedialog, messagebox

import numpy as np

import win32com.client

import pythoncom

from datetime import datetime

import html

import queue

import webbrowser

from decimal import Decimal, InvalidOperation

from urllib.parse import urlencode

import win32com.client.dynamic

import win32con

import win32clipboard

import win32gui

from tkinter import messagebox
import xml.etree.ElementTree as ET

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CACHE_PATH = os.path.join(SCRIPT_DIR, "mch_cache.json")

SETTINGS_PATH = os.path.join(SCRIPT_DIR, "settings.json")

CHANGED_LIST_PATH = os.path.join(SCRIPT_DIR, "changed_files_list.json")

LOG_PATH = os.path.join(SCRIPT_DIR, "monitorcad_debug.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    filename=LOG_PATH,
    filemode="a",
    force=True,
)

_original_print = print

def print(*args, **kwargs):
    message = " ".join(str(arg) for arg in args)
    logging.info(message)
    return _original_print(*args, **kwargs)

MODEL_EXTENSIONS = {'.m3d', '.m3d_'}

DRAWING_EXTENSIONS = {'.cdw'}

# Интервал опроса app.Documents (секунды). Нужен как подстраховка к
# событию OnOpenDocument: если КОМПАС не прислал событие открытия,
# документ всё равно попадёт в обработку при следующем опросе.
DOCUMENT_POLL_INTERVAL = 3.0

# Пауза перед повторной попыткой регистрации документа, если первая
# попытка завершилась ошибкой (например, не удалось посчитать МЦХ).
DOCUMENT_POLL_RETRY_INTERVAL = 30.0

DOCUMENT_CALCULATION_MAX_RETRIES = 3

# Имена внутренних документов КОМПАС, которые нельзя обрабатывать
# как пользовательские модели (например, служебный calculate_mch).
INTERNAL_DOCUMENT_NAMES = {"calculate_mch"}

GEOMETRY_FIELDS = [
    "Масса (г)", "Объем (мм³)", "Площадь (мм²)",
    "Центр масс Xc (мм)", "Центр масс Yc (мм)", "Центр масс Zc (мм)",
    "Jx центр. (г·мм²)", "Jy центр. (г·мм²)", "Jz центр. (г·мм²)",
    "Jxy центробежный (г·мм²)", "Jxz центробежный (г·мм²)", "Jyz центробежный (г·мм²)"
]

def is_path_valid(file_path):
    try:
        os.stat(file_path)
        return True
    except Exception:
        return False

def load_cache():
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_cache(cache):
    try:
        os.makedirs(SCRIPT_DIR, exist_ok=True)
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"   [ПРЕДУПРЕЖДЕНИЕ] Не удалось сохранить кэш: {e}")

def load_changed_list():
    """
    Загружает персистентный список изменённых файлов ("Изменённые файлы" в GUI).
    Хранится отдельно от mch_cache.json и очищается ТОЛЬКО по кнопке
    "Очистить список" - переживает закрытие программы, сбой питания и т.п.
    """
    if os.path.exists(CHANGED_LIST_PATH):
        try:
            with open(CHANGED_LIST_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("records", []), data.get("total_changes", 0)
        except Exception:
            return [], 0
    return [], 0

def save_changed_list(records, total_changes):
    try:
        os.makedirs(SCRIPT_DIR, exist_ok=True)
        with open(CHANGED_LIST_PATH, "w", encoding="utf-8") as f:
            json.dump({"records": records, "total_changes": total_changes},
                      f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"   [ПРЕДУПРЕЖДЕНИЕ] Не удалось сохранить список изменений: {e}")

def load_settings():
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_settings(settings):
    try:
        os.makedirs(SCRIPT_DIR, exist_ok=True)
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"   [ПРЕДУПРЕЖДЕНИЕ] Не удалось сохранить настройки: {e}")

def get_file_hash(file_path):
    try:
        with open(file_path, 'rb') as f:
            return hashlib.md5(f.read()).hexdigest()
    except Exception:
        return None

def file_markers_changed(file_path, cache):
    try:
        stat = os.stat(file_path)
        current_mtime = stat.st_mtime
        current_size = stat.st_size
        current_hash = get_file_hash(file_path)
    except Exception:
        return False, None, 0, 0, None

    key = os.path.normcase(os.path.abspath(file_path))
    prev = cache.get(key)

    if prev is None:
        return True, key, current_mtime, current_size, current_hash

    changed = (
        prev.get("mtime") != current_mtime or
        prev.get("size") != current_size or
        prev.get("hash") != current_hash
    )

    return changed, key, current_mtime, current_size, current_hash

def geometry_actually_changed(prev_mch, new_mch, rel_tol=1e-6, abs_tol=1e-6):
    if prev_mch is None:
        return True, GEOMETRY_FIELDS

    diffs = []
    for field in GEOMETRY_FIELDS:
        old_val = prev_mch.get(field)
        new_val = new_mch.get(field)
        if old_val is None or new_val is None:
            diffs.append(field)
            continue
        if not math.isclose(old_val, new_val, rel_tol=rel_tol, abs_tol=abs_tol):
            diffs.append(field)

    return (len(diffs) > 0), diffs

def init_win32com_environment():
    """Подготавливает pywin32/COM к работе с KOMPAS и снижает вероятность падения на cache/CLSID."""
    try:
        win32com.client.gencache.is_readonly = False
    except Exception:
        pass

    try:
        win32com.client.gencache.Rebuild()
    except Exception:
        pass

    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

def get_kompas_connection():
    init_win32com_environment()

    try:
        kompas_app = win32com.client.GetActiveObject("KOMPAS.Application.7")
        print("   [ИНФО] Найден уже запущенный процесс КОМПАС-3D — используем его.")
        return kompas_app, False
    except Exception as e:
        print(f"   [ОШИБКА] Не удалось подключиться к запущенному КОМПАС: {e}")
        print(traceback.format_exc())

    try:
        kompas_app = win32com.client.Dispatch("KOMPAS.Application.7")
        kompas_app.Visible = False
        print("   [ИНФО] Запущен новый фоновый процесс КОМПАС-3D.")
        return kompas_app, True
    except Exception as e:
        print(f"   [ОШИБКА] Не удалось запустить новый процесс КОМПАС: {e}")
        print(traceback.format_exc())
        return None, False

def find_open_document(kompas_app, file_path):
    target = os.path.normcase(os.path.abspath(file_path))
    try:
        for doc in kompas_app.Documents:
            try:
                doc_path = doc.PathName
                if doc_path and os.path.normcase(os.path.abspath(doc_path)) == target:
                    return doc
            except Exception:
                continue
    except Exception as e:
        print(f"   [ОШИБКА] Не удалось перечислить документы КОМПАС: {e}")
        print(traceback.format_exc())
    return None

def determine_part_type(top_part):
    """
    Определяет тип детали: листовая или обычная.
    Использует ISheetMetalContainer и SheetMetalBodies.Count.
    """
    try:
        # Приводим к ISheetMetalContainer
        sheet_metal = win32com.client.CastTo(top_part, 'ISheetMetalContainer')
        if sheet_metal is None:
            return "Обычная деталь"
        
        # Получаем коллекцию листовых тел
        sheet_bodies = sheet_metal.SheetMetalBodies
        count = sheet_bodies.Count
        
        if count > 0:
            return "Листовая деталь"
        else:
            return "Обычная деталь"
    
    except Exception:
        return "Обычная деталь"

def calculate_mch(kompas_app, file_path):
    doc = find_open_document(kompas_app, file_path)
    doc_opened_by_us = False

    try:
        if doc is not None:
            print(f"   [ИНФО] Используем уже открытый документ: {file_path}")
        else:
            print(f"   [ИНФО] Открываем документ КОМПАС: {file_path}")
            doc = kompas_app.Documents.Open(file_path, False, True)
            doc_opened_by_us = True
            if not doc:
                print(f"   [ОШИБКА] Не удалось открыть файл: {file_path}")
                return None, doc_opened_by_us, None

        doc3d = win32com.client.CastTo(doc, 'IKompasDocument3D')
        if not doc3d:
            print(f"   [ОШИБКА] Файл {file_path} не является 3D-моделью")
            return None, doc_opened_by_us, None

        top_part = doc3d.TopPart
        if not top_part:
            print(f"   [ОШИБКА] Нет детали в {file_path}")
            return None, doc_opened_by_us, None

        try:
            top_part.Update()
        except Exception as e:
            print(f"   [ПРЕДУПРЕЖДЕНИЕ] Update() не сработал для {file_path}: {e}")

        part_type = determine_part_type(top_part)
        print(f"   [ИНФО] Тип детали для {os.path.basename(file_path)}: {part_type}")

        mip = win32com.client.CastTo(top_part, 'IMassInertiaParam7')
        if mip is None:
            print(f"   [ОШИБКА] Не удалось получить IMassInertiaParam7 для {file_path}")
            return None, doc_opened_by_us, None

        if not mip.Actual:
            try:
                ok = mip.Calculate()
                if not ok:
                    print(f"   [ОШИБКА] Calculate() вернул False для {file_path}")
                    return None, doc_opened_by_us, None
            except Exception as e:
                print(f"   [ОШИБКА] Исключение в Calculate() для {file_path}: {e}")
                print(traceback.format_exc())
                return None, doc_opened_by_us, None

        mass_g = float(mip.Mass)
        density_g_cm3 = float(mip.Density)
        volume_mm3 = float(mip.Volume) * 1000.0
        area_mm2 = float(mip.Area) * 100.0
        xc_mm = float(mip.Xc) * 10.0
        yc_mm = float(mip.Yc) * 10.0
        zc_mm = float(mip.Zc) * 10.0

        jx = float(mip.Jx) * 100.0
        jy = float(mip.Jy) * 100.0
        jz = float(mip.Jz) * 100.0
        jxy = float(mip.Jxy) * 100.0
        jxz = float(mip.Jxz) * 100.0
        jyz = float(mip.Jyz) * 100.0

        inertia_tensor = np.array([
            [jx, -jxy, -jxz],
            [-jxy, jy, -jyz],
            [-jxz, -jyz, jz],
        ])
        j1, j2, j3 = sorted(np.linalg.eigvalsh(inertia_tensor), reverse=True)

        mch_data = {
            "Масса (г)": mass_g,
            "Объем (мм³)": round(volume_mm3, 3),
            "Площадь (мм²)": round(area_mm2, 3),
            "Центр масс Xc (мм)": round(xc_mm, 4),
            "Центр масс Yc (мм)": round(yc_mm, 4),
            "Центр масс Zc (мм)": round(zc_mm, 4),
            "Jx центр. (г·мм²)": round(jx, 3),
            "Jy центр. (г·мм²)": round(jy, 3),
            "Jz центр. (г·мм²)": round(jz, 3),
            "Jxy центробежный (г·мм²)": round(jxy, 3),
            "Jxz центробежный (г·мм²)": round(jxz, 3),
            "Jyz центробежный (г·мм²)": round(jyz, 3),
            "J1 главный центр. (г·мм²)": round(j1, 3),
            "J2 главный центр. (г·мм²)": round(j2, 3),
            "J3 главный центр. (г·мм²)": round(j3, 3),
        }

        print(f"   [ИНФО] Расчёт МЦХ успешно завершён для {os.path.basename(file_path)}")

        if doc_opened_by_us:
            try:
                doc.Close(2)
            except Exception:
                pass

        return mch_data, doc_opened_by_us, part_type

    except Exception as e:
        print(f"   [ОШИБКА] Ошибка при расчете МЦХ для {file_path}: {e}")
        print(traceback.format_exc())
        if doc_opened_by_us and doc is not None:
            try:
                doc.Close(2)
            except Exception:
                pass
        return None, doc_opened_by_us, None

def find_model_files(base_dir):
    model_files = []
    try:
        for root, dirs, files in os.walk(base_dir):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in MODEL_EXTENSIONS:
                    file_path = os.path.join(root, file)
                    if is_path_valid(file_path):
                        model_files.append(file_path)
    except Exception as e:
        print(f"   [ОШИБКА] Не удалось просканировать папку: {e}")
        print(traceback.format_exc())
    return model_files

def find_drawing_for_model(model_path):
    """Ищет чертёж .cdw с тем же именем в той же папке"""
    base = os.path.splitext(model_path)[0]
    drawing_path = base + ".cdw"
    if os.path.exists(drawing_path):
        return drawing_path
    return None

def document_designation(file_name):
    """Возвращает обозначение и имя без числа толщины листа."""
    stem = os.path.basename(file_name)
    lower_name = stem.casefold()
    for suffix in (".dxf.frw", ".dxf", ".m3d", ".cdw"):
        if lower_name.endswith(suffix):
            stem = stem[:-len(suffix)]
            break
    parts = re.split(r"\s*[_-]\s*", stem, maxsplit=1)
    if len(parts) != 2:
        return re.sub(r"\s+", " ", stem).strip().casefold()

    code, name = parts
    name = re.sub(r"^\d+(?:[.,]\d+)?\s*", "", name).strip()
    return f"{code.strip()} {name}".casefold()

def filename_with_disk_case(file_path):
    """Возвращает имя файла в том регистре, в котором оно хранится на диске."""
    filename = os.path.basename(file_path)
    directory = os.path.dirname(file_path)
    try:
        disk_filenames = os.listdir(directory)
        for disk_filename in disk_filenames:
            if disk_filename.casefold() == filename.casefold():
                return disk_filename

        normalized_filename = re.sub(r"\s+", " ", filename).strip().casefold()
        normalized_matches = [
            disk_filename
            for disk_filename in disk_filenames
            if re.sub(r"\s+", " ", disk_filename).strip().casefold()
            == normalized_filename
        ]
        if len(normalized_matches) == 1:
            return normalized_matches[0]
        if len(normalized_matches) > 1:
            logging.warning(
                "[WARNING] Неоднозначное имя файла после нормализации пробелов: %s",
                file_path,
            )
            return filename

        stem, extension = os.path.splitext(filename)
        if extension.casefold() == ".cdw":
            normalized_stem = re.sub(r"\s+", " ", stem).strip().casefold()
            model_matches = [
                os.path.splitext(disk_filename)[0]
                for disk_filename in disk_filenames
                if os.path.splitext(disk_filename)[1].casefold()
                in MODEL_EXTENSIONS
                and re.sub(
                    r"\s+",
                    " ",
                    os.path.splitext(disk_filename)[0],
                ).strip().casefold()
                == normalized_stem
            ]
            if len(model_matches) == 1:
                return model_matches[0] + extension
    except OSError as error:
        logging.warning(
            "[WARNING] Не удалось определить регистр имени файла %s: %s",
            file_path,
            error,
        )
    return filename

def find_unfolding_for_model(model_path, unfolding_dir):
    """Ищет DXF-развертку с тем же обозначением в отдельной папке."""
    if not unfolding_dir or not os.path.isdir(unfolding_dir):
        return None

    designation = document_designation(model_path)
    for root, _, files in os.walk(unfolding_dir):
        for file_name in files:
            if not file_name.casefold().endswith((".dxf", ".dxf.frw")):
                continue
            if document_designation(file_name) == designation:
                return os.path.join(root, file_name)
    return None

def path_belongs_to_directory(file_path, directory):
    """Проверяет, находится ли файл внутри выбранной папки."""
    try:
        file_path = os.path.normcase(os.path.abspath(file_path))
        directory = os.path.normcase(os.path.abspath(directory))
        return os.path.commonpath((file_path, directory)) == directory
    except (ValueError, TypeError):
        return False

def open_folder_and_select_file(full_path):
    """
    Открывает проводник Windows и выделяет указанный файл.
    Работает с кириллицей и пробелами в пути.
    """
    if not os.path.exists(full_path):
        print(f"   [ОШИБКА] Файл не найден: {full_path}")
        return False

    try:
        norm_path = os.path.normpath(full_path)
        cmd = f'explorer /select,"{norm_path}"'
        subprocess.Popen(cmd)
        return True
    except Exception as e:
        print(f"   [ОШИБКА] Не удалось открыть проводник: {e}")
        return False

class _MonitorCADBaseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MonitorCAD — Мониторинг листовых деталей")
        self.root.geometry("900x650")
        self.root.minsize(850, 600)
        self.root.resizable(True, True)
        
        self.monitor_dir = tk.StringVar()
        self.unfolding_dir = tk.StringVar()
        self.settings = load_settings()
        self.mail_recipients = tk.StringVar(
            value=self.settings.get("mail_recipients", CHANGE_LIST_RECIPIENT)
        )
        self.auto_mode = tk.BooleanVar()
        self.interval = tk.StringVar(value="30")
        self.is_running = False
        self.monitor_thread = None
        self.stop_event = threading.Event()
        self.cache = load_cache()
        self.changed_files = {}  # {filename: (full_path, part_type, status)}
        self.total_changes = 0
        
        # Галочки для выбора типов документов
        self.track_detali = tk.BooleanVar(value=True)
        self.track_chertezh = tk.BooleanVar(value=True)
        self.track_listovaya = tk.BooleanVar(value=True)
        self.track_unfolding = tk.BooleanVar(value=True)
        
        if "monitor_dir" in self.settings:
            self.monitor_dir.set(self.settings["monitor_dir"])
        if "unfolding_dir" in self.settings:
            self.unfolding_dir.set(self.settings["unfolding_dir"])
        if "auto_mode" in self.settings:
            self.auto_mode.set(self.settings["auto_mode"])
        if "interval" in self.settings:
            try:
                interval_val = int(self.settings["interval"])
                if interval_val >= 1:
                    self.interval.set(str(interval_val))
            except (ValueError, TypeError):
                pass  # Некорректное значение — оставляем по умолчанию "30"
        if "track_unfolding" in self.settings:
            self.track_unfolding.set(self.settings["track_unfolding"])
        
        self.setup_ui()
        self.apply_saved_settings()
        self.restore_changed_list()
        
        # Автоматический запуск, если галка уже была отмечена
        if self.auto_mode.get() and self.monitor_dir.get():
            self.root.after(500, self.start_monitoring)
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        top_frame = ttk.LabelFrame(main_frame, text="Настройки", padding="10")
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        folder_frame = ttk.Frame(top_frame)
        folder_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(folder_frame, text="Папка для мониторинга:").pack(side=tk.LEFT, padx=(0, 5))
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.monitor_dir, width=60)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(folder_frame, text="Обзор...", command=self.browse_folder).pack(side=tk.RIGHT)

        unfolding_frame = ttk.Frame(top_frame)
        unfolding_frame.pack(fill=tk.X, pady=5)

        ttk.Label(unfolding_frame, text="Папка разверток DXF:").pack(side=tk.LEFT, padx=(0, 5))
        self.unfolding_entry = ttk.Entry(unfolding_frame, textvariable=self.unfolding_dir, width=60)
        self.unfolding_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(unfolding_frame, text="Обзор...", command=self.browse_unfolding_folder).pack(side=tk.RIGHT)

        recipients_frame = ttk.Frame(top_frame)
        recipients_frame.pack(fill=tk.X, pady=5)
        ttk.Label(recipients_frame, text="Адреса рассылки:").pack(
            side=tk.LEFT,
            padx=(0, 5),
        )
        self.mail_recipients_entry = ttk.Entry(
            recipients_frame,
            textvariable=self.mail_recipients,
            width=60,
        )
        self.mail_recipients_entry.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
        )
        self.mail_recipients_entry.bind(
            "<Return>",
            lambda event: self.save_settings(),
        )
        self.mail_recipients_entry.bind(
            "<FocusOut>",
            lambda event: self.save_settings(),
        )
        
        # Блок выбора типов документов
        types_frame = ttk.Frame(top_frame)
        types_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(types_frame, text="Отслеживать:").pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Checkbutton(types_frame, text="Деталь (.m3d)", variable=self.track_detali).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(types_frame, text="Чертёж (.cdw)", variable=self.track_chertezh).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(types_frame, text="Листовое тело (.m3d)", variable=self.track_listovaya).pack(side=tk.LEFT)
        ttk.Checkbutton(types_frame, text="Развертка (.dxf)", variable=self.track_unfolding).pack(side=tk.LEFT, padx=(10, 0))
        
        auto_frame = ttk.Frame(top_frame)
        auto_frame.pack(fill=tk.X, pady=5)
        
        self.auto_check = ttk.Checkbutton(auto_frame, text="Автоматический режим (по интервалу)", 
                                          variable=self.auto_mode, command=self.on_auto_mode_change)
        self.auto_check.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(auto_frame, text="Интервал (секунд):").pack(side=tk.LEFT, padx=(0, 5))
        self.interval_entry = ttk.Entry(auto_frame, textvariable=self.interval, width=10)
        self.interval_entry.pack(side=tk.LEFT)
        self.interval_entry.bind('<Return>', self.on_interval_enter)
        self.interval_entry.bind('<FocusOut>', self.on_interval_focus_out)
        
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.start_btn = ttk.Button(button_frame, text="▶ ПУСК", command=self.start_monitoring, width=15)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_btn = ttk.Button(button_frame, text="■ СТОП", command=self.stop_monitoring, width=15, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT)
        
        indicator_frame = ttk.Frame(top_frame)
        indicator_frame.pack(fill=tk.X, pady=5)
        
        self.indicator_canvas = tk.Canvas(indicator_frame, width=20, height=20, bg='white')
        self.indicator_canvas.pack(side=tk.LEFT, padx=(0, 10))
        self.indicator_oval = self.indicator_canvas.create_oval(3, 3, 17, 17, fill='gray', outline='black')
        
        self.status_label = ttk.Label(indicator_frame, text="ОСТАНОВЛЕН", font=("Arial", 10, "bold"))
        self.status_label.pack(side=tk.LEFT)

        # Бегущая строка состояния — показывает, что именно сейчас делает
        # поток мониторинга (сканирование, подключение к КОМПАС, расчёт
        # конкретного файла, обратный отсчёт до следующей проверки и т.д.)
        status_line_frame = ttk.Frame(top_frame)
        status_line_frame.pack(fill=tk.X, pady=(2, 0))

        self.status_line_var = tk.StringVar(value="Готово к запуску.")
        self.status_line_label = ttk.Label(
            status_line_frame,
            textvariable=self.status_line_var,
            font=("Consolas", 9),
            foreground="#555555",
        )
        self.status_line_label.pack(side=tk.LEFT)

        # Полоса прогресса: детерминированная с процентами. Показывает
        # количество обработанных файлов при сканировании, процент выполнения
        # расчёта МЦХ, а также обратный отсчёт до следующей проверки.
        self.progress_bar = ttk.Progressbar(
            top_frame,
            orient=tk.HORIZONTAL,
            mode="determinate",
            maximum=100,
            length=300,
        )
        self.progress_bar.pack(fill=tk.X, pady=(4, 0))

        list_frame = ttk.LabelFrame(main_frame, text="Изменённые файлы", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        columns = ("filename", "part_type", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        
        self.tree.heading("filename", text="Файл")
        self.tree.heading("part_type", text="Тип")
        self.tree.heading("status", text="Статус")
        
        self.tree.column("filename", width=440, anchor=tk.W)
        self.tree.column("part_type", width=180, anchor=tk.CENTER)
        self.tree.column("status", width=220, anchor=tk.CENTER)
        self.group_colors = ("#eaf3ff", "#fff4df", "#eaf8ed", "#f5eaff", "#ffecef")
        for index, color in enumerate(self.group_colors):
            self.tree.tag_configure(f"group_{index}", background=color)
        
        self.tree.bind("<Double-1>", self.open_file_folder)
        self.tree.bind("<<TreeviewSelect>>", self.update_clear_selected_button)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X)
        
        self.clear_selected_btn = ttk.Button(
            bottom_frame,
            text="Снять выбранную",
            command=self.clear_selected_record,
            state=tk.DISABLED,
        )
        self.clear_selected_btn.pack(side=tk.RIGHT)
        
    
    def restore_changed_list(self):
        """
        Восстанавливает список "Изменённые файлы" из персистентного файла
        на диске (changed_files_list.json). Вызывается один раз при старте
        программы, чтобы список не терялся при закрытии программы,
        сбое питания и т.п. Очищается список только через clear_list().
        """
        records, total_changes = load_changed_list()

        for rec in records:
            status = rec.get("status", "")
            if status == "ПЕРВОЕ ИЗМЕНЕНИЕ":
                status = "ИЗМЕНЕН"
            full_path = rec.get("full_path", "")
            filename = (
                filename_with_disk_case(full_path)
                if full_path
                else rec.get("filename", "")
            )
            self.tree.insert(
                "",
                tk.END,
                values=(filename, rec.get("part_type", ""), status),
            )
            self.changed_files[filename] = (
                full_path,
                rec.get("part_type", ""),
                status,
            )

        self.total_changes = total_changes
        self.refresh_tree_groups()

    def persist_changed_list(self):
        """
        Сохраняет текущее содержимое tree на диск в том же порядке, в котором
        оно отображается. Вызывается после каждого добавления/обновления
        строки, чтобы список переживал закрытие программы.
        """
        records = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if len(values) < 3:
                continue
            filename, part_type, status = values
            full_path = self.changed_files.get(filename, ("", "", ""))[0]
            records.append({
                "timestamp": datetime.now().strftime("%H:%M:%S"),
                "filename": filename,
                "part_type": part_type,
                "status": status,
                "full_path": full_path,
            })
        save_changed_list(records, self.total_changes)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку для мониторинга")
        if folder:
            self.monitor_dir.set(os.path.normpath(folder))
            self.save_settings()

    def browse_unfolding_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с развертками DXF")
        if folder:
            self.unfolding_dir.set(os.path.normpath(folder))
            self.save_settings()

    def use_state_for_selected_folder(self):
        """Оставляет только кэш и записи, относящиеся к выбранной папке."""
        selected_dir = self.monitor_dir.get()
        self.cache = {
            key: value for key, value in self.cache.items()
            if path_belongs_to_directory(key, selected_dir)
        }

        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            filename = values[0] if values else ""
            full_path = self.changed_files.get(filename, ("", "", ""))[0]
            record_type = self.changed_files.get(filename, ("", "", ""))[1]
            record_dir = self.unfolding_dir.get() if record_type == "Развертка DXF" else selected_dir
            if not path_belongs_to_directory(full_path, record_dir):
                self.tree.delete(item)
                self.changed_files.pop(filename, None)

        self.total_changes = len(self.tree.get_children())
        self.update_clear_selected_button()
        self.persist_changed_list()
    
    def on_auto_mode_change(self):
        if self.auto_mode.get():
            self.start_btn.config(state=tk.DISABLED)
            self.interval_entry.config(state=tk.NORMAL)
            if self.is_running:
                self.stop_monitoring()
                self.start_monitoring()
            elif self.interval.get() and self.monitor_dir.get():
                self.start_monitoring()
        else:
            self.start_btn.config(state=tk.NORMAL)
            self.interval_entry.config(state=tk.DISABLED)
            # При выключении автоматического режима останавливаем текущий цикл
            if self.is_running:
                self.stop_monitoring()
        
        self.save_settings()
    
    def on_interval_enter(self, event):
        self.save_settings()
        if self.auto_mode.get() and self.is_running:
            self.stop_monitoring()
            self.start_monitoring()
    
    def on_interval_focus_out(self, event):
        self.save_settings()
        if self.auto_mode.get() and self.is_running:
            self.stop_monitoring()
            self.start_monitoring()
    
    def save_settings(self):
        # Поле интервала не должно оставаться пустым/битым — нормализуем перед сохранением
        interval_str = self.interval.get().strip()
        try:
            interval_val = int(interval_str)
            if interval_val < 1:
                raise ValueError
        except (ValueError, TypeError):
            interval_str = "30"
        self.interval.set(interval_str)
        settings = {
            "monitor_dir": self.monitor_dir.get(),
            "unfolding_dir": self.unfolding_dir.get(),
            "auto_mode": self.auto_mode.get(),
            "interval": interval_str,
            "track_detali": self.track_detali.get(),
            "track_chertezh": self.track_chertezh.get(),
            "track_listovaya": self.track_listovaya.get(),
            "track_unfolding": self.track_unfolding.get(),
            "mail_recipients": self.mail_recipients.get().strip(),
        }
        save_settings(settings)
    
    def apply_saved_settings(self):
        if self.auto_mode.get():
            self.start_btn.config(state=tk.DISABLED)
            self.interval_entry.config(state=tk.NORMAL)
        else:
            self.start_btn.config(state=tk.NORMAL)
            self.interval_entry.config(state=tk.DISABLED)
    
    def set_status(self, text):
        """
        Потокобезопасно обновляет "бегущую строку" состояния.
        Можно вызывать как из главного потока, так и из monitor_loop
        (внутри всегда уходит через root.after на главный поток).
        """
        self.root.after(0, lambda t=text: self.status_line_var.set(t))

    def set_progress(self, value):
        """Потокобезопасно выставляет процент заполнения детерминированной полосы (0-100)."""
        self.root.after(0, lambda v=value: self.progress_bar.config(value=v))

    def clear_selected_record(self):
        """Удаляет выбранную запись и связанное сообщение об отсутствии чертежа."""
        selected = self.tree.selection()
        if not selected:
            self.status_line_var.set("Сначала выберите запись в списке.")
            self.tree.focus_set()
            return

        values = self.tree.item(selected[0], "values")
        if values:
            self.clear_single_record(values[0])

    def update_clear_selected_button(self, event=None):
        """Делает кнопку снятия доступной только при выбранной записи."""
        state = tk.NORMAL if self.tree.selection() else tk.DISABLED
        self.clear_selected_btn.config(state=state)

    def check_unfolding_for_model(self, model_path, part_type, drawing_path):
        """Добавляет предупреждение, если DXF-развертка отсутствует или устарела."""
        if (
            part_type != "Листовая деталь"
            or not self.track_unfolding.get()
            or not self.unfolding_dir.get()
        ):
            return

        unfolding_path = find_unfolding_for_model(model_path, self.unfolding_dir.get())
        if unfolding_path is None:
            unfolding_filename = os.path.splitext(os.path.basename(model_path))[0] + ".dxf"
            record_status = "НЕТ РАЗВЕРТКИ"
            record_path = os.path.join(self.unfolding_dir.get(), unfolding_filename)
        elif drawing_path is None:
            return
        else:
            model_mtime = os.path.getmtime(model_path)
            drawing_mtime = os.path.getmtime(drawing_path)
            unfolding_mtime = os.path.getmtime(unfolding_path)
            if model_mtime > drawing_mtime:
                unfolding_filename = os.path.basename(unfolding_path)
                record_status = "ТРЕБУЕТ ПЕРЕСТРОЙКИ"
                record_path = unfolding_path
            elif unfolding_mtime > drawing_mtime:
                unfolding_filename = os.path.basename(unfolding_path)
                record_status = "ИЗМЕНЕН"
                record_path = unfolding_path
            else:
                unfolding_filename = os.path.basename(unfolding_path)
                record_status = "РАЗВЕРТКА УСТАРЕЛА"
                record_path = unfolding_path

        self.root.after(
            0,
            self.set_unfolding_record,
            model_path,
            unfolding_filename,
            record_path,
            record_status,
        )

    def set_unfolding_record(self, model_path, filename, full_path, status):
        """Сохраняет событие развертки, заменяя прежнее предупреждение."""
        filename = filename_with_disk_case(full_path)
        designation = document_designation(model_path)
        old_records = {
            old_filename for old_filename, (_, part_type, _) in self.changed_files.items()
            if part_type == "Развертка DXF"
            and document_designation(old_filename) == designation
            and old_filename != filename
        }

        for old_filename in old_records:
            self.changed_files.pop(old_filename, None)

        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if values and values[0] in old_records:
                self.tree.delete(item)

        self.total_changes = max(0, self.total_changes - len(old_records))
        already_listed = filename in self.changed_files
        self.changed_files[filename] = (full_path, "Развертка DXF", status)
        if not already_listed:
            self.total_changes += 1
        self.add_to_list(filename, "Развертка DXF", status)

    def refresh_pending_unfoldings(self):
        """Повторно проверяет DXF для листовых моделей с предупреждениями."""
        pending_designations = {
            document_designation(filename)
            for filename, (_, part_type, _) in self.changed_files.items()
            if part_type in {"Листовая деталь", "Чертёж", "Развертка DXF"}
        }
        if not pending_designations:
            return

        for model_path in find_model_files(self.monitor_dir.get()):
            cache_entry = self.cache.get(
                os.path.normcase(os.path.abspath(model_path)),
                {},
            )
            model_record = self.changed_files.get(os.path.basename(model_path), ("", "", ""))
            part_type = cache_entry.get("part_type") or model_record[1]
            if part_type != "Листовая деталь":
                continue
            if document_designation(model_path) not in pending_designations:
                continue
            self.check_unfolding_for_model(
                model_path,
                "Листовая деталь",
                cache_entry.get("drawing_path") or find_drawing_for_model(model_path),
            )

    def clear_single_record(self, filename):
        """Удаляет запись и связанное сообщение об отсутствии чертежа."""
        filenames_to_remove = {filename}
        base_name, extension = os.path.splitext(filename)

        if extension.lower() == ".m3d":
            drawing_filename = base_name + ".cdw"
            drawing_record = self.changed_files.get(drawing_filename)
            if drawing_record and drawing_record[2] == "НЕТ ЧЕРТЕЖА":
                filenames_to_remove.add(drawing_filename)
        elif extension.lower() == ".cdw":
            model_filename = base_name + ".m3d"
            drawing_record = self.changed_files.get(filename)
            if drawing_record and drawing_record[2] == "НЕТ ЧЕРТЕЖА":
                filenames_to_remove.add(model_filename)

        for record_filename in filenames_to_remove:
            self.changed_files.pop(record_filename, None)

        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            if values and values[0] in filenames_to_remove:
                self.tree.delete(item)

        self.total_changes = max(0, self.total_changes - len(filenames_to_remove))
        self.update_clear_selected_button()
        self.persist_changed_list()

    def open_file_folder(self, event):
        """Открывает папку с файлом при двойном клике и выделяет файл"""
        selected = self.tree.selection()
        if selected:
            item = selected[0]
            values = self.tree.item(item, 'values')
            filename = values[0]
            
            if filename in self.changed_files:
                full_path, part_type, status = self.changed_files[filename]
                # Открываем проводник и выделяем файл через Windows API (из главного потока GUI)
                open_folder_and_select_file(full_path)
    
    def start_monitoring(self):
        if not self.monitor_dir.get():
            messagebox.showerror("Ошибка", "Выберите папку для мониторинга!")
            return
        
        if not os.path.exists(self.monitor_dir.get()):
            messagebox.showerror("Ошибка", "Указанная папка не существует!")
            return
        
        if self.is_running:
            self.stop_monitoring()
        
        try:
            if self.interval.get().strip():
                interval = int(self.interval.get().strip())
                if interval < 1:
                    raise ValueError
            else:
                interval = 30
                self.interval.set("30")
        except ValueError:
            messagebox.showerror("Ошибка", "Интервал должен быть положительным числом!")
            return
        
        # Ручной режим выполняет один полный цикл; авто — бесконечный по интервалу
        one_shot = not self.auto_mode.get()
        
        self.cache = load_cache()
        self.use_state_for_selected_folder()
        
        if not self.cache:
            print("Создание первичного слепка...")
            self.status_line_var.set("Создание первичного слепка папки...")
            self.root.update_idletasks()
            model_files = find_model_files(self.monitor_dir.get())
            total = len(model_files)
            for idx, file_path in enumerate(model_files, start=1):
                key = os.path.normcase(os.path.abspath(file_path))
                try:
                    stat = os.stat(file_path)
                    drawing_path = find_drawing_for_model(file_path)
                    drawing_mtime = None
                    if drawing_path:
                        drawing_mtime = os.stat(drawing_path).st_mtime
                    
                    self.cache[key] = {
                        "mtime": stat.st_mtime,
                        "size": stat.st_size,
                        "hash": get_file_hash(file_path),
                        "last_mch": None,
                        "mch_calculated": False,
                        "last_check": datetime.now().isoformat(),
                        "drawing_path": drawing_path,
                        "drawing_mtime": drawing_mtime,
                        "part_type": None
                    }
                except Exception:
                    continue
                
                # Обновляем прогресс
                percent = int(idx / total * 100) if total > 0 else 0
                self.set_status(f"Создание слепка... {idx} из {total} файлов ({percent}%)")
                self.set_progress(percent)
            save_cache(self.cache)
        
        self.is_running = True
        self.stop_event.clear()
        self.stop_btn.config(state=tk.NORMAL)
        self.start_btn.config(state=tk.DISABLED)
        self.status_label.config(text="РАБОТАЕТ")
        self.indicator_canvas.itemconfig(self.indicator_oval, fill='green')
        self.status_line_var.set("Мониторинг запущен.")
        
        self.monitor_thread = threading.Thread(target=self.monitor_loop, args=(interval, one_shot), daemon=True)
        self.monitor_thread.start()
        
        self.save_settings()
    
    def stop_monitoring(self):
        if not self.is_running:
            return
        
        self.is_running = False
        self.stop_event.set()
        self.stop_btn.config(state=tk.DISABLED)
        self.start_btn.config(state=tk.NORMAL)
        self.status_label.config(text="ОСТАНОВЛЕН")
        self.indicator_canvas.itemconfig(self.indicator_oval, fill='gray')
        self.status_line_var.set("Мониторинг остановлен.")
    
    def _finish_one_shot(self):
        """Завершает разовый прогон ручного режима: возвращает интерфейс в состояние 'ОСТАНОВЛЕН'."""
        if not self.is_running:
            return
        self.stop_monitoring()
        self.status_line_var.set("Проверка завершена (ручной режим).")
    
    def clear_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.update_clear_selected_button()
        
        self.changed_files = {}
        self.total_changes = 0
        # Персистентный список изменений очищается ТОЛЬКО здесь - по нажатию кнопки.
        save_changed_list([], 0)
        
        # Обновляем слепок ТОЛЬКО для файлов, которые были в списке
        if self.is_running:
            save_cache(self.cache)
            print("Слепок обновлён — изменения приняты в работу")

        self.status_line_var.set("Список очищен.")
    
    def monitor_loop(self, interval, one_shot=False):
        # Инициализируем COM в этом потоке
        pythoncom.CoInitialize()
        
        # В ручном режиме (один цикл) сканирование начинается сразу,
        # в авто — по истечении интервала
        next_check_time = time.monotonic() if one_shot else time.monotonic() + interval
        
        try:
            while not self.stop_event.is_set():
                try:
                    current_time = time.monotonic()
                    
                    if current_time >= next_check_time:

                        # Сканирование папки с живым прогрессом
                        self.set_status("Сканирование папки на изменения...")
                        model_files = find_model_files(self.monitor_dir.get())
                        total_files = len(model_files)
                        
                        changed_models = []
                        for idx, file_path in enumerate(model_files, start=1):
                            try:
                                key = os.path.normcase(os.path.abspath(file_path))
                                if key not in self.cache:
                                    # Новая деталь — заносим в реестр
                                    try:
                                        stat = os.stat(file_path)
                                        drawing_path = find_drawing_for_model(file_path)
                                        drawing_mtime = None
                                        if drawing_path:
                                            drawing_mtime = os.stat(drawing_path).st_mtime
                                        
                                        self.cache[key] = {
                                            "mtime": stat.st_mtime,
                                            "size": stat.st_size,
                                            "hash": get_file_hash(file_path),
                                            "last_mch": None,
                                            "mch_calculated": False,
                                            "last_check": datetime.now().isoformat(),
                                            "drawing_path": drawing_path,
                                            "drawing_mtime": drawing_mtime,
                                            "part_type": None
                                        }
                                        
                                        # Сразу добавляем в список изменений как НОВЫЙ
                                        filename = os.path.basename(file_path)
                                        part_type = "Неизвестно"
                                        status = "НОВЫЙ"
                                        
                                        self.total_changes += 1
                                        self.changed_files[filename] = (file_path, part_type, status)
                                        self.root.after(0, self.add_to_list, filename, part_type, status)
                                        
                                        # Проверяем, есть ли чертёж
                                        if drawing_path is None:
                                            # Нет чертежа — добавляем строку
                                            drawing_filename = os.path.splitext(filename)[0] + ".cdw"
                                            drawing_status = "НЕТ ЧЕРТЕЖА"
                                            self.total_changes += 1
                                            self.changed_files[drawing_filename] = (os.path.join(os.path.dirname(file_path), drawing_filename), "Чертёж", drawing_status)
                                            self.root.after(0, self.add_to_list, drawing_filename, "Чертёж", drawing_status)
                                    except Exception:
                                        continue
                                
                                markers_changed, _, _, _, _ = file_markers_changed(file_path, self.cache)
                                if markers_changed:
                                    changed_models.append(file_path)
                                
                                # Обновляем прогресс сканирования
                                percent = int(idx / total_files * 100) if total_files > 0 else 0
                                self.set_status(f"Сканирование... {idx} из {total_files} файлов ({percent}%)")
                                self.set_progress(percent)
                            except Exception:
                                continue
                        
                        # Проверяем, не появились ли чертежи для файлов со статусом "НЕТ ЧЕРТЕЖА"
                        for filename, (file_path, part_type, status) in list(self.changed_files.items()):
                            if status == "НЕТ ЧЕРТЕЖА":
                                # Находим модель для этого чертежа
                                model_path = os.path.splitext(file_path)[0] + ".m3d"
                                if os.path.exists(model_path) and os.path.isfile(file_path):
                                    # Чертёж появился — обновляем статус
                                    drawing_path = file_path
                                    drawing_mtime = os.stat(drawing_path).st_mtime
                                    model_mtime = self.cache.get(os.path.normcase(os.path.abspath(model_path)), {}).get("mtime", 0)
                                    
                                    if drawing_mtime >= model_mtime:
                                        # Чертёж актуален — меняем статус на ИЗМЕНЕН
                                        self.changed_files[filename] = (drawing_path, "Чертёж", "ИЗМЕНЕН")
                                        self.root.after(0, self.update_status_in_list, filename, "ИЗМЕНЕН")
                                    else:
                                        # Чертёж устарел — меняем статус на ТРЕБУЕТ ПЕРЕСТРОЙКИ
                                        self.changed_files[filename] = (drawing_path, "Чертёж", "ТРЕБУЕТ ПЕРЕСТРОЙКИ")
                                        self.root.after(0, self.update_status_in_list, filename, "ТРЕБУЕТ ПЕРЕСТРОЙКИ")
                                    
                                    # Обновляем drawing_path в кэше модели
                                    model_key = os.path.normcase(os.path.abspath(model_path))
                                    if model_key in self.cache:
                                        self.cache[model_key]["drawing_path"] = drawing_path
                                        self.cache[model_key]["drawing_mtime"] = drawing_mtime
                                        save_cache(self.cache)

                        # Проверяем, не появились ли ожидаемые развертки DXF.
                        self.refresh_pending_unfoldings()
                                
                        if changed_models:
                            self.root.after(0, self.update_indicator, 'red')
                            self.set_status(f"Обнаружено изменённых файлов: {len(changed_models)}. Подключение к КОМПАС-3D...")
                            
                            kompas_app = None
                            app_started_by_us = False
                            
                            try:
                                kompas_app, app_started_by_us = get_kompas_connection()
                                
                                if kompas_app is None:
                                    print("   [ОШИБКА] Не удалось подключиться к КОМПАС — пропускаем проверку")
                                    self.set_status("Не удалось подключиться к КОМПАС-3D.")
                                    continue
                                
                                for idx, file_path in enumerate(changed_models, start=1):
                                    if self.stop_event.is_set():
                                        break

                                    percent = int(idx / len(changed_models) * 100)
                                    self.set_status(
                                        f"Расчёт МЦХ ({idx}/{len(changed_models)}, {percent}%): "
                                        f"{os.path.basename(file_path)}"
                                    )
                                    self.set_progress(percent)
                                    
                                    try:
                                        key = os.path.normcase(os.path.abspath(file_path))
                                        prev_mch = self.cache[key].get("last_mch")
                                        mch_calculated = self.cache[key].get("mch_calculated", False)
                                        
                                        mch_data, _, part_type = calculate_mch(kompas_app, file_path)
                                        
                                        if mch_data is not None:
                                            changed, diff_fields = geometry_actually_changed(prev_mch, mch_data)
                                            
                                            filename = os.path.basename(file_path)
                                            
                                            # Получаем тип детали из calculate_mch
                                            if part_type is None:
                                                part_type = self.cache[key].get("part_type", "Обычная деталь")
                                            
                                            # Определяем статус
                                            if not mch_calculated:
                                                if prev_mch is None:
                                                    status = "ИЗМЕНЕН"
                                                else:
                                                    status = "ИЗМЕНЕН"
                                            else:
                                                if changed:
                                                    status = "ИЗМЕНЕН"
                                                else:
                                                    # Геометрия не изменилась — не выводим в список
                                                    stat = os.stat(file_path)
                                                    self.cache[key]["mtime"] = stat.st_mtime
                                                    self.cache[key]["size"] = stat.st_size
                                                    self.cache[key]["hash"] = get_file_hash(file_path)
                                                    self.cache[key]["last_check"] = datetime.now().isoformat()
                                                    save_cache(self.cache)
                                                    continue
                                            
                                            self.total_changes += 1
                                            self.changed_files[filename] = (file_path, part_type, status)
                                            
                                            self.root.after(0, self.add_to_list, filename, part_type, status)
                                            
                                            stat = os.stat(file_path)
                                            model_mtime = stat.st_mtime
                                            
                                            # Проверяем чертёж
                                            drawing_path = self.cache[key].get("drawing_path")
                                            drawing_mtime = self.cache[key].get("drawing_mtime")
                                            
                                            # Проверяем наличие чертежа
                                            if drawing_path is None:
                                                # Нет чертежа — добавляем строку
                                                drawing_filename = os.path.splitext(filename)[0] + ".cdw"
                                                drawing_status = "НЕТ ЧЕРТЕЖА"
                                                self.total_changes += 1
                                                self.changed_files[drawing_filename] = (os.path.join(os.path.dirname(file_path), drawing_filename), "Чертёж", drawing_status)
                                                self.root.after(0, self.add_to_list, drawing_filename, "Чертёж", drawing_status)
                                            elif drawing_mtime is not None and drawing_mtime < model_mtime:
                                                # Чертёж устарел
                                                if self.track_chertezh.get():
                                                    drawing_filename = os.path.basename(drawing_path)
                                                    drawing_status = "ТРЕБУЕТ ПЕРЕСТРОЙКИ"
                                                    self.total_changes += 1
                                                    self.changed_files[drawing_filename] = (drawing_path, "Чертёж", drawing_status)
                                                    self.root.after(0, self.add_to_list, drawing_filename, "Чертёж", drawing_status)

                                            self.check_unfolding_for_model(file_path, part_type, drawing_path)
                                            
                                            self.cache[key] = {
                                                "mtime": stat.st_mtime,
                                                "size": stat.st_size,
                                                "hash": get_file_hash(file_path),
                                                "last_mch": mch_data,
                                                "mch_calculated": True,
                                                "last_change": datetime.now().isoformat(),
                                                "change_count": self.cache[key].get("change_count", 0) + 1,
                                                "drawing_path": drawing_path,
                                                "drawing_mtime": drawing_mtime,
                                                "part_type": part_type
                                            }
                                            
                                        else:
                                            filename = os.path.basename(file_path)
                                            status = "ОШИБКА"
                                            reason = "Не удалось рассчитать МЦХ"
                                            print(f"   [ОШИБКА] Не удалось рассчитать МЦХ для {filename}. Причина: {reason}")
                                            self.total_changes += 1
                                            self.changed_files[filename] = (file_path, "Неизвестно", status)
                                            
                                            self.root.after(0, self.add_to_list, filename, "Неизвестно", status)
                                            
                                            stat = os.stat(file_path)
                                            self.cache[key]["mtime"] = stat.st_mtime
                                            self.cache[key]["size"] = stat.st_size
                                            self.cache[key]["hash"] = get_file_hash(file_path)
                                            self.cache[key]["last_check"] = datetime.now().isoformat()
                                            
                                    except Exception:
                                        continue
                                
                            except Exception as e:
                                print(f"   [ОШИБКА] Проблема при работе с КОМПАС: {e}")
                                print(traceback.format_exc())
                            finally:
                                if app_started_by_us and kompas_app:
                                    try:
                                        kompas_app.Quit()
                                    except Exception:
                                        pass
                                
                                save_cache(self.cache)
                            
                            self.root.after(0, self.update_indicator, 'green')
                            self.set_progress(0)

                        self.set_status("Проверка статуса чертежей...")

                        # Проверяем, обновился ли чертёж для файлов со статусом "ТРЕБУЕТ ПЕРЕСТРОЙКИ"
                        for filename, (file_path, part_type, status) in list(self.changed_files.items()):
                            if status == "ТРЕБУЕТ ПЕРЕСТРОЙКИ":
                                try:
                                    # Находим модель для этого чертежа
                                    drawing_path = file_path
                                    model_path = os.path.splitext(drawing_path)[0] + ".m3d"
                                    
                                    if os.path.exists(model_path):
                                        model_key = os.path.normcase(os.path.abspath(model_path))
                                        model_entry = self.cache.get(model_key, {})
                                        model_mtime = model_entry.get("mtime", 0)
                                        drawing_mtime = os.stat(drawing_path).st_mtime
                                        
                                        if drawing_mtime >= model_mtime:
                                            # Чертёж перестроен — меняем статус на ИЗМЕНЕН
                                            self.changed_files[filename] = (drawing_path, "Чертёж", "ИЗМЕНЕН")
                                            self.root.after(0, self.update_status_in_list, filename, "ИЗМЕНЕН")
                                            
                                            # Обновляем drawing_mtime в кэше модели
                                            if model_key in self.cache:
                                                self.cache[model_key]["drawing_mtime"] = drawing_mtime
                                                save_cache(self.cache)
                                except Exception:
                                    continue
                        
                        self.root.after(0, self.update_status_label)

                        if changed_models:
                            self.set_status(
                                f"Проверка завершена. Обработано изменений: {len(changed_models)}. "
                                f"Следующая проверка через {interval} сек."
                            )
                        else:
                            self.set_status(f"Изменений не обнаружено. Следующая проверка через {interval} сек.")

                        next_check_time = time.monotonic() + interval
                        
                        # В ручном режиме (одиночная проверка) завершаем цикл
                        if one_shot:
                            break
                        
                except Exception as e:
                    print(f"Ошибка в цикле мониторинга: {e}")
                    print(traceback.format_exc())
                    next_check_time = time.monotonic() + interval
                    if one_shot:
                        break

                # Пока ждём следующей проверки - показываем тикающий обратный отсчёт
                # и заполняющуюся полосу (чем ближе следующая проверка - тем полнее),
                # чтобы было видно, что поток жив и не завис.
                remaining = max(0, int(next_check_time - time.monotonic()))
                if remaining > 0:
                    self.set_status(f"Ожидание... следующая проверка через {remaining} сек.")
                    cooldown_percent = int(100 * (1 - remaining / interval)) if interval > 0 else 0
                    self.set_progress(cooldown_percent)

                self.stop_event.wait(timeout=1)
        finally:
            # Освобождаем COM при завершении потока
            pythoncom.CoUninitialize()
            
            # В ручном режиме после одного цикла возвращаем интерфейс в состояние "ОСТАНОВЛЕН"
            if one_shot and not self.stop_event.is_set():
                self.root.after(0, self._finish_one_shot)
    
    def update_status_in_list(self, filename, new_status):
        """Обновляет статус файла в списке"""
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            if values and values[0].casefold() == filename.casefold():
                row_values = list(values)
                row_values[2] = new_status
                self.tree.item(item, values=tuple(row_values))
                break
        self.persist_changed_list()
    
    def add_to_list(self, filename, part_type, status):
        record = self.changed_files.get(filename)
        if record and record[0]:
            disk_filename = filename_with_disk_case(record[0])
            if disk_filename != filename:
                self.changed_files.pop(filename, None)
                filename = disk_filename
                self.changed_files[filename] = record

        # Ищем уже существующую запись с таким файлом
        existing_item = None
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            if values and values[0].casefold() == filename.casefold():
                existing_item = item
                break
        
        # Если есть — удаляем старую
        if existing_item:
            self.tree.delete(existing_item)
        
        # Добавляем новую
        self.tree.insert("", tk.END, values=(filename, part_type, status))
        self.refresh_tree_groups()
        self.persist_changed_list()

    def refresh_tree_groups(self):
        """Группирует связанные документы рядом в списке."""
        group_order = {}
        items = list(self.tree.get_children())
        for item in items:
            values = self.tree.item(item, "values")
            if not values:
                continue
            filename = values[0]
            group_key = document_designation(filename)
            if group_key not in group_order:
                group_order[group_key] = len(group_order)
            color_index = hash(group_key) % len(self.group_colors)
            self.tree.item(item, tags=(f"group_{color_index}",))

        ordered_items = sorted(
            items,
            key=lambda item: group_order.get(
                document_designation(self.tree.item(item, "values")[0]),
                0,
            ),
        )
        for index, item in enumerate(ordered_items):
            self.tree.move(item, "", index)
    
    def update_indicator(self, color):
        self.indicator_canvas.itemconfig(self.indicator_oval, fill=color)
    
    def update_status_label(self):
        self.status_label.config(text="РАБОТАЕТ")

EVENT_CACHE_PATH = os.path.join(SCRIPT_DIR, "mch_cache_events.json")

EVENT_CHANGED_LIST_PATH = os.path.join(
    SCRIPT_DIR, "changed_files_events.json"
)

CHANGE_JOURNAL_PATH = os.path.join(
    SCRIPT_DIR, "change_journal_events.json"
)

APPLICATION_ICON_PATH = os.path.join(
    SCRIPT_DIR, "change-management.ico"
)

CHANGE_LIST_RECIPIENT = "ationis@mail.ru"

def load_event_cache():
    if not os.path.exists(EVENT_CACHE_PATH):
        return {}
    try:
        with open(EVENT_CACHE_PATH, "r", encoding="utf-8") as cache_file:
            data = json.load(cache_file)
        if not isinstance(data, dict):
            raise ValueError("Корень файла кэша должен быть объектом JSON.")
        return data
    except (OSError, json.JSONDecodeError, ValueError) as error:
        print(f"[ОШИБКА] Не удалось загрузить событийный кэш: {error}")
        return {}

def save_event_cache(cache):
    try:
        with open(EVENT_CACHE_PATH, "w", encoding="utf-8") as cache_file:
            json.dump(cache, cache_file, ensure_ascii=False, indent=2)
    except OSError as error:
        print(f"[ОШИБКА] Не удалось сохранить событийный кэш: {error}")

def load_event_changed_list():
    if not os.path.exists(EVENT_CHANGED_LIST_PATH):
        return [], 0
    try:
        with open(EVENT_CHANGED_LIST_PATH, "r", encoding="utf-8") as records_file:
            data = json.load(records_file)
        if not isinstance(data, dict):
            raise ValueError("Корень файла списка изменений должен быть объектом JSON.")
        records = data.get("records", [])
        total_changes = data.get("total_changes", 0)
        if not isinstance(records, list) or not isinstance(total_changes, int):
            raise ValueError("Неверный формат списка изменений.")
        return records, total_changes
    except (OSError, json.JSONDecodeError, ValueError) as error:
        print(f"[ОШИБКА] Не удалось загрузить список изменений: {error}")
        return [], 0

def save_event_changed_list(records, total_changes):
    try:
        with open(EVENT_CHANGED_LIST_PATH, "w", encoding="utf-8") as records_file:
            json.dump(
                {"records": records, "total_changes": total_changes},
                records_file,
                ensure_ascii=False,
                indent=2,
            )
    except OSError as error:
        print(f"[ОШИБКА] Не удалось сохранить список изменений: {error}")

def load_change_journal():
    if not os.path.exists(CHANGE_JOURNAL_PATH):
        return []
    try:
        with open(CHANGE_JOURNAL_PATH, "r", encoding="utf-8") as journal_file:
            data = json.load(journal_file)
        if not isinstance(data, dict):
            raise ValueError("Корень файла журнала должен быть объектом JSON.")
        records = data.get("records", [])
        if not isinstance(records, list):
            raise ValueError("Записи журнала должны быть списком.")
        changes = [
            record
            for record in records
            if isinstance(record, dict)
            and all(
                isinstance(record.get(field), str)
                for field in ("timestamp", "filename", "full_path")
            )
            and record.get("status", "ИЗМЕНЕН") == "ИЗМЕНЕН"
        ]
        unique_changes = []
        seen_signatures = set()
        last_change_by_path = {}
        for record in sorted(changes, key=lambda item: item["timestamp"]):
            path_key = document_key(record["full_path"])
            file_signature = record.get("file_signature")
            identity = (
                path_key,
                file_signature if file_signature else record["timestamp"],
            )
            if identity in seen_signatures:
                continue

            previous = last_change_by_path.get(path_key)
            if previous is not None:
                try:
                    delta = abs(
                        (
                            datetime.fromisoformat(record["timestamp"])
                            - datetime.fromisoformat(previous["timestamp"])
                        ).total_seconds()
                    )
                except ValueError:
                    delta = None
                if (
                    delta is not None
                    and delta <= 5
                    and not file_signature
                    and not previous.get("file_signature")
                ):
                    continue

            unique_changes.append(record)
            seen_signatures.add(identity)
            last_change_by_path[path_key] = record
        if unique_changes != records:
            save_change_journal(unique_changes)
        return unique_changes
    except (OSError, json.JSONDecodeError, ValueError) as error:
        print(f"[ОШИБКА] Не удалось загрузить журнал изменений: {error}")
        return []

def save_change_journal(records):
    try:
        with open(CHANGE_JOURNAL_PATH, "w", encoding="utf-8") as journal_file:
            json.dump(
                {"records": records},
                journal_file,
                ensure_ascii=False,
                indent=2,
            )
    except OSError as error:
        print(f"[ОШИБКА] Не удалось сохранить журнал изменений: {error}")

def document_key(file_path):
    return os.path.normcase(os.path.abspath(file_path))

def normalized_stamp_mass(model_mass_g, stamp_mass_unit, constants):
    unit_factors = {
        constants.ksMUnGR: Decimal("1"),
        constants.ksMUnKG: Decimal("1000"),
        constants.ksMUnT: Decimal("1000000"),
    }
    if stamp_mass_unit not in unit_factors:
        raise ValueError(
            f"Не поддерживаются единицы массы штампа: {stamp_mass_unit}"
        )
    return Decimal(str(model_mass_g)) / unit_factors[stamp_mass_unit]

def parse_stamp_mass(value):
    normalized = re.sub(r"\s+", "", str(value)).replace(",", ".")
    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?", normalized):
        raise ValueError(f"В поле массы чертежа не число: {value!r}")
    try:
        return Decimal(normalized)
    except InvalidOperation as error:
        raise ValueError(f"Не удалось прочитать массу чертежа: {value!r}") from error

def drawing_is_current(
    model_mass,
    drawing_mass,
):
    """Match against the precision exposed by the drawing's mass property."""
    if model_mass is None or drawing_mass is None:
        return False
    property_resolution = Decimal(1).scaleb(drawing_mass.as_tuple().exponent)
    return abs(model_mass - drawing_mass) <= property_resolution / 2

def cached_model_change_mtime(cache_entry):
    model_change_mtime = cache_entry.get("last_geometry_mtime")
    if model_change_mtime is not None:
        return model_change_mtime

    last_change = cache_entry.get("last_change")
    if not isinstance(last_change, str):
        return None
    try:
        return datetime.fromisoformat(last_change).timestamp()
    except ValueError:
        logging.warning(
            "[WARNING] Invalid cached model-change timestamp: %s",
            last_change,
        )
        return None

class WindowsTrayIcon:
    CALLBACK_MESSAGE = win32con.WM_APP + 1
    MENU_OPEN = 1
    MENU_EXIT = 2

    def __init__(self, root, on_open, on_exit):
        self.root = root
        self.on_open = on_open
        self.on_exit = on_exit
        self.instance = win32gui.GetModuleHandle(None)
        self.icon_handle = win32gui.LoadImage(
            0,
            APPLICATION_ICON_PATH,
            win32con.IMAGE_ICON,
            16,
            16,
            win32con.LR_LOADFROMFILE,
        )
        self.class_name = (
            f"MonitorCADTrayWindow-{os.getpid()}-{id(self)}"
        )
        self.window = None
        self.icon_added = False
        self.closing = threading.Event()
        self.ready = threading.Event()
        self.startup_error = None
        self.thread = threading.Thread(
            target=self._run_message_loop,
            name="MonitorCAD tray icon",
            daemon=True,
        )
        self.thread.start()
        if not self.ready.wait(timeout=5):
            self.closing.set()
            if self.window is not None and win32gui.IsWindow(self.window):
                win32gui.PostMessage(self.window, win32con.WM_CLOSE, 0, 0)
            self.thread.join(timeout=2)
            raise RuntimeError(
                "Windows не создал окно значка MonitorCAD в области уведомлений."
            )
        if self.startup_error is not None:
            raise RuntimeError(
                f"Не удалось создать значок MonitorCAD: {self.startup_error}"
            ) from self.startup_error

    def _run_message_loop(self):
        try:
            window_class = win32gui.WNDCLASS()
            window_class.hInstance = self.instance
            window_class.lpszClassName = self.class_name
            window_class.lpfnWndProc = self._window_proc
            win32gui.RegisterClass(window_class)

            self.window = win32gui.CreateWindow(
                self.class_name,
                "MonitorCAD notification icon",
                win32con.WS_OVERLAPPED,
                0,
                0,
                0,
                0,
                0,
                0,
                self.instance,
                None,
            )
            self._add_icon()
            self.ready.set()
            win32gui.PumpMessages()
        except Exception as error:
            if not self.ready.is_set():
                self.startup_error = error
                self.ready.set()
            else:
                logging.exception(
                    "[ERROR] Ошибка окна значка MonitorCAD в области уведомлений."
                )
        finally:
            if self.icon_added:
                try:
                    win32gui.Shell_NotifyIcon(
                        win32gui.NIM_DELETE,
                        (self.window, 1),
                    )
                except win32gui.error:
                    logging.exception(
                        "[ERROR] Не удалось удалить значок MonitorCAD из области уведомлений."
                    )
                self.icon_added = False
            if self.window is not None and win32gui.IsWindow(self.window):
                win32gui.DestroyWindow(self.window)
            self.window = None
            if self.icon_handle:
                win32gui.DestroyIcon(self.icon_handle)
                self.icon_handle = None

    def _notification_data(self):
        flags = (
            win32gui.NIF_ICON
            | win32gui.NIF_MESSAGE
            | win32gui.NIF_TIP
        )
        return (
            self.window,
            1,
            flags,
            self.CALLBACK_MESSAGE,
            self.icon_handle,
            "MonitorCAD — Мониторинг листовых деталей",
        )

    def _add_icon(self):
        win32gui.Shell_NotifyIcon(
            win32gui.NIM_ADD,
            self._notification_data(),
        )
        self.icon_added = True

    def _window_proc(self, window, message, wparam, lparam):
        if message == self.CALLBACK_MESSAGE:
            if lparam in (win32con.WM_LBUTTONUP, win32con.WM_LBUTTONDBLCLK):
                self._post_to_root(self.on_open)
            elif lparam == win32con.WM_RBUTTONUP:
                self._show_menu()
            return 0
        if message == win32con.WM_COMMAND:
            command = win32gui.LOWORD(wparam)
            if command == self.MENU_OPEN:
                self._post_to_root(self.on_open)
            elif command == self.MENU_EXIT:
                self._post_to_root(self.on_exit)
            return 0
        if message == win32con.WM_DESTROY:
            if self.icon_added:
                win32gui.Shell_NotifyIcon(
                    win32gui.NIM_DELETE,
                    (window, 1),
                )
                self.icon_added = False
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(window, message, wparam, lparam)

    def _post_to_root(self, callback):
        if self.closing.is_set():
            return
        try:
            self.root.after(0, callback)
        except tk.TclError:
            self.closing.set()

    def _show_menu(self):
        menu = win32gui.CreatePopupMenu()
        try:
            win32gui.AppendMenu(
                menu,
                win32con.MF_STRING,
                self.MENU_OPEN,
                "Открыть MonitorCAD",
            )
            win32gui.AppendMenu(
                menu,
                win32con.MF_STRING,
                self.MENU_EXIT,
                "Выход",
            )
            x, y = win32gui.GetCursorPos()
            win32gui.SetForegroundWindow(self.window)
            win32gui.TrackPopupMenu(
                menu,
                win32con.TPM_RIGHTBUTTON,
                x,
                y,
                0,
                self.window,
                None,
            )
            win32gui.PostMessage(self.window, win32con.WM_NULL, 0, 0)
        finally:
            win32gui.DestroyMenu(menu)

    def close(self):
        self.closing.set()
        if self.window is not None and win32gui.IsWindow(self.window):
            win32gui.PostMessage(self.window, win32con.WM_CLOSE, 0, 0)
        self.thread.join(timeout=2)
        if self.thread.is_alive():
            logging.error(
                "[ERROR] Поток значка MonitorCAD не завершился после запроса выхода."
            )

class MonitorCADApp(_MonitorCADBaseApp):
    def __init__(self, root):
        super().__init__(root)
        self.cache = load_event_cache()
        self.cache_lock = threading.RLock()
        self.event_queue = queue.Queue()
        self.document_sinks = {}
        self.internal_documents = set()
        self.known_open_document_keys = set()
        self.ignored_document_keys = set()
        self.finished_document_keys = set()
        self.deferred_document_retries = {}
        self.document_calculation_retry_counts = {}
        self.monitor_options = {}
        self.journal_records = load_change_journal()
        self.journal_window = None
        self.journal_tree = None
        self.observed_file_signatures = {}
        self.tray_icon = None
        self.root.bind("<Unmap>", self._on_window_minimized, add="+")
        self.root.protocol("WM_DELETE_WINDOW", self._quit_app)
        self.root.after(0, self._initialize_tray_icon)
        self.cache_count_var = tk.StringVar()
        self.cache_count_label = ttk.Label(
            self.status_label.master,
            textvariable=self.cache_count_var,
            font=("Arial", 9),
        )
        self.cache_count_label.pack(side=tk.RIGHT, padx=(8, 0))
        self.journal_button = ttk.Button(
            self.status_label.master,
            text="Журнал",
            command=self.show_journal,
        )
        self.journal_button.pack(side=tk.RIGHT, padx=(8, 0))
        self.scan_running = False
        button_frame = self.start_btn.master
        self.start_btn.destroy()
        self.stop_btn.destroy()
        self.scan_btn = ttk.Button(
            button_frame,
            text="Проверить изменения",
            command=self.request_manual_scan,
            state=tk.DISABLED,
            width=24,
        )
        self.scan_btn.pack(side=tk.LEFT)
        self.send_list_btn = ttk.Button(
            self.clear_selected_btn.master,
            text="Отправить список",
            command=self.send_changed_files_list,
        )
        self.send_list_btn.pack(side=tk.RIGHT, padx=(0, 8))

        for setting_name in (
            "track_detali",
            "track_chertezh",
            "track_listovaya",
            "track_unfolding",
        ):
            if setting_name in self.settings:
                getattr(self, setting_name).set(self.settings[setting_name])

        auto_frame = self.auto_check.master
        self.auto_check.destroy()
        self.interval_entry.config(state="disabled")
        for widget in auto_frame.winfo_children():
            if isinstance(widget, ttk.Label) and widget.cget("text").startswith(
                "Интервал"
            ):
                widget.destroy()
        self.interval_entry.destroy()
        self.auto_mode.set(True)
        auto_frame.destroy()

        self._migrate_current_changes_to_journal()
        self.update_cache_count()
        self.root.after(500, self._refresh_counts)
        if self.monitor_dir.get():
            self.root.after(500, self.start_monitoring)

    def _initialize_tray_icon(self):
        try:
            self.root.iconbitmap(APPLICATION_ICON_PATH)
            self.tray_icon = WindowsTrayIcon(
                self.root,
                self._restore_from_tray,
                self._quit_app,
            )
        except Exception as error:
            logging.exception(
                "[ERROR] Не удалось создать значок MonitorCAD в области уведомлений."
            )
            self.set_status(
                f"Не удалось создать значок в области уведомлений: {error}"
            )

    def _on_window_minimized(self, event):
        if event.widget is self.root and self.root.state() == "iconic":
            self.root.after_idle(self.root.withdraw)

    def _restore_from_tray(self):
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        self.root.focus_force()

    def _quit_app(self):
        if self.tray_icon is not None:
            self.tray_icon.close()
            self.tray_icon = None
        self.stop_monitoring()
        self.root.destroy()

    def _refresh_counts(self):
        try:
            self.update_cache_count()
            if self.journal_window is not None and self.journal_window.winfo_exists():
                self._populate_journal_tree()
            self.root.after(500, self._refresh_counts)
        except tk.TclError:
            return

    def update_cache_count(self):
        with self.cache_lock:
            count = sum(
                1
                for entry in self.cache.values()
                if self._cache_has_baseline(entry)
            )
        self.cache_count_var.set(f"Деталей в кэше: {count}")

    def _migrate_current_changes_to_journal(self):
        known_entries = {
            document_key(record["full_path"])
            for record in self.journal_records
            if record["full_path"]
        }
        migrated = False
        for filename, (file_path, part_type, status) in self.changed_files.items():
            if (
                not file_path
                or status != "ИЗМЕНЕН"
                or os.path.splitext(file_path)[1].lower()
                not in MODEL_EXTENSIONS
            ):
                continue
            model_path = file_path
            with self.cache_lock:
                cache_entry = self.cache.get(document_key(model_path), {})
            timestamp = cache_entry.get("last_change")
            if timestamp is None and os.path.isfile(file_path):
                timestamp = datetime.fromtimestamp(
                    os.path.getmtime(file_path)
                ).isoformat(timespec="seconds")
            if timestamp is None:
                timestamp = datetime.now().isoformat(timespec="seconds")
            identity = document_key(file_path)
            if identity in known_entries:
                continue
            self.journal_records.append(
                {
                    "timestamp": timestamp,
                    "filename": filename,
                    "part_type": part_type,
                    "status": status,
                    "full_path": file_path,
                }
            )
            known_entries.add(identity)
            migrated = True
        if migrated:
            save_change_journal(self.journal_records)

    def _populate_journal_tree(self):
        if self.journal_tree is None or not self.journal_tree.winfo_exists():
            return
        for item in self.journal_tree.get_children():
            self.journal_tree.delete(item)

        for record in sorted(
            self.journal_records,
            key=lambda item: item["timestamp"],
            reverse=True,
        ):
            timestamp = record["timestamp"]
            try:
                date_text = datetime.fromisoformat(timestamp).strftime(
                    "%d.%m.%Y %H:%M:%S"
                )
            except ValueError:
                date_text = timestamp
            self.journal_tree.insert(
                "",
                tk.END,
                values=(
                    date_text,
                    record["filename"],
                ),
            )

    def show_journal(self):
        if self.journal_window is not None and self.journal_window.winfo_exists():
            self.journal_window.deiconify()
            self.journal_window.lift()
            self._populate_journal_tree()
            return

        window = tk.Toplevel(self.root)
        window.title("Журнал изменений MonitorCAD")
        window.geometry("850x420")
        window.transient(self.root)
        self.journal_window = window

        frame = ttk.Frame(window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        columns = ("timestamp", "filename")
        self.journal_tree = ttk.Treeview(
            frame,
            columns=columns,
            show="headings",
        )
        self.journal_tree.heading("timestamp", text="Дата изменения")
        self.journal_tree.heading("filename", text="Файл")
        self.journal_tree.column("timestamp", width=150, anchor=tk.CENTER)
        self.journal_tree.column("filename", width=650, anchor=tk.W)
        self.journal_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(
            frame,
            orient=tk.VERTICAL,
            command=self.journal_tree.yview,
        )
        self.journal_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        window.protocol("WM_DELETE_WINDOW", self._close_journal)
        self._populate_journal_tree()

    def _close_journal(self):
        if self.journal_window is not None:
            self.journal_window.destroy()
        self.journal_window = None
        self.journal_tree = None

    def browse_folder(self):
        super().browse_folder()
        if self.monitor_dir.get() and not self.is_running:
            self.root.after(100, self.start_monitoring)

    def request_manual_scan(self):
        if not self.is_running:
            self.status_line_var.set(
                "Сначала подключитесь к КОМПАСу — мониторинг запускается автоматически."
            )
            return
        if self.scan_running:
            return
        self.scan_running = True
        self.scan_btn.config(state=tk.DISABLED)
        self.set_status("Ручная сверка базы ожидает COM-поток...")
        self.event_queue.put(("manual_scan", None))

    def send_changed_files_list(self):
        recipients = []
        invalid_recipients = []
        for recipient in re.split(r"[;,]", self.mail_recipients.get()):
            recipient = recipient.strip()
            if not recipient:
                continue
            if not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+", recipient):
                invalid_recipients.append(recipient)
                continue
            if recipient.casefold() not in {
                existing.casefold() for existing in recipients
            }:
                recipients.append(recipient)

        if invalid_recipients:
            messagebox.showerror(
                "MonitorCAD",
                "Проверьте адреса рассылки: "
                + ", ".join(invalid_recipients),
                parent=self.root,
            )
            return
        if not recipients:
            messagebox.showerror(
                "MonitorCAD",
                "Укажите хотя бы один адрес в поле «Адреса рассылки».",
                parent=self.root,
            )
            return

        self.save_settings()
        rows_by_name = {}
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if len(values) < 3:
                continue
            filename, _, status = (str(value) for value in values[:3])
            if status != "ИЗМЕНЕН":
                continue
            detail_name = os.path.splitext(filename)[0]
            rows_by_name.setdefault(
                detail_name.casefold(),
                (detail_name, status),
            )
        rows = list(rows_by_name.values())

        if not rows:
            messagebox.showinfo(
                "MonitorCAD",
                "В списке «Изменённые файлы» нет документов со статусом «ИЗМЕНЕН».",
                parent=self.root,
            )
            return

        table_rows = "".join(
            "<tr>"
            + "".join(
                f"<td>{html.escape(value)}</td>"
                for value in row
            )
            + "</tr>"
            for row in rows
        )
        table = (
            '<table style="border-collapse:collapse;width:100%">'
            "<thead><tr>"
            '<th style="text-align:center">Деталь</th>'
            '<th style="text-align:center">Статус</th>'
            "</tr></thead>"
            f"<tbody>{table_rows}</tbody>"
            "</table>"
        )
        fragment_before = "<html><body><!--StartFragment-->"
        fragment_after = "<!--EndFragment--></body></html>"
        fragment_bytes = table.encode("utf-8")
        before_bytes = fragment_before.encode("utf-8")
        after_bytes = fragment_after.encode("utf-8")
        clipboard_header = (
            b"Version:0.9\r\n"
            b"StartHTML:0000000000\r\n"
            b"EndHTML:0000000000\r\n"
            b"StartFragment:0000000000\r\n"
            b"EndFragment:0000000000\r\n"
        )
        start_html = len(clipboard_header)
        start_fragment = start_html + len(before_bytes)
        end_fragment = start_fragment + len(fragment_bytes)
        end_html = end_fragment + len(after_bytes)
        clipboard_header = (
            f"Version:0.9\r\n"
            f"StartHTML:{start_html:010d}\r\n"
            f"EndHTML:{end_html:010d}\r\n"
            f"StartFragment:{start_fragment:010d}\r\n"
            f"EndFragment:{end_fragment:010d}\r\n"
        ).encode("ascii")
        html_clipboard_data = (
            clipboard_header
            + before_bytes
            + fragment_bytes
            + after_bytes
        )
        plain_text = "\n".join(
            ["\t".join(("Деталь", "Статус"))]
            + ["\t".join(row) for row in rows]
        )
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(
                    win32clipboard.RegisterClipboardFormat("HTML Format"),
                    html_clipboard_data,
                )
                win32clipboard.SetClipboardData(
                    win32con.CF_UNICODETEXT,
                    plain_text,
                )
            finally:
                win32clipboard.CloseClipboard()
        except Exception as error:
            logging.exception(
                "[ERROR] Не удалось скопировать список изменений в буфер обмена."
            )
            messagebox.showerror(
                "MonitorCAD",
                f"Не удалось подготовить текст письма: {error}",
                parent=self.root,
            )
            return

        compose_url = "https://e.mail.ru/compose/?" + urlencode(
            {
                "to": ",".join(recipients),
                "subject": "MonitorCAD — список изменений",
            }
        )
        try:
            browser_opened = webbrowser.open(compose_url, new=2)
        except OSError as error:
            logging.exception(
                "[ERROR] Не удалось открыть страницу создания письма Mail.ru."
            )
            messagebox.showerror(
                "MonitorCAD",
                f"Не удалось открыть Mail.ru: {error}",
                parent=self.root,
            )
            return

        if not browser_opened:
            logging.error(
                "[ERROR] Не удалось открыть страницу создания письма Mail.ru."
            )
            messagebox.showerror(
                "MonitorCAD",
                "Не удалось открыть браузер для создания письма Mail.ru.",
                parent=self.root,
            )
            return

        self.set_status(
            "Таблица изменений скопирована. Вставьте её в письмо Mail.ru."
        )

    def _manual_scan_finished(self):
        self.scan_running = False
        if self.is_running:
            self.scan_btn.config(state=tk.NORMAL)

    def _scan_cached_database(self, kompas_app, event_module):
        with self.cache_lock:
            entries = [
                (key, dict(entry))
                for key, entry in self.cache.items()
                if self._cache_has_baseline(entry)
            ]

        total = len(entries)
        inspected = 0
        changed_files = 0
        changed_geometry = 0
        errors = 0
        skipped_open = 0
        logging.info(
            "[INFO] Manual cache scan started; cached model count: %d",
            total,
        )

        try:
            for index, (key, cached_entry) in enumerate(entries, start=1):
                if self.stop_event.is_set():
                    break

                filename = os.path.basename(key)
                self.set_status(
                    f"Сверка базы ({index}/{total}): {filename}"
                )
                self.set_progress(int(index / total * 100) if total else 0)

                try:
                    if not os.path.isfile(display_path):
                        errors += 1
                        logging.warning(
                            "[WARNING] Cached model file is missing: %s", key
                        )
                        continue

                    stat = os.stat(display_path)
                    current_hash = get_file_hash(display_path)
                    if current_hash is None:
                        errors += 1
                        logging.error(
                            "[ERROR] Could not calculate file hash: %s", key
                        )
                        continue

                    markers_changed = (
                        cached_entry.get("mtime") != stat.st_mtime
                        or cached_entry.get("size") != stat.st_size
                        or cached_entry.get("hash") != current_hash
                    )
                    inspected += 1
                    if not markers_changed:
                        continue

                    if self._path_is_open(kompas_app, key):
                        skipped_open += 1
                        logging.info(
                            "[INFO] Manual scan skipped an open document: %s",
                            key,
                        )
                        continue

                    changed_files += 1
                    logging.info(
                        "[INFO] Changed file markers found during manual scan: %s",
                        key,
                    )
                    self.internal_documents.add(key)
                    try:
                        mch_data, _, part_type = calculate_mch(
                            kompas_app,
                            key,
                        )
                    finally:
                        self.internal_documents.discard(key)

                    if mch_data is None:
                        errors += 1
                        self._schedule_change(
                            filename,
                            key,
                            cached_entry.get("part_type") or "Неизвестно",
                            "ОШИБКА",
                        )
                        logging.error(
                            "[ERROR] MCH calculation failed during manual scan: %s",
                            key,
                        )
                        continue

                    geometry_changed, _ = geometry_actually_changed(
                        cached_entry["last_mch"],
                        mch_data,
                    )
                    type_changed = cached_entry.get("part_type") != part_type
                    updated_entry = dict(cached_entry)
                    updated_entry.update(
                        {
                            "mtime": stat.st_mtime,
                            "size": stat.st_size,
                            "hash": current_hash,
                            "mch_calculated": True,
                            "last_check": datetime.now().isoformat(),
                        }
                    )
                    if geometry_changed or type_changed:
                        changed_geometry += 1
                        updated_entry["last_mch"] = mch_data
                        updated_entry["part_type"] = part_type
                        updated_entry["last_change"] = datetime.now().isoformat()
                        updated_entry["last_geometry_mtime"] = stat.st_mtime
                        updated_entry["change_count"] = (
                            cached_entry.get("change_count", 0) + 1
                        )
                        self._schedule_change(
                            filename,
                            key,
                            part_type or "Неизвестно",
                            "ИЗМЕНЕН",
                        )
                        if (
                            part_type == "Листовая деталь"
                            and self.monitor_options["track_unfolding"]
                        ):
                            self.root.after(
                                0,
                                self.check_unfolding_for_model,
                                key,
                                part_type,
                                updated_entry.get("drawing_path"),
                            )
                        logging.info(
                            "[INFO] MCH changed during manual scan: %s",
                            key,
                        )
                    else:
                        logging.info(
                            "[INFO] File markers changed but MCH did not: %s",
                            key,
                        )

                    self._check_drawing(
                        key,
                        updated_entry,
                        kompas_app=kompas_app,
                    )
                    with self.cache_lock:
                        self.cache[key] = updated_entry
                        save_event_cache(self.cache)
                except Exception as error:
                    errors += 1
                    logging.exception(
                        "[ERROR] Manual scan failed for %s: %s",
                        key,
                        error,
                    )
                finally:
                    self._process_pending_events(kompas_app, event_module)
        finally:
            self.set_progress(0)
            self.root.after(0, self._manual_scan_finished)
            summary = (
                f"Сверка завершена: проверено {inspected} из {total}, "
                f"признаки изменения {changed_files}, изменились МЦХ "
                f"{changed_geometry}, ошибок {errors}, открытых пропущено "
                f"{skipped_open}."
            )
            self.set_status(summary)
            logging.info("[INFO] %s", summary)

    def _process_pending_events(self, kompas_app, event_module):
        while True:
            try:
                event_name, event_value = self.event_queue.get_nowait()
            except queue.Empty:
                return

            if event_name == "manual_scan":
                logging.warning(
                    "[WARNING] Ignoring a duplicate manual scan request."
                )
            else:
                self._process_event(
                    kompas_app,
                    event_module,
                    event_name,
                    event_value,
                )

    def restore_changed_list(self):
        records, total_changes = load_event_changed_list()
        selected_dir = self.monitor_dir.get()

        for record in records:
            full_path = record.get("full_path", "")
            filename = (
                filename_with_disk_case(full_path)
                if full_path
                else record.get("filename", "")
            )
            part_type = record.get("part_type", "")
            record_dir = (
                self.unfolding_dir.get()
                if part_type == "Развертка DXF"
                else selected_dir
            )
            if selected_dir and not path_belongs_to_directory(
                full_path, record_dir
            ):
                continue

            status = record.get("status", "")
            self.tree.insert(
                "",
                tk.END,
                values=(filename, part_type, status),
            )
            self.changed_files[filename] = (full_path, part_type, status)

        self.total_changes = len(self.tree.get_children())
        self.refresh_tree_groups()

    def persist_changed_list(self):
        records = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if len(values) < 3:
                continue
            filename, part_type, status = values
            full_path = self.changed_files.get(filename, ("", "", ""))[0]
            records.append(
                {
                    "timestamp": datetime.now().strftime("%H:%M:%S"),
                    "filename": filename,
                    "part_type": part_type,
                    "status": status,
                    "full_path": full_path,
                }
            )
        save_event_changed_list(records, self.total_changes)

    def save_event_cache(self):
        with self.cache_lock:
            save_event_cache(self.cache)

    def start_monitoring(self):
        monitor_dir = self.monitor_dir.get().strip()
        if not monitor_dir:
            messagebox.showerror("Ошибка", "Выберите папку для мониторинга!")
            return
        if not os.path.isdir(monitor_dir):
            messagebox.showerror("Ошибка", "Указанная папка не существует!")
            return
        if self.is_running:
            return
        if self.monitor_thread is not None and self.monitor_thread.is_alive():
            self.set_status("Ожидание завершения предыдущего потока COM...")
            return

        self.cache = load_event_cache()
        self.update_cache_count()
        self.monitor_options = {
            "monitor_dir": os.path.normpath(monitor_dir),
            "unfolding_dir": os.path.normpath(self.unfolding_dir.get().strip())
            if self.unfolding_dir.get().strip()
            else "",
            "track_detali": self.track_detali.get(),
            "track_chertezh": self.track_chertezh.get(),
            "track_listovaya": self.track_listovaya.get(),
            "track_unfolding": self.track_unfolding.get(),
        }
        self.is_running = True
        self.stop_event.clear()
        self.scan_btn.config(state=tk.DISABLED)
        self.status_label.config(text="ПОДКЛЮЧЕНИЕ")
        self.indicator_canvas.itemconfig(self.indicator_oval, fill="orange")
        self.set_status("Подключение к КОМПАС-3D и установка слежения за документами...")
        self.set_progress(0)
        self.save_settings()

        self.monitor_thread = threading.Thread(
            target=self.monitor_loop,
            daemon=True,
        )
        self.monitor_thread.start()

    def stop_monitoring(self):
        if not self.is_running:
            return
        self.is_running = False
        self.stop_event.set()
        self.scan_btn.config(state=tk.DISABLED)
        self.status_label.config(text="ОСТАНОВЛЕН")
        self.indicator_canvas.itemconfig(self.indicator_oval, fill="gray")
        self.set_status("Отслеживание событий КОМПАС остановлено.")

    def clear_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.update_clear_selected_button()
        self.changed_files = {}
        self.total_changes = 0
        save_event_changed_list([], 0)
        self.status_line_var.set("Список очищен.")

    def _monitoring_failed(self, message):
        self.is_running = False
        self.stop_event.set()
        self.scan_btn.config(state=tk.DISABLED)
        self.status_label.config(text="ОШИБКА")
        self.indicator_canvas.itemconfig(self.indicator_oval, fill="red")
        self.status_line_var.set(message)

    def _waiting_for_kompas(self):
        if not self.is_running:
            return
        self.status_label.config(text="ОЖИДАЕТ КОМПАС")
        self.indicator_canvas.itemconfig(self.indicator_oval, fill="orange")
        self.set_status(
            "Откройте КОМПАС-3D — MonitorCAD автоматически подключится."
        )

    def _event_path_is_tracked(self, file_path):
        extension = os.path.splitext(file_path)[1].lower()
        if extension not in MODEL_EXTENSIONS:
            return False
        return path_belongs_to_directory(
            file_path,
            self.monitor_options["monitor_dir"],
        )

    def _cache_has_baseline(self, entry):
        return (
            isinstance(entry, dict)
            and entry.get("mch_calculated") is True
            and isinstance(entry.get("last_mch"), dict)
            and entry.get("mass_is_exact") is True
        )

    def _part_type_is_tracked(self, part_type):
        if part_type == "Листовая деталь":
            return self.monitor_options["track_listovaya"]
        return self.monitor_options["track_detali"]

    def _defer_document_registration(self, key, document):
        """Планирует отложенную повторную попытку расчёта МЦХ."""
        retry_count = self.document_calculation_retry_counts.get(key, 0)
        if retry_count >= DOCUMENT_CALCULATION_MAX_RETRIES:
            self.deferred_document_retries.pop(key, None)
            logging.error(
                "[ERROR] Исчерпаны повторные попытки расчёта МЦХ (%s): %s",
                DOCUMENT_CALCULATION_MAX_RETRIES,
                key,
            )
            return

        retry_count += 1
        self.document_calculation_retry_counts[key] = retry_count
        self.deferred_document_retries[key] = (
            time.monotonic() + DOCUMENT_POLL_RETRY_INTERVAL,
            document,
        )
        logging.warning(
            "[WARNING] Повторная попытка расчёта МЦХ %s/%s через %s сек.: %s",
            retry_count,
            DOCUMENT_CALCULATION_MAX_RETRIES,
            int(DOCUMENT_POLL_RETRY_INTERVAL),
            key,
        )

    def _register_open_document(self, kompas_app, document, event_module):
        """Подписывается на события документа.

        Возвращает нормализованный ключ документа, если он был успешно
        обработан, иначе None.
        """
        try:
            file_path = document.PathName
        except Exception as error:
            print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось получить путь документа: {error}")
            return None

        if not file_path:
            return None

        if os.path.basename(file_path).lower() in INTERNAL_DOCUMENT_NAMES:
            return None

        extension = os.path.splitext(file_path)[1].lower()
        is_drawing = extension in DRAWING_EXTENSIONS
        if is_drawing:
            if (
                not self.monitor_options["track_chertezh"]
                or not path_belongs_to_directory(
                    file_path,
                    self.monitor_options["monitor_dir"],
                )
            ):
                return None
        elif not self._event_path_is_tracked(file_path):
            return None

        key = document_key(file_path)
        if key in self.internal_documents:
            return None
        if key in self.finished_document_keys:
            # Документ уже полностью обработан после закрытия — повторно
            # открывать его не нужно, пока он реально не закроется.
            return None

        if key not in self.document_sinks:
            owner = self

            class DocumentEvents(event_module.ksDocumentFileNotify):
                def __init__(self, dispatch, document_path):
                    self.document_key = key
                    self.document_path = document_path
                    super().__init__(dispatch)

                def OnCloseDocument(self):
                    if self.document_key not in owner.internal_documents:
                        owner.event_queue.put(
                            ("closed", self.document_key, self.document_path)
                        )

                def OnSaveDocument(self):
                    owner.event_queue.put(
                        ("saved", self.document_key, self.document_path)
                    )

            try:
                sink = DocumentEvents(document, file_path)
            except Exception as error:
                print(
                    f"[ОШИБКА] Не удалось подписаться на сохранение {file_path}: {error}"
                )
                return None
            self.document_sinks[key] = (document, sink)
            if is_drawing:
                signature = self._get_file_signature(file_path)
                if signature is not None:
                    self.observed_file_signatures[key] = signature
                logging.info(
                    "[INFO] Подписка на сохранение чертежа установлена: %s",
                    file_path,
                )

        if is_drawing:
            return key

        if key in self.deferred_document_retries:
            return None
        if (
            self.document_calculation_retry_counts.get(key, 0)
            >= DOCUMENT_CALCULATION_MAX_RETRIES
        ):
            return None

        with self.cache_lock:
            entry = self.cache.get(key)
            has_baseline = self._cache_has_baseline(entry)
            if has_baseline and not self._part_type_is_tracked(
                entry.get("part_type")
            ):
                self.document_calculation_retry_counts.pop(key, None)
                self.deferred_document_retries.pop(key, None)
                sink_record = self.document_sinks.pop(key, None)
                if sink_record is not None:
                    sink_record[1].close()
                return None

        if has_baseline:
            return key

        self.set_status(f"Первичный расчёт МЦХ: {os.path.basename(file_path)}")
        try:
            mch_data, _, part_type = calculate_mch(kompas_app, file_path)
            if mch_data is None:
                self._schedule_change(
                    os.path.basename(file_path),
                    file_path,
                    "Неизвестно",
                    "ОШИБКА",
                )
                print(
                    f"[ОШИБКА] Не удалось записать исходные МЦХ для {file_path}."
                )
                self._defer_document_registration(key, document)
                return None

            if not self._part_type_is_tracked(part_type):
                was_retry = self.document_calculation_retry_counts.pop(
                    key,
                    0,
                )
                self.deferred_document_retries.pop(key, None)
                sink_record = self.document_sinks.pop(key, None)
                if sink_record is not None:
                    sink_record[1].close()
                self.ignored_document_keys.add(key)
                if was_retry:
                    self.root.after(
                        0,
                        self._clear_initial_calculation_error,
                        key,
                    )
                return None

            stat = os.stat(file_path)
            drawing_path = find_drawing_for_model(file_path)
            drawing_mtime = (
                os.stat(drawing_path).st_mtime if drawing_path else None
            )
            new_entry = {
                "mtime": stat.st_mtime,
                "size": stat.st_size,
                "hash": get_file_hash(file_path),
                "last_mch": mch_data,
                "mch_calculated": True,
                "mass_is_exact": True,
                "last_check": datetime.now().isoformat(),
                "drawing_path": drawing_path,
                "drawing_mtime": drawing_mtime,
                "part_type": part_type,
            }
            with self.cache_lock:
                self.cache[key] = new_entry
                save_event_cache(self.cache)
            print(
                f"[ИНФО] Исходные МЦХ сохранены в событийный кэш: {file_path}"
            )
            self.set_status(
                f"Исходная точка сохранена: {os.path.basename(file_path)}"
            )
            if self.document_calculation_retry_counts.pop(key, 0):
                self.deferred_document_retries.pop(key, None)
                logging.info(
                    "[INFO] Повторный расчёт МЦХ завершился успешно: %s",
                    file_path,
                )
                self.root.after(
                    0,
                    self._clear_initial_calculation_error,
                    key,
                )
        except Exception as error:
            print(f"[ОШИБКА] Не удалось обработать открытый документ {file_path}: {error}")
            print(traceback.format_exc())
            self._defer_document_registration(key, document)
            return None
        return key

    def _process_due_document_retries(self, kompas_app, event_module):
        """Запускает отложенные попытки в COM-потоке."""
        now = time.monotonic()
        due_keys = [
            key
            for key, (deadline, _) in self.deferred_document_retries.items()
            if deadline <= now
        ]
        for key in due_keys:
            retry_record = self.deferred_document_retries.pop(key, None)
            if retry_record is None:
                continue
            _, document = retry_record
            try:
                if document_key(document.PathName) != key:
                    raise ValueError("Путь открытого документа изменился.")
            except Exception as error:
                self.document_calculation_retry_counts.pop(key, None)
                logging.warning(
                    "[WARNING] Повтор расчёта отменён: документ закрыт "
                    "или недоступен (%s): %s",
                    key,
                    error,
                )
                continue

            retry_count = self.document_calculation_retry_counts.get(key, 0)
            logging.info(
                "[INFO] Выполняется повторная попытка расчёта МЦХ %s/%s: %s",
                retry_count,
                DOCUMENT_CALCULATION_MAX_RETRIES,
                key,
            )
            registered_key = self._register_open_document(
                kompas_app,
                document,
                event_module,
            )
            if registered_key is not None:
                self.known_open_document_keys.add(registered_key)

    def _clear_initial_calculation_error(self, key):
        """Убирает из списка ошибку первичного расчёта после успешного повтора."""
        filenames_to_remove = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if not values or len(values) < 3 or values[2] != "ОШИБКА":
                continue
            filename = values[0]
            full_path = self.changed_files.get(
                filename,
                ("", "", ""),
            )[0]
            if full_path and document_key(full_path) == key:
                filenames_to_remove.append(filename)
                self.tree.delete(item)

        for filename in filenames_to_remove:
            self.changed_files.pop(filename, None)
        if filenames_to_remove:
            self.total_changes = max(
                0,
                self.total_changes - len(filenames_to_remove),
            )
            self.update_clear_selected_button()
            self.persist_changed_list()

    def _register_open_documents(self, kompas_app, event_module):
        try:
            documents = list(kompas_app.Documents)
        except Exception as error:
            print(f"[ОШИБКА] Не удалось получить список документов КОМПАС: {error}")
            print(traceback.format_exc())
            return
        for document in documents:
            registered_key = self._register_open_document(
                kompas_app,
                document,
                event_module,
            )
            if registered_key is not None:
                self.known_open_document_keys.add(registered_key)

    def _poll_open_documents(self, kompas_app, event_module):
        """Опрашивает app.Documents и регистрирует документы, для которых
        не пришло событие открытия.

        Работает как подстраховка к событию OnOpenDocument: уже известные
        документы повторно не обрабатываются.
        """
        try:
            documents = list(kompas_app.Documents)
        except Exception as error:
            print(f"[ОШИБКА] Не удалось опросить документы КОМПАС: {error}")
            return

        current_keys = set()
        for document in documents:
            try:
                file_path = document.PathName
            except Exception:
                continue
            if not file_path:
                continue
            key = document_key(file_path)
            current_keys.add(key)

            if key in self.known_open_document_keys:
                if key in self.document_sinks:
                    continue
                # Подписка на события потеряна — переподписываемся.
                self.known_open_document_keys.discard(key)

            registered_key = self._register_open_document(
                kompas_app,
                document,
                event_module,
            )
            if registered_key is None:
                continue
            self.known_open_document_keys.add(registered_key)
            logging.info(
                "[INFO] Документ найден опросом app.Documents: %s",
                file_path,
            )

        # Документы, которые закрылись (событие могло не прийти).
        for stale_key in self.known_open_document_keys - current_keys:
            self.known_open_document_keys.discard(stale_key)
            sink_record = self.document_sinks.pop(stale_key, None)
            if sink_record is not None:
                try:
                    sink_record[1].close()
                except Exception as error:
                    print(
                        f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось снять COM-подписку: {error}"
                    )

    def _path_is_open(self, kompas_app, target_key):
        for document in kompas_app.Documents:
            try:
                path = document.PathName
            except Exception:
                continue
            if path and document_key(path) == target_key:
                return True
        return False

    def set_unfolding_record(self, model_path, filename, full_path, status):
        super().set_unfolding_record(model_path, filename, full_path, status)
        self._append_journal_record(
            filename,
            full_path,
            "Развертка DXF",
            status,
        )

    def _get_file_signature(self, file_path):
        try:
            stat = os.stat(file_path)
        except OSError as error:
            logging.warning(
                "[WARNING] Не удалось прочитать файл для журнала %s: %s",
                file_path,
                error,
            )
            return None
        file_hash = get_file_hash(file_path)
        if file_hash is None:
            logging.warning(
                "[WARNING] Не удалось вычислить хэш для записи в журнал: %s",
                file_path,
            )
            return None
        return f"{stat.st_size}:{file_hash}"

    def _record_saved_drawing_change(self, drawing_path):
        key = document_key(drawing_path)
        signature = self._get_file_signature(drawing_path)
        if signature is None:
            return
        previous_signature = self.observed_file_signatures.get(key)
        self.observed_file_signatures[key] = signature
        if previous_signature is None or previous_signature == signature:
            return
        self._append_journal_record(
            os.path.basename(drawing_path),
            drawing_path,
            "Чертёж",
            "ИЗМЕНЕН",
        )

    def _append_journal_record(self, filename, file_path, part_type, status):
        if status != "ИЗМЕНЕН" or not os.path.isfile(file_path):
            return

        file_signature = self._get_file_signature(file_path)
        if file_signature is None:
            return
        path_key = document_key(file_path)
        for record in reversed(self.journal_records):
            if document_key(record["full_path"]) == path_key:
                if (
                    file_signature
                    and record.get("file_signature") == file_signature
                ):
                    return
                break

        self.journal_records.append(
            {
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "filename": filename,
                "part_type": part_type,
                "status": status,
                "full_path": file_path,
                "file_signature": file_signature,
            }
        )
        save_change_journal(self.journal_records)
        if self.journal_window is not None and self.journal_window.winfo_exists():
            self._populate_journal_tree()

    def _record_change(
        self,
        filename,
        file_path,
        part_type,
        status,
        record_journal=True,
    ):
        filename = filename_with_disk_case(file_path) if file_path else filename
        existing = self.changed_files.get(filename)
        if existing is None:
            self.total_changes += 1
        self.changed_files[filename] = (file_path, part_type, status)
        self.add_to_list(filename, part_type, status)
        if record_journal and file_path:
            self._append_journal_record(filename, file_path, part_type, status)

    def _schedule_change(
        self,
        filename,
        file_path,
        part_type,
        status,
        record_journal=True,
    ):
        self.root.after(
            0,
            self._record_change,
            filename,
            file_path,
            part_type,
            status,
            record_journal,
        )

    def _check_drawing(
        self,
        model_path,
        cache_entry,
        record_journal=False,
        kompas_app=None,
    ):
        if not self.monitor_options["track_chertezh"]:
            return
        drawing_path = find_drawing_for_model(model_path)
        if drawing_path is None:
            drawing_filename = os.path.splitext(os.path.basename(model_path))[0] + ".cdw"
            drawing_path = os.path.join(os.path.dirname(model_path), drawing_filename)
            self._schedule_change(
                drawing_filename,
                drawing_path,
                "Чертёж",
                "НЕТ ЧЕРТЕЖА",
                record_journal,
            )
            cache_entry["drawing_path"] = None
            cache_entry["drawing_mtime"] = None
            return

        cache_entry["drawing_path"] = drawing_path
        if kompas_app is None:
            return
        if self._path_is_open(kompas_app, document_key(drawing_path)):
            logging.info(
                "[INFO] Проверка актуальности чертежа отложена до закрытия: %s",
                drawing_path,
            )
            return

        try:
            model_stat = os.stat(model_path)
            drawing_stat = os.stat(drawing_path)
            model_change_mtime = cached_model_change_mtime(cache_entry)
            if model_change_mtime is None:
                logging.info(
                    "[INFO] Актуальность чертежа не пересматривается: "
                    "с момента исходного МЦХ нет зафиксированного изменения детали: %s",
                    drawing_path,
                )
                return

            drawing_mass, mass_unit = self._read_drawing_mass(
                kompas_app,
                drawing_path,
            )
            model_mass = normalized_stamp_mass(
                cache_entry["last_mch"]["Масса (г)"],
                mass_unit,
                self._kompas_constants(),
            )
            cache_entry["drawing_mtime"] = drawing_stat.st_mtime
            cache_entry["drawing_mass"] = str(drawing_mass)
            cache_entry["drawing_mass_unit"] = mass_unit

            is_current = drawing_is_current(
                model_mass,
                drawing_mass,
            )
            status = "ИЗМЕНЕН" if is_current else "ТРЕБУЕТ ПЕРЕСТРОЙКИ"
            self._schedule_change(
                os.path.basename(drawing_path),
                drawing_path,
                "Чертёж",
                status,
                record_journal,
            )
            logging.info(
                "[INFO] Drawing actuality checked: %s; model saved=%s, "
                "drawing saved=%s, model mass=%s, drawing property mass=%s %s, "
                "result=%s",
                drawing_path,
                model_stat.st_mtime,
                drawing_stat.st_mtime,
                model_mass,
                drawing_mass,
                mass_unit,
                status,
            )
        except Exception as error:
            logging.exception(
                "[ERROR] Не удалось проверить актуальность чертежа %s: %s",
                drawing_path,
                error,
            )
            self.set_status(
                f"Ошибка проверки чертежа {os.path.basename(drawing_path)}: {error}"
            )
            self._schedule_change(
                os.path.basename(drawing_path),
                drawing_path,
                "Чертёж",
                "ОШИБКА",
                record_journal=False,
            )

    def _kompas_constants(self):
        return win32com.client.gencache.EnsureModule(
            "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}",
            0,
            1,
            0,
        ).constants

    def _read_drawing_mass(self, kompas_app, drawing_path):
        api = win32com.client.gencache.GetModuleForProgID(
            "KOMPAS.Application.7"
        )
        drawing_key = document_key(drawing_path)
        self.internal_documents.add(drawing_key)
        try:
            drawing_document = kompas_app.Documents.Open(
                drawing_path,
                False,
                True,
            )
            if drawing_document is None:
                raise RuntimeError(f"КОМПАС не открыл чертёж: {drawing_path}")

            try:
                property_keeper = api.IPropertyKeeper(drawing_document)
                properties = ET.fromstring(property_keeper.Properties)
                mass_values = [
                    element.get("value")
                    for element in properties.iter("property")
                    if element.get("id") == "mass"
                ]
                if len(mass_values) != 1 or not mass_values[0]:
                    raise ValueError(
                        "Не удалось однозначно получить свойство массы "
                        f"чертежа: {drawing_path}"
                    )
                return (
                    parse_stamp_mass(mass_values[0]),
                    self._kompas_constants().ksMUnKG,
                )
            finally:
                drawing_document.Close(False)
        finally:
            self.internal_documents.discard(drawing_key)

    def _refresh_model_cache_for_drawing_check(
        self,
        kompas_app,
        model_path,
        cache_entry,
    ):
        stat = os.stat(model_path)
        file_hash = get_file_hash(model_path)
        if file_hash is None:
            raise RuntimeError(f"Не удалось вычислить хэш детали: {model_path}")

        markers_changed = (
            cache_entry.get("mtime") != stat.st_mtime
            or cache_entry.get("size") != stat.st_size
            or cache_entry.get("hash") != file_hash
        )
        if not markers_changed:
            return
        if self._path_is_open(kompas_app, document_key(model_path)):
            raise RuntimeError(
                f"Деталь открыта в КОМПАСе, проверка отложена: {model_path}"
            )

        self.internal_documents.add(document_key(model_path))
        try:
            mch_data, _, part_type = calculate_mch(
                kompas_app,
                model_path,
            )
        finally:
            self.internal_documents.discard(document_key(model_path))
        if mch_data is None:
            raise RuntimeError(f"Не удалось пересчитать МЦХ детали: {model_path}")

        geometry_changed, diff_fields = geometry_actually_changed(
            cache_entry["last_mch"],
            mch_data,
        )
        type_changed = cache_entry.get("part_type") != part_type
        if geometry_changed or type_changed:
            cache_entry["last_mch"] = mch_data
            cache_entry["part_type"] = part_type
            cache_entry["last_change"] = datetime.now().isoformat()
            cache_entry["last_geometry_mtime"] = stat.st_mtime
            cache_entry["change_count"] = cache_entry.get("change_count", 0) + 1
            self._schedule_change(
                os.path.basename(model_path),
                model_path,
                part_type or "Неизвестно",
                "ИЗМЕНЕН",
            )
            logging.info(
                "[INFO] Model geometry changed while refreshing drawing check: %s; %s",
                model_path,
                ", ".join(diff_fields),
            )

        cache_entry.update(
            {
                "mtime": stat.st_mtime,
                "size": stat.st_size,
                "hash": file_hash,
                "mass_is_exact": True,
                "mch_calculated": True,
                "last_check": datetime.now().isoformat(),
            }
        )
        with self.cache_lock:
            self.cache[model_path] = cache_entry
            save_event_cache(self.cache)

    def _reconcile_stale_drawings(self, kompas_app):
        if not self.monitor_options["track_chertezh"]:
            return

        with self.cache_lock:
            entries = [
                (model_path, dict(entry))
                for model_path, entry in self.cache.items()
                if self._cache_has_baseline(entry)
            ]

        reconciled = 0
        for model_path, entry in entries:
            drawing_path = (
                entry.get("drawing_path")
                or os.path.splitext(model_path)[0] + ".cdw"
            )
            model_change = self.changed_files.get(
                os.path.basename(model_path)
            )
            model_is_changed = (
                model_change is not None
                and model_change[2] == "ИЗМЕНЕН"
                and document_key(model_change[0]) == document_key(model_path)
            )
            drawing_change = self.changed_files.get(
                os.path.basename(drawing_path)
            )
            drawing_needs_reconciliation = (
                drawing_change is not None
                and drawing_change[1] == "Чертёж"
                and drawing_change[2]
                in ("ТРЕБУЕТ ПЕРЕСТРОЙКИ", "АКТУАЛЕН", "ИЗМЕНЕН")
                and document_key(drawing_change[0]) == document_key(drawing_path)
            )
            if not (model_is_changed or drawing_needs_reconciliation):
                continue
            if not os.path.isfile(model_path):
                continue

            self._check_drawing(
                model_path,
                entry,
                record_journal=False,
                kompas_app=kompas_app,
            )
            with self.cache_lock:
                self.cache[model_path] = entry
                save_event_cache(self.cache)
            reconciled += 1

        if reconciled:
            logging.info(
                "[INFO] При запуске сверены статусы связанных чертежей: %s",
                reconciled,
            )

    def _refresh_saved_drawing(self, kompas_app, drawing_path):
        drawing_key = document_key(drawing_path)
        with self.cache_lock:
            related_models = [
                (model_path, dict(entry))
                for model_path, entry in self.cache.items()
                if self._cache_has_baseline(entry)
                and document_key(
                    entry.get("drawing_path")
                    or os.path.splitext(model_path)[0] + ".cdw"
                )
                == drawing_key
            ]

        if not related_models:
            logging.warning(
                "[WARNING] Для сохраненного чертежа не найдена модель в кэше: %s",
                drawing_path,
            )
            self.set_status(
                f"Чертёж сохранён, но модель не найдена в кэше: "
                f"{os.path.basename(drawing_path)}"
            )
            return

        for model_path, entry in related_models:
            if not os.path.isfile(model_path):
                continue
            try:
                self._refresh_model_cache_for_drawing_check(
                    kompas_app,
                    model_path,
                    entry,
                )
            except Exception as error:
                logging.exception(
                    "[ERROR] Не удалось обновить данные детали для проверки %s: %s",
                    drawing_path,
                    error,
                )
                self.set_status(
                    f"Не удалось подготовить проверку чертежа "
                    f"{os.path.basename(drawing_path)}: {error}"
                )
                continue
            self._check_drawing(
                model_path,
                entry,
                kompas_app=kompas_app,
            )
            with self.cache_lock:
                current_entry = self.cache.get(model_path)
                if current_entry is not None:
                    current_entry["drawing_path"] = entry.get("drawing_path")
                    for field in (
                        "drawing_mtime",
                        "drawing_mass",
                        "drawing_mass_unit",
                    ):
                        if field in entry:
                            current_entry[field] = entry[field]
                    save_event_cache(self.cache)

    def _finish_document(self, kompas_app, key, file_path=None):
        display_path = file_path or key
        sink_record = self.document_sinks.pop(key, None)
        if sink_record is not None:
            try:
                sink_record[1].close()
            except Exception as error:
                print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось снять COM-подписку: {error}")

        deadline = time.monotonic() + 5
        try:
            while self._path_is_open(kompas_app, key):
                if self.stop_event.is_set() or time.monotonic() >= deadline:
                    print(
                        f"[ПРЕДУПРЕЖДЕНИЕ] Документ ещё открыт после события закрытия: {key}"
                    )
                    return
                pythoncom.PumpWaitingMessages()
                self.stop_event.wait(0.05)
        except Exception as error:
            print(f"[ОШИБКА] Не удалось подтвердить закрытие {key}: {error}")
            return

        self.deferred_document_retries.pop(key, None)
        self.document_calculation_retry_counts.pop(key, None)
        if not os.path.isfile(display_path):
            print(f"[ПРЕДУПРЕЖДЕНИЕ] Закрытый файл не найден: {key}")
            return

        if os.path.splitext(display_path)[1].lower() in DRAWING_EXTENSIONS:
            self._record_saved_drawing_change(key)
            self._refresh_saved_drawing(kompas_app, key)
            print(
                f"[ИНФО] Сохранённый чертёж повторно проверен: {key}"
            )
            return

        with self.cache_lock:
            previous = self.cache.get(key)
            if not self._cache_has_baseline(previous):
                return
            previous_mch = previous["last_mch"]

        self.internal_documents.add(key)
        try:
            self.set_status(f"Проверка закрытой детали: {os.path.basename(display_path)}")
            mch_data, _, part_type = calculate_mch(kompas_app, key)
            if mch_data is None:
                self._schedule_change(
                    os.path.basename(display_path),
                    key,
                    previous.get("part_type") or "Неизвестно",
                    "ОШИБКА",
                )
                print(
                    f"[ОШИБКА] Не удалось пересчитать МЦХ после закрытия: {key}"
                )
                return

            type_changed = previous.get("part_type") != part_type
            geometry_changed, diff_fields = geometry_actually_changed(
                previous_mch,
                mch_data,
            )
            stat = os.stat(display_path)
            updated_entry = dict(previous)
            updated_entry.update(
                {
                    "mtime": stat.st_mtime,
                    "size": stat.st_size,
                    "hash": get_file_hash(display_path),
                    "last_mch": mch_data,
                    "mass_is_exact": True,
                    "mch_calculated": True,
                    "last_check": datetime.now().isoformat(),
                    "change_count": previous.get("change_count", 0),
                }
            )
            if geometry_changed or type_changed:
                updated_entry["last_mch"] = mch_data
                updated_entry["part_type"] = part_type
                updated_entry["last_change"] = datetime.now().isoformat()
                updated_entry["last_geometry_mtime"] = stat.st_mtime
                updated_entry["change_count"] += 1
                change_details = (
                    ", ".join(diff_fields) if geometry_changed else "тип детали"
                )
                safe_change_details = change_details.encode(
                    "ascii",
                    errors="backslashreplace",
                ).decode("ascii")
                logging.info(
                    f"[ИНФО] Изменена геометрия {key}; поля: "
                    f"{safe_change_details}"
                )
                self._schedule_change(
                    os.path.basename(display_path),
                    key,
                    part_type or "Неизвестно",
                    "ИЗМЕНЕН",
                )
                if part_type == "Листовая деталь" and self.monitor_options[
                    "track_unfolding"
                ]:
                    self.root.after(
                        0,
                        self.check_unfolding_for_model,
                        key,
                        part_type,
                        updated_entry.get("drawing_path"),
                    )
            else:
                print(f"[ИНФО] МЦХ не изменились после закрытия: {key}")

            self._check_drawing(
                key,
                updated_entry,
                kompas_app=kompas_app,
            )
            with self.cache_lock:
                self.cache[key] = updated_entry
                save_event_cache(self.cache)
        except Exception as error:
            print(f"[ОШИБКА] Ошибка проверки закрытой детали {key}: {error}")
            print(traceback.format_exc())
            self._schedule_change(
                os.path.basename(display_path),
                key,
                previous.get("part_type") or "Неизвестно",
                "ОШИБКА",
            )
        finally:
            self.internal_documents.discard(key)
            self.set_progress(0)

    def monitor_loop(self):
        try:
            pythoncom.CoInitialize()
        except Exception as error:
            print(f"[ОШИБКА] Не удалось инициализировать COM-поток: {error}")
            self.root.after(
                0,
                self._monitoring_failed,
                "Не удалось инициализировать COM для подключения к КОМПАС. "
                "Техническая причина записана в журнал.",
            )
            return
        app_sink = None
        kompas_app = None
        try:
            connection_warning_logged = False
            while not self.stop_event.is_set():
                try:
                    active_object = pythoncom.GetActiveObject(
                        "KOMPAS.Application.7"
                    )
                    dispatch = active_object.QueryInterface(
                        pythoncom.IID_IDispatch
                    )
                    kompas_app = win32com.client.dynamic.Dispatch(dispatch)
                    break
                except Exception as error:
                    if not connection_warning_logged:
                        logging.warning(
                            "[WARNING] КОМПАС пока недоступен; ожидание запуска: %s",
                            error,
                        )
                        self.root.after(0, self._waiting_for_kompas)
                        connection_warning_logged = True
                    if self.stop_event.wait(5):
                        return

            if self.stop_event.is_set():
                return

            type_info = kompas_app._oleobj_.GetTypeInfo()
            typelib, _ = type_info.GetContainingTypeLib()
            event_module = (
                win32com.client.gencache.EnsureModuleForTypelibInterface(
                    typelib,
                    bForDemand=False,
                    bBuildHidden=1,
                )
            )
            if event_module is None:
                typelib_attr = typelib.GetLibAttr()
                event_module = win32com.client.gencache.MakeModuleForTypelib(
                    typelib_attr[0],
                    typelib_attr[1],
                    typelib_attr[3],
                    typelib_attr[4],
                    bForDemand=False,
                    bBuildHidden=1,
                )
            if event_module is None:
                raise RuntimeError(
                    "Не удалось подготовить интерфейсы событий КОМПАС. "
                    "Проверьте установку КОМПАС-3D и перезапустите MonitorCAD."
                )
            if not all(
                hasattr(event_module, interface_name)
                for interface_name in (
                    "ksKompasObjectNotify",
                    "ksDocumentFileNotify",
                )
            ):
                raise RuntimeError(
                    "В установленной версии КОМПАС отсутствуют необходимые "
                    "интерфейсы событий. Проверьте версию КОМПАС-3D."
                )
            owner = self

            class ApplicationEvents(event_module.ksKompasObjectNotify):
                def __init__(self, dispatch):
                    super().__init__(dispatch)

                def OnOpenDocument(self, new_doc, doc_type):
                    owner.event_queue.put(("opened", None))

                def OnApplicationDestroy(self):
                    owner.event_queue.put(("application_destroyed", None))

            app_sink = ApplicationEvents(kompas_app)
            self.set_status("Подключено. Ожидание открытия или закрытия деталей...")
            self.root.after(
                0,
                lambda: self.status_label.config(text="ОТСЛЕЖИВАЕТ"),
            )
            self.root.after(
                0,
                lambda: self.indicator_canvas.itemconfig(
                    self.indicator_oval,
                    fill="green",
                ),
            )
            self.root.after(0, lambda: self.scan_btn.config(state=tk.NORMAL))

            self._reconcile_stale_drawings(kompas_app)
            self._register_open_documents(kompas_app, event_module)
            while not self.stop_event.is_set():
                pythoncom.PumpWaitingMessages()
                self._process_due_document_retries(
                    kompas_app,
                    event_module,
                )
                try:
                    event_data = self.event_queue.get(timeout=0.05)
                except queue.Empty:
                    continue
                event_name = event_data[0]
                event_value = event_data[1] if len(event_data) == 2 else event_data[1:]

                if event_name == "manual_scan":
                    self._scan_cached_database(kompas_app, event_module)
                elif event_name == "saved":
                    if isinstance(event_value, tuple):
                        _, saved_path = event_value
                    else:
                        saved_path = event_value
                    if (
                        os.path.splitext(saved_path)[1].lower()
                        in DRAWING_EXTENSIONS
                    ):
                        self._record_saved_drawing_change(saved_path)
                        logging.info(
                            "[INFO] Чертеж сохранен; проверка актуальности будет "
                            "после закрытия: %s",
                            event_value,
                        )
                else:
                    self._process_event(
                        kompas_app,
                        event_module,
                        event_name,
                        event_value,
                    )

        except Exception as error:
            print(f"[ОШИБКА] Отслеживание событий КОМПАС остановлено: {error}")
            print(traceback.format_exc())
            user_message = (
                str(error)
                if isinstance(error, RuntimeError)
                else "Не удалось запустить слежение за документами КОМПАС. "
                "Техническая причина записана в журнал."
            )
            self.root.after(0, self._monitoring_failed, user_message)
        finally:
            for _, sink in self.document_sinks.values():
                try:
                    sink.close()
                except Exception as error:
                    print(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка снятия COM-подписки: {error}")
            self.document_sinks.clear()
            if app_sink is not None:
                try:
                    app_sink.close()
                except Exception as error:
                    print(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка снятия COM-подписки: {error}")
            pythoncom.CoUninitialize()

    def _process_event(
        self,
        kompas_app,
        event_module,
        event_name,
        event_value,
    ):
        if event_name == "opened":
            self._register_open_documents(kompas_app, event_module)
        elif event_name == "closed":
            if isinstance(event_value, tuple):
                doc_key, doc_path = event_value
            else:
                doc_key = doc_path = event_value
            self._finish_document(kompas_app, doc_key, doc_path)
        elif event_name == "application_destroyed":
            raise RuntimeError("Работающий процесс КОМПАС-3D был закрыт.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MonitorCADApp(root)
    root.mainloop()
