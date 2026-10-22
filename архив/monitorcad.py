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


# Пути — теперь определяются относительно расположения скрипта
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

# Расширения 3D-моделей — только детали
MODEL_EXTENSIONS = {'.m3d', '.m3d_'}

# Расширения чертежей
DRAWING_EXTENSIONS = {'.cdw'}

# Поля для сравнения МЦХ
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
            "Масса (г)": round(mass_g, 4),
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


class MonitorCADApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MonitorCAD — Мониторинг листовых деталей")
        self.root.geometry("900x650")
        self.root.minsize(850, 600)
        self.root.resizable(True, True)
        
        self.monitor_dir = tk.StringVar()
        self.unfolding_dir = tk.StringVar()
        self.auto_mode = tk.BooleanVar()
        self.interval = tk.StringVar(value="30")
        self.is_running = False
        self.monitor_thread = None
        self.stop_event = threading.Event()
        self.cache = load_cache()
        self.settings = load_settings()
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
            self.tree.insert("", tk.END, values=(
                rec.get("filename", ""),
                rec.get("part_type", ""),
                status,
            ))
            self.changed_files[rec.get("filename", "")] = (
                rec.get("full_path", ""),
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
            "track_unfolding": self.track_unfolding.get()
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
            return

        values = self.tree.item(selected[0], "values")
        if values:
            self.clear_single_record(values[0])

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
            if values and values[0] == filename:
                row_values = list(values)
                row_values[2] = new_status
                self.tree.item(item, values=tuple(row_values))
                break
        self.persist_changed_list()
    
    def add_to_list(self, filename, part_type, status):
        # Ищем уже существующую запись с таким файлом
        existing_item = None
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            if values and values[0] == filename:
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


if __name__ == "__main__":
    os.makedirs(SCRIPT_DIR, exist_ok=True)
    
    root = tk.Tk()
    app = MonitorCADApp(root)
    root.mainloop()