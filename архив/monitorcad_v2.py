import json
import os
import queue
import threading
import time
import traceback
from datetime import datetime

import pythoncom
import win32com.client
import win32com.client.dynamic
import win32con
import win32gui
from tkinter import messagebox

import monitorcad as legacy


EVENT_CACHE_PATH = os.path.join(legacy.SCRIPT_DIR, "mch_cache_events.json")
EVENT_CHANGED_LIST_PATH = os.path.join(
    legacy.SCRIPT_DIR, "changed_files_events.json"
)
CHANGE_JOURNAL_PATH = os.path.join(
    legacy.SCRIPT_DIR, "change_journal_events.json"
)
APPLICATION_ICON_PATH = os.path.join(
    legacy.SCRIPT_DIR, "change-management.ico"
)


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
        legacy.print(f"[ОШИБКА] Не удалось загрузить событийный кэш: {error}")
        return {}


def save_event_cache(cache):
    try:
        with open(EVENT_CACHE_PATH, "w", encoding="utf-8") as cache_file:
            json.dump(cache, cache_file, ensure_ascii=False, indent=2)
    except OSError as error:
        legacy.print(f"[ОШИБКА] Не удалось сохранить событийный кэш: {error}")


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
        legacy.print(f"[ОШИБКА] Не удалось загрузить список изменений: {error}")
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
        legacy.print(f"[ОШИБКА] Не удалось сохранить список изменений: {error}")


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
        legacy.print(f"[ОШИБКА] Не удалось загрузить журнал изменений: {error}")
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
        legacy.print(f"[ОШИБКА] Не удалось сохранить журнал изменений: {error}")


def document_key(file_path):
    return os.path.normcase(os.path.abspath(file_path))


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
                legacy.logging.exception(
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
                    legacy.logging.exception(
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
        except legacy.tk.TclError:
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
            legacy.logging.error(
                "[ERROR] Поток значка MonitorCAD не завершился после запроса выхода."
            )


class MonitorCADApp(legacy.MonitorCADApp):
    def __init__(self, root):
        super().__init__(root)
        self.cache = load_event_cache()
        self.cache_lock = threading.RLock()
        self.event_queue = queue.Queue()
        self.document_sinks = {}
        self.internal_documents = set()
        self.monitor_options = {}
        self.journal_records = load_change_journal()
        self.journal_window = None
        self.journal_tree = None
        self.observed_file_signatures = {}
        self.tray_icon = None
        self.root.bind("<Unmap>", self._on_window_minimized, add="+")
        self.root.protocol("WM_DELETE_WINDOW", self._quit_app)
        self.root.after(0, self._initialize_tray_icon)
        self.cache_count_var = legacy.tk.StringVar()
        self.cache_count_label = legacy.ttk.Label(
            self.status_label.master,
            textvariable=self.cache_count_var,
            font=("Arial", 9),
        )
        self.cache_count_label.pack(side=legacy.tk.RIGHT, padx=(8, 0))
        self.journal_button = legacy.ttk.Button(
            self.status_label.master,
            text="Журнал",
            command=self.show_journal,
        )
        self.journal_button.pack(side=legacy.tk.RIGHT, padx=(8, 0))
        self.scan_running = False
        button_frame = self.start_btn.master
        self.start_btn.destroy()
        self.stop_btn.destroy()
        self.scan_btn = legacy.ttk.Button(
            button_frame,
            text="Проверить изменения",
            command=self.request_manual_scan,
            state=legacy.tk.DISABLED,
            width=24,
        )
        self.scan_btn.pack(side=legacy.tk.LEFT)

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
            if isinstance(widget, legacy.ttk.Label) and widget.cget("text").startswith(
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
            legacy.logging.exception(
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
        except legacy.tk.TclError:
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
                not in legacy.MODEL_EXTENSIONS
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
                legacy.tk.END,
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

        window = legacy.tk.Toplevel(self.root)
        window.title("Журнал изменений MonitorCAD")
        window.geometry("850x420")
        window.transient(self.root)
        self.journal_window = window

        frame = legacy.ttk.Frame(window, padding=10)
        frame.pack(fill=legacy.tk.BOTH, expand=True)
        columns = ("timestamp", "filename")
        self.journal_tree = legacy.ttk.Treeview(
            frame,
            columns=columns,
            show="headings",
        )
        self.journal_tree.heading("timestamp", text="Дата изменения")
        self.journal_tree.heading("filename", text="Файл")
        self.journal_tree.column("timestamp", width=150, anchor=legacy.tk.CENTER)
        self.journal_tree.column("filename", width=650, anchor=legacy.tk.W)
        self.journal_tree.pack(side=legacy.tk.LEFT, fill=legacy.tk.BOTH, expand=True)

        scrollbar = legacy.ttk.Scrollbar(
            frame,
            orient=legacy.tk.VERTICAL,
            command=self.journal_tree.yview,
        )
        self.journal_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=legacy.tk.RIGHT, fill=legacy.tk.Y)
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
        self.scan_btn.config(state=legacy.tk.DISABLED)
        self.set_status("Ручная сверка базы ожидает COM-поток...")
        self.event_queue.put(("manual_scan", None))

    def _manual_scan_finished(self):
        self.scan_running = False
        if self.is_running:
            self.scan_btn.config(state=legacy.tk.NORMAL)

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
        legacy.logging.info(
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
                    if not os.path.isfile(key):
                        errors += 1
                        legacy.logging.warning(
                            "[WARNING] Cached model file is missing: %s", key
                        )
                        continue

                    stat = os.stat(key)
                    current_hash = legacy.get_file_hash(key)
                    if current_hash is None:
                        errors += 1
                        legacy.logging.error(
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
                        legacy.logging.info(
                            "[INFO] Manual scan skipped an open document: %s",
                            key,
                        )
                        continue

                    changed_files += 1
                    legacy.logging.info(
                        "[INFO] Changed file markers found during manual scan: %s",
                        key,
                    )
                    self.internal_documents.add(key)
                    try:
                        mch_data, _, part_type = legacy.calculate_mch(
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
                        legacy.logging.error(
                            "[ERROR] MCH calculation failed during manual scan: %s",
                            key,
                        )
                        continue

                    geometry_changed, _ = legacy.geometry_actually_changed(
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
                        updated_entry["change_count"] = (
                            cached_entry.get("change_count", 0) + 1
                        )
                        self._schedule_change(
                            filename,
                            key,
                            part_type or "Неизвестно",
                            "ИЗМЕНЕН",
                        )
                        self._check_drawing(key, updated_entry)
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
                        legacy.logging.info(
                            "[INFO] MCH changed during manual scan: %s",
                            key,
                        )
                    else:
                        legacy.logging.info(
                            "[INFO] File markers changed but MCH did not: %s",
                            key,
                        )

                    with self.cache_lock:
                        self.cache[key] = updated_entry
                        save_event_cache(self.cache)
                except Exception as error:
                    errors += 1
                    legacy.logging.exception(
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
            legacy.logging.info("[INFO] %s", summary)

    def _process_pending_events(self, kompas_app, event_module):
        while True:
            try:
                event_name, event_value = self.event_queue.get_nowait()
            except queue.Empty:
                return

            if event_name == "manual_scan":
                legacy.logging.warning(
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
            filename = record.get("filename", "")
            full_path = record.get("full_path", "")
            part_type = record.get("part_type", "")
            record_dir = (
                self.unfolding_dir.get()
                if part_type == "Развертка DXF"
                else selected_dir
            )
            if selected_dir and not legacy.path_belongs_to_directory(
                full_path, record_dir
            ):
                continue

            status = record.get("status", "")
            self.tree.insert(
                "",
                legacy.tk.END,
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
        self._reconcile_stale_drawings()

        self.is_running = True
        self.stop_event.clear()
        self.scan_btn.config(state=legacy.tk.DISABLED)
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
        self.scan_btn.config(state=legacy.tk.DISABLED)
        self.status_label.config(text="ОСТАНОВЛЕН")
        self.indicator_canvas.itemconfig(self.indicator_oval, fill="gray")
        self.set_status("Отслеживание событий КОМПАС остановлено.")

    def clear_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.changed_files = {}
        self.total_changes = 0
        save_event_changed_list([], 0)
        self.status_line_var.set("Список очищен.")

    def _monitoring_failed(self, message):
        self.is_running = False
        self.stop_event.set()
        self.scan_btn.config(state=legacy.tk.DISABLED)
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
        if extension not in legacy.MODEL_EXTENSIONS:
            return False
        return legacy.path_belongs_to_directory(
            file_path,
            self.monitor_options["monitor_dir"],
        )

    def _cache_has_baseline(self, entry):
        return (
            isinstance(entry, dict)
            and entry.get("mch_calculated") is True
            and isinstance(entry.get("last_mch"), dict)
        )

    def _part_type_is_tracked(self, part_type):
        if part_type == "Листовая деталь":
            return self.monitor_options["track_listovaya"]
        return self.monitor_options["track_detali"]

    def _register_open_document(self, kompas_app, document, event_module):
        try:
            file_path = document.PathName
        except Exception as error:
            legacy.print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось получить путь документа: {error}")
            return

        if not file_path:
            return

        extension = os.path.splitext(file_path)[1].lower()
        is_drawing = extension in legacy.DRAWING_EXTENSIONS
        if is_drawing:
            if (
                not self.monitor_options["track_chertezh"]
                or not legacy.path_belongs_to_directory(
                    file_path,
                    self.monitor_options["monitor_dir"],
                )
            ):
                return
        elif not self._event_path_is_tracked(file_path):
            return

        key = document_key(file_path)
        if key in self.internal_documents:
            return

        if key not in self.document_sinks:
            owner = self

            class DocumentEvents(event_module.ksDocumentFileNotify):
                def __init__(self, dispatch):
                    self.document_key = key
                    super().__init__(dispatch)

                def OnCloseDocument(self):
                    owner.event_queue.put(("closed", self.document_key))

                def OnSaveDocument(self):
                    owner.event_queue.put(("saved", self.document_key))

            try:
                sink = DocumentEvents(document)
            except Exception as error:
                legacy.print(
                    f"[ОШИБКА] Не удалось подписаться на сохранение {file_path}: {error}"
                )
                return
            self.document_sinks[key] = (document, sink)
            if is_drawing:
                signature = self._get_file_signature(file_path)
                if signature is not None:
                    self.observed_file_signatures[key] = signature
                legacy.logging.info(
                    "[INFO] Подписка на сохранение чертежа установлена: %s",
                    file_path,
                )

        if is_drawing:
            return

        with self.cache_lock:
            entry = self.cache.get(key)
            has_baseline = self._cache_has_baseline(entry)
            if has_baseline and not self._part_type_is_tracked(
                entry.get("part_type")
            ):
                sink_record = self.document_sinks.pop(key, None)
                if sink_record is not None:
                    sink_record[1].close()
                return

        if has_baseline:
            return

        self.set_status(f"Первичный расчёт МЦХ: {os.path.basename(file_path)}")
        try:
            mch_data, _, part_type = legacy.calculate_mch(kompas_app, file_path)
            if mch_data is None:
                self._schedule_change(
                    os.path.basename(file_path),
                    file_path,
                    "Неизвестно",
                    "ОШИБКА",
                )
                legacy.print(
                    f"[ОШИБКА] Не удалось записать исходные МЦХ для {file_path}."
                )
                return

            if not self._part_type_is_tracked(part_type):
                sink_record = self.document_sinks.pop(key, None)
                if sink_record is not None:
                    sink_record[1].close()
                return

            stat = os.stat(file_path)
            drawing_path = legacy.find_drawing_for_model(file_path)
            drawing_mtime = (
                os.stat(drawing_path).st_mtime if drawing_path else None
            )
            new_entry = {
                "mtime": stat.st_mtime,
                "size": stat.st_size,
                "hash": legacy.get_file_hash(file_path),
                "last_mch": mch_data,
                "mch_calculated": True,
                "last_check": datetime.now().isoformat(),
                "drawing_path": drawing_path,
                "drawing_mtime": drawing_mtime,
                "part_type": part_type,
            }
            with self.cache_lock:
                self.cache[key] = new_entry
                save_event_cache(self.cache)
            legacy.print(
                f"[ИНФО] Исходные МЦХ сохранены в событийный кэш: {file_path}"
            )
            self.set_status(
                f"Исходная точка сохранена: {os.path.basename(file_path)}"
            )
        except Exception as error:
            legacy.print(f"[ОШИБКА] Не удалось обработать открытый документ {file_path}: {error}")
            legacy.print(traceback.format_exc())

    def _register_open_documents(self, kompas_app, event_module):
        try:
            documents = list(kompas_app.Documents)
        except Exception as error:
            legacy.print(f"[ОШИБКА] Не удалось получить список документов КОМПАС: {error}")
            legacy.print(traceback.format_exc())
            return
        for document in documents:
            self._register_open_document(kompas_app, document, event_module)

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
            legacy.logging.warning(
                "[WARNING] Не удалось прочитать файл для журнала %s: %s",
                file_path,
                error,
            )
            return None
        file_hash = legacy.get_file_hash(file_path)
        if file_hash is None:
            legacy.logging.warning(
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

    def _check_drawing(self, model_path, cache_entry, record_journal=False):
        if not self.monitor_options["track_chertezh"]:
            return
        drawing_path = legacy.find_drawing_for_model(model_path)
        model_mtime = os.path.getmtime(model_path)
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
        else:
            drawing_mtime = os.path.getmtime(drawing_path)
            cache_entry["drawing_path"] = drawing_path
            cache_entry["drawing_mtime"] = drawing_mtime
            if drawing_mtime < model_mtime:
                self._schedule_change(
                    os.path.basename(drawing_path),
                    drawing_path,
                    "Чертёж",
                    "ТРЕБУЕТ ПЕРЕСТРОЙКИ",
                    record_journal,
                )
            else:
                self._schedule_change(
                    os.path.basename(drawing_path),
                    drawing_path,
                    "Чертёж",
                    "ИЗМЕНЕН",
                    record_journal,
                )

    def _reconcile_stale_drawings(self):
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
                and drawing_change[2] == "ТРЕБУЕТ ПЕРЕСТРОЙКИ"
                and document_key(drawing_change[0]) == document_key(drawing_path)
            )
            if not (model_is_changed or drawing_needs_reconciliation):
                continue
            if not os.path.isfile(model_path):
                continue

            self._check_drawing(model_path, entry, record_journal=False)
            with self.cache_lock:
                self.cache[model_path] = entry
                save_event_cache(self.cache)
            reconciled += 1

        if reconciled:
            legacy.logging.info(
                "[INFO] При запуске сверены статусы связанных чертежей: %s",
                reconciled,
            )

    def _refresh_saved_drawing(self, drawing_path):
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
            legacy.logging.warning(
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
            self._check_drawing(model_path, entry)
            with self.cache_lock:
                current_entry = self.cache.get(model_path)
                if current_entry is not None:
                    current_entry["drawing_path"] = entry.get("drawing_path")
                    current_entry["drawing_mtime"] = entry.get("drawing_mtime")
                    save_event_cache(self.cache)

    def _finish_document(self, kompas_app, key):
        sink_record = self.document_sinks.pop(key, None)
        if sink_record is not None:
            try:
                sink_record[1].close()
            except Exception as error:
                legacy.print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось снять COM-подписку: {error}")

        deadline = time.monotonic() + 5
        try:
            while self._path_is_open(kompas_app, key):
                if self.stop_event.is_set() or time.monotonic() >= deadline:
                    legacy.print(
                        f"[ПРЕДУПРЕЖДЕНИЕ] Документ ещё открыт после события закрытия: {key}"
                    )
                    return
                pythoncom.PumpWaitingMessages()
                self.stop_event.wait(0.05)
        except Exception as error:
            legacy.print(f"[ОШИБКА] Не удалось подтвердить закрытие {key}: {error}")
            return

        if not os.path.isfile(key):
            legacy.print(f"[ПРЕДУПРЕЖДЕНИЕ] Закрытый файл не найден: {key}")
            return

        if os.path.splitext(key)[1].lower() in legacy.DRAWING_EXTENSIONS:
            self._record_saved_drawing_change(key)
            self._refresh_saved_drawing(key)
            legacy.print(
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
            self.set_status(f"Проверка закрытой детали: {os.path.basename(key)}")
            mch_data, _, part_type = legacy.calculate_mch(kompas_app, key)
            if mch_data is None:
                self._schedule_change(
                    os.path.basename(key),
                    key,
                    previous.get("part_type") or "Неизвестно",
                    "ОШИБКА",
                )
                legacy.print(
                    f"[ОШИБКА] Не удалось пересчитать МЦХ после закрытия: {key}"
                )
                return

            type_changed = previous.get("part_type") != part_type
            geometry_changed, diff_fields = legacy.geometry_actually_changed(
                previous_mch,
                mch_data,
            )
            stat = os.stat(key)
            updated_entry = dict(previous)
            updated_entry.update(
                {
                    "mtime": stat.st_mtime,
                    "size": stat.st_size,
                    "hash": legacy.get_file_hash(key),
                    "mch_calculated": True,
                    "last_check": datetime.now().isoformat(),
                    "change_count": previous.get("change_count", 0),
                }
            )
            if geometry_changed or type_changed:
                updated_entry["last_mch"] = mch_data
                updated_entry["part_type"] = part_type
                updated_entry["last_change"] = datetime.now().isoformat()
                updated_entry["change_count"] += 1
                change_details = (
                    ", ".join(diff_fields) if geometry_changed else "тип детали"
                )
                safe_change_details = change_details.encode(
                    "ascii",
                    errors="backslashreplace",
                ).decode("ascii")
                legacy.logging.info(
                    f"[ИНФО] Изменена геометрия {key}; поля: "
                    f"{safe_change_details}"
                )
                self._schedule_change(
                    os.path.basename(key),
                    key,
                    part_type or "Неизвестно",
                    "ИЗМЕНЕН",
                )
                self._check_drawing(key, updated_entry)
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
                legacy.print(f"[ИНФО] МЦХ не изменились после закрытия: {key}")

            with self.cache_lock:
                self.cache[key] = updated_entry
                save_event_cache(self.cache)
        except Exception as error:
            legacy.print(f"[ОШИБКА] Ошибка проверки закрытой детали {key}: {error}")
            legacy.print(traceback.format_exc())
            self._schedule_change(
                os.path.basename(key),
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
            legacy.print(f"[ОШИБКА] Не удалось инициализировать COM-поток: {error}")
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
                        legacy.logging.warning(
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
            self.root.after(0, lambda: self.scan_btn.config(state=legacy.tk.NORMAL))

            self._register_open_documents(kompas_app, event_module)
            while not self.stop_event.is_set():
                pythoncom.PumpWaitingMessages()
                try:
                    event_name, event_value = self.event_queue.get(timeout=0.05)
                except queue.Empty:
                    continue

                if event_name == "manual_scan":
                    self._scan_cached_database(kompas_app, event_module)
                elif event_name == "saved":
                    if (
                        os.path.splitext(event_value)[1].lower()
                        in legacy.DRAWING_EXTENSIONS
                    ):
                        self._record_saved_drawing_change(event_value)
                        legacy.logging.info(
                            "[INFO] Сохранен чертеж, проверяю его статус: %s",
                            event_value,
                        )
                        self._refresh_saved_drawing(event_value)
                else:
                    self._process_event(
                        kompas_app,
                        event_module,
                        event_name,
                        event_value,
                    )

        except Exception as error:
            legacy.print(f"[ОШИБКА] Отслеживание событий КОМПАС остановлено: {error}")
            legacy.print(traceback.format_exc())
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
                    legacy.print(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка снятия COM-подписки: {error}")
            self.document_sinks.clear()
            if app_sink is not None:
                try:
                    app_sink.close()
                except Exception as error:
                    legacy.print(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка снятия COM-подписки: {error}")
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
            self._finish_document(kompas_app, event_value)
        elif event_name == "application_destroyed":
            raise RuntimeError("Работающий процесс КОМПАС-3D был закрыт.")


if __name__ == "__main__":
    root = legacy.tk.Tk()
    app = MonitorCADApp(root)
    root.mainloop()
