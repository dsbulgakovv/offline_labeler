from __future__ import annotations

import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Callable, Dict, List, Optional, Tuple

from .config import (
    APP_NAME,
    ModesConfigError,
    ensure_runtime_dirs,
    latest_modes_backup,
    load_app_settings,
    load_modes,
    restore_latest_modes_backup,
    save_app_settings,
)
from .io_utils import (
    DialogueRecord,
    InputFileFormatError,
    build_draft_path,
    build_output_path,
    inspect_input_file,
    list_input_files,
    load_dialogues,
    load_draft_json,
    normalize_annotations_for_records,
    save_draft_json,
    save_results_excel,
    split_dialogue_text,
)
from .mode_editor import ModeEditorWindow
from .models import ValidationIssue, empty_annotation
from .rules import disabled_field_keys, is_annotation_complete, validate_annotation
from .storage import (
    build_current_file_history,
    get_field_suggestions,
    get_grouped_history,
    load_history,
    save_history,
    update_history_from_annotations,
)


class HistoryDialog(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        field_label: str,
        grouped_history: Dict[str, List[str]],
        on_pick: Callable[[str], None],
    ) -> None:
        super().__init__(master)
        self.title(f"История значений — {field_label}")
        self.geometry("760x420")
        self.transient(master)
        self.grab_set()
        self.on_pick = on_pick

        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        tabs = {
            "current_file": "Текущий файл",
            "global_top": "Частые значения",
            "recent": "Недавние",
        }
        self.listboxes: Dict[str, tk.Listbox] = {}
        for key, title in tabs.items():
            frame = ttk.Frame(notebook, padding=8)
            notebook.add(frame, text=title)
            lb = tk.Listbox(frame)
            lb.pack(fill=tk.BOTH, expand=True)
            for value in grouped_history.get(key, []):
                lb.insert(tk.END, value)
            lb.bind("<Double-Button-1>", lambda e, listbox=lb: self._apply(listbox))
            self.listboxes[key] = lb

        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Button(btns, text="Применить", command=self.apply_current).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Закрыть", command=self.destroy).pack(side=tk.RIGHT, padx=(0, 8))

    def _apply(self, listbox: tk.Listbox) -> None:
        selection = listbox.curselection()
        if not selection:
            return
        value = listbox.get(selection[0])
        self.on_pick(value)
        self.destroy()

    def apply_current(self) -> None:
        for lb in self.listboxes.values():
            if lb.curselection():
                self._apply(lb)
                return


class StatsWindow(tk.Toplevel):
    def __init__(self, master: "LabelerApp") -> None:
        super().__init__(master)
        self.app = master
        self.title("Статистика по разметке")
        self.geometry("980x720")
        self.minsize(880, 620)

        container = ttk.Frame(self, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(container)
        top.pack(fill=tk.X)
        ttk.Button(top, text="Обновить", command=self.refresh).pack(side=tk.RIGHT)

        summary_frame = ttk.LabelFrame(container, text="Сводка", padding=10)
        summary_frame.pack(fill=tk.X, pady=(8, 10))
        self.summary_text = tk.Text(summary_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        self.summary_text.pack(fill=tk.X)

        notebook = ttk.Notebook(container)
        notebook.pack(fill=tk.BOTH, expand=True)

        progress_tab = ttk.Frame(notebook, padding=10)
        value_tab = ttk.Frame(notebook, padding=10)
        notebook.add(progress_tab, text="По полям")
        notebook.add(value_tab, text="Распределения")

        self.field_tree = ttk.Treeview(progress_tab, columns=("field", "filled", "required", "type"), show="headings")
        for name, width in [("field", 260), ("filled", 160), ("required", 100), ("type", 120)]:
            self.field_tree.heading(name, text=name)
            self.field_tree.column(name, width=width, anchor="w")
        self.field_tree.pack(fill=tk.BOTH, expand=True)

        self.value_tree = ttk.Treeview(value_tab, columns=("field", "value", "count"), show="headings")
        for name, width in [("field", 240), ("value", 340), ("count", 100)]:
            self.value_tree.heading(name, text=name)
            self.value_tree.column(name, width=width, anchor="w")
        self.value_tree.pack(fill=tk.BOTH, expand=True)
        self.refresh()

    def refresh(self) -> None:
        summary = self.app.compute_summary()
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete("1.0", tk.END)
        lines = [
            f"Всего записей: {summary['total_records']}",
            f"Полностью размечено: {summary['completed_records']}",
            f"Осталось: {summary['remaining_records']}",
            f"Помечено 'на проверку': {summary['needs_review_count']}",
            f"Низкая/средняя уверенность: {summary['low_or_medium_confidence_count']}",
            f"Текущий режим: {summary['mode_name']} (v{summary['mode_version']})",
            f"Фильтр 'только неразмеченные': {'включён' if self.app.filter_unlabeled_var.get() else 'выключен'}",
            f"Текущий файл: {summary['source_file']}",
        ]
        self.summary_text.insert("1.0", "\n".join(lines))
        self.summary_text.config(state=tk.DISABLED)

        for tree in (self.field_tree, self.value_tree):
            for item in tree.get_children():
                tree.delete(item)

        for row in summary["field_completion_rows"]:
            self.field_tree.insert("", tk.END, values=row)
        for row in summary["value_distribution_rows"]:
            self.value_tree.insert("", tk.END, values=row)


class LabelerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.paths = ensure_runtime_dirs()
        self.base_dir = self.paths["base_dir"]
        self.input_dir = self.paths["input_dir"]
        self.output_dir = self.paths["output_dir"]
        self.meta_dir = self.paths["meta_dir"]
        self.backups_dir = self.paths["backups_dir"]
        self.modes_path = self.paths["modes_path"]
        self.app_settings_path = self.paths["app_settings_path"]

        self.settings = load_app_settings(self.app_settings_path)
        self.modes_payload = load_modes(self.modes_path)
        self.modes = self.modes_payload["modes"]
        self.mode_map = {mode["id"]: mode for mode in self.modes}
        self.history_payload = load_history(self.meta_dir)
        self.current_file_history: Dict[str, Dict[str, int]] = {}

        self.title(APP_NAME)
        self.geometry(self.settings.get("window_geometry", "1500x980"))
        self.minsize(1260, 820)

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Muted.TLabel", foreground="#5b6470")
        style.configure("Accent.TButton", padding=(12, 8))

        default_mode_id = self.settings.get("last_mode_id") or self.modes[0]["id"]
        if default_mode_id not in self.mode_map:
            default_mode_id = self.modes[0]["id"]
        self.selected_mode_id = tk.StringVar(value=default_mode_id)
        self.annotator_name_var = tk.StringVar(value=self.settings.get("last_annotator_name", ""))
        self.filter_unlabeled_var = tk.BooleanVar(value=False)
        self.jump_session_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Выберите режим и файл из input, затем начните разметку.")
        self.progress_var = tk.StringVar(value="0 / 0")
        self.done_var = tk.StringVar(value="Размечено: 0")
        self.remaining_var = tk.StringVar(value="Осталось: 0")
        self.validation_var = tk.StringVar(value="Проверка: нет данных")
        self.last_autosave_var = tk.StringVar(value="Автосохранение: ещё не выполнялось")

        self.records: List[DialogueRecord] = []
        self.annotations: Dict[str, Dict[str, Any]] = {}
        self.current_index = 0
        self.current_input_path: Optional[Path] = None
        self.file_display_map: Dict[str, Path] = {}
        self.widget_bindings: Dict[str, Dict[str, Any]] = {}
        self.entry_edit_state: Dict[str, Dict[str, List[str]]] = {}
        self.prev_mode_id = default_mode_id
        self.filtered_indices: List[int] = []
        self.dirty = False
        self.autosave_job: Optional[str] = None
        self.currently_disabled_fields: List[str] = []
        self.suspend_reactive_updates = False

        self._build_ui()
        self._bind_hotkeys()
        self._attach_global_edit_support()
        self.refresh_modes_ui()
        self.refresh_files()
        self.schedule_autosave()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def current_mode(self) -> Dict[str, Any]:
        return self.mode_map[self.selected_mode_id.get()]

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=14)
        container.pack(fill=tk.BOTH, expand=True)

        top = ttk.LabelFrame(container, text="Параметры разметки", padding=12)
        top.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(top, text="Режим:", style="Header.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.mode_combo = ttk.Combobox(top, values=[], state="readonly", width=44)
        self.mode_combo.grid(row=0, column=1, sticky="w")
        self.mode_combo.bind("<<ComboboxSelected>>", self.on_mode_changed)

        ttk.Button(top, text="Конструктор режимов", command=self.open_mode_editor).grid(row=0, column=2, padx=(12, 0))
        ttk.Button(top, text="Восстановить backup", command=self.restore_modes_backup).grid(row=0, column=3, padx=(8, 0))
        ttk.Button(top, text="Статистика", command=self.open_stats_window).grid(row=0, column=4, padx=(8, 0))

        ttk.Label(top, text="Имя разметчика:", style="Header.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(10, 0))
        annotator_entry = ttk.Entry(top, textvariable=self.annotator_name_var, width=28)
        annotator_entry.grid(row=1, column=1, sticky="w", pady=(10, 0))
        self._attach_entry_history(annotator_entry)
        ttk.Checkbutton(top, text="Показывать только неразмеченные", variable=self.filter_unlabeled_var, command=self.on_filter_toggled).grid(row=1, column=2, columnspan=2, sticky="w", padx=(12, 0), pady=(10, 0))

        ttk.Label(top, text="Файл из input:", style="Header.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=(10, 0))
        self.file_combo = ttk.Combobox(top, values=[], state="readonly", width=78)
        self.file_combo.grid(row=2, column=1, columnspan=4, sticky="we", pady=(10, 0))
        ttk.Button(top, text="Обновить список", command=self.refresh_files).grid(row=2, column=5, padx=(12, 0), pady=(10, 0))
        ttk.Button(top, text="Загрузить", command=self.load_selected_dataset, style="Accent.TButton").grid(row=2, column=6, padx=(8, 0), pady=(10, 0))

        ttk.Label(top, text="Переход к session_id:", style="Header.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=(10, 0))
        jump_entry = ttk.Entry(top, textvariable=self.jump_session_var, width=28)
        jump_entry.grid(row=3, column=1, sticky="w", pady=(10, 0))
        self._attach_entry_history(jump_entry)
        ttk.Button(top, text="Перейти", command=self.jump_to_session).grid(row=3, column=2, sticky="w", padx=(12, 0), pady=(10, 0))
        ttk.Label(top, textvariable=self.last_autosave_var, style="Muted.TLabel").grid(row=3, column=3, columnspan=4, sticky="e", pady=(10, 0))
        top.columnconfigure(4, weight=1)

        middle = ttk.Panedwindow(container, orient=tk.HORIZONTAL)
        middle.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.LabelFrame(middle, text="Диалог", padding=12)
        right_frame = ttk.LabelFrame(middle, text="Разметка", padding=12)
        middle.add(left_frame, weight=2)
        middle.add(right_frame, weight=2)

        dialogue_header = ttk.Frame(left_frame)
        dialogue_header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(dialogue_header, text="Session ID:", style="Header.TLabel").pack(side=tk.LEFT)
        self.session_value = ttk.Label(dialogue_header, text="-", font=("Segoe UI", 10, "bold"))
        self.session_value.pack(side=tk.LEFT, padx=(6, 20))
        ttk.Label(dialogue_header, text="Прогресс:", style="Header.TLabel").pack(side=tk.LEFT)
        ttk.Label(dialogue_header, textvariable=self.progress_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(6, 18))
        ttk.Label(dialogue_header, textvariable=self.done_var, style="Muted.TLabel").pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(dialogue_header, textvariable=self.remaining_var, style="Muted.TLabel").pack(side=tk.LEFT)

        ttk.Label(left_frame, text="Реплики берутся из поля text и разделяются по символу //", style="Muted.TLabel").pack(anchor="w", pady=(0, 8))
        self.dialogue_text = tk.Text(left_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Segoe UI", 11), padx=14, pady=12, bg="#fbfcfe", relief=tk.FLAT)
        self.dialogue_text.pack(fill=tk.BOTH, expand=True)

        nav = ttk.Frame(left_frame)
        nav.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(nav, text="← Предыдущий", command=self.go_prev).pack(side=tk.LEFT)
        ttk.Button(nav, text="Сохранить черновик (Ctrl+S)", command=self.save_draft).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(nav, text="Следующий (Ctrl+Enter) →", command=self.go_next).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(nav, text="Выгрузить в Excel", command=self.export_results, style="Accent.TButton").pack(side=tk.RIGHT)

        right_top = ttk.Frame(right_frame)
        right_top.pack(fill=tk.BOTH, expand=True)

        guidance_frame = ttk.LabelFrame(right_top, text="Инструкции и примеры", padding=10)
        guidance_frame.pack(fill=tk.X, pady=(0, 10))
        self.guidance_text = tk.Text(guidance_frame, height=7, wrap=tk.WORD, state=tk.DISABLED, bg="#fbfcfe", relief=tk.FLAT)
        self.guidance_text.pack(fill=tk.X)

        form_container = ttk.Frame(right_top)
        form_container.pack(fill=tk.BOTH, expand=True)
        self.form_canvas = tk.Canvas(form_container, highlightthickness=0, bg="#ffffff")
        self.form_scrollbar = ttk.Scrollbar(form_container, orient="vertical", command=self.form_canvas.yview)
        self.form_inner = ttk.Frame(self.form_canvas, padding=(2, 2, 8, 2))
        self.form_inner.bind("<Configure>", lambda e: self.form_canvas.configure(scrollregion=self.form_canvas.bbox("all")))
        self.form_window_id = self.form_canvas.create_window((0, 0), window=self.form_inner, anchor="nw")
        self.form_canvas.bind("<Configure>", self._on_form_canvas_configure, add="+")
        self.form_canvas.configure(yscrollcommand=self.form_scrollbar.set)
        self.form_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.form_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        validation_frame = ttk.LabelFrame(right_frame, text="Проверки", padding=10)
        validation_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(validation_frame, textvariable=self.validation_var, style="Header.TLabel").pack(anchor="w", pady=(0, 6))
        self.validation_list = tk.Listbox(validation_frame, height=5)
        self.validation_list.pack(fill=tk.X)

        bottom = ttk.Frame(container)
        bottom.pack(fill=tk.X, pady=(12, 0))
        ttk.Label(bottom, textvariable=self.status_var, foreground="#1f4e79").pack(side=tk.LEFT)
        ttk.Label(bottom, text="Горячие клавиши: Ctrl+Enter — следующая, Ctrl+S — черновик, Alt+←/→ — навигация", style="Muted.TLabel").pack(side=tk.RIGHT)

        self.rebuild_form()
        self.render_current_record()

    def _bind_hotkeys(self) -> None:
        self.bind_all("<Control-s>", self._hotkey_save, add="+")
        self.bind_all("<Control-Return>", self._hotkey_next, add="+")
        self.bind_all("<Alt-Left>", self._hotkey_prev, add="+")
        self.bind_all("<Alt-Right>", self._hotkey_next, add="+")
        self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

    def _event_in_main_context(self, event: Optional[Any] = None) -> bool:
        modal = self.grab_current()
        if modal is not None and modal is not self:
            return False
        if event is None:
            return True
        widget = getattr(event, "widget", None)
        if widget is None:
            return True
        try:
            return widget.winfo_toplevel() is self
        except Exception:
            return False

    def _widget_is_descendant_of(self, widget: Any, ancestor: tk.Misc) -> bool:
        current = widget
        while current is not None:
            if current is ancestor:
                return True
            try:
                parent_name = current.winfo_parent()
            except Exception:
                return False
            if not parent_name:
                return False
            try:
                current = current.nametowidget(parent_name)
            except Exception:
                return False
        return False

    def _hotkey_save(self, event=None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        self.save_draft()
        return "break"

    def _hotkey_next(self, event=None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        self.go_next()
        return "break"

    def _hotkey_prev(self, event=None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        self.go_prev()
        return "break"

    def _on_mousewheel(self, event: Any) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        widget = self.winfo_containing(event.x_root, event.y_root)
        if widget is None:
            return None
        if not (widget is self.form_canvas or self._widget_is_descendant_of(widget, self.form_inner)):
            return None
        delta = getattr(event, "delta", 0)
        if not delta:
            return None
        self.form_canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        return "break"

    def _attach_global_edit_support(self) -> None:
        self.bind_all("<Control-KeyPress>", self._edit_control_keypress, add="+")
        self.bind_all("<Control-Insert>", self._edit_copy, add="+")
        self.bind_all("<Shift-Insert>", self._edit_paste, add="+")
        self.bind_all("<Shift-Delete>", self._edit_cut, add="+")
        for class_name in ("Entry", "TEntry", "TCombobox", "Text"):
            self.bind_class(class_name, "<Button-3>", self._show_edit_menu, add="+")

    def _control_action_from_event(self, event: Any) -> Optional[str]:
        keysym = (getattr(event, "keysym", "") or "").lower()
        keycode = getattr(event, "keycode", None)
        keysym_map = {
            "c": "copy",
            "с": "copy",
            "x": "cut",
            "ч": "cut",
            "v": "paste",
            "м": "paste",
            "a": "select_all",
            "ф": "select_all",
            "z": "undo",
            "я": "undo",
            "y": "redo",
            "н": "redo",
        }
        keycode_map = {
            65: "select_all",
            67: "copy",
            86: "paste",
            88: "cut",
            89: "redo",
            90: "undo",
        }
        if keysym in keysym_map:
            return keysym_map[keysym]
        return keycode_map.get(keycode)

    def _edit_control_keypress(self, event: Optional[Any] = None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        widget = self._editable_widget(event)
        if widget is None:
            return None
        action = self._control_action_from_event(event)
        if action == "copy":
            return self._edit_copy(event)
        if action == "cut":
            return self._edit_cut(event)
        if action == "paste":
            return self._edit_paste(event)
        if action == "select_all":
            return self._edit_select_all(event)
        if action == "undo":
            return self._edit_undo(event)
        if action == "redo":
            return self._edit_redo(event)
        return None

    def _is_edit_widget(self, widget: Any) -> bool:
        return isinstance(widget, (tk.Entry, ttk.Entry, ttk.Combobox, tk.Text))

    def _editable_widget(self, event: Optional[Any] = None) -> Optional[Any]:
        widget = None
        if event is not None and self._is_edit_widget(getattr(event, "widget", None)):
            widget = event.widget
        if widget is None:
            widget = self.focus_get()
        if self._is_edit_widget(widget):
            return widget
        return None

    def _widget_is_writable(self, widget: Any) -> bool:
        try:
            state = str(widget.cget("state"))
        except Exception:
            return True
        return state not in {"disabled", "readonly"}

    def _get_widget_text(self, widget: Any) -> str:
        if isinstance(widget, tk.Text):
            return widget.get("1.0", tk.END).rstrip("\n")
        return widget.get()

    def _set_widget_text(self, widget: Any, value: str) -> None:
        if isinstance(widget, tk.Text):
            widget.delete("1.0", tk.END)
            widget.insert("1.0", value)
            try:
                widget.edit_separator()
            except tk.TclError:
                pass
            return
        prev_state = None
        try:
            prev_state = str(widget.cget("state"))
        except Exception:
            prev_state = None
        if prev_state == "readonly":
            widget.configure(state="normal")
        widget.delete(0, tk.END)
        widget.insert(0, value)
        try:
            widget.icursor(tk.END)
            widget.selection_clear()
        except Exception:
            pass
        if prev_state == "readonly":
            widget.configure(state="readonly")

    def _reset_entry_history(self, widget: Any) -> None:
        self.entry_edit_state[str(widget)] = {"undo": [self._get_widget_text(widget)], "redo": []}

    def _capture_entry_history(self, widget: Any) -> None:
        key = str(widget)
        value = self._get_widget_text(widget)
        state = self.entry_edit_state.setdefault(key, {"undo": [value], "redo": []})
        undo_stack = state["undo"]
        if not undo_stack or undo_stack[-1] != value:
            undo_stack.append(value)
            if len(undo_stack) > 120:
                del undo_stack[:-120]
            state["redo"] = []

    def _attach_entry_history(self, widget: Any) -> None:
        self._reset_entry_history(widget)
        widget.bind("<KeyRelease>", lambda e, w=widget: self.after_idle(lambda: self._capture_entry_history(w)), add="+")
        widget.bind("<FocusIn>", lambda e, w=widget: self.after_idle(lambda: self._capture_entry_history(w)), add="+")

    def _perform_copy(self, widget: Any) -> None:
        try:
            widget.event_generate("<<Copy>>")
        except tk.TclError:
            pass

    def _perform_cut(self, widget: Any) -> None:
        if not self._widget_is_writable(widget):
            return
        try:
            widget.event_generate("<<Cut>>")
        except tk.TclError:
            return
        if not isinstance(widget, tk.Text):
            self.after_idle(lambda w=widget: self._capture_entry_history(w))

    def _perform_paste(self, widget: Any) -> None:
        if not self._widget_is_writable(widget):
            return
        try:
            widget.event_generate("<<Paste>>")
        except tk.TclError:
            return
        if not isinstance(widget, tk.Text):
            self.after_idle(lambda w=widget: self._capture_entry_history(w))

    def _perform_select_all(self, widget: Any) -> None:
        if isinstance(widget, tk.Text):
            widget.tag_add(tk.SEL, "1.0", "end-1c")
            widget.mark_set(tk.INSERT, "1.0")
            widget.see(tk.INSERT)
            return
        widget.selection_range(0, tk.END)
        widget.icursor(tk.END)

    def _perform_undo(self, widget: Any) -> None:
        if isinstance(widget, tk.Text):
            if not self._widget_is_writable(widget):
                return
            try:
                widget.edit_undo()
            except tk.TclError:
                pass
            return
        key = str(widget)
        state = self.entry_edit_state.get(key)
        if not state or len(state["undo"]) <= 1:
            return
        current = state["undo"].pop()
        state["redo"].append(current)
        self._set_widget_text(widget, state["undo"][-1])

    def _perform_redo(self, widget: Any) -> None:
        if isinstance(widget, tk.Text):
            if not self._widget_is_writable(widget):
                return
            try:
                widget.edit_redo()
            except tk.TclError:
                pass
            return
        key = str(widget)
        state = self.entry_edit_state.get(key)
        if not state or not state["redo"]:
            return
        value = state["redo"].pop()
        state["undo"].append(value)
        self._set_widget_text(widget, value)

    def _edit_copy(self, event: Optional[Any] = None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        widget = self._editable_widget(event)
        if widget is None:
            return None
        self._perform_copy(widget)
        return "break"

    def _edit_cut(self, event: Optional[Any] = None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        widget = self._editable_widget(event)
        if widget is None:
            return None
        self._perform_cut(widget)
        return "break"

    def _edit_paste(self, event: Optional[Any] = None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        widget = self._editable_widget(event)
        if widget is None:
            return None
        self._perform_paste(widget)
        return "break"

    def _edit_select_all(self, event: Optional[Any] = None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        widget = self._editable_widget(event)
        if widget is None:
            return None
        self._perform_select_all(widget)
        return "break"

    def _edit_undo(self, event: Optional[Any] = None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        widget = self._editable_widget(event)
        if widget is None:
            return None
        self._perform_undo(widget)
        return "break"

    def _edit_redo(self, event: Optional[Any] = None) -> Optional[str]:
        if not self._event_in_main_context(event):
            return None
        widget = self._editable_widget(event)
        if widget is None:
            return None
        self._perform_redo(widget)
        return "break"

    def _show_edit_menu(self, event: Any) -> Optional[str]:
        widget = self._editable_widget(event)
        if widget is None:
            return None
        widget.focus_set()
        menu = tk.Menu(self, tearoff=0)
        writable = self._widget_is_writable(widget)
        menu.add_command(label="Отменить\tCtrl+Z", command=lambda w=widget: self._perform_undo(w), state=("normal" if writable else "disabled"))
        menu.add_command(label="Повторить\tCtrl+Y", command=lambda w=widget: self._perform_redo(w), state=("normal" if writable else "disabled"))
        menu.add_separator()
        menu.add_command(label="Вырезать\tCtrl+X", command=lambda w=widget: self._perform_cut(w), state=("normal" if writable else "disabled"))
        menu.add_command(label="Копировать\tCtrl+C", command=lambda w=widget: self._perform_copy(w))
        menu.add_command(label="Вставить\tCtrl+V", command=lambda w=widget: self._perform_paste(w), state=("normal" if writable else "disabled"))
        menu.add_separator()
        menu.add_command(label="Выделить всё\tCtrl+A", command=lambda w=widget: self._perform_select_all(w))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def refresh_modes_ui(self) -> None:
        self.modes = self.modes_payload["modes"]
        self.mode_map = {mode["id"]: mode for mode in self.modes}
        values = [f"{mode['name']} ({mode['id']}) v{mode.get('version', '')}" for mode in self.modes]
        self.mode_combo["values"] = values
        mode_ids = [mode["id"] for mode in self.modes]
        selected_id = self.selected_mode_id.get()
        if selected_id not in mode_ids:
            selected_id = self.modes[0]["id"]
            self.selected_mode_id.set(selected_id)
        idx = mode_ids.index(selected_id)
        self.mode_combo.current(idx)
        self.prev_mode_id = selected_id
        self.rebuild_form()
        self.update_guidance_panel()

    def _on_form_canvas_configure(self, event: Any) -> None:
        try:
            self.form_canvas.itemconfigure(self.form_window_id, width=max(200, int(event.width) - 2))
        except Exception:
            pass

    def restore_modes_backup(self) -> None:
        backup_path = latest_modes_backup(self.backups_dir)
        if backup_path is None:
            messagebox.showwarning("Backup не найден", "В папке config/_history нет резервных копий modes.json.", parent=self)
            return
        if not messagebox.askyesno("Восстановить backup", "Восстановить последний backup modes.json? Текущий modes.json будет заменён.", parent=self):
            return
        try:
            restored = restore_latest_modes_backup(self.modes_path, self.backups_dir)
            self.modes_payload = load_modes(self.modes_path)
        except Exception as exc:
            messagebox.showerror("Ошибка восстановления", str(exc), parent=self)
            return
        self.refresh_modes_ui()
        self.status_var.set(f"Восстановлен backup: {Path(restored).name}")
        messagebox.showinfo("Готово", f"Восстановлен backup: {Path(restored).name}", parent=self)

    def refresh_files(self) -> None:
        files = list_input_files(self.input_dir)
        self.file_display_map = {}
        values: List[str] = []
        preferred_display = None
        preferred_name = self.settings.get("last_input_file")
        for path in files:
            status = inspect_input_file(path)
            display = path.name
            if status.warning:
                display = u"⚠ {0} — подозрительный формат".format(path.name)
            self.file_display_map[display] = path
            values.append(display)
            if preferred_name and path.name == preferred_name:
                preferred_display = display
        self.file_combo["values"] = values
        if preferred_display in values:
            self.file_combo.current(values.index(preferred_display))
        elif values:
            self.file_combo.current(0)
        self.status_var.set(f"Файлов в input: {len(values)}")

    def on_mode_changed(self, event=None) -> None:
        idx = self.mode_combo.current()
        if idx < 0:
            return
        new_mode = self.modes[idx]
        new_mode_id = new_mode["id"]
        if self.records and self.current_input_path is not None and new_mode_id != self.prev_mode_id:
            if not messagebox.askyesno(
                "Смена режима",
                "Сейчас загружен файл. При смене режима текущие загруженные данные и черновик в памяти будут пересозданы для нового режима. Продолжить?",
                parent=self,
            ):
                prev_idx = [mode["id"] for mode in self.modes].index(self.prev_mode_id)
                self.mode_combo.current(prev_idx)
                return
            self.selected_mode_id.set(new_mode_id)
            self.prev_mode_id = new_mode_id
            self.annotations = {}
            self.widget_bindings.clear()
            self.rebuild_form()
            self.update_guidance_panel()
            self.load_selected_dataset(reload_same_file=True)
            return
        self.selected_mode_id.set(new_mode_id)
        self.prev_mode_id = new_mode_id
        self.rebuild_form()
        self.update_guidance_panel()
        self.render_current_record()
        self.status_var.set(f"Выбран режим: {new_mode['name']}")

    def update_guidance_panel(self) -> None:
        mode = self.current_mode()
        lines = []
        if mode.get("description"):
            lines.append(f"Описание: {mode['description']}")
        if mode.get("instructions"):
            lines.append(f"Инструкции: {mode['instructions']}")
        examples = mode.get("examples", [])
        if examples:
            lines.append("Примеры:")
            lines.extend([f"• {item}" for item in examples])
        text = "\n".join(lines) if lines else "Для этого режима нет дополнительных инструкций."
        self.guidance_text.config(state=tk.NORMAL)
        self.guidance_text.delete("1.0", tk.END)
        self.guidance_text.insert("1.0", text)
        self.guidance_text.config(state=tk.DISABLED)

    def rebuild_form(self) -> None:
        for widget in self.form_inner.winfo_children():
            widget.destroy()
        self.widget_bindings.clear()

        mode = self.current_mode()
        for field in mode.get("fields", []):
            self._build_field_widget(field)

        meta_frame = ttk.LabelFrame(self.form_inner, text="Метаданные записи", padding=10)
        meta_frame.pack(fill=tk.X, pady=(8, 12))
        ttk.Label(meta_frame, text="Уверенность").grid(row=0, column=0, sticky="w")
        confidence_var = tk.StringVar(value="")
        confidence_widget = ttk.Combobox(meta_frame, textvariable=confidence_var, values=["", "Высокая", "Средняя", "Низкая"], state="readonly", width=20)
        confidence_widget.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=4)
        confidence_var.trace_add("write", lambda *_: self.mark_dirty())

        needs_review_var = tk.BooleanVar(value=False)
        needs_review_widget = ttk.Checkbutton(meta_frame, text="Нужна проверка", variable=needs_review_var, command=self.mark_dirty)
        needs_review_widget.grid(row=0, column=2, sticky="w", padx=(16, 0), pady=4)

        ttk.Label(meta_frame, text="Общий комментарий").grid(row=1, column=0, sticky="nw")
        comment_text = tk.Text(meta_frame, height=4, width=52, undo=True, autoseparators=True, maxundo=-1)
        comment_text.grid(row=1, column=1, columnspan=2, sticky="we", padx=(8, 0), pady=4)
        comment_text.bind("<KeyRelease>", lambda e: self.mark_dirty())
        comment_text.bind("<FocusOut>", lambda e: self.update_validation_panel())
        meta_frame.columnconfigure(2, weight=1)

        self.widget_bindings["__meta__confidence"] = {"type": "select", "var": confidence_var, "widget": confidence_widget}
        self.widget_bindings["__meta__needs_review"] = {"type": "checkbox", "var": needs_review_var, "widget": needs_review_widget}
        self.widget_bindings["__meta__annotator_comment"] = {"type": "textarea", "widget": comment_text}

    def _build_field_widget(self, field: Dict[str, Any]) -> None:
        frame = ttk.Frame(self.form_inner)
        frame.pack(fill=tk.X, pady=(0, 12))
        label_text = field["label"] + (" *" if field.get("required") else "")
        ttk.Label(frame, text=label_text, style="Header.TLabel").pack(anchor="w", pady=(0, 4))
        if field.get("help"):
            ttk.Label(frame, text=field["help"], style="Muted.TLabel", wraplength=620, justify=tk.LEFT).pack(anchor="w")
        if field.get("examples"):
            ttk.Label(frame, text="Например: " + " | ".join(field["examples"]), style="Muted.TLabel", wraplength=620, justify=tk.LEFT).pack(anchor="w", pady=(0, 4))

        field_type = field["type"]
        if field_type in {"text", "number"}:
            var = tk.StringVar(value="")
            row = ttk.Frame(frame)
            row.pack(fill=tk.X, pady=(4, 0))
            widget = ttk.Combobox(row, textvariable=var)
            widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self._attach_entry_history(widget)
            widget.bind("<KeyRelease>", lambda e, key=field["key"]: self.refresh_text_suggestions(key))
            widget.bind("<FocusIn>", lambda e, key=field["key"]: self.refresh_text_suggestions(key))
            var.trace_add("write", lambda *_: self.mark_dirty())
            history_button = ttk.Button(row, text="История", command=lambda key=field["key"], label=field["label"]: self.open_history_dialog(key, label))
            history_button.pack(side=tk.LEFT, padx=(8, 0))
            self.widget_bindings[field["key"]] = {
                "type": field_type,
                "var": var,
                "widget": widget,
                "field": field,
                "controls": [history_button],
            }
        elif field_type == "textarea":
            suggestion_row = ttk.Frame(frame)
            suggestion_row.pack(fill=tk.X, pady=(4, 2))
            sugg_var = tk.StringVar(value="")
            sugg_widget = ttk.Combobox(suggestion_row, textvariable=sugg_var)
            sugg_widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self._attach_entry_history(sugg_widget)
            sugg_widget.bind("<KeyRelease>", lambda e, key=field["key"]: self.refresh_text_suggestions(key, for_textarea=True))
            sugg_widget.bind("<FocusIn>", lambda e, key=field["key"]: self.refresh_text_suggestions(key, for_textarea=True))
            insert_button = ttk.Button(suggestion_row, text="Вставить", command=lambda key=field["key"]: self.apply_textarea_suggestion(key))
            insert_button.pack(side=tk.LEFT, padx=(8, 0))
            history_button = ttk.Button(suggestion_row, text="История", command=lambda key=field["key"], label=field["label"]: self.open_history_dialog(key, label))
            history_button.pack(side=tk.LEFT, padx=(8, 0))
            widget = tk.Text(frame, height=5, wrap=tk.WORD, undo=True, autoseparators=True, maxundo=-1)
            widget.pack(fill=tk.X, pady=(0, 2))
            widget.bind("<KeyRelease>", lambda e: self.mark_dirty())
            widget.bind("<FocusOut>", lambda e: self.update_validation_panel())
            self.widget_bindings[field["key"]] = {
                "type": field_type,
                "widget": widget,
                "field": field,
                "suggestion_var": sugg_var,
                "suggestion_widget": sugg_widget,
                "controls": [insert_button, history_button],
            }
        elif field_type == "select":
            var = tk.StringVar(value="")
            widget = ttk.Combobox(frame, textvariable=var, values=field.get("options", []), state="readonly")
            widget.pack(fill=tk.X, pady=(4, 0))
            var.trace_add("write", lambda *_: self.mark_dirty())
            self.widget_bindings[field["key"]] = {"type": field_type, "var": var, "widget": widget, "field": field}
        elif field_type == "multiselect":
            list_frame = ttk.Frame(frame)
            list_frame.pack(fill=tk.X, pady=(4, 0))
            vars_list: List[Tuple[str, tk.BooleanVar]] = []
            widgets_list: List[ttk.Checkbutton] = []
            for idx, option in enumerate(field.get("options", [])):
                var = tk.BooleanVar(value=False)
                cb = ttk.Checkbutton(list_frame, text=option, variable=var, command=self.mark_dirty)
                cb.grid(row=idx // 2, column=idx % 2, sticky="w", padx=(0, 18), pady=2)
                vars_list.append((option, var))
                widgets_list.append(cb)
            self.widget_bindings[field["key"]] = {"type": field_type, "vars": vars_list, "widgets": widgets_list, "field": field}
        elif field_type == "checkbox":
            var = tk.BooleanVar(value=False)
            widget = ttk.Checkbutton(frame, text="Да", variable=var, command=self.mark_dirty)
            widget.pack(anchor="w", pady=(4, 0))
            self.widget_bindings[field["key"]] = {"type": field_type, "var": var, "widget": widget, "field": field}

    def refresh_text_suggestions(self, field_key: str, for_textarea: bool = False) -> None:
        binding = self.widget_bindings.get(field_key)
        if not binding:
            return
        prefix = ""
        if for_textarea:
            prefix = binding["suggestion_var"].get()
            widget = binding["suggestion_widget"]
        else:
            prefix = binding["var"].get()
            widget = binding["widget"]
        suggestions = get_field_suggestions(
            self.history_payload,
            self.current_file_history,
            self.current_mode()["id"],
            field_key,
            prefix=prefix,
            max_items=40,
        )
        widget["values"] = suggestions

    def open_history_dialog(self, field_key: str, field_label: str) -> None:
        grouped = get_grouped_history(self.history_payload, self.current_file_history, self.current_mode()["id"], field_key)

        def _apply(value: str) -> None:
            binding = self.widget_bindings[field_key]
            if binding["type"] in {"text", "number"}:
                self._capture_entry_history(binding["widget"])
                binding["var"].set(value)
                self._capture_entry_history(binding["widget"])
            elif binding["type"] == "textarea":
                binding["widget"].delete("1.0", tk.END)
                binding["widget"].insert("1.0", value)
                try:
                    binding["widget"].edit_separator()
                except tk.TclError:
                    pass
            self.mark_dirty()

        HistoryDialog(self, field_label, grouped, _apply)

    def apply_textarea_suggestion(self, field_key: str) -> None:
        binding = self.widget_bindings[field_key]
        value = binding["suggestion_var"].get().strip()
        if not value:
            return
        binding["widget"].delete("1.0", tk.END)
        binding["widget"].insert("1.0", value)
        try:
            binding["widget"].edit_separator()
        except tk.TclError:
            pass
        self.mark_dirty()

    def load_selected_dataset(self, reload_same_file: bool = False) -> None:
        selected = self.file_combo.get().strip()
        if not selected:
            messagebox.showwarning("Нет файла", "Выберите файл из списка.", parent=self)
            return
        file_path = self.file_display_map.get(selected, self.input_dir / selected)
        if not file_path.exists():
            messagebox.showerror("Файл не найден", f"Файл {selected} не найден в input.", parent=self)
            return
        if self.records and self.dirty and not reload_same_file:
            self.save_draft(silent=True)
        try:
            records, _ = load_dialogues(file_path)
        except InputFileFormatError as exc:
            if exc.can_try_csv and messagebox.askyesno(
                "Похоже на CSV под видом XLSX",
                str(exc) + "\n\nОткрыть этот файл как CSV с разделителем ';'?",
                parent=self,
            ):
                try:
                    records, _ = load_dialogues(file_path, treat_as_csv=True)
                except Exception as inner_exc:
                    messagebox.showerror("Ошибка загрузки", str(inner_exc), parent=self)
                    return
            else:
                messagebox.showerror("Ошибка загрузки", str(exc), parent=self)
                return
        except Exception as exc:
            messagebox.showerror("Ошибка загрузки", str(exc), parent=self)
            return
        self.records = records
        self.current_input_path = file_path
        self.annotations = {}
        self.current_index = 0
        self.filtered_indices = []
        draft_path = build_draft_path(self.output_dir, file_path, self.current_mode()["id"])
        draft = load_draft_json(draft_path)
        if draft:
            if messagebox.askyesno("Черновик найден", "Найден черновик для этого файла и режима. Восстановить его?", parent=self):
                self.annotations = normalize_annotations_for_records(self.records, draft.get("annotations", {}))
                self.current_index = int(draft.get("current_index", 0))
                self.annotator_name_var.set(draft.get("annotator_name", self.annotator_name_var.get()))
                self.filter_unlabeled_var.set(bool(draft.get("filter_unlabeled", False)))
        self.current_file_history = build_current_file_history(self.current_mode().get("fields", []), self.annotations)
        self.recompute_filtered_indices()
        if self.filter_unlabeled_var.get() and self.filtered_indices:
            self.current_index = self.filtered_indices[0]
        self.dirty = False
        self.update_status_counts()
        self.render_current_record()
        self.status_var.set(f"Загружен файл: {file_path.name}. Записей: {len(self.records)}")

    def render_current_record(self) -> None:
        if not self.records:
            self.session_value.config(text="-")
            self._render_dialogue([])
            self._clear_form_values()
            self.update_validation_panel()
            self.update_status_counts()
            return
        self.current_index = max(0, min(self.current_index, len(self.records) - 1))
        record = self.records[self.current_index]
        self.session_value.config(text=record.session_id)
        self._render_dialogue(split_dialogue_text(record.text))
        values = self.annotations.get(record.annotation_key)
        if values is None:
            values = empty_annotation(self.current_mode())
            self.annotations[record.annotation_key] = values
        self._populate_form_values(values)
        self.update_status_counts()
        self.update_validation_panel()

    def _render_dialogue(self, replicas: List[str]) -> None:
        self.dialogue_text.config(state=tk.NORMAL)
        self.dialogue_text.delete("1.0", tk.END)
        if not replicas:
            self.dialogue_text.insert("1.0", "Нет данных для отображения.")
        else:
            lines = [f"- {replica}" for replica in replicas]
            self.dialogue_text.insert("1.0", "\n".join(lines))
        self.dialogue_text.config(state=tk.DISABLED)

    def _clear_form_values(self) -> None:
        self.suspend_reactive_updates = True
        try:
            for key, binding in self.widget_bindings.items():
                if key.startswith("__meta__"):
                    continue
                self._set_binding_value(binding, "")
            self.widget_bindings["__meta__confidence"]["var"].set("")
            self.widget_bindings["__meta__needs_review"]["var"].set(False)
            comment_widget = self.widget_bindings["__meta__annotator_comment"]["widget"]
            comment_widget.delete("1.0", tk.END)
        finally:
            self.suspend_reactive_updates = False
        self._apply_rule_driven_field_states(allow_dirty=False)

    def _populate_form_values(self, values: Dict[str, Any]) -> None:
        self.suspend_reactive_updates = True
        try:
            for field in self.current_mode().get("fields", []):
                self._set_binding_value(self.widget_bindings[field["key"]], values.get(field["key"], field.get("default", "")))
            self.widget_bindings["__meta__confidence"]["var"].set(values.get("confidence", ""))
            self.widget_bindings["__meta__needs_review"]["var"].set(bool(values.get("needs_review", False)))
            comment_widget = self.widget_bindings["__meta__annotator_comment"]["widget"]
            comment_widget.delete("1.0", tk.END)
            comment_widget.insert("1.0", values.get("annotator_comment", ""))
            try:
                comment_widget.edit_reset()
            except tk.TclError:
                pass
        finally:
            self.suspend_reactive_updates = False
        self._apply_rule_driven_field_states()

    def _set_binding_value(self, binding: Dict[str, Any], value: Any) -> None:
        field_type = binding["type"]
        if field_type in {"text", "number", "select"}:
            binding["var"].set("" if value is None else str(value))
            self._reset_entry_history(binding["widget"])
        elif field_type == "textarea":
            prev_state = str(binding["widget"].cget("state"))
            if prev_state == "disabled":
                binding["widget"].configure(state="normal")
            binding["widget"].delete("1.0", tk.END)
            binding["widget"].insert("1.0", "" if value is None else str(value))
            binding["suggestion_var"].set("")
            try:
                binding["widget"].edit_reset()
            except tk.TclError:
                pass
            if prev_state == "disabled":
                binding["widget"].configure(state="disabled")
        elif field_type == "checkbox":
            binding["var"].set(bool(value))
        elif field_type == "multiselect":
            selected = set(value if isinstance(value, list) else [])
            for option, var in binding["vars"]:
                var.set(option in selected)

    def _set_widget_state(self, widget: Any, state: str) -> None:
        if widget is None:
            return
        try:
            widget.configure(state=state)
        except Exception:
            pass

    def _set_binding_disabled_state(self, binding: Dict[str, Any], disabled: bool) -> None:
        field_type = binding["type"]
        if field_type in {"text", "number"}:
            self._set_widget_state(binding.get("widget"), "disabled" if disabled else "normal")
        elif field_type == "textarea":
            self._set_widget_state(binding.get("suggestion_widget"), "disabled" if disabled else "normal")
            self._set_widget_state(binding.get("widget"), "disabled" if disabled else "normal")
        elif field_type == "select":
            self._set_widget_state(binding.get("widget"), "disabled" if disabled else "readonly")
        elif field_type == "checkbox":
            self._set_widget_state(binding.get("widget"), "disabled" if disabled else "normal")
        elif field_type == "multiselect":
            for widget in binding.get("widgets", []):
                self._set_widget_state(widget, "disabled" if disabled else "normal")
        for control in binding.get("controls", []):
            self._set_widget_state(control, "disabled" if disabled else "normal")

    def _value_is_filled(self, value: Any) -> bool:
        if value is None:
            return False
        if isinstance(value, bool):
            return value
        if isinstance(value, (list, tuple, set, dict)):
            return bool(value)
        return str(value).strip() != ""

    def _apply_rule_driven_field_states(self, allow_dirty: bool = True) -> None:
        if self.suspend_reactive_updates or not self.widget_bindings:
            return
        changed = False
        final_disabled = set()
        self.suspend_reactive_updates = True
        try:
            while True:
                current_values = self.collect_current_values()
                disabled_now = set(disabled_field_keys(self.current_mode(), current_values))
                newly_disabled = [field_key for field_key in disabled_now if field_key not in final_disabled]
                if not newly_disabled:
                    final_disabled = disabled_now
                    break
                final_disabled = disabled_now
                for field_key in newly_disabled:
                    binding = self.widget_bindings.get(field_key)
                    if not binding:
                        continue
                    if self._value_is_filled(current_values.get(field_key)):
                        self._set_binding_value(binding, "")
                        changed = True
            final_disabled = set(disabled_field_keys(self.current_mode(), self.collect_current_values()))
            for field in self.current_mode().get("fields", []):
                key = field["key"]
                binding = self.widget_bindings.get(key)
                if not binding:
                    continue
                self._set_binding_disabled_state(binding, key in final_disabled)
        finally:
            self.suspend_reactive_updates = False
        self.currently_disabled_fields = sorted(final_disabled)
        if changed and allow_dirty:
            self.dirty = True

    def collect_current_values(self) -> Dict[str, Any]:
        values: Dict[str, Any] = {}
        for field in self.current_mode().get("fields", []):
            binding = self.widget_bindings[field["key"]]
            field_type = binding["type"]
            if field_type in {"text", "number", "select"}:
                values[field["key"]] = binding["var"].get().strip()
            elif field_type == "textarea":
                values[field["key"]] = binding["widget"].get("1.0", tk.END).strip()
            elif field_type == "checkbox":
                values[field["key"]] = bool(binding["var"].get())
            elif field_type == "multiselect":
                values[field["key"]] = [option for option, var in binding["vars"] if var.get()]
        values["confidence"] = self.widget_bindings["__meta__confidence"]["var"].get().strip()
        values["needs_review"] = bool(self.widget_bindings["__meta__needs_review"]["var"].get())
        values["annotator_comment"] = self.widget_bindings["__meta__annotator_comment"]["widget"].get("1.0", tk.END).strip()
        values["updated_at"] = datetime.now().isoformat(timespec="seconds")
        return values

    def save_current_annotation(self, interactive: bool = False) -> bool:
        if not self.records:
            return True
        record = self.records[self.current_index]
        self._apply_rule_driven_field_states()
        values = self.collect_current_values()
        issues = validate_annotation(self.current_mode(), values)
        values["completed"] = is_annotation_complete(self.current_mode(), values)
        self.annotations[record.annotation_key] = values
        self.current_file_history = build_current_file_history(self.current_mode().get("fields", []), self.annotations)
        self.update_validation_panel(issues)
        self.update_status_counts()
        self.dirty = False
        if interactive and any(issue.level == "error" for issue in issues):
            messages = "\n".join(f"• {issue.message}" for issue in issues if issue.level == "error")
            return messagebox.askyesno(
                "Есть ошибки валидации",
                f"У текущей записи есть ошибки:\n\n{messages}\n\nПерейти всё равно?",
                parent=self,
            )
        if interactive and any(issue.level == "warning" for issue in issues):
            messages = "\n".join(f"• {issue.message}" for issue in issues if issue.level == "warning")
            return messagebox.askyesno(
                "Есть предупреждения",
                f"У текущей записи есть предупреждения:\n\n{messages}\n\nПерейти всё равно?",
                parent=self,
            )
        return True

    def update_validation_panel(self, issues: Optional[List[ValidationIssue]] = None) -> None:
        if issues is None:
            if self.records:
                issues = validate_annotation(self.current_mode(), self.collect_current_values())
            else:
                issues = []
        self.validation_list.delete(0, tk.END)
        if not issues:
            self.validation_var.set("Проверка: всё в порядке")
            self.validation_list.insert(tk.END, "Ошибок и предупреждений нет.")
            return
        error_count = sum(1 for issue in issues if issue.level == "error")
        warning_count = sum(1 for issue in issues if issue.level == "warning")
        self.validation_var.set(f"Проверка: ошибок {error_count}, предупреждений {warning_count}")
        for issue in issues:
            prefix = "[ERROR]" if issue.level == "error" else "[WARN]"
            self.validation_list.insert(tk.END, f"{prefix} {issue.message}")

    def recompute_filtered_indices(self) -> None:
        if not self.records:
            self.filtered_indices = []
            return
        if not self.filter_unlabeled_var.get():
            self.filtered_indices = list(range(len(self.records)))
            return
        filtered: List[int] = []
        for idx, record in enumerate(self.records):
            ann = self.annotations.get(record.annotation_key)
            if ann is None or not bool(ann.get("completed", False)):
                filtered.append(idx)
        self.filtered_indices = filtered

    def on_filter_toggled(self) -> None:
        if self.records:
            self.save_current_annotation(interactive=False)
        self.recompute_filtered_indices()
        if self.filter_unlabeled_var.get() and self.records:
            current_ann = self.annotations.get(self.records[self.current_index].annotation_key, {})
            if current_ann.get("completed"):
                if self.filtered_indices:
                    self.current_index = self.filtered_indices[0]
                else:
                    self.status_var.set("Неразмеченных записей не осталось.")
        self.render_current_record()

    def _navigation_indices(self) -> List[int]:
        if not self.records:
            return []
        if self.filter_unlabeled_var.get():
            return self.filtered_indices
        return list(range(len(self.records)))

    def _move_relative(self, step: int) -> None:
        if not self.records:
            return
        if not self.save_current_annotation(interactive=True):
            return
        nav_indices = self._navigation_indices()
        if not nav_indices:
            self.status_var.set("В текущем фильтре записей нет.")
            return
        if self.current_index not in nav_indices:
            candidate = nav_indices[0 if step > 0 else -1]
        else:
            pos = nav_indices.index(self.current_index)
            new_pos = max(0, min(len(nav_indices) - 1, pos + step))
            candidate = nav_indices[new_pos]
        self.current_index = candidate
        self.render_current_record()

    def go_next(self) -> None:
        self._move_relative(1)

    def go_prev(self) -> None:
        self._move_relative(-1)

    def jump_to_session(self) -> None:
        target = self.jump_session_var.get().strip()
        if not target:
            return
        if not self.records:
            return
        for idx, record in enumerate(self.records):
            if record.session_id == target:
                ann = self.annotations.get(record.annotation_key, {})
                if self.filter_unlabeled_var.get() and ann.get("completed"):
                    self.filter_unlabeled_var.set(False)
                    self.recompute_filtered_indices()
                    self.status_var.set("Фильтр отключён, потому что запись уже размечена.")
                if self.dirty:
                    self.save_current_annotation(interactive=False)
                self.current_index = idx
                self.render_current_record()
                return
        messagebox.showinfo("Не найдено", f"session_id {target} не найден в текущем файле.", parent=self)

    def update_status_counts(self) -> None:
        total = len(self.records)
        completed = 0
        needs_review = 0
        for record in self.records:
            ann = self.annotations.get(record.annotation_key, {})
            if ann.get("completed"):
                completed += 1
            if ann.get("needs_review"):
                needs_review += 1
        remaining = max(0, total - completed)
        current_position = self.current_index + 1 if total else 0
        if self.filter_unlabeled_var.get() and self.records:
            visible_total = len(self.filtered_indices)
            current_position = self.filtered_indices.index(self.current_index) + 1 if self.current_index in self.filtered_indices and self.filtered_indices else 0
            self.progress_var.set(f"{current_position} / {visible_total} (фильтр)")
        else:
            self.progress_var.set(f"{current_position} / {total}")
        self.done_var.set(f"Размечено: {completed}")
        self.remaining_var.set(f"Осталось: {remaining}; на проверку: {needs_review}")

    def mark_dirty(self, *_args) -> None:
        if self.suspend_reactive_updates:
            return
        self.dirty = True
        self._apply_rule_driven_field_states()
        self.update_validation_panel()

    def save_draft(self, silent: bool = False) -> None:
        if not self.current_input_path or not self.records:
            return
        self.save_current_annotation(interactive=False)
        draft_path = build_draft_path(self.output_dir, self.current_input_path, self.current_mode()["id"])
        save_draft_json(
            draft_path,
            self.current_input_path,
            self.current_mode(),
            self.annotations,
            self.current_index,
            self.annotator_name_var.get().strip(),
            self.filter_unlabeled_var.get(),
        )
        self.history_payload = update_history_from_annotations(
            self.history_payload,
            self.current_mode()["id"],
            self.current_mode().get("fields", []),
            self.annotations,
            recent_limit=int(self.settings.get("history_recent_limit", 20)),
            global_limit=int(self.settings.get("history_global_limit", 200)),
        )
        save_history(self.meta_dir, self.history_payload)
        self.last_autosave_var.set(f"Сохранено: {datetime.now().strftime('%H:%M:%S')}")
        if not silent:
            self.status_var.set(f"Черновик сохранён: {draft_path.name}")

    def schedule_autosave(self) -> None:
        if self.autosave_job is not None:
            self.after_cancel(self.autosave_job)
        interval = max(5, int(self.settings.get("autosave_seconds", 30))) * 1000
        self.autosave_job = self.after(interval, self._autosave_tick)

    def _autosave_tick(self) -> None:
        try:
            if self.dirty and self.current_input_path and self.records:
                self.save_draft(silent=True)
                self.last_autosave_var.set(f"Автосохранение: {datetime.now().strftime('%H:%M:%S')}")
        finally:
            self.schedule_autosave()

    def export_results(self) -> None:
        if not self.current_input_path or not self.records:
            messagebox.showwarning("Нет данных", "Сначала загрузите файл и начните разметку.", parent=self)
            return
        self.save_current_annotation(interactive=False)
        summary = self.compute_summary()
        if summary["remaining_records"] > 0:
            if not messagebox.askyesno(
                "Есть неразмеченные записи",
                f"Осталось неразмеченных записей: {summary['remaining_records']}.\nВыгрузить результат всё равно?",
                parent=self,
            ):
                return
        output_path = build_output_path(
            self.output_dir,
            self.current_input_path,
            self.current_mode()["id"],
            completed_count=summary["completed_records"],
            total_count=summary["total_records"],
        )
        save_results_excel(
            output_path,
            self.current_input_path,
            self.current_mode(),
            self.records,
            self.annotations,
            self.annotator_name_var.get().strip(),
            summary={
                "total_records": summary["total_records"],
                "completed_records": summary["completed_records"],
                "remaining_records": summary["remaining_records"],
                "needs_review_count": summary["needs_review_count"],
                "mode_name": summary["mode_name"],
                "mode_version": summary["mode_version"],
            },
        )
        self.history_payload = update_history_from_annotations(
            self.history_payload,
            self.current_mode()["id"],
            self.current_mode().get("fields", []),
            self.annotations,
            recent_limit=int(self.settings.get("history_recent_limit", 20)),
            global_limit=int(self.settings.get("history_global_limit", 200)),
        )
        save_history(self.meta_dir, self.history_payload)
        self.status_var.set(f"Результат сохранён: {output_path.name}")
        messagebox.showinfo("Готово", f"Результат сохранён в {output_path}", parent=self)

    def compute_summary(self) -> Dict[str, Any]:
        total = len(self.records)
        completed = 0
        needs_review = 0
        low_medium = 0
        field_completion_rows: List[Tuple[str, str, str, str]] = []
        value_distribution_rows: List[Tuple[str, str, int]] = []
        mode = self.current_mode()

        for record in self.records:
            ann = self.annotations.get(record.annotation_key, {})
            if ann.get("completed"):
                completed += 1
            if ann.get("needs_review"):
                needs_review += 1
            if ann.get("confidence") in {"Средняя", "Низкая"}:
                low_medium += 1

        for field in mode.get("fields", []):
            filled = 0
            counter: Dict[str, int] = {}
            for record in self.records:
                ann = self.annotations.get(record.annotation_key, {})
                value = ann.get(field["key"], "")
                non_empty = False
                if isinstance(value, list):
                    non_empty = bool(value)
                    for item in value:
                        counter[str(item)] = counter.get(str(item), 0) + 1
                elif isinstance(value, bool):
                    non_empty = value is True
                    if value:
                        counter["Да"] = counter.get("Да", 0) + 1
                else:
                    non_empty = str(value).strip() != ""
                    if non_empty:
                        counter[str(value).strip()] = counter.get(str(value).strip(), 0) + 1
                if non_empty:
                    filled += 1
            field_completion_rows.append((field["label"], f"{filled} / {total}" if total else "0 / 0", "Да" if field.get("required") else "Нет", field["type"]))
            for value, count in sorted(counter.items(), key=lambda x: (-x[1], x[0].lower()))[:10]:
                value_distribution_rows.append((field["label"], value, count))

        return {
            "total_records": total,
            "completed_records": completed,
            "remaining_records": max(0, total - completed),
            "needs_review_count": needs_review,
            "low_or_medium_confidence_count": low_medium,
            "mode_name": mode["name"],
            "mode_version": mode.get("version", ""),
            "source_file": self.current_input_path.name if self.current_input_path else "",
            "field_completion_rows": field_completion_rows,
            "value_distribution_rows": value_distribution_rows,
        }

    def open_stats_window(self) -> None:
        if self.records:
            self.save_current_annotation(interactive=False)
        StatsWindow(self)

    def open_mode_editor(self) -> None:
        ModeEditorWindow(
            self,
            self.modes_payload,
            self.modes_path,
            self.backups_dir,
            self._handle_modes_saved,
        )

    def _handle_modes_saved(self, payload: Dict[str, Any]) -> None:
        self.modes_payload = payload
        self.refresh_modes_ui()
        self.status_var.set("Конфиг режимов обновлён.")

    def on_close(self) -> None:
        if self.current_input_path and self.records and self.dirty:
            self.save_draft(silent=True)
        self.settings["last_mode_id"] = self.selected_mode_id.get()
        self.settings["last_input_file"] = self.current_input_path.name if self.current_input_path else ""
        self.settings["last_annotator_name"] = self.annotator_name_var.get().strip()
        self.settings["window_geometry"] = self.geometry()
        save_app_settings(self.settings, self.app_settings_path)
        self.destroy()


def run() -> None:
    try:
        app = LabelerApp()
    except ModesConfigError as exc:
        root = tk.Tk()
        root.withdraw()
        message = str(exc)
        backup_path = latest_modes_backup(ensure_runtime_dirs()["backups_dir"])
        if backup_path is not None:
            restore = messagebox.askyesno(
                "Ошибка modes.json",
                message + "\n\nВосстановить последний backup автоматически?",
                parent=root,
            )
            if restore:
                try:
                    paths = ensure_runtime_dirs()
                    restored = restore_latest_modes_backup(paths["modes_path"], paths["backups_dir"])
                    messagebox.showinfo("Backup восстановлен", f"Восстановлен backup: {Path(restored).name}", parent=root)
                    root.destroy()
                    app = LabelerApp()
                    app.mainloop()
                    return
                except Exception as inner_exc:
                    messagebox.showerror("Ошибка восстановления", str(inner_exc), parent=root)
        else:
            messagebox.showerror("Ошибка modes.json", message, parent=root)
        root.destroy()
        return
    app.mainloop()
