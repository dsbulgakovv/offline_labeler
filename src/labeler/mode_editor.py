from __future__ import annotations

import tkinter as tk
from copy import deepcopy
from datetime import datetime
from tkinter import messagebox, ttk
from typing import Any, Callable, Dict, List, Optional, Tuple

from .config import bump_mode_version, latest_modes_backup, restore_latest_modes_backup, save_modes, validate_mode, validate_modes_payload, load_modes
from .models import ALLOWED_FIELD_TYPES, ALLOWED_RULE_ACTIONS, ALLOWED_RULE_OPERATORS, clone_mode


def _copy_title(value: str, fallback: str) -> str:
    text = str(value or "").strip()
    if text.endswith(" (копия)"):
        return text
    if not text:
        return fallback
    return "{0} (копия)".format(text)


def _unique_copy_key(base_value: str, existing_values: List[str], fallback: str) -> str:
    base = str(base_value or "").strip() or fallback
    existing = set(str(item or "").strip() for item in existing_values if str(item or "").strip())
    candidate = "{0}_copy".format(base)
    index = 2
    while candidate in existing:
        candidate = "{0}_copy_{1}".format(base, index)
        index += 1
    return candidate


class FieldEditorDialog(tk.Toplevel):
    def __init__(self, master: tk.Misc, field: Optional[Dict[str, Any]], on_save: Callable[[Dict[str, Any]], None]) -> None:
        super().__init__(master)
        self.title("Поле режима")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.on_save = on_save
        self.result: Optional[Dict[str, Any]] = None

        field = deepcopy(field or {})
        self.key_var = tk.StringVar(value=field.get("key", ""))
        self.label_var = tk.StringVar(value=field.get("label", ""))
        self.type_var = tk.StringVar(value=field.get("type", "text"))
        self.required_var = tk.BooleanVar(value=bool(field.get("required", False)))
        self.default_var = tk.StringVar(value="" if field.get("default") is None else str(field.get("default", "")))

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="key").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.key_var, width=32).grid(row=0, column=1, sticky="we", padx=(8, 0), pady=4)
        ttk.Label(frame, text="label").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.label_var, width=32).grid(row=1, column=1, sticky="we", padx=(8, 0), pady=4)
        ttk.Label(frame, text="type").grid(row=2, column=0, sticky="w")
        ttk.Combobox(frame, textvariable=self.type_var, values=sorted(ALLOWED_FIELD_TYPES), state="readonly", width=29).grid(row=2, column=1, sticky="we", padx=(8, 0), pady=4)
        ttk.Checkbutton(frame, text="Обязательное", variable=self.required_var).grid(row=3, column=1, sticky="w", pady=4)
        ttk.Label(frame, text="default").grid(row=4, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.default_var, width=32).grid(row=4, column=1, sticky="we", padx=(8, 0), pady=4)

        ttk.Label(frame, text="help").grid(row=5, column=0, sticky="nw")
        self.help_text = tk.Text(frame, width=50, height=4)
        self.help_text.grid(row=5, column=1, sticky="we", padx=(8, 0), pady=4)
        self.help_text.insert("1.0", field.get("help", ""))

        ttk.Label(frame, text="examples (по одному на строку)").grid(row=6, column=0, sticky="nw")
        self.examples_text = tk.Text(frame, width=50, height=4)
        self.examples_text.grid(row=6, column=1, sticky="we", padx=(8, 0), pady=4)
        self.examples_text.insert("1.0", "\n".join(field.get("examples", [])))

        ttk.Label(frame, text="options (по одному на строку)").grid(row=7, column=0, sticky="nw")
        self.options_text = tk.Text(frame, width=50, height=5)
        self.options_text.grid(row=7, column=1, sticky="we", padx=(8, 0), pady=4)
        self.options_text.insert("1.0", "\n".join(field.get("options", [])))

        btns = ttk.Frame(frame)
        btns.grid(row=8, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Сохранить", command=self._save).pack(side=tk.RIGHT, padx=(0, 8))

        frame.columnconfigure(1, weight=1)
        self.bind("<Escape>", lambda e: self.destroy())

    def _save(self) -> None:
        field_type = self.type_var.get().strip()
        payload: Dict[str, Any] = {
            "key": self.key_var.get().strip(),
            "label": self.label_var.get().strip(),
            "type": field_type,
            "required": bool(self.required_var.get()),
            "help": self.help_text.get("1.0", tk.END).strip(),
            "examples": [line.strip() for line in self.examples_text.get("1.0", tk.END).splitlines() if line.strip()],
        }
        default_value = self.default_var.get().strip()
        if field_type == "checkbox":
            payload["default"] = default_value.lower() in {"1", "true", "yes", "да"}
        elif field_type == "multiselect":
            payload["default"] = [line.strip() for line in default_value.split(";") if line.strip()]
        else:
            payload["default"] = default_value
        if field_type in {"select", "multiselect"}:
            payload["options"] = [line.strip() for line in self.options_text.get("1.0", tk.END).splitlines() if line.strip()]
        if not payload["key"] or not payload["label"]:
            messagebox.showerror("Ошибка", "У поля должны быть заполнены key и label.", parent=self)
            return
        if field_type not in ALLOWED_FIELD_TYPES:
            messagebox.showerror("Ошибка", "Неподдерживаемый тип поля.", parent=self)
            return
        if field_type in {"select", "multiselect"} and not payload.get("options"):
            messagebox.showerror("Ошибка", "Для select/multiselect нужно указать options.", parent=self)
            return
        self.result = payload
        self.on_save(payload)
        self.destroy()


class RuleEditorDialog(tk.Toplevel):
    def __init__(self, master: tk.Misc, field_keys: List[str], rule: Optional[Dict[str, Any]], on_save: Callable[[Dict[str, Any]], None]) -> None:
        super().__init__(master)
        self.title("Правило")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.on_save = on_save
        self.field_keys = field_keys

        rule = deepcopy(rule or {})
        condition = (rule.get("if") or [{}])[0]
        actions = [item for item in (rule.get("then") or []) if isinstance(item, dict)]
        action = actions[0] if actions else {}
        action_type = action.get("type", "require_filled")
        action_values = [str(item).strip() for item in (action.get("values", []) or []) if str(item).strip()]
        selected_action_fields: List[str] = []
        self.extra_actions: List[Dict[str, Any]] = []
        if actions:
            for item in actions:
                if item.get("type") != action_type:
                    self.extra_actions.append(deepcopy(item))
                    continue
                candidate_values = [str(value).strip() for value in (item.get("values", []) or []) if str(value).strip()]
                if candidate_values != action_values:
                    self.extra_actions.append(deepcopy(item))
                    continue
                field_key = str(item.get("field", "")).strip()
                if field_key and field_key in field_keys and field_key not in selected_action_fields:
                    selected_action_fields.append(field_key)
                elif field_key:
                    self.extra_actions.append(deepcopy(item))

        self.id_var = tk.StringVar(value=rule.get("id", ""))
        self.description_var = tk.StringVar(value=rule.get("description", ""))
        self.severity_var = tk.StringVar(value=rule.get("severity", "error"))
        self.message_var = tk.StringVar(value=rule.get("message", ""))
        self.cond_field_var = tk.StringVar(value=condition.get("field", field_keys[0] if field_keys else ""))
        self.cond_op_var = tk.StringVar(value=condition.get("op", "equals"))
        self.cond_value_var = tk.StringVar(value="" if condition.get("value") is None else str(condition.get("value", "")))
        self.action_type_var = tk.StringVar(value=action_type)
        self.action_values_var = tk.StringVar(value="; ".join(action_values))

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        rows = [
            ("id", ttk.Entry(frame, textvariable=self.id_var, width=40)),
            ("Описание", ttk.Entry(frame, textvariable=self.description_var, width=40)),
            ("Severity", ttk.Combobox(frame, textvariable=self.severity_var, values=["error", "warning"], state="readonly", width=37)),
            ("Сообщение", ttk.Entry(frame, textvariable=self.message_var, width=40)),
            ("Если: поле", ttk.Combobox(frame, textvariable=self.cond_field_var, values=field_keys, state="readonly", width=37)),
            ("Если: оператор", ttk.Combobox(frame, textvariable=self.cond_op_var, values=sorted(ALLOWED_RULE_OPERATORS), state="readonly", width=37)),
            ("Если: значение", ttk.Entry(frame, textvariable=self.cond_value_var, width=40)),
            ("Тогда: действие", ttk.Combobox(frame, textvariable=self.action_type_var, values=sorted(ALLOWED_RULE_ACTIONS), state="readonly", width=37)),
        ]
        for row_idx, (label, widget) in enumerate(rows):
            ttk.Label(frame, text=label).grid(row=row_idx, column=0, sticky="w", pady=4)
            widget.grid(row=row_idx, column=1, sticky="we", padx=(8, 0), pady=4)

        fields_row = len(rows)
        ttk.Label(frame, text="Тогда: поля").grid(row=fields_row, column=0, sticky="nw", pady=4)
        action_fields_frame = ttk.Frame(frame)
        action_fields_frame.grid(row=fields_row, column=1, sticky="we", padx=(8, 0), pady=4)
        action_fields_frame.columnconfigure(0, weight=1)
        self.action_fields_listbox = tk.Listbox(
            action_fields_frame,
            selectmode=tk.MULTIPLE,
            exportselection=False,
            height=min(8, max(4, len(field_keys) or 1)),
        )
        action_fields_scroll = ttk.Scrollbar(action_fields_frame, orient="vertical", command=self.action_fields_listbox.yview)
        self.action_fields_listbox.configure(yscrollcommand=action_fields_scroll.set)
        self.action_fields_listbox.grid(row=0, column=0, sticky="nsew")
        action_fields_scroll.grid(row=0, column=1, sticky="ns")
        for field_key in field_keys:
            self.action_fields_listbox.insert(tk.END, field_key)
        for index, field_key in enumerate(field_keys):
            if field_key in selected_action_fields:
                self.action_fields_listbox.selection_set(index)
        if not selected_action_fields and field_keys:
            self.action_fields_listbox.selection_set(0)
        ttk.Label(frame, text="Можно выбрать несколько полей.").grid(
            row=fields_row + 1,
            column=1,
            sticky="w",
            padx=(8, 0),
        )

        values_row = fields_row + 2
        ttk.Label(frame, text="Тогда: values (;)").grid(row=values_row, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.action_values_var, width=40).grid(row=values_row, column=1, sticky="we", padx=(8, 0), pady=4)

        btns = ttk.Frame(frame)
        btns.grid(row=values_row + 1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Сохранить", command=self._save).pack(side=tk.RIGHT, padx=(0, 8))
        frame.columnconfigure(1, weight=1)

    def _save(self) -> None:
        selected_indices = self.action_fields_listbox.curselection()
        selected_fields = [self.field_keys[int(index)] for index in selected_indices]
        payload = {
            "id": self.id_var.get().strip(),
            "description": self.description_var.get().strip(),
            "severity": self.severity_var.get().strip() or "error",
            "message": self.message_var.get().strip(),
            "if": [
                {
                    "field": self.cond_field_var.get().strip(),
                    "op": self.cond_op_var.get().strip(),
                    "value": self.cond_value_var.get().strip(),
                }
            ],
            "then": [],
        }
        if not payload["id"] or not payload["description"]:
            messagebox.showerror("Ошибка", "У правила должны быть заполнены id и описание.", parent=self)
            return
        if not selected_fields:
            messagebox.showerror("Ошибка", "Выбери хотя бы одно поле в секции 'Тогда'.", parent=self)
            return
        action_values = [item.strip() for item in self.action_values_var.get().split(";") if item.strip()]
        action_type = self.action_type_var.get().strip()
        for field_key in selected_fields:
            payload["then"].append(
                {
                    "type": action_type,
                    "field": field_key,
                    "values": list(action_values),
                }
            )
        payload["then"].extend(deepcopy(self.extra_actions))
        self.on_save(payload)
        self.destroy()


class ModeEditorWindow(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        modes_payload: Dict[str, Any],
        modes_path,
        backups_dir,
        on_modes_saved: Callable[[Dict[str, Any]], None],
    ) -> None:
        super().__init__(master)
        self.title("Конструктор режимов")
        self.geometry("1200x760")
        self.minsize(1040, 680)
        self.transient(master)
        self.on_modes_saved = on_modes_saved
        self.modes_path = modes_path
        self.backups_dir = backups_dir
        self.payload = deepcopy(modes_payload)
        self.mode_list: List[Dict[str, Any]] = self.payload.get("modes", [])
        self.current_mode_index = 0

        container = ttk.Frame(self, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        toolbar = ttk.Frame(container)
        toolbar.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(toolbar, text="Новый режим", command=self.new_mode).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Дублировать", command=self.duplicate_mode).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(toolbar, text="Удалить", command=self.delete_mode).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(toolbar, text="Восстановить последний backup", command=self.restore_latest_backup).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Button(toolbar, text="Проверить конфиг", command=self.validate_all).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(toolbar, text="Сохранить все", command=self.save_all).pack(side=tk.RIGHT)

        body = ttk.Panedwindow(container, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(body)
        right = ttk.Frame(body)
        body.add(left, weight=1)
        body.add(right, weight=3)

        self.mode_listbox = tk.Listbox(left, exportselection=False)
        self.mode_listbox.pack(fill=tk.BOTH, expand=True)
        self.mode_listbox.bind("<<ListboxSelect>>", self.on_mode_select)

        self.general_id_var = tk.StringVar()
        self.general_name_var = tk.StringVar()
        self.general_version_var = tk.StringVar()
        self.change_comment_var = tk.StringVar()

        notebook = ttk.Notebook(right)
        notebook.pack(fill=tk.BOTH, expand=True)
        general_tab = ttk.Frame(notebook, padding=10)
        fields_tab = ttk.Frame(notebook, padding=10)
        rules_tab = ttk.Frame(notebook, padding=10)
        notebook.add(general_tab, text="Общее")
        notebook.add(fields_tab, text="Поля")
        notebook.add(rules_tab, text="Правила")

        ttk.Label(general_tab, text="Mode ID").grid(row=0, column=0, sticky="w")
        ttk.Entry(general_tab, textvariable=self.general_id_var, width=40).grid(row=0, column=1, sticky="we", padx=(8, 0), pady=4)
        ttk.Label(general_tab, text="Название").grid(row=1, column=0, sticky="w")
        ttk.Entry(general_tab, textvariable=self.general_name_var, width=40).grid(row=1, column=1, sticky="we", padx=(8, 0), pady=4)
        ttk.Label(general_tab, text="Версия").grid(row=2, column=0, sticky="w")
        ttk.Entry(general_tab, textvariable=self.general_version_var, width=20).grid(row=2, column=1, sticky="w", padx=(8, 0), pady=4)
        ttk.Label(general_tab, text="Комментарий изменения").grid(row=3, column=0, sticky="w")
        ttk.Entry(general_tab, textvariable=self.change_comment_var, width=40).grid(row=3, column=1, sticky="we", padx=(8, 0), pady=4)

        ttk.Label(general_tab, text="Описание").grid(row=4, column=0, sticky="nw")
        self.description_text = tk.Text(general_tab, height=4, width=60)
        self.description_text.grid(row=4, column=1, sticky="we", padx=(8, 0), pady=4)
        ttk.Label(general_tab, text="Инструкции").grid(row=5, column=0, sticky="nw")
        self.instructions_text = tk.Text(general_tab, height=6, width=60)
        self.instructions_text.grid(row=5, column=1, sticky="we", padx=(8, 0), pady=4)
        ttk.Label(general_tab, text="Примеры\n(по одному на строку)").grid(row=6, column=0, sticky="nw")
        self.examples_text = tk.Text(general_tab, height=6, width=60)
        self.examples_text.grid(row=6, column=1, sticky="we", padx=(8, 0), pady=4)
        ttk.Button(general_tab, text="Применить изменения вкладки", command=self.apply_general_changes).grid(row=7, column=1, sticky="e", pady=(8, 0))
        general_tab.columnconfigure(1, weight=1)

        fields_toolbar = ttk.Frame(fields_tab)
        fields_toolbar.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(fields_toolbar, text="Добавить поле", command=self.add_field).pack(side=tk.LEFT)
        ttk.Button(fields_toolbar, text="Редактировать", command=self.edit_field).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(fields_toolbar, text="Дублировать", command=self.duplicate_field).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(fields_toolbar, text="Удалить", command=self.delete_field).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(fields_toolbar, text="↑", width=4, command=lambda: self.move_field(-1)).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Button(fields_toolbar, text="↓", width=4, command=lambda: self.move_field(1)).pack(side=tk.LEFT, padx=(4, 0))
        self.fields_tree = ttk.Treeview(fields_tab, columns=("key", "label", "type", "required"), show="headings", height=16)
        for name, width in [("key", 160), ("label", 260), ("type", 100), ("required", 90)]:
            self.fields_tree.heading(name, text=name)
            self.fields_tree.column(name, width=width, anchor="w")
        self.fields_tree.pack(fill=tk.BOTH, expand=True)

        rules_toolbar = ttk.Frame(rules_tab)
        rules_toolbar.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(rules_toolbar, text="Добавить правило", command=self.add_rule).pack(side=tk.LEFT)
        ttk.Button(rules_toolbar, text="Редактировать", command=self.edit_rule).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(rules_toolbar, text="Дублировать", command=self.duplicate_rule).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(rules_toolbar, text="Удалить", command=self.delete_rule).pack(side=tk.LEFT, padx=(8, 0))
        self.rules_tree = ttk.Treeview(rules_tab, columns=("id", "description", "severity"), show="headings", height=16)
        for name, width in [("id", 180), ("description", 420), ("severity", 90)]:
            self.rules_tree.heading(name, text=name)
            self.rules_tree.column(name, width=width, anchor="w")
        self.rules_tree.pack(fill=tk.BOTH, expand=True)

        self.refresh_mode_list()
        if self.mode_list:
            self.mode_listbox.selection_set(0)
            self.load_mode_to_editor(0)

    def refresh_mode_list(self) -> None:
        self.mode_listbox.delete(0, tk.END)
        for mode in self.mode_list:
            self.mode_listbox.insert(tk.END, f"{mode['name']} ({mode['id']}) v{mode.get('version', '')}")

    def on_mode_select(self, event=None) -> None:
        selection = self.mode_listbox.curselection()
        if not selection:
            return
        self.load_mode_to_editor(selection[0])

    def current_mode(self) -> Dict[str, Any]:
        return self.mode_list[self.current_mode_index]

    def _normalized_change_log(self, mode: Dict[str, Any]) -> List[Dict[str, Any]]:
        result: List[Dict[str, Any]] = []
        for item in mode.get("change_log", []) or []:
            if isinstance(item, dict):
                result.append({
                    "changed_at": item.get("changed_at", ""),
                    "comment": str(item.get("comment", "")).strip(),
                })
            elif str(item).strip():
                result.append({
                    "changed_at": "",
                    "comment": str(item).strip(),
                })
        return result

    def _format_change_log_item(self, item: Dict[str, Any]) -> str:
        stamp = str(item.get("changed_at", "")).strip()
        comment = str(item.get("comment", "")).strip()
        return ("{0} — {1}".format(stamp, comment) if stamp else comment) or "(пусто)"

    def refresh_change_log_list(self) -> None:
        mode = self.current_mode()
        items = self._normalized_change_log(mode)
        mode["change_log"] = items
        self.change_log_listbox.delete(0, tk.END)
        for item in items:
            self.change_log_listbox.insert(tk.END, self._format_change_log_item(item))

    def bump_selected_mode_version(self) -> None:
        self.general_version_var.set(bump_mode_version(self.general_version_var.get().strip() or '1.0'))

    def add_or_update_change_log_entry(self) -> None:
        mode = self.current_mode()
        items = self._normalized_change_log(mode)
        comment = self.change_comment_var.get().strip()
        if not comment:
            messagebox.showwarning("Пустой комментарий", "Введите комментарий для change_log.", parent=self)
            return
        payload = {
            "changed_at": datetime.now().isoformat(timespec="seconds"),
            "comment": comment,
        }
        sel = self.change_log_listbox.curselection()
        if sel:
            items[sel[0]] = payload
        else:
            items.insert(0, payload)
        mode["change_log"] = items
        self.refresh_change_log_list()
        self.change_comment_var.set("")

    def keep_only_selected_change_log(self) -> None:
        sel = self.change_log_listbox.curselection()
        if not sel:
            return
        mode = self.current_mode()
        items = self._normalized_change_log(mode)
        mode["change_log"] = [items[sel[0]]]
        self.refresh_change_log_list()
        self.change_log_listbox.selection_set(0)

    def delete_selected_change_log(self) -> None:
        sel = self.change_log_listbox.curselection()
        if not sel:
            return
        mode = self.current_mode()
        items = self._normalized_change_log(mode)
        del items[sel[0]]
        mode["change_log"] = items
        self.refresh_change_log_list()

    def load_mode_to_editor(self, idx: int) -> None:
        self.current_mode_index = idx
        mode = self.mode_list[idx]
        self.general_id_var.set(mode.get("id", ""))
        self.general_name_var.set(mode.get("name", ""))
        self.general_version_var.set(mode.get("version", "1.0"))
        change_log = self._normalized_change_log(mode)
        self.change_comment_var.set(change_log[0].get("comment", "") if change_log else "")
        self.description_text.delete("1.0", tk.END)
        self.description_text.insert("1.0", mode.get("description", ""))
        self.instructions_text.delete("1.0", tk.END)
        self.instructions_text.insert("1.0", mode.get("instructions", ""))
        self.examples_text.delete("1.0", tk.END)
        self.examples_text.insert("1.0", "\n".join(mode.get("examples", [])))
        self.refresh_change_log_list()
        self.refresh_fields_tree()
        self.refresh_rules_tree()

    def apply_general_changes(self) -> None:
        mode = self.current_mode()
        old_snapshot = clone_mode(mode)
        mode["id"] = self.general_id_var.get().strip()
        mode["name"] = self.general_name_var.get().strip()
        mode["version"] = self.general_version_var.get().strip() or mode.get("version", "1.0")
        mode["description"] = self.description_text.get("1.0", tk.END).strip()
        mode["instructions"] = self.instructions_text.get("1.0", tk.END).strip()
        mode["examples"] = [line.strip() for line in self.examples_text.get("1.0", tk.END).splitlines() if line.strip()]
        mode["change_log"] = self._normalized_change_log(mode)
        if old_snapshot != mode:
            mode["updated_at"] = datetime.now().isoformat(timespec="seconds")
        try:
            validate_mode(mode)
        except Exception as exc:
            messagebox.showerror("Ошибка режима", str(exc), parent=self)
            self.mode_list[self.current_mode_index] = old_snapshot
            self.load_mode_to_editor(self.current_mode_index)
            return
        self.refresh_mode_list()
        self.mode_listbox.selection_clear(0, tk.END)
        self.mode_listbox.selection_set(self.current_mode_index)
        messagebox.showinfo("Готово", "Изменения вкладки применены.", parent=self)

    def new_mode(self) -> None:
        base = {
            "id": f"mode_{len(self.mode_list) + 1}",
            "name": "Новый режим",
            "version": "1.0",
            "description": "",
            "instructions": "",
            "examples": [],
            "change_log": [],
            "fields": [
                {
                    "key": "field_1",
                    "label": "Поле 1",
                    "type": "text",
                    "required": False,
                    "help": "",
                    "examples": [],
                    "default": "",
                }
            ],
            "rules": [],
        }
        self.mode_list.append(base)
        self.refresh_mode_list()
        idx = len(self.mode_list) - 1
        self.mode_listbox.selection_clear(0, tk.END)
        self.mode_listbox.selection_set(idx)
        self.load_mode_to_editor(idx)

    def duplicate_mode(self) -> None:
        mode = clone_mode(self.current_mode())
        mode["id"] = f"{mode['id']}_copy"
        mode["name"] = f"{mode['name']} (копия)"
        mode["version"] = "1.0"
        mode["change_log"] = []
        self.mode_list.append(mode)
        self.refresh_mode_list()
        idx = len(self.mode_list) - 1
        self.mode_listbox.selection_clear(0, tk.END)
        self.mode_listbox.selection_set(idx)
        self.load_mode_to_editor(idx)

    def restore_latest_backup(self) -> None:
        backup_path = latest_modes_backup(self.backups_dir)
        if backup_path is None:
            messagebox.showwarning("Backup не найден", "В папке config/_history нет резервных копий.", parent=self)
            return
        if not messagebox.askyesno("Восстановить backup", "Восстановить последний backup modes.json? Несохранённые изменения в конструкторе будут потеряны.", parent=self):
            return
        try:
            restore_latest_modes_backup(self.modes_path, self.backups_dir)
            self.payload = load_modes(self.modes_path)
            self.mode_list = self.payload.get("modes", [])
            self.refresh_mode_list()
            if self.mode_list:
                self.mode_listbox.selection_clear(0, tk.END)
                self.mode_listbox.selection_set(0)
                self.load_mode_to_editor(0)
            self.on_modes_saved(self.payload)
        except Exception as exc:
            messagebox.showerror("Ошибка восстановления", str(exc), parent=self)
            return
        messagebox.showinfo("Готово", "Последний backup восстановлен.", parent=self)

    def delete_mode(self) -> None:
        if not self.mode_list:
            return
        if not messagebox.askyesno("Подтверждение", "Удалить выбранный режим?", parent=self):
            return
        del self.mode_list[self.current_mode_index]
        if not self.mode_list:
            self.new_mode()
            return
        self.refresh_mode_list()
        new_idx = max(0, self.current_mode_index - 1)
        self.mode_listbox.selection_set(new_idx)
        self.load_mode_to_editor(new_idx)

    def refresh_fields_tree(self) -> None:
        for item in self.fields_tree.get_children():
            self.fields_tree.delete(item)
        for idx, field in enumerate(self.current_mode().get("fields", [])):
            self.fields_tree.insert("", tk.END, iid=str(idx), values=(field["key"], field["label"], field["type"], "Да" if field.get("required") else "Нет"))

    def refresh_rules_tree(self) -> None:
        for item in self.rules_tree.get_children():
            self.rules_tree.delete(item)
        for idx, rule in enumerate(self.current_mode().get("rules", [])):
            self.rules_tree.insert("", tk.END, iid=str(idx), values=(rule["id"], rule.get("description", ""), rule.get("severity", "error")))

    def add_field(self) -> None:
        def _save(field: Dict[str, Any]) -> None:
            self.current_mode().setdefault("fields", []).append(field)
            self.refresh_fields_tree()

        FieldEditorDialog(self, None, _save)

    def _selected_tree_index(self, tree: ttk.Treeview) -> Optional[int]:
        selected = tree.selection()
        if not selected:
            return None
        return int(selected[0])

    def edit_field(self) -> None:
        idx = self._selected_tree_index(self.fields_tree)
        if idx is None:
            return
        fields = self.current_mode().setdefault("fields", [])

        def _save(field: Dict[str, Any]) -> None:
            fields[idx] = field
            self.refresh_fields_tree()

        FieldEditorDialog(self, fields[idx], _save)

    def delete_field(self) -> None:
        idx = self._selected_tree_index(self.fields_tree)
        if idx is None:
            return
        del self.current_mode().setdefault("fields", [])[idx]
        self.refresh_fields_tree()

    def duplicate_field(self) -> None:
        idx = self._selected_tree_index(self.fields_tree)
        if idx is None:
            return
        fields = self.current_mode().setdefault("fields", [])
        field = deepcopy(fields[idx])
        field["key"] = _unique_copy_key(field.get("key", "field"), [item.get("key", "") for item in fields], "field")
        field["label"] = _copy_title(field.get("label", ""), "Копия поля")
        fields.insert(idx + 1, field)
        self.refresh_fields_tree()
        self.fields_tree.selection_set(str(idx + 1))

    def move_field(self, delta: int) -> None:
        idx = self._selected_tree_index(self.fields_tree)
        if idx is None:
            return
        fields = self.current_mode().setdefault("fields", [])
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(fields):
            return
        fields[idx], fields[new_idx] = fields[new_idx], fields[idx]
        self.refresh_fields_tree()
        self.fields_tree.selection_set(str(new_idx))

    def add_rule(self) -> None:
        field_keys = [field["key"] for field in self.current_mode().get("fields", [])]

        def _save(rule: Dict[str, Any]) -> None:
            self.current_mode().setdefault("rules", []).append(rule)
            self.refresh_rules_tree()

        RuleEditorDialog(self, field_keys, None, _save)

    def edit_rule(self) -> None:
        idx = self._selected_tree_index(self.rules_tree)
        if idx is None:
            return
        rules = self.current_mode().setdefault("rules", [])
        field_keys = [field["key"] for field in self.current_mode().get("fields", [])]

        def _save(rule: Dict[str, Any]) -> None:
            rules[idx] = rule
            self.refresh_rules_tree()

        RuleEditorDialog(self, field_keys, rules[idx], _save)

    def delete_rule(self) -> None:
        idx = self._selected_tree_index(self.rules_tree)
        if idx is None:
            return
        del self.current_mode().setdefault("rules", [])[idx]
        self.refresh_rules_tree()

    def duplicate_rule(self) -> None:
        idx = self._selected_tree_index(self.rules_tree)
        if idx is None:
            return
        rules = self.current_mode().setdefault("rules", [])
        rule = deepcopy(rules[idx])
        rule["id"] = _unique_copy_key(rule.get("id", "rule"), [item.get("id", "") for item in rules], "rule")
        rule["description"] = _copy_title(rule.get("description", ""), "Копия правила")
        rules.insert(idx + 1, rule)
        self.refresh_rules_tree()
        self.rules_tree.selection_set(str(idx + 1))

    def validate_all(self) -> None:
        try:
            validate_modes_payload(self.payload)
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc), parent=self)
            return
        messagebox.showinfo("Готово", "Конфиг режимов валиден.", parent=self)

    def save_all(self) -> None:
        self.apply_general_changes()
        try:
            validate_modes_payload(self.payload)
            save_modes(self.payload, self.modes_path, self.backups_dir)
        except Exception as exc:
            messagebox.showerror("Ошибка при сохранении", str(exc), parent=self)
            return
        self.on_modes_saved(self.payload)
        messagebox.showinfo("Готово", "Режимы сохранены. Главный экран обновлён.", parent=self)
