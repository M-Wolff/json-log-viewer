from __future__ import annotations

import fnmatch
import json
from pathlib import Path
import re
import statistics
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import math

from .model import JsonDocument, Record


class JsonLogViewerApp:
    def __init__(self, root: tk.Tk, initial_path: str | None = None) -> None:
        self.root = root
        self.root.title("JSON Log Viewer")
        self.root.geometry("1600x950")

        self.document: JsonDocument | None = None
        self.filtered_records: list[Record] = []
        self.row_map: dict[str, int] = {}
        self.display_columns: list[str] = []
        self.column_search_var = tk.StringVar()
        self.display_column_vars: dict[str, tk.BooleanVar] = {}
        self.filter_vars: dict[str, tk.StringVar] = {}
        self.filter_entries: dict[str, ttk.Entry] = {}
        self.derived_columns: dict[str, str] = {}
        self.derived_values: dict[int, dict[str, str]] = {}
        self.sort_column: str = ""
        self.sort_descending = False

        self.file_label_var = tk.StringVar(value="No file loaded")
        self.status_var = tk.StringVar(value="Open a JSON file to begin.")
        self.global_search_var = tk.StringVar()
        self.show_deleted_var = tk.BooleanVar(value=True)
        self.derived_name_var = tk.StringVar()

        self._build_layout()

        if initial_path:
            self.load_file(initial_path)

    def _build_layout(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        top_bar = ttk.Frame(self.root, padding=8)
        top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.columnconfigure(1, weight=1)

        ttk.Button(top_bar, text="Open JSON", command=self.open_file).grid(row=0, column=0, padx=(0, 8))
        ttk.Label(top_bar, textvariable=self.file_label_var).grid(row=0, column=1, sticky="w")
        ttk.Button(top_bar, text="Reload", command=self.reload_file).grid(row=0, column=2, padx=4)
        ttk.Button(top_bar, text="Preview Changes", command=self.show_diff_preview).grid(row=0, column=3, padx=4)
        ttk.Button(top_bar, text="Save With Backup", command=self.save_changes).grid(row=0, column=4, padx=(4, 0))

        main_frame = ttk.Frame(self.root, padding=8)
        main_frame.grid(row=1, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=0, minsize=460)
        main_frame.columnconfigure(1, weight=2, minsize=420)
        main_frame.columnconfigure(2, weight=2, minsize=420)
        main_frame.rowconfigure(0, weight=3)
        main_frame.rowconfigure(1, weight=2)

        left_sidebar = ttk.Frame(main_frame, padding=(0, 0, 8, 0))
        left_sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew")
        left_sidebar.columnconfigure(0, weight=1)

        table_container = ttk.Frame(main_frame)
        table_container.grid(row=0, column=1, columnspan=2, sticky="nsew")
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(0, weight=1)

        middle_bottom = ttk.Frame(main_frame, padding=(8, 8, 8, 0))
        middle_bottom.grid(row=1, column=1, sticky="nsew")
        middle_bottom.columnconfigure(0, weight=1)
        middle_bottom.rowconfigure(0, weight=0)
        middle_bottom.rowconfigure(1, weight=1)

        right_bottom = ttk.Frame(main_frame, padding=(8, 8, 0, 0))
        right_bottom.grid(row=1, column=2, sticky="nsew")
        right_bottom.columnconfigure(0, weight=1)
        right_bottom.rowconfigure(0, weight=1)

        self._build_controls(left_sidebar)
        self._build_table(table_container)
        self._build_middle_bottom(middle_bottom)
        self._build_detail(right_bottom)

        status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor="w", padding=8)
        status_bar.grid(row=2, column=0, sticky="ew")

    def _build_controls(self, parent: ttk.Frame) -> None:
        search_frame = ttk.LabelFrame(parent, text="Search", padding=8)
        search_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        search_frame.columnconfigure(0, weight=1)

        ttk.Label(search_frame, text="Global regex").grid(row=0, column=0, sticky="w")
        search_entry = ttk.Entry(search_frame, textvariable=self.global_search_var)
        search_entry.grid(row=1, column=0, sticky="ew", pady=(2, 6))
        search_entry.bind("<KeyRelease>", lambda _event: self.refresh_table())
        ttk.Checkbutton(
            search_frame,
            text="Show pending deletions in table",
            variable=self.show_deleted_var,
            command=self.refresh_table,
        ).grid(row=2, column=0, sticky="w")

        columns_frame = ttk.LabelFrame(parent, text="Columns", padding=8)
        columns_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.rowconfigure(3, weight=1)
        parent.rowconfigure(1, weight=1)

        ttk.Label(columns_frame, text="Select visible columns").grid(row=0, column=0, sticky="w")
        ttk.Label(columns_frame, text="Search column names").grid(row=1, column=0, sticky="w", pady=(2, 0))
        column_search_entry = ttk.Entry(columns_frame, textvariable=self.column_search_var)
        column_search_entry.grid(row=2, column=0, sticky="ew", pady=(2, 6))
        column_search_entry.bind("<KeyRelease>", lambda _event: self.render_display_column_checkboxes())
        ttk.Label(columns_frame, text="Checkboxes keep existing selections when you add more.").grid(row=4, column=0, sticky="w", pady=(6, 6))

        self.display_columns_container, self.display_columns_canvas = self._build_scrollable_checkbox_panel(columns_frame, row=3, height=320)

        display_buttons = ttk.Frame(columns_frame)
        display_buttons.grid(row=5, column=0, sticky="ew", pady=(6, 0))
        display_buttons.columnconfigure(0, weight=1)
        display_buttons.columnconfigure(1, weight=1)
        ttk.Button(display_buttons, text="Select All", command=self.select_all_columns).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(display_buttons, text="Reset Defaults", command=self.reset_default_columns).grid(row=0, column=1, sticky="ew")

        filter_frame = ttk.LabelFrame(parent, text="Active Filters", padding=8)
        filter_frame.grid(row=2, column=0, sticky="nsew")
        filter_frame.columnconfigure(0, weight=1)
        filter_frame.rowconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        ttk.Label(filter_frame, text="Regex filters. Any non-empty field is active.").grid(row=1, column=0, sticky="w", pady=(0, 6))

        self.filter_entries_container, self.filter_entries_canvas = self._build_scrollable_checkbox_panel(filter_frame, row=0, height=340)

    def _build_table(self, table_frame: ttk.Frame) -> None:
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame, columns=(), show="headings", selectmode="extended")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", lambda _event: self.update_detail_panel())
        self.tree.tag_configure("deleted", foreground="#a00000")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

    def _build_middle_bottom(self, parent: ttk.Frame) -> None:
        actions = ttk.LabelFrame(parent, text="Actions", padding=8)
        actions.grid(row=0, column=0, sticky="new")
        actions.columnconfigure(0, weight=1)

        ttk.Button(actions, text="Delete Selected Rows", command=self.delete_selected_rows).grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ttk.Button(actions, text="Restore Selected Rows", command=self.restore_selected_rows).grid(row=1, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Restore All Pending", command=self.restore_all_rows).grid(row=2, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Preview Changes", command=self.show_diff_preview).grid(row=3, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Save With Backup", command=self.save_changes).grid(row=4, column=0, sticky="ew", pady=(4, 0))

        derived_frame = ttk.LabelFrame(parent, text="Derived Columns", padding=8)
        derived_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        derived_frame.columnconfigure(0, weight=1)
        derived_frame.rowconfigure(3, weight=1)

        ttk.Label(derived_frame, text="Column name").grid(row=0, column=0, sticky="w")
        ttk.Entry(derived_frame, textvariable=self.derived_name_var).grid(row=1, column=0, sticky="ew", pady=(2, 6))
        ttk.Label(
            derived_frame,
            text="Script returns `result`. Helpers: value(name), num(name), values(pattern), nums(pattern), mean(...).",
            justify="left",
        ).grid(row=2, column=0, sticky="w")

        script_panel = ttk.Frame(derived_frame)
        script_panel.grid(row=3, column=0, sticky="nsew", pady=(6, 6))
        script_panel.columnconfigure(0, weight=1)
        script_panel.rowconfigure(0, weight=1)

        self.derived_script_text = tk.Text(script_panel, height=12, wrap="none")
        self.derived_script_text.grid(row=0, column=0, sticky="nsew")
        self.derived_script_text.bind("<KeyRelease>", lambda _event: self.highlight_script_box())
        script_y = ttk.Scrollbar(script_panel, orient="vertical", command=self.derived_script_text.yview)
        script_y.grid(row=0, column=1, sticky="ns")
        script_x = ttk.Scrollbar(script_panel, orient="horizontal", command=self.derived_script_text.xview)
        script_x.grid(row=1, column=0, sticky="ew")
        self.derived_script_text.configure(yscrollcommand=script_y.set, xscrollcommand=script_x.set)
        self._configure_script_tags()

        derived_buttons = ttk.Frame(derived_frame)
        derived_buttons.grid(row=4, column=0, sticky="ew")
        derived_buttons.columnconfigure(0, weight=1)
        derived_buttons.columnconfigure(1, weight=1)
        derived_buttons.columnconfigure(2, weight=1)
        ttk.Button(derived_buttons, text="Add / Update", command=self.save_derived_column).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(derived_buttons, text="Load Selected", command=self.load_selected_derived_column).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(derived_buttons, text="Remove", command=self.remove_derived_column).grid(row=0, column=2, sticky="ew", padx=(4, 0))

        list_panel = ttk.Frame(derived_frame)
        list_panel.grid(row=5, column=0, sticky="nsew", pady=(6, 0))
        list_panel.columnconfigure(0, weight=1)
        list_panel.rowconfigure(0, weight=1)
        derived_frame.rowconfigure(5, weight=1)

        self.derived_listbox = tk.Listbox(list_panel, exportselection=False, height=5)
        self.derived_listbox.grid(row=0, column=0, sticky="nsew")
        self.derived_listbox.bind("<<ListboxSelect>>", lambda _event: self.load_selected_derived_column())
        derived_list_scroll = ttk.Scrollbar(list_panel, orient="vertical", command=self.derived_listbox.yview)
        derived_list_scroll.grid(row=0, column=1, sticky="ns")
        self.derived_listbox.configure(yscrollcommand=derived_list_scroll.set)

    def _build_detail(self, detail_frame: ttk.Frame) -> None:
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(1, weight=1)

        ttk.Label(detail_frame, text="Selected Record JSON").grid(row=0, column=0, sticky="w")
        self.detail_text = tk.Text(detail_frame, wrap="none")
        self.detail_text.grid(row=1, column=0, sticky="nsew")
        self.detail_text.configure(state="disabled")
        detail_y = ttk.Scrollbar(detail_frame, orient="vertical", command=self.detail_text.yview)
        detail_y.grid(row=1, column=1, sticky="ns")
        detail_x = ttk.Scrollbar(detail_frame, orient="horizontal", command=self.detail_text.xview)
        detail_x.grid(row=2, column=0, sticky="ew")
        self.detail_text.configure(yscrollcommand=detail_y.set, xscrollcommand=detail_x.set)

    def _build_scrollable_checkbox_panel(self, parent: ttk.Frame, row: int, height: int) -> tuple[ttk.Frame, tk.Canvas]:
        panel = ttk.Frame(parent)
        panel.grid(row=row, column=0, sticky="nsew")
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(0, weight=1)

        canvas = tk.Canvas(panel, height=height, highlightthickness=0)
        scrollbar = ttk.Scrollbar(panel, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        inner = ttk.Frame(canvas)
        inner.columnconfigure(0, weight=1)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _resize_canvas(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _resize_inner(event: tk.Event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        inner.bind("<Configure>", _resize_canvas)
        canvas.bind("<Configure>", _resize_inner)
        self._attach_mousewheel(canvas, canvas)
        self._attach_mousewheel(inner, canvas)
        return inner, canvas

    def _attach_mousewheel(self, widget: tk.Widget, scroll_target: tk.Canvas) -> None:
        def _on_mousewheel(event: tk.Event) -> None:
            if event.delta:
                scroll_target.yview_scroll(int(-event.delta / 120), "units")
            elif getattr(event, "num", None) == 4:
                scroll_target.yview_scroll(-3, "units")
            elif getattr(event, "num", None) == 5:
                scroll_target.yview_scroll(3, "units")

        def _bind_global(_event: tk.Event) -> None:
            self.root.bind_all("<MouseWheel>", _on_mousewheel, add="+")
            self.root.bind_all("<Button-4>", _on_mousewheel, add="+")
            self.root.bind_all("<Button-5>", _on_mousewheel, add="+")

        def _unbind_global(_event: tk.Event) -> None:
            self.root.unbind_all("<MouseWheel>")
            self.root.unbind_all("<Button-4>")
            self.root.unbind_all("<Button-5>")

        widget.bind("<Enter>", _bind_global, add="+")
        widget.bind("<Leave>", _unbind_global, add="+")

    def open_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Open JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if selected:
            self.load_file(selected)

    def load_file(self, path: str) -> None:
        try:
            self.document = JsonDocument.load(path)
        except Exception as exc:  # pragma: no cover - UI error path
            messagebox.showerror("Open failed", str(exc))
            return

        self.file_label_var.set(str(Path(path)))
        self.global_search_var.set("")
        self.column_search_var.set("")
        self.show_deleted_var.set(True)
        self.derived_columns.clear()
        self.derived_values.clear()
        self.derived_name_var.set("")
        self.derived_script_text.delete("1.0", tk.END)
        self.derived_listbox.delete(0, tk.END)
        self.sort_column = ""
        self.sort_descending = False
        self._populate_column_selectors()
        self.refresh_table()
        self.status_var.set(
            f"Loaded {len(self.document.record_models)} records from {Path(path).name}. "
            f"Editing list path: {self._record_path_label()}"
        )

    def reload_file(self) -> None:
        if not self.document:
            return
        self.load_file(str(self.document.path))

    def _record_path_label(self) -> str:
        if not self.document:
            return "$"
        if not self.document.record_path:
            return "$"
        return "$." + ".".join(str(part) for part in self.document.record_path)

    def _populate_column_selectors(self) -> None:
        if not self.document:
            return

        columns = self.all_columns()
        previous_display = {
            column for column, variable in self.display_column_vars.items() if variable.get()
        }
        self.display_column_vars = {}

        for child in self.display_columns_container.winfo_children():
            child.destroy()

        default_display = set(default_columns(columns))
        for column in columns:
            display_var = tk.BooleanVar(value=column in previous_display if previous_display else column in default_display)
            self.display_column_vars[column] = display_var

        self.render_display_column_checkboxes()
        self.apply_display_columns()
        self.rebuild_filter_entries()

    def render_display_column_checkboxes(self) -> None:
        for child in self.display_columns_container.winfo_children():
            child.destroy()

        search_text = self.column_search_var.get().strip().lower()
        visible_columns = [
            column for column in self.display_column_vars if not search_text or search_text in column.lower()
        ]

        if not visible_columns:
            ttk.Label(self.display_columns_container, text="No matching columns.").grid(row=0, column=0, sticky="w")
            self.display_columns_canvas.yview_moveto(0.0)
            return

        for row_index, column in enumerate(visible_columns):
            ttk.Checkbutton(
                self.display_columns_container,
                text=column,
                variable=self.display_column_vars[column],
                command=self.apply_display_columns,
            ).grid(row=row_index, column=0, sticky="w")

        self.display_columns_canvas.yview_moveto(0.0)

    def apply_display_columns(self) -> None:
        self.display_columns = [
            column for column, variable in self.display_column_vars.items() if variable.get()
        ]
        self.refresh_table()

    def select_all_columns(self) -> None:
        for variable in self.display_column_vars.values():
            variable.set(True)
        self.apply_display_columns()

    def reset_default_columns(self) -> None:
        if not self.document:
            return
        defaults = set(default_columns(self.document.columns))
        for column, variable in self.display_column_vars.items():
            variable.set(column in defaults)
        self.apply_display_columns()

    def rebuild_filter_entries(self) -> None:
        existing_values = {column: var.get() for column, var in self.filter_vars.items()}

        for child in self.filter_entries_container.winfo_children():
            child.destroy()

        self.filter_vars.clear()
        self.filter_entries.clear()

        if not self.document:
            return

        for row_index, column in enumerate(self.all_columns()):
            ttk.Label(self.filter_entries_container, text=column).grid(row=row_index, column=0, sticky="nw", padx=(0, 8), pady=2)
            variable = tk.StringVar(value=existing_values.get(column, ""))
            variable.trace_add("write", lambda *_args: self.refresh_table())
            entry = ttk.Entry(self.filter_entries_container, textvariable=variable)
            entry.grid(row=row_index, column=1, sticky="ew", pady=2)
            self.filter_vars[column] = variable
            self.filter_entries[column] = entry

        self.filter_entries_canvas.yview_moveto(0.0)
        self.refresh_table()

    def current_column_filters(self) -> dict[str, str]:
        return {column: variable.get() for column, variable in self.filter_vars.items()}

    def all_columns(self) -> list[str]:
        if not self.document:
            return []
        return [*self.document.columns, *sorted(self.derived_columns)]

    def record_values(self, record: Record) -> dict[str, str]:
        derived = self.derived_values.get(record.index, {})
        return {**record.flattened, **derived}

    def compile_regex(self, pattern: str, label: str) -> re.Pattern[str] | None:
        normalized = pattern.strip()
        if not normalized:
            return None
        try:
            return re.compile(normalized, re.IGNORECASE)
        except re.error as exc:
            raise ValueError(f"Invalid regex in {label}: {exc}") from exc

    def filtered_records_for_view(self) -> list[Record]:
        if not self.document:
            return []

        global_pattern = self.compile_regex(self.global_search_var.get(), "global search")
        column_patterns = {
            column: self.compile_regex(value, f"filter '{column}'")
            for column, value in self.current_column_filters().items()
            if value.strip()
        }

        source_records = self.document.record_models if self.show_deleted_var.get() else self.document.active_records()
        filtered: list[Record] = []
        for record in source_records:
            values = self.record_values(record)
            if global_pattern:
                haystack = " | ".join(values.values())
                if not global_pattern.search(haystack):
                    continue

            include = True
            for column, pattern in column_patterns.items():
                if not pattern.search(values.get(column, "")):
                    include = False
                    break

            if include:
                filtered.append(record)

        return filtered

    def refresh_table(self) -> None:
        if not self.document:
            self.tree.delete(*self.tree.get_children())
            self.update_detail_panel()
            return

        try:
            records = self.filtered_records_for_view()
        except ValueError as exc:
            self.tree.delete(*self.tree.get_children())
            self.row_map.clear()
            self.update_detail_panel()
            self.status_var.set(str(exc))
            return

        if self.sort_column:
            records = sorted(
                records,
                key=lambda record: sort_key_for_values(
                    self.record_values(record),
                    record,
                    self.sort_column,
                    self.document.deleted_indices,
                ),
                reverse=self.sort_descending,
            )
        self.filtered_records = records

        columns = ["__status__", *(self.display_columns or default_columns(self.document.columns))]
        self.tree.configure(columns=columns)
        for column in columns:
            label = "status" if column == "__status__" else column
            if column == self.sort_column:
                label = f"{label} {'▼' if self.sort_descending else '▲'}"
            width = 90 if column == "__status__" else 170
            stretch = False if column == "__status__" else True
            self.tree.heading(column, text=label, command=lambda selected=column: self.toggle_sort(selected))
            self.tree.column(column, width=width, anchor="w", stretch=stretch)

        self.tree.delete(*self.tree.get_children())
        self.row_map.clear()

        for record in records:
            item_id = str(record.index)
            status = "deleted" if record.index in self.document.deleted_indices else "active"
            record_values = self.record_values(record)
            values = [status, *[record_values.get(column, "") for column in columns[1:]]]
            tags = ("deleted",) if status == "deleted" else ()
            self.tree.insert("", tk.END, iid=item_id, values=values, tags=tags)
            self.row_map[item_id] = record.index

        deleted_count = len(self.document.deleted_indices)
        self.status_var.set(
            f"Showing {len(records)} records. Pending deletions: {deleted_count}. "
            f"Backups will be written to {self.document.path.parent / '.backups'}"
        )
        self.update_detail_panel()

    def toggle_sort(self, column: str) -> None:
        if self.sort_column == column:
            self.sort_descending = not self.sort_descending
        else:
            self.sort_column = column
            self.sort_descending = False
        self.refresh_table()

    def selected_record_indices(self) -> list[int]:
        selected_ids = self.tree.selection()
        return [self.row_map[item_id] for item_id in selected_ids if item_id in self.row_map]

    def delete_selected_rows(self) -> None:
        if not self.document:
            return

        selected = self.selected_record_indices()
        if not selected:
            messagebox.showinfo("No selection", "Select one or more rows to mark for deletion.")
            return

        self.document.mark_deleted(selected)
        self.refresh_table()

    def restore_selected_rows(self) -> None:
        if not self.document:
            return

        selected = self.selected_record_indices()
        if not selected:
            messagebox.showinfo("No selection", "Select one or more currently visible rows to restore.")
            return

        self.document.restore_deleted(selected)
        self.refresh_table()

    def restore_all_rows(self) -> None:
        if not self.document:
            return
        self.document.restore_all()
        self.refresh_table()

    def update_detail_panel(self) -> None:
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", tk.END)

        selected_ids = self.tree.selection()
        if selected_ids and self.document:
            index = self.row_map.get(selected_ids[0])
            if index is not None:
                record = self.document.record_models[index].original
                self.detail_text.insert("1.0", json.dumps(record, indent=2, ensure_ascii=False, sort_keys=True))

        self.detail_text.configure(state="disabled")

    def show_diff_preview(self) -> None:
        if not self.document:
            return

        deleted = self.document.deleted_records()
        altered_summary = build_altered_rows_summary(deleted)

        window = tk.Toplevel(self.root)
        window.title("Pending Changes")
        window.geometry("1200x800")
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)

        summary = (
            f"Pending deletions: {len(deleted)} / {len(self.document.record_models)} records\n"
            f"Backup target directory: {self.document.path.parent / '.backups'}"
        )
        ttk.Label(window, text=summary, padding=8, justify="left").grid(row=0, column=0, sticky="ew")

        notebook = ttk.Notebook(window)
        notebook.grid(row=1, column=0, sticky="nsew")

        deleted_frame = ttk.Frame(notebook, padding=8)
        deleted_frame.columnconfigure(0, weight=1)
        deleted_frame.rowconfigure(0, weight=1)
        notebook.add(deleted_frame, text="Affected Rows")

        deleted_text = tk.Text(deleted_frame, wrap="none")
        deleted_text.grid(row=0, column=0, sticky="nsew")
        deleted_y = ttk.Scrollbar(deleted_frame, orient="vertical", command=deleted_text.yview)
        deleted_y.grid(row=0, column=1, sticky="ns")
        deleted_x = ttk.Scrollbar(deleted_frame, orient="horizontal", command=deleted_text.xview)
        deleted_x.grid(row=1, column=0, sticky="ew")
        deleted_text.configure(yscrollcommand=deleted_y.set, xscrollcommand=deleted_x.set)
        deleted_text.insert(
            "1.0",
            altered_summary,
        )
        deleted_text.configure(state="disabled")

        deleted_table_frame = ttk.Frame(notebook, padding=8)
        deleted_table_frame.columnconfigure(0, weight=1)
        deleted_table_frame.rowconfigure(0, weight=1)
        notebook.add(deleted_table_frame, text="Deleted Rows Table")

        preview_columns = ["__status__", *(self.display_columns or default_columns(self.document.columns))]
        deleted_tree = ttk.Treeview(deleted_table_frame, columns=preview_columns, show="headings")
        deleted_tree.grid(row=0, column=0, sticky="nsew")
        for column in preview_columns:
            label = "status" if column == "__status__" else column
            width = 90 if column == "__status__" else 180
            deleted_tree.heading(column, text=label)
            deleted_tree.column(column, width=width, anchor="w", stretch=True)

        deleted_tree.tag_configure("deleted", foreground="#a00000")
        for record in deleted:
            record_values = self.record_values(record)
            values = ["deleted", *[record_values.get(column, "") for column in preview_columns[1:]]]
            deleted_tree.insert("", tk.END, values=values, tags=("deleted",))

        deleted_table_y = ttk.Scrollbar(deleted_table_frame, orient="vertical", command=deleted_tree.yview)
        deleted_table_y.grid(row=0, column=1, sticky="ns")
        deleted_table_x = ttk.Scrollbar(deleted_table_frame, orient="horizontal", command=deleted_tree.xview)
        deleted_table_x.grid(row=1, column=0, sticky="ew")
        deleted_tree.configure(yscrollcommand=deleted_table_y.set, xscrollcommand=deleted_table_x.set)

        diff_frame = ttk.Frame(notebook, padding=8)
        diff_frame.columnconfigure(0, weight=1)
        diff_frame.rowconfigure(0, weight=1)
        notebook.add(diff_frame, text="Unified Diff")

        diff_widget = tk.Text(diff_frame, wrap="none")
        diff_widget.grid(row=0, column=0, sticky="nsew")
        diff_y = ttk.Scrollbar(diff_frame, orient="vertical", command=diff_widget.yview)
        diff_y.grid(row=0, column=1, sticky="ns")
        diff_x = ttk.Scrollbar(diff_frame, orient="horizontal", command=diff_widget.xview)
        diff_x.grid(row=1, column=0, sticky="ew")
        diff_widget.configure(yscrollcommand=diff_y.set, xscrollcommand=diff_x.set)
        diff_widget.insert("1.0", "Open this tab to generate the unified diff.")
        diff_widget.configure(state="disabled")

        diff_loaded = False

        def load_diff_if_needed() -> None:
            nonlocal diff_loaded
            if diff_loaded:
                return
            diff_loaded = True
            diff_widget.configure(state="normal")
            diff_widget.delete("1.0", tk.END)
            diff_widget.insert("1.0", "Generating diff...")
            diff_widget.update_idletasks()
            diff_text = self.document.preview_diff_text()
            diff_widget.delete("1.0", tk.END)
            diff_widget.insert("1.0", diff_text or "No changes.")
            diff_widget.configure(state="disabled")

        def on_tab_changed(_event: tk.Event) -> None:
            current_tab = notebook.tab(notebook.select(), "text")
            if current_tab == "Unified Diff":
                window.after(10, load_diff_if_needed)

        notebook.bind("<<NotebookTabChanged>>", on_tab_changed)

    def save_changes(self) -> None:
        if not self.document:
            return
        if not self.document.deleted_indices:
            messagebox.showinfo("No changes", "There are no pending deletions to save.")
            return

        deleted_count = len(self.document.deleted_indices)
        answer = messagebox.askyesno(
            "Confirm save",
            f"Create a backup and permanently remove {deleted_count} record(s) from the JSON file?",
        )
        if not answer:
            return

        try:
            result = self.document.save_with_backup()
        except Exception as exc:  # pragma: no cover - UI error path
            messagebox.showerror("Save failed", str(exc))
            return

        self._populate_column_selectors()
        self.refresh_table()
        messagebox.showinfo(
            "Saved",
            f"Backup created at:\n{result.backup_path}\n\nRemoved {result.deleted_count} record(s).",
        )

    def save_derived_column(self) -> None:
        if not self.document:
            return

        name = self.derived_name_var.get().strip()
        script = self.derived_script_text.get("1.0", tk.END).strip()
        if not name:
            messagebox.showinfo("Missing name", "Provide a name for the derived column.")
            return
        if name == "__status__":
            messagebox.showinfo("Reserved name", "`__status__` is reserved.")
            return
        if name in self.document.columns:
            messagebox.showinfo("Name conflict", "This name already exists as a real column.")
            return
        if not script:
            messagebox.showinfo("Missing script", "Provide a Python expression or script that sets `result`.")
            return

        try:
            computed = self.compute_derived_column_values(name, script)
        except Exception as exc:
            messagebox.showerror("Derived column failed", str(exc))
            return

        self.derived_columns[name] = script
        for record_index, value in computed.items():
            self.derived_values.setdefault(record_index, {})[name] = value

        self.refresh_derived_column_list()
        self._populate_column_selectors()
        self.status_var.set(f"Derived column '{name}' updated for {len(computed)} records.")

    def refresh_derived_column_list(self) -> None:
        current = self.derived_name_var.get().strip()
        self.derived_listbox.delete(0, tk.END)
        names = sorted(self.derived_columns)
        for name in names:
            self.derived_listbox.insert(tk.END, name)
        if current:
            if current in names:
                self.derived_listbox.selection_set(names.index(current))

    def load_selected_derived_column(self) -> None:
        selection = self.derived_listbox.curselection()
        if not selection:
            return
        name = self.derived_listbox.get(selection[0])
        self.derived_name_var.set(name)
        self.derived_script_text.delete("1.0", tk.END)
        self.derived_script_text.insert("1.0", self.derived_columns.get(name, ""))
        self.highlight_script_box()

    def remove_derived_column(self) -> None:
        selection = self.derived_listbox.curselection()
        if not selection:
            messagebox.showinfo("No selection", "Select a derived column to remove.")
            return
        name = self.derived_listbox.get(selection[0])
        self.derived_columns.pop(name, None)
        for values in self.derived_values.values():
            values.pop(name, None)
        self.refresh_derived_column_list()
        self._populate_column_selectors()
        self.status_var.set(f"Derived column '{name}' removed.")

    def compute_derived_column_values(self, name: str, script: str) -> dict[int, str]:
        if not self.document:
            return {}

        compiled: dict[int, str] = {}
        for record in self.document.record_models:
            merged = self.record_values(record)
            result = execute_derived_script(script, merged)
            compiled[record.index] = format_derived_result(result)
        return compiled

    def _configure_script_tags(self) -> None:
        self.derived_script_text.tag_configure("keyword", foreground="#7c3aed")
        self.derived_script_text.tag_configure("string", foreground="#047857")
        self.derived_script_text.tag_configure("comment", foreground="#6b7280")
        self.derived_script_text.tag_configure("number", foreground="#b45309")

    def highlight_script_box(self) -> None:
        text = self.derived_script_text.get("1.0", tk.END)
        for tag in ("keyword", "string", "comment", "number"):
            self.derived_script_text.tag_remove(tag, "1.0", tk.END)

        for match in re.finditer(r"\b(?:and|as|assert|break|class|continue|def|elif|else|except|False|finally|for|from|if|import|in|is|lambda|None|not|or|pass|raise|return|True|try|while|with|yield|result)\b", text):
            self._tag_script_match("keyword", match.start(), match.end())
        for match in re.finditer(r"#.*$", text, re.MULTILINE):
            self._tag_script_match("comment", match.start(), match.end())
        for match in re.finditer(r"(?:'[^'\\]*(?:\\.[^'\\]*)*'|\"[^\"\\]*(?:\\.[^\"\\]*)*\")", text):
            self._tag_script_match("string", match.start(), match.end())
        for match in re.finditer(r"\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b", text):
            self._tag_script_match("number", match.start(), match.end())

    def _tag_script_match(self, tag: str, start: int, end: int) -> None:
        start_index = f"1.0+{start}c"
        end_index = f"1.0+{end}c"
        self.derived_script_text.tag_add(tag, start_index, end_index)


def default_columns(columns: list[str], limit: int = 8) -> list[str]:
    scalar_like = [column for column in columns if "." not in column and "[" not in column]
    preferred = scalar_like or columns
    return preferred[:limit]


def sort_key_for_values(
    values: dict[str, str],
    record: Record,
    column: str,
    deleted_indices: set[int],
) -> tuple[int, int, str | float]:
    if column == "__status__":
        status = "deleted" if record.index in deleted_indices else "active"
        return (0, 0, status)

    raw_value = values.get(column, "")
    normalized = raw_value.strip().lower()

    try:
        numeric = float(normalized)
        return (0, 1, numeric)
    except ValueError:
        return (1, 1, normalized)


def build_altered_rows_summary(deleted: list[Record]) -> str:
    sections: list[str] = []

    if deleted:
        deleted_text = "\n\n".join(
            f"Status: deleted\nOriginal index: {record.index}\n"
            f"{json.dumps(record.original, indent=2, ensure_ascii=False, sort_keys=True)}"
            for record in deleted
        )
        sections.append(f"Deleted rows ({len(deleted)}):\n\n{deleted_text}")
    else:
        sections.append("Deleted rows (0):\n\nNone.")

    sections.append(
        "Altered rows (0):\n\n"
        "None. The current viewer supports deletion workflows only, so rows are either active or marked deleted."
    )

    return "\n\n".join(sections)


def execute_derived_script(script: str, record_values: dict[str, str]) -> object:
    def value(name: str, default: str = "") -> str:
        return record_values.get(name, default)

    def values(pattern: str) -> list[str]:
        return [val for key, val in record_values.items() if fnmatch.fnmatch(key, pattern)]

    def num(name: str, default: float | None = None) -> float | None:
        raw = record_values.get(name, "")
        if raw == "":
            return default
        return float(raw)

    def nums(pattern: str) -> list[float]:
        result: list[float] = []
        for item in values(pattern):
            if item == "":
                continue
            result.append(float(item))
        return result

    def mean(items: list[float], default: float | None = None) -> float | None:
        return statistics.fmean(items) if items else default

    safe_builtins = {
        "abs": abs,
        "all": all,
        "any": any,
        "bool": bool,
        "dict": dict,
        "enumerate": enumerate,
        "float": float,
        "int": int,
        "len": len,
        "list": list,
        "max": max,
        "min": min,
        "round": round,
        "set": set,
        "sorted": sorted,
        "str": str,
        "sum": sum,
        "tuple": tuple,
        "zip": zip,
    }

    local_scope = {
        "record": dict(record_values),
        "value": value,
        "values": values,
        "num": num,
        "nums": nums,
        "mean": mean,
        "math": math,
        "statistics": statistics,
        "result": None,
    }

    # Derived columns are intentionally GUI-only, but we still run user code inside a
    # small namespace so common calculations are easy while accidental side effects stay limited.
    try:
        code = compile(script, "<derived-column>", "eval")
        return eval(code, {"__builtins__": safe_builtins}, local_scope)
    except SyntaxError:
        code = compile(script, "<derived-column>", "exec")
        exec(code, {"__builtins__": safe_builtins}, local_scope)
        if local_scope.get("result") is None:
            raise ValueError("Derived-column script must evaluate to a value or assign to `result`.")
        return local_scope["result"]


def format_derived_result(result: object) -> str:
    if result is None:
        return ""
    if isinstance(result, float):
        return f"{result:.12g}"
    if isinstance(result, (list, dict, tuple, set)):
        return json.dumps(result, ensure_ascii=False)
    return str(result)
