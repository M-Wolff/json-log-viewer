from __future__ import annotations

import json
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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
        self.filter_vars: dict[str, tk.StringVar] = {}
        self.filter_entries: dict[str, ttk.Entry] = {}

        self.file_label_var = tk.StringVar(value="No file loaded")
        self.status_var = tk.StringVar(value="Open a JSON file to begin.")
        self.global_search_var = tk.StringVar()
        self.show_deleted_var = tk.BooleanVar(value=False)

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

        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.grid(row=1, column=0, sticky="nsew")

        controls = ttk.Frame(main_pane, padding=8)
        controls.columnconfigure(0, weight=1)
        main_pane.add(controls, weight=0)

        content = ttk.PanedWindow(main_pane, orient=tk.VERTICAL)
        main_pane.add(content, weight=1)

        self._build_controls(controls)
        self._build_content(content)

        status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor="w", padding=8)
        status_bar.grid(row=2, column=0, sticky="ew")

    def _build_controls(self, parent: ttk.Frame) -> None:
        search_frame = ttk.LabelFrame(parent, text="Search", padding=8)
        search_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        search_frame.columnconfigure(0, weight=1)

        ttk.Label(search_frame, text="Global contains").grid(row=0, column=0, sticky="w")
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
        columns_frame.rowconfigure(1, weight=1)
        parent.rowconfigure(1, weight=1)

        ttk.Label(columns_frame, text="Select visible columns").grid(row=0, column=0, sticky="w")

        self.display_listbox = tk.Listbox(columns_frame, selectmode=tk.EXTENDED, exportselection=False, height=14)
        self.display_listbox.grid(row=1, column=0, sticky="nsew")
        self.display_listbox.bind("<<ListboxSelect>>", lambda _event: self.apply_display_columns())
        display_scroll = ttk.Scrollbar(columns_frame, orient="vertical", command=self.display_listbox.yview)
        display_scroll.grid(row=1, column=1, sticky="ns")
        self.display_listbox.configure(yscrollcommand=display_scroll.set)

        display_buttons = ttk.Frame(columns_frame)
        display_buttons.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        display_buttons.columnconfigure(0, weight=1)
        display_buttons.columnconfigure(1, weight=1)
        ttk.Button(display_buttons, text="Select All", command=self.select_all_columns).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(display_buttons, text="Reset Defaults", command=self.reset_default_columns).grid(row=0, column=1, sticky="ew")

        filter_select_frame = ttk.LabelFrame(parent, text="Filter Columns", padding=8)
        filter_select_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        filter_select_frame.columnconfigure(0, weight=1)
        filter_select_frame.rowconfigure(1, weight=1)
        parent.rowconfigure(2, weight=1)

        ttk.Label(filter_select_frame, text="Select columns to filter on").grid(row=0, column=0, sticky="w")

        self.filter_column_listbox = tk.Listbox(filter_select_frame, selectmode=tk.EXTENDED, exportselection=False, height=10)
        self.filter_column_listbox.grid(row=1, column=0, sticky="nsew")
        self.filter_column_listbox.bind("<<ListboxSelect>>", lambda _event: self.rebuild_filter_entries())
        filter_select_scroll = ttk.Scrollbar(filter_select_frame, orient="vertical", command=self.filter_column_listbox.yview)
        filter_select_scroll.grid(row=1, column=1, sticky="ns")
        self.filter_column_listbox.configure(yscrollcommand=filter_select_scroll.set)

        filter_frame = ttk.LabelFrame(parent, text="Active Filters", padding=8)
        filter_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        filter_frame.columnconfigure(0, weight=1)

        canvas = tk.Canvas(filter_frame, height=220, highlightthickness=0)
        scrollbar = ttk.Scrollbar(filter_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        filter_frame.rowconfigure(0, weight=1)

        self.filter_entries_container = ttk.Frame(canvas)
        self.filter_entries_container.columnconfigure(1, weight=1)
        canvas_window = canvas.create_window((0, 0), window=self.filter_entries_container, anchor="nw")

        def _resize_canvas(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _resize_inner(event: tk.Event) -> None:
            canvas.itemconfigure(canvas_window, width=event.width)

        self.filter_entries_container.bind("<Configure>", _resize_canvas)
        canvas.bind("<Configure>", _resize_inner)

        actions = ttk.LabelFrame(parent, text="Actions", padding=8)
        actions.grid(row=4, column=0, sticky="ew")
        actions.columnconfigure(0, weight=1)

        ttk.Button(actions, text="Delete Selected Rows", command=self.delete_selected_rows).grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ttk.Button(actions, text="Restore Selected Rows", command=self.restore_selected_rows).grid(row=1, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Restore All Pending", command=self.restore_all_rows).grid(row=2, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Preview Changes", command=self.show_diff_preview).grid(row=3, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Save With Backup", command=self.save_changes).grid(row=4, column=0, sticky="ew", pady=(4, 0))

    def _build_content(self, parent: ttk.PanedWindow) -> None:
        table_frame = ttk.Frame(parent, padding=8)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        parent.add(table_frame, weight=3)

        detail_frame = ttk.Frame(parent, padding=8)
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(1, weight=1)
        parent.add(detail_frame, weight=1)

        self.tree = ttk.Treeview(table_frame, columns=(), show="headings", selectmode="extended")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", lambda _event: self.update_detail_panel())
        self.tree.tag_configure("deleted", foreground="#a00000")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        ttk.Label(detail_frame, text="Selected Record JSON").grid(row=0, column=0, sticky="w")
        self.detail_text = tk.Text(detail_frame, wrap="none")
        self.detail_text.grid(row=1, column=0, sticky="nsew")
        self.detail_text.configure(state="disabled")
        detail_y = ttk.Scrollbar(detail_frame, orient="vertical", command=self.detail_text.yview)
        detail_y.grid(row=1, column=1, sticky="ns")
        detail_x = ttk.Scrollbar(detail_frame, orient="horizontal", command=self.detail_text.xview)
        detail_x.grid(row=2, column=0, sticky="ew")
        self.detail_text.configure(yscrollcommand=detail_y.set, xscrollcommand=detail_x.set)

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
        self.show_deleted_var.set(False)
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

        columns = self.document.columns
        self.display_listbox.delete(0, tk.END)
        self.filter_column_listbox.delete(0, tk.END)
        for column in columns:
            self.display_listbox.insert(tk.END, column)
            self.filter_column_listbox.insert(tk.END, column)

        self.display_columns = default_columns(columns)
        self._select_listbox_values(self.display_listbox, self.display_columns)
        self.apply_display_columns()

        filter_defaults = self.display_columns[: min(3, len(self.display_columns))]
        self._select_listbox_values(self.filter_column_listbox, filter_defaults)
        self.rebuild_filter_entries()

    def _select_listbox_values(self, listbox: tk.Listbox, values: list[str]) -> None:
        options = listbox.get(0, tk.END)
        desired = set(values)
        for index, value in enumerate(options):
            if value in desired:
                listbox.selection_set(index)
            else:
                listbox.selection_clear(index)

    def apply_display_columns(self) -> None:
        selected_indices = self.display_listbox.curselection()
        self.display_columns = [self.display_listbox.get(index) for index in selected_indices]
        self.refresh_table()

    def select_all_columns(self) -> None:
        self.display_listbox.selection_set(0, tk.END)
        self.apply_display_columns()

    def reset_default_columns(self) -> None:
        if not self.document:
            return
        self.display_columns = default_columns(self.document.columns)
        self._select_listbox_values(self.display_listbox, self.display_columns)
        self.apply_display_columns()

    def rebuild_filter_entries(self) -> None:
        selected_columns = [self.filter_column_listbox.get(index) for index in self.filter_column_listbox.curselection()]
        existing_values = {column: var.get() for column, var in self.filter_vars.items()}

        for child in self.filter_entries_container.winfo_children():
            child.destroy()

        self.filter_vars.clear()
        self.filter_entries.clear()

        if not selected_columns:
            ttk.Label(self.filter_entries_container, text="No filter columns selected.").grid(row=0, column=0, sticky="w")
            self.refresh_table()
            return

        for row_index, column in enumerate(selected_columns):
            ttk.Label(self.filter_entries_container, text=column).grid(row=row_index, column=0, sticky="nw", padx=(0, 8), pady=2)
            variable = tk.StringVar(value=existing_values.get(column, ""))
            variable.trace_add("write", lambda *_args: self.refresh_table())
            entry = ttk.Entry(self.filter_entries_container, textvariable=variable)
            entry.grid(row=row_index, column=1, sticky="ew", pady=2)
            self.filter_vars[column] = variable
            self.filter_entries[column] = entry

        self.refresh_table()

    def current_column_filters(self) -> dict[str, str]:
        return {column: variable.get() for column, variable in self.filter_vars.items()}

    def refresh_table(self) -> None:
        if not self.document:
            self.tree.delete(*self.tree.get_children())
            self.update_detail_panel()
            return

        records = self.document.filtered_records(
            global_search=self.global_search_var.get(),
            column_filters=self.current_column_filters(),
            include_deleted=self.show_deleted_var.get(),
        )
        self.filtered_records = records

        columns = ["__status__", *(self.display_columns or default_columns(self.document.columns))]
        self.tree.configure(columns=columns)
        for column in columns:
            label = "status" if column == "__status__" else column
            width = 90 if column == "__status__" else 170
            stretch = False if column == "__status__" else True
            self.tree.heading(column, text=label)
            self.tree.column(column, width=width, anchor="w", stretch=stretch)

        self.tree.delete(*self.tree.get_children())
        self.row_map.clear()

        for record in records:
            item_id = str(record.index)
            status = "deleted" if record.index in self.document.deleted_indices else "active"
            values = [status, *[record.flattened.get(column, "") for column in columns[1:]]]
            tags = ("deleted",) if status == "deleted" else ()
            self.tree.insert("", tk.END, iid=item_id, values=values, tags=tags)
            self.row_map[item_id] = record.index

        deleted_count = len(self.document.deleted_indices)
        self.status_var.set(
            f"Showing {len(records)} records. Pending deletions: {deleted_count}. "
            f"Backups will be written to {self.document.path.parent / '.backups'}"
        )
        self.update_detail_panel()

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

        diff_text = self.document.preview_diff_text()
        deleted = self.document.deleted_records()

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
        notebook.add(deleted_frame, text="Removed Records")

        deleted_text = tk.Text(deleted_frame, wrap="none")
        deleted_text.grid(row=0, column=0, sticky="nsew")
        deleted_y = ttk.Scrollbar(deleted_frame, orient="vertical", command=deleted_text.yview)
        deleted_y.grid(row=0, column=1, sticky="ns")
        deleted_x = ttk.Scrollbar(deleted_frame, orient="horizontal", command=deleted_text.xview)
        deleted_x.grid(row=1, column=0, sticky="ew")
        deleted_text.configure(yscrollcommand=deleted_y.set, xscrollcommand=deleted_x.set)
        deleted_text.insert(
            "1.0",
            "\n\n".join(
                f"Original index: {record.index}\n{json.dumps(record.original, indent=2, ensure_ascii=False, sort_keys=True)}"
                for record in deleted
            )
            if deleted
            else "No pending deletions.",
        )
        deleted_text.configure(state="disabled")

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
        diff_widget.insert("1.0", diff_text or "No changes.")
        diff_widget.configure(state="disabled")

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


def default_columns(columns: list[str], limit: int = 8) -> list[str]:
    scalar_like = [column for column in columns if "." not in column and "[" not in column]
    preferred = scalar_like or columns
    return preferred[:limit]
