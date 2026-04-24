from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import difflib
import json
from pathlib import Path
import re
from typing import Any


JsonValue = dict[str, Any] | list[Any] | str | int | float | bool | None


@dataclass(slots=True)
class Record:
    index: int
    original: JsonValue
    flattened: dict[str, str]


@dataclass(slots=True)
class SaveResult:
    backup_path: Path
    deleted_count: int


class JsonDocument:
    """Stateful model for a JSON file whose primary editable unit is a list of records.

    The viewer edits a single list inside the JSON document. We support three common
    shapes reliably:
    1. a top-level list of records
    2. a top-level object with one or more nested lists of records
    3. a top-level object treated as a single record when no list is present

    For nested objects containing several candidate record lists, we choose the largest
    list of objects/scalars. This keeps the UI automatic for experiment logs while still
    preserving the full original JSON structure on write.
    """

    def __init__(self, path: Path, root_data: JsonValue, record_path: tuple[Any, ...], records: list[JsonValue]):
        self.path = path
        self.root_data = root_data
        self.record_path = record_path
        self.records = records
        self.deleted_indices: set[int] = set()
        self.record_models = [
            Record(index=index, original=record, flattened=flatten_json(record))
            for index, record in enumerate(records)
        ]
        self.columns = sorted({key for record in self.record_models for key in record.flattened})

    @classmethod
    def load(cls, path: str | Path) -> "JsonDocument":
        file_path = Path(path)
        with file_path.open("r", encoding="utf-8") as handle:
            root_data = json.load(handle)

        record_path, records = locate_record_list(root_data)
        return cls(path=file_path, root_data=root_data, record_path=record_path, records=records)

    def active_records(self) -> list[Record]:
        return [record for record in self.record_models if record.index not in self.deleted_indices]

    def filtered_records(
        self,
        global_search: str,
        column_filters: dict[str, str],
        include_deleted: bool = False,
    ) -> list[Record]:
        global_pattern = compile_search_pattern(global_search, "global search")
        compiled_filters = {
            column: compile_search_pattern(value, f"filter '{column}'")
            for column, value in column_filters.items()
            if value.strip()
        }

        filtered: list[Record] = []
        source_records = self.record_models if include_deleted else self.active_records()
        for record in source_records:
            if global_pattern:
                haystack = " | ".join(record.flattened.values())
                if not global_pattern.search(haystack):
                    continue

            include = True
            for column, pattern in compiled_filters.items():
                actual = record.flattened.get(column, "")
                if not pattern.search(actual):
                    include = False
                    break

            if include:
                filtered.append(record)

        return filtered

    def mark_deleted(self, indices: list[int]) -> None:
        self.deleted_indices.update(indices)

    def restore_deleted(self, indices: list[int]) -> None:
        self.deleted_indices.difference_update(indices)

    def restore_all(self) -> None:
        self.deleted_indices.clear()

    def deleted_records(self) -> list[Record]:
        return [record for record in self.record_models if record.index in self.deleted_indices]

    def pending_data(self) -> JsonValue:
        kept_records = [record for index, record in enumerate(self.records) if index not in self.deleted_indices]
        return set_at_path(self.root_data, self.record_path, kept_records)

    def preview_diff_text(self) -> str:
        original_text = json.dumps(self.root_data, indent=2, ensure_ascii=False)
        updated_text = json.dumps(self.pending_data(), indent=2, ensure_ascii=False)
        diff = difflib.unified_diff(
            original_text.splitlines(),
            updated_text.splitlines(),
            fromfile=str(self.path),
            tofile=str(self.path),
            lineterm="",
        )
        return "\n".join(diff)

    def save_with_backup(self) -> SaveResult:
        backup_dir = self.path.parent / ".backups"
        backup_dir.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_path = backup_dir / f"{self.path.stem}-{timestamp}{self.path.suffix}"

        deleted_count = len(self.deleted_indices)

        # The backup is written before the original file is touched so that an operator can
        # always recover the prior state even if the later write fails or is interrupted.
        backup_path.write_text(
            json.dumps(self.root_data, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )

        updated_data = self.pending_data()
        self.path.write_text(
            json.dumps(updated_data, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )

        self.root_data = updated_data
        self.records = [record.original for record in self.active_records()]
        self.deleted_indices.clear()
        self.record_models = [
            Record(index=index, original=record, flattened=flatten_json(record))
            for index, record in enumerate(self.records)
        ]
        self.columns = sorted({key for record in self.record_models for key in record.flattened})

        return SaveResult(backup_path=backup_path, deleted_count=deleted_count)


def locate_record_list(root_data: JsonValue) -> tuple[tuple[Any, ...], list[JsonValue]]:
    if isinstance(root_data, list):
        return (), root_data

    candidates: list[tuple[tuple[Any, ...], list[JsonValue]]] = []
    collect_list_candidates(root_data, (), candidates)
    if candidates:
        candidates.sort(key=lambda item: candidate_score(item[1]), reverse=True)
        return candidates[0]

    if isinstance(root_data, dict):
        return (), [root_data]

    raise ValueError("JSON root must be a list or an object containing records.")


def collect_list_candidates(current: JsonValue, path: tuple[Any, ...], candidates: list[tuple[tuple[Any, ...], list[JsonValue]]]) -> None:
    if isinstance(current, list):
        candidates.append((path, current))
        for index, value in enumerate(current):
            if isinstance(value, (dict, list)):
                collect_list_candidates(value, path + (index,), candidates)
        return

    if isinstance(current, dict):
        for key, value in current.items():
            if isinstance(value, (dict, list)):
                collect_list_candidates(value, path + (key,), candidates)


def candidate_score(records: list[JsonValue]) -> tuple[int, int, int]:
    structured_items = sum(isinstance(item, (dict, list)) for item in records)
    dict_items = sum(isinstance(item, dict) for item in records)
    return (dict_items, structured_items, len(records))


def set_at_path(root_data: JsonValue, path: tuple[Any, ...], replacement: JsonValue) -> JsonValue:
    if not path:
        return replacement

    if isinstance(root_data, dict):
        clone = dict(root_data)
        head = path[0]
        clone[head] = set_at_path(clone[head], path[1:], replacement)
        return clone

    if isinstance(root_data, list):
        clone = list(root_data)
        head = path[0]
        clone[head] = set_at_path(clone[head], path[1:], replacement)
        return clone

    raise ValueError("Cannot replace a nested path inside a scalar value.")


def flatten_json(value: JsonValue, prefix: str = "") -> dict[str, str]:
    """Flatten nested dict/list data into dotted columns for table display.

    The GUI needs stable, comparable column names even when records contain nested
    dictionaries or lists. We therefore convert nested structures into path-like keys such
    as `train_losses.mean.mse` or `layers[0]` while preserving scalar values as strings.
    """

    flattened: dict[str, str] = {}

    if isinstance(value, dict):
        if not value:
            flattened[prefix or "$"] = "{}"
        for key, child in value.items():
            child_prefix = f"{prefix}.{key}" if prefix else str(key)
            flattened.update(flatten_json(child, child_prefix))
        return flattened

    if isinstance(value, list):
        if not value:
            flattened[prefix or "$"] = "[]"
        for index, child in enumerate(value):
            child_prefix = f"{prefix}[{index}]" if prefix else f"[{index}]"
            flattened.update(flatten_json(child, child_prefix))
        return flattened

    flattened[prefix or "$"] = format_scalar(value)
    return flattened


def format_scalar(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, float):
        return f"{value:.12g}"
    return str(value)


def compile_search_pattern(pattern: str, label: str) -> re.Pattern[str] | None:
    normalized = pattern.strip()
    if not normalized:
        return None

    try:
        return re.compile(normalized, re.IGNORECASE)
    except re.error as exc:
        raise ValueError(f"Invalid regex in {label}: {exc}") from exc
