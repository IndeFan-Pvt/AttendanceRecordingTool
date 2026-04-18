from __future__ import annotations

import argparse
import calendar
import html
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import win32com.client as win32
from win32com.client import dynamic
import xlrd
from ortools.sat.python import cp_model


def configure_stdio() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


configure_stdio()


BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = BASE_DIR / "akanecco_shift_config.json"
DEFAULT_REPORT_PATH = BASE_DIR / "akanecco_generated_validation.html"
DEFAULT_SHIFT_KINDS = {
    "": "rest",
    "早": "early",
    "遅": "late",
    "日": "day",
    "夜": "night",
    "夜休": "night_rest",
    "休": "rest",
    "特": "rest",
    "5.5": "day",
    "6.0": "day",
}
DEFAULT_COUNT_SYMBOLS = {"夜": "夜", "早": "早", "遅": "遅", "休": "休"}
JAPANESE_WEEKDAYS = ["月", "火", "水", "木", "金", "土", "日"]
XLS_EMPLOYEE_SETTING_HEADER_ROW_INDEX = 4
XLS_EMPLOYEE_SETTING_COLUMNS = {
    "night_fairness_target": (41, "夜勤公平化対象"),
    "required_double_night_target": (42, "夜夜必須対象"),
    "required_double_night_min_count": (43, "夜夜必須回数"),
    "weekend_fairness_target": (44, "土日休公平化対象"),
    "max_consecutive_work_limit": (45, "個別連勤上限"),
    "max_four_day_streak_count": (46, "4連勤許容回数"),
}
XLS_MONTHLY_SETTING_HEADER_CELL = (0, 47)
XLS_MONTHLY_SETTING_CELLS = {
    "fairness_night_spread": (1, 47),
    "fairness_weekend_spread": (1, 48),
    "weekend_rest_count_mode": (1, 49),
}
TRUE_MARKERS = {"○", "〇", "1", "true", "yes", "y", "on"}
FALSE_MARKERS = {"×", "✕", "x", "0", "false", "no", "n", "off"}
WEEKEND_REST_COUNT_MODE_ALIASES = {
    "休のみ": "rest_only",
    "休": "rest_only",
    "rest_only": "rest_only",
    "休+夜休": "rest_and_night_rest",
    "休+夜": "rest_and_night_rest",
    "rest_and_night_rest": "rest_and_night_rest",
    "休+特+夜休": "rest_special_night_rest",
    "休+特+夜": "rest_special_night_rest",
    "rest_special_night_rest": "rest_special_night_rest",
}


@dataclass(frozen=True)
class EmployeeConfig:
    employee_id: str
    display_name: str
    unit: str
    employment: str
    row: int
    allowed_shifts: tuple[str, ...]
    aliases: tuple[str, ...] = ()
    weekday_allowed_shifts: dict[int, tuple[str, ...]] = field(default_factory=dict)
    date_allowed_shift_overrides: dict[int, tuple[str, ...]] = field(default_factory=dict)
    require_weekend_pair_rest: bool = False
    night_fairness_target: bool = False
    required_double_night_min_count: int | None = None
    weekend_fairness_target: bool = False
    min_counts: dict[str, int] = field(default_factory=dict)
    max_counts: dict[str, int] = field(default_factory=dict)
    max_consecutive_work_limit: int | None = None
    max_four_day_streak_count: int | None = None
    exact_rest_days: int | None = None
    min_rest_days: int | None = None
    max_rest_days: int | None = None
    specified_holidays: tuple[int, ...] = ()
    fixed_assignments: dict[int, str] = field(default_factory=dict)
    previous_tail: tuple[str, ...] = ()


@dataclass(frozen=True)
class SchedulerConfig:
    config_path: Path
    target_path: Path
    manual_source: Path
    sheet_index: int
    year: int
    month: int
    days_in_month: int
    unit_name: str
    shift_kinds: dict[str, str]
    count_symbols: dict[str, str]
    employees: tuple[EmployeeConfig, ...]
    required_per_day: dict[str, dict[str, int]]
    night_total_per_day: int
    day_requirements: dict[int, dict[str, dict[str, int]]]
    max_consecutive_work: int
    max_consecutive_night: int
    max_consecutive_rest: int
    max_consecutive_rest_with_special: int
    fairness_night_spread: int | None
    fairness_weekend_spread: int | None
    weekend_rest_count_mode: str
    require_weekend_pair_rest: bool
    prefer_double_night: bool


@dataclass(frozen=True)
class ScheduleSolveResult:
    schedule: dict[str, list[str]]
    is_partial: bool = False
    message: str = ""
    diagnostics: dict[str, object] = field(default_factory=dict)


def default_config_dict() -> dict[str, object]:
    return {
        "year": 2026,
        "month": 2,
        "unit_name": "あかねっこ",
        "target_path": "【白紙】あかねっこ.xls",
        "manual_source": "元データ/シフト自動作成、特養いちご/シフト自動作成、特養いちご/R7年勤務表/R7.2月/【統一書式】あかねっこ2月.xls",
        "sheet_index": 1,
        "shift_kinds": DEFAULT_SHIFT_KINDS,
        "count_symbols": DEFAULT_COUNT_SYMBOLS,
        "rules": {
            "required_per_day": {
                "南": {"early": 1, "late": 1},
                "北": {"early": 1, "late": 1},
                "night_total": 1,
            },
            "max_consecutive_work": 5,
            "max_consecutive_night": 2,
            "max_consecutive_rest": 3,
            "max_consecutive_rest_with_special": 5,
            "fairness_night_spread": 1,
            "fairness_weekend_spread": 1,
            "require_weekend_pair_rest": True,
            "prefer_double_night": True,
            "day_requirements": {},
        },
        "employees": [
            {
                "name": "笹谷 正人",
                "unit": "南",
                "employment": "full",
                "row": 6,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "川部 春香",
                "unit": "南",
                "employment": "full",
                "row": 8,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "都築 良弘",
                "unit": "南",
                "employment": "full",
                "row": 10,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "嶋﨑 光子",
                "unit": "南",
                "employment": "full",
                "row": 12,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "八重 美貴",
                "unit": "南",
                "employment": "part",
                "row": 14,
                "allowed_shifts": ["", "5.5", "休"],
                "min_rest_days": 11,
                "max_rest_days": 18,
            },
            {
                "name": "稲富茜里",
                "unit": "南",
                "employment": "part",
                "row": 16,
                "allowed_shifts": ["", "6.0", "休"],
                "min_rest_days": 11,
                "max_rest_days": 18,
            },
            {
                "name": "駒倉 利光",
                "unit": "北",
                "employment": "full",
                "row": 21,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "竹内孝寛",
                "unit": "北",
                "employment": "full",
                "row": 23,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "岡 虹光",
                "unit": "北",
                "employment": "full",
                "row": 25,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "鈴木 美奈",
                "unit": "北",
                "employment": "full",
                "row": 27,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "菅野 実里",
                "unit": "北",
                "employment": "full",
                "row": 29,
                "allowed_shifts": ["早", "遅", "日", "夜", "休", "特"],
                "min_counts": {"夜": 3, "日": 1},
                "max_counts": {"夜": 4},
                "min_rest_days": 6,
                "max_rest_days": 10,
            },
            {
                "name": "石井 都子",
                "unit": "北",
                "employment": "part",
                "row": 31,
                "allowed_shifts": ["", "6.0", "休"],
                "min_rest_days": 11,
                "max_rest_days": 18,
            },
        ],
    }


def normalize_cell_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("　", " ").strip()


def normalize_employee_name(value: object) -> str:
    return re.sub(r"\s+", "", normalize_cell_text(value))


def fallback_employee_id(display_name: str, employee_index: int) -> str:
    normalized_name = normalize_employee_name(display_name)
    if normalized_name:
        return f"legacy:{normalized_name}"
    return f"legacy:employee-{employee_index + 1:03d}"


def normalize_employee_aliases(employee_raw: dict[str, object], display_name: str, employee_id: str) -> tuple[str, ...]:
    raw_aliases = employee_raw.get("aliases", [])
    if raw_aliases is None:
        raw_aliases = []
    if not isinstance(raw_aliases, list):
        raise ValueError(f"{display_name} ({employee_id}) の aliases は配列で指定してください。")

    aliases: list[str] = []
    normalized_seen: set[str] = set()
    legacy_name = normalize_cell_text(employee_raw.get("name"))
    alias_sources: list[object] = [display_name]
    if legacy_name and normalize_employee_name(legacy_name) != normalize_employee_name(display_name):
        alias_sources.append(legacy_name)
    alias_sources.extend(raw_aliases)

    for raw_alias in alias_sources:
        alias = normalize_cell_text(raw_alias)
        normalized_alias = normalize_employee_name(alias)
        if not normalized_alias or normalized_alias in normalized_seen:
            continue
        aliases.append(alias)
        normalized_seen.add(normalized_alias)
    return tuple(aliases)


def validate_employee_identity_constraints(employees: list[EmployeeConfig]) -> None:
    employee_names_by_id: dict[str, str] = {}
    employee_ids_by_alias: dict[str, str] = {}

    for employee in employees:
        existing_name = employee_names_by_id.get(employee.employee_id)
        if existing_name is not None:
            raise ValueError(
                f"employee_id が重複しています: {employee.employee_id} ({existing_name}, {employee.display_name})"
            )
        employee_names_by_id[employee.employee_id] = employee.display_name

        for alias in employee.aliases:
            normalized_alias = normalize_employee_name(alias)
            if not normalized_alias:
                continue
            existing_employee_id = employee_ids_by_alias.get(normalized_alias)
            if existing_employee_id is not None and existing_employee_id != employee.employee_id:
                raise ValueError(
                    f"aliases の正規化値が重複しています: {alias} ({employee.employee_id}, {existing_employee_id})"
                )
            employee_ids_by_alias[normalized_alias] = employee.employee_id


def normalize_employee_rest_day_targets(
    employee_raw: dict[str, object],
    display_name: str,
    days_in_month: int,
) -> tuple[int | None, int | None, int | None]:
    exact_rest_days = employee_raw.get("exact_rest_days")
    min_rest_days = employee_raw.get("min_rest_days")
    max_rest_days = employee_raw.get("max_rest_days")

    exact_value = None if exact_rest_days is None else int(exact_rest_days)
    min_value = None if min_rest_days is None else int(min_rest_days)
    max_value = None if max_rest_days is None else int(max_rest_days)

    for label, value in (("exact_rest_days", exact_value), ("min_rest_days", min_value), ("max_rest_days", max_value)):
        if value is None:
            continue
        if value < 0 or value > days_in_month:
            raise ValueError(f"{display_name} の {label} は 0 以上 {days_in_month} 以下で指定してください。")

    if min_value is not None and max_value is not None and min_value > max_value:
        raise ValueError(f"{display_name} の min_rest_days は max_rest_days 以下で指定してください。")

    if exact_value is not None:
        if min_value is not None and min_value != exact_value:
            raise ValueError(f"{display_name} の exact_rest_days と min_rest_days が一致していません。")
        if max_value is not None and max_value != exact_value:
            raise ValueError(f"{display_name} の exact_rest_days と max_rest_days が一致していません。")

    return exact_value, min_value, max_value


def display_symbol(symbol: str) -> str:
    return "空欄" if symbol == "" else symbol


def sheet_cell_text(worksheet, row_index: int, column_index: int) -> str:
    if row_index < 0 or column_index < 0:
        return ""
    if row_index >= worksheet.nrows or column_index >= worksheet.ncols:
        return ""
    return normalize_cell_text(worksheet.cell_value(row_index, column_index))


def parse_workbook_bool(value: object, label: str) -> bool:
    text = normalize_cell_text(value)
    if not text:
        return False
    normalized = text.lower()
    if text in TRUE_MARKERS or normalized in TRUE_MARKERS:
        return True
    if text in FALSE_MARKERS or normalized in FALSE_MARKERS:
        return False
    raise ValueError(f"勤務表の {label} は ○ / × / TRUE / FALSE / 1 / 0 で指定してください。入力値: {text}")


def parse_workbook_optional_int(value: object, label: str) -> int | None:
    text = normalize_cell_text(value)
    if not text:
        return None
    try:
        parsed = int(float(text))
    except Exception as exc:
        raise ValueError(f"勤務表の {label} は整数で指定してください。入力値: {text}") from exc
    if str(parsed) != text and not re.fullmatch(r"\d+\.0+", text):
        raise ValueError(f"勤務表の {label} は整数で指定してください。入力値: {text}")
    return parsed


def normalize_weekend_rest_count_mode(value: object) -> str:
    text = normalize_cell_text(value)
    if not text:
        return "rest_special_night_rest"
    normalized = WEEKEND_REST_COUNT_MODE_ALIASES.get(text)
    if normalized is None:
        raise ValueError(
            "weekend_rest_count_mode は 休のみ / 休+夜休 / 休+特+夜休 のいずれかで指定してください。"
        )
    return normalized


def weekend_rest_symbols_for_mode(shift_kinds: dict[str, str], mode: str) -> tuple[str, ...]:
    night_rest_symbols = tuple(symbol_names_by_kind(shift_kinds, "night_rest"))
    if mode == "rest_only":
        return (("休",) if "休" in shift_kinds else tuple())
    if mode == "rest_and_night_rest":
        base = ("休",) if "休" in shift_kinds else tuple()
        return tuple(dict.fromkeys((*base, *night_rest_symbols)))
    if mode == "rest_special_night_rest":
        rest_symbols = tuple(symbol_names_by_kind(shift_kinds, "rest"))
        return tuple(dict.fromkeys((*rest_symbols, *night_rest_symbols)))
    raise ValueError(f"未対応の weekend_rest_count_mode です: {mode}")


def employee_max_consecutive_work(employee: EmployeeConfig, config: SchedulerConfig) -> int:
    return employee.max_consecutive_work_limit if employee.max_consecutive_work_limit is not None else config.max_consecutive_work


def count_consecutive_work_windows(shifts: list[str], work_symbols: set[str], window_size: int) -> int:
    if window_size <= 0 or len(shifts) < window_size:
        return 0
    total = 0
    for start in range(len(shifts) - window_size + 1):
        if all(shift in work_symbols for shift in shifts[start : start + window_size]):
            total += 1
    return total


def selected_night_fairness_employee_ids(config: SchedulerConfig, night_symbols: list[str]) -> list[str]:
    explicit_targets = [
        employee.employee_id
        for employee in config.employees
        if employee.night_fairness_target and any(shift in night_symbols for shift in employee.allowed_shifts)
    ]
    if explicit_targets:
        return explicit_targets
    return [
        employee.employee_id
        for employee in config.employees
        if any(shift in night_symbols for shift in employee.allowed_shifts)
    ]


def selected_weekend_fairness_employee_ids(config: SchedulerConfig) -> list[str]:
    explicit_targets = [employee.employee_id for employee in config.employees if employee.weekend_fairness_target]
    if explicit_targets:
        return explicit_targets
    return [employee.employee_id for employee in config.employees if employee.employment == "full"]


def read_workbook_employee_settings(
    workbook_path: Path,
    sheet_index: int,
    employees: tuple[EmployeeConfig, ...],
) -> dict[str, dict[str, object]]:
    worksheet = xlrd.open_workbook(str(workbook_path)).sheet_by_index(sheet_index - 1)
    active_columns = {
        field_name: column_index
        for field_name, (column_index, header_label) in XLS_EMPLOYEE_SETTING_COLUMNS.items()
        if sheet_cell_text(worksheet, XLS_EMPLOYEE_SETTING_HEADER_ROW_INDEX, column_index) == header_label
    }
    if not active_columns:
        return {}

    row_map = build_employee_row_map(worksheet)
    results: dict[str, dict[str, object]] = {}
    for employee in employees:
        row_index = resolve_employee_row_index(worksheet, employee, row_map)
        if row_index is None:
            continue
        current: dict[str, object] = {}
        if "night_fairness_target" in active_columns:
            column_index = active_columns["night_fairness_target"]
            current["night_fairness_target"] = parse_workbook_bool(
                worksheet.cell_value(row_index, column_index),
                f"{employee.display_name} / 夜勤公平化対象",
            )
        required_double_night_target = False
        if "required_double_night_target" in active_columns:
            column_index = active_columns["required_double_night_target"]
            required_double_night_target = parse_workbook_bool(
                worksheet.cell_value(row_index, column_index),
                f"{employee.display_name} / 夜夜必須対象",
            )
        required_double_night_min_count = None
        if "required_double_night_min_count" in active_columns:
            column_index = active_columns["required_double_night_min_count"]
            required_double_night_min_count = parse_workbook_optional_int(
                worksheet.cell_value(row_index, column_index),
                f"{employee.display_name} / 夜夜必須回数",
            )
            if required_double_night_min_count is not None and required_double_night_min_count < 0:
                raise ValueError(f"{employee.display_name} の 夜夜必須回数 は 0 以上で指定してください。")
        if required_double_night_target or required_double_night_min_count is not None:
            current["required_double_night_min_count"] = (
                required_double_night_min_count if required_double_night_min_count is not None else 1
            )
        if "weekend_fairness_target" in active_columns:
            column_index = active_columns["weekend_fairness_target"]
            current["weekend_fairness_target"] = parse_workbook_bool(
                worksheet.cell_value(row_index, column_index),
                f"{employee.display_name} / 土日休公平化対象",
            )
        if "max_consecutive_work_limit" in active_columns:
            column_index = active_columns["max_consecutive_work_limit"]
            max_consecutive_work_limit = parse_workbook_optional_int(
                worksheet.cell_value(row_index, column_index),
                f"{employee.display_name} / 個別連勤上限",
            )
            if max_consecutive_work_limit is not None:
                if max_consecutive_work_limit <= 0:
                    raise ValueError(f"{employee.display_name} の 個別連勤上限 は 1 以上で指定してください。")
                current["max_consecutive_work_limit"] = max_consecutive_work_limit
        if "max_four_day_streak_count" in active_columns:
            column_index = active_columns["max_four_day_streak_count"]
            max_four_day_streak_count = parse_workbook_optional_int(
                worksheet.cell_value(row_index, column_index),
                f"{employee.display_name} / 4連勤許容回数",
            )
            if max_four_day_streak_count is not None:
                if max_four_day_streak_count < 0:
                    raise ValueError(f"{employee.display_name} の 4連勤許容回数 は 0 以上で指定してください。")
                current["max_four_day_streak_count"] = max_four_day_streak_count
        if current:
            results[employee.employee_id] = current
    return results


def read_workbook_monthly_settings(workbook_path: Path, sheet_index: int) -> dict[str, object]:
    worksheet = xlrd.open_workbook(str(workbook_path)).sheet_by_index(sheet_index - 1)
    header_text = sheet_cell_text(worksheet, *XLS_MONTHLY_SETTING_HEADER_CELL)
    has_monthly_settings = header_text == "月次設定" or any(
        sheet_cell_text(worksheet, row_index, column_index)
        for row_index, column_index in XLS_MONTHLY_SETTING_CELLS.values()
    )
    if not has_monthly_settings:
        return {}

    overrides: dict[str, object] = {}
    fairness_night_spread = parse_workbook_optional_int(
        worksheet.cell_value(*XLS_MONTHLY_SETTING_CELLS["fairness_night_spread"]),
        "AV2 / 夜勤公平化許容差",
    )
    if fairness_night_spread is not None:
        overrides["fairness_night_spread"] = fairness_night_spread

    fairness_weekend_spread = parse_workbook_optional_int(
        worksheet.cell_value(*XLS_MONTHLY_SETTING_CELLS["fairness_weekend_spread"]),
        "AW2 / 土日休公平化許容差",
    )
    if fairness_weekend_spread is not None:
        overrides["fairness_weekend_spread"] = fairness_weekend_spread

    weekend_rest_count_mode = sheet_cell_text(worksheet, *XLS_MONTHLY_SETTING_CELLS["weekend_rest_count_mode"])
    if weekend_rest_count_mode:
        overrides["weekend_rest_count_mode"] = normalize_weekend_rest_count_mode(weekend_rest_count_mode)
    return overrides


def resolve_path(base_dir: Path, raw_path: str | Path) -> Path:
    path = Path(raw_path)
    if path.is_absolute():
        return path
    return (base_dir / path).resolve()


def create_excel_application():
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        return dynamic.DumbDispatch(dynamic.Dispatch("Excel.Application"))


def weekend_day_indexes(year: int, month: int, days_in_month: int) -> list[int]:
    return [day - 1 for day in range(1, days_in_month + 1) if calendar.weekday(year, month, day) >= 5]


def weekend_pair_day_indexes(year: int, month: int, days_in_month: int) -> list[tuple[int, int]]:
    weekend_pairs: list[tuple[int, int]] = []
    for day in range(1, days_in_month):
        if calendar.weekday(year, month, day) == 5 and calendar.weekday(year, month, day + 1) == 6:
            weekend_pairs.append((day - 1, day))
    return weekend_pairs


def symbol_names_by_kind(shift_kinds: dict[str, str], kind: str) -> list[str]:
    return [symbol for symbol, symbol_kind in shift_kinds.items() if symbol_kind == kind]


def first_symbol_by_kind(shift_kinds: dict[str, str], kind: str) -> str | None:
    symbols = symbol_names_by_kind(shift_kinds, kind)
    return symbols[0] if symbols else None


def standard_day_symbols(shift_kinds: dict[str, str]) -> list[str]:
    return ["日"] if shift_kinds.get("日") == "day" else []


def normalize_night_rest_sequence(
    sequence: list[str] | tuple[str, ...],
    shift_kinds: dict[str, str],
    previous_shift: str | None = None,
) -> list[str]:
    night_symbols = set(symbol_names_by_kind(shift_kinds, "night"))
    night_rest_symbol = first_symbol_by_kind(shift_kinds, "night_rest")
    if not night_symbols or night_rest_symbol is None:
        return list(sequence)

    normalized: list[str] = []
    prior_shift = previous_shift
    for shift in sequence:
        current_shift = shift
        if current_shift == "休" and prior_shift in night_symbols:
            current_shift = night_rest_symbol
        normalized.append(current_shift)
        prior_shift = current_shift
    return normalized


def normalize_night_rest_assignments(
    assignments: dict[int, str],
    shift_kinds: dict[str, str],
    days_in_month: int,
    previous_shift: str | None = None,
) -> dict[int, str]:
    night_symbols = set(symbol_names_by_kind(shift_kinds, "night"))
    night_rest_symbol = first_symbol_by_kind(shift_kinds, "night_rest")
    if not night_symbols or night_rest_symbol is None:
        return dict(assignments)

    normalized = dict(assignments)
    prior_shift = previous_shift
    for day in range(1, days_in_month + 1):
        if day not in normalized:
            prior_shift = None
            continue
        current_shift = normalized[day]
        if current_shift == "休" and prior_shift in night_symbols:
            current_shift = night_rest_symbol
            normalized[day] = current_shift
        prior_shift = current_shift
    return normalized


def parse_weekday_key(value: object) -> int:
    normalized = normalize_cell_text(value)
    weekday_map = {
        "0": 0,
        "1": 1,
        "2": 2,
        "3": 3,
        "4": 4,
        "5": 5,
        "6": 6,
        "月": 0,
        "火": 1,
        "水": 2,
        "木": 3,
        "金": 4,
        "土": 5,
        "日": 6,
        "mon": 0,
        "tue": 1,
        "wed": 2,
        "thu": 3,
        "fri": 4,
        "sat": 5,
        "sun": 6,
    }
    key = normalized.lower()
    if key in weekday_map:
        return weekday_map[key]
    raise ValueError(f"曜日指定が不正です: {value}")


def normalize_allowed_shift_rule(
    shifts: object,
    shift_kinds: dict[str, str],
    employee_name: str,
    rule_label: str,
    allow_night_rest: bool,
) -> tuple[str, ...]:
    if not isinstance(shifts, list):
        raise ValueError(f"{employee_name} の {rule_label} は配列で指定してください。")

    normalized: list[str] = []
    for raw_shift in shifts:
        shift = str(raw_shift)
        if shift not in shift_kinds:
            raise ValueError(f"{employee_name} の {rule_label} に未定義の勤務記号があります: {shift}")
        normalized.append(shift)

    night_rest_symbol = first_symbol_by_kind(shift_kinds, "night_rest")
    if allow_night_rest and night_rest_symbol and "休" in normalized and night_rest_symbol not in normalized:
        normalized.append(night_rest_symbol)
    return tuple(normalized)


def effective_allowed_shifts_for_day(config: SchedulerConfig, employee: EmployeeConfig, day: int) -> tuple[str, ...] | None:
    if day in employee.date_allowed_shift_overrides:
        return employee.date_allowed_shift_overrides[day]

    weekday = calendar.weekday(config.year, config.month, day)
    if weekday in employee.weekday_allowed_shifts:
        return employee.weekday_allowed_shifts[weekday]
    return None


def parse_excel_day_number(value: object) -> int | None:
    if isinstance(value, float):
        return int(value) if value.is_integer() else None
    if isinstance(value, int):
        return value
    text = normalize_cell_text(value)
    match = re.fullmatch(r"(\d+)(?:\.0+)?", text)
    return int(match.group(1)) if match else None


def open_workbook(excel, path: Path):
    workbooks = excel.Workbooks
    try:
        workbooks = dynamic.DumbDispatch(workbooks)
    except Exception:
        pass
    try:
        workbook = workbooks.Open(str(path))
        return workbook
    except Exception as exc:
        if "保護ビュー" not in str(exc):
            raise
        normalized_target = str(path).lower()
        protected_windows = excel.ProtectedViewWindows
        try:
            protected_windows = dynamic.DumbDispatch(protected_windows)
        except Exception:
            pass
        for window in protected_windows:
            try:
                protected_path = str(window.Workbook.FullName).lower()
            except Exception:
                continue
            if protected_path == normalized_target:
                return window.Edit()
        raise


def detect_template_period(template_path: Path, sheet_index: int = 1) -> tuple[int | None, int | None, int | None]:
    worksheet = xlrd.open_workbook(str(template_path)).sheet_by_index(sheet_index - 1)
    title_text = ""
    if worksheet.nrows >= 1 and worksheet.ncols >= 1:
        title_text = normalize_cell_text(worksheet.cell_value(0, 0))
    match = re.search(r"R\s*(\d+)\s*年\s*(\d+)\s*月", title_text)
    detected_year = 2018 + int(match.group(1)) if match else None
    detected_month = int(match.group(2)) if match else None

    day_numbers: list[int] = []
    row_index = 3
    max_detectable_day = calendar.monthrange(detected_year, detected_month)[1] if detected_year and detected_month else 31
    for column_index in range(4, 35):
        if row_index < worksheet.nrows and column_index < worksheet.ncols:
            day_number = parse_excel_day_number(worksheet.cell_value(row_index, column_index))
            if day_number is not None and 1 <= day_number <= max_detectable_day:
                day_numbers.append(day_number)
    detected_days = max(day_numbers) if day_numbers else None
    return detected_year, detected_month, detected_days


def normalize_days_in_month(year: int, month: int, requested_days: int | None, source_label: str) -> int:
    actual_days = calendar.monthrange(year, month)[1]
    if requested_days is None:
        return actual_days

    normalized_days = int(requested_days)
    if normalized_days < 1:
        raise ValueError(
            f"{source_label} の日数が不正です: {normalized_days}。{year}年{month}月の日数は 1 以上 {actual_days} 以下で指定してください。"
        )
    if normalized_days > actual_days:
        print(
            f"[warn] {source_label} の日数 {normalized_days} は {year}年{month}月の末日 {actual_days} を超えているため、{actual_days}日に補正しました。",
            file=sys.stderr,
        )
        return actual_days
    return normalized_days


def build_employee_row_map(worksheet) -> dict[str, int]:
    row_map: dict[str, int] = {}
    if worksheet.ncols < 4:
        return row_map
    for row_index in range(worksheet.nrows):
        normalized_name = normalize_employee_name(worksheet.cell_value(row_index, 3))
        if normalized_name and normalized_name != normalize_employee_name("名前") and normalized_name not in row_map:
            row_map[normalized_name] = row_index
    return row_map


def resolve_employee_row_index(worksheet, employee: EmployeeConfig, row_map: dict[str, int]) -> int | None:
    normalized_aliases: list[str] = []
    for alias in (employee.aliases or (employee.display_name,)):
        normalized_alias = normalize_employee_name(alias)
        if normalized_alias and normalized_alias not in normalized_aliases:
            normalized_aliases.append(normalized_alias)

    for normalized_alias in normalized_aliases:
        row_index = row_map.get(normalized_alias)
        if row_index is not None:
            return row_index

    fallback_index = employee.row - 1
    if 0 <= fallback_index < worksheet.nrows:
        fallback_name = normalize_employee_name(worksheet.cell_value(fallback_index, 3))
        if fallback_name in normalized_aliases:
            return fallback_index
    return None


def read_fixed_assignments_from_workbook(
    workbook_path: Path,
    sheet_index: int,
    employees: tuple[EmployeeConfig, ...],
    shift_kinds: dict[str, str],
    days_in_month: int,
) -> dict[str, dict[int, str]]:
    worksheet = xlrd.open_workbook(str(workbook_path)).sheet_by_index(sheet_index - 1)
    results: dict[str, dict[int, str]] = {}
    valid_symbols = set(shift_kinds)
    row_map = build_employee_row_map(worksheet)
    for employee in employees:
        fixed: dict[int, str] = {}
        row_index = resolve_employee_row_index(worksheet, employee, row_map)
        if row_index is None:
            results[employee.employee_id] = fixed
            continue
        for day in range(1, days_in_month + 1):
            column_index = day + 3
            if row_index < worksheet.nrows and column_index < worksheet.ncols:
                symbol = normalize_cell_text(worksheet.cell_value(row_index, column_index))
                if symbol in valid_symbols:
                    fixed[day] = symbol
        results[employee.employee_id] = normalize_night_rest_assignments(fixed, shift_kinds, days_in_month)
    return results


def is_completed_schedule_like_fixed_assignments(
    fixed_assignments: dict[str, dict[int, str]],
    days_in_month: int,
) -> bool:
    non_empty_assignments = [assignments for assignments in fixed_assignments.values() if assignments]
    if not non_empty_assignments or days_in_month <= 0:
        return False

    fully_filled_count = sum(len(assignments) >= days_in_month for assignments in non_empty_assignments)
    total_assignment_count = sum(len(assignments) for assignments in non_empty_assignments)
    employee_count = len(non_empty_assignments)
    coverage_threshold = employee_count * days_in_month * 4
    return fully_filled_count * 5 >= employee_count * 4 and total_assignment_count * 5 >= coverage_threshold


def read_specified_holidays_from_workbook(
    workbook_path: Path,
    sheet_index: int,
    employees: tuple[EmployeeConfig, ...],
    days_in_month: int,
    holiday_symbol: str = "休",
) -> dict[str, tuple[int, ...]]:
    worksheet = xlrd.open_workbook(str(workbook_path)).sheet_by_index(sheet_index - 1)
    results: dict[str, tuple[int, ...]] = {}
    row_map = build_employee_row_map(worksheet)
    normalized_holiday_symbol = normalize_cell_text(holiday_symbol)
    for employee in employees:
        holidays: list[int] = []
        row_index = resolve_employee_row_index(worksheet, employee, row_map)
        if row_index is None:
            results[employee.employee_id] = ()
            continue
        for day in range(1, days_in_month + 1):
            column_index = day + 3
            if row_index < worksheet.nrows and column_index < worksheet.ncols:
                symbol = normalize_cell_text(worksheet.cell_value(row_index, column_index))
                if symbol == normalized_holiday_symbol:
                    holidays.append(day)
        results[employee.employee_id] = tuple(holidays)
    return results


def read_previous_tail_from_workbook(
    workbook_path: Path,
    sheet_index: int,
    employees: tuple[EmployeeConfig, ...],
    shift_kinds: dict[str, str],
    tail_length: int,
) -> dict[str, tuple[str, ...]]:
    worksheet = xlrd.open_workbook(str(workbook_path)).sheet_by_index(sheet_index - 1)
    _, _, detected_days = detect_template_period(workbook_path, sheet_index)
    if detected_days is None or detected_days <= 0:
        return {}

    valid_symbols = set(shift_kinds)
    row_map = build_employee_row_map(worksheet)
    tail_by_employee: dict[str, tuple[str, ...]] = {}
    start_day = max(1, detected_days - tail_length + 1)
    for employee in employees:
        row_index = resolve_employee_row_index(worksheet, employee, row_map)
        if row_index is None:
            continue

        tail: list[str] = []
        for day in range(start_day, detected_days + 1):
            column_index = day + 3
            if row_index < worksheet.nrows and column_index < worksheet.ncols:
                symbol = normalize_cell_text(worksheet.cell_value(row_index, column_index))
                tail.append(symbol if symbol in valid_symbols else "")
        tail_by_employee[employee.employee_id] = tuple(normalize_night_rest_sequence(tail, shift_kinds))
    return tail_by_employee


def resolve_reference_source(target_path: Path, fallback_source: Path) -> Path:
    stem = target_path.stem
    resolved_target_path = target_path.resolve()
    if stem.endswith("_temp"):
        sibling_source = target_path.with_name(stem.removesuffix("_temp") + target_path.suffix)
        if sibling_source.exists():
            return sibling_source.resolve()

    preferred_names: list[str] = [target_path.name]
    if stem.endswith("_temp"):
        non_temp_name = target_path.with_name(stem.removesuffix("_temp") + target_path.suffix).name
        preferred_names.insert(0, non_temp_name)
    fallback_names = [fallback_source.name]

    search_roots: list[Path] = []
    for candidate_root in [target_path.parent, *target_path.parents]:
        resolved_root = candidate_root.resolve()
        if resolved_root not in search_roots:
            search_roots.append(resolved_root)

    for candidate_name in preferred_names:
        for search_root in search_roots:
            direct_candidate = (search_root / candidate_name).resolve()
            if direct_candidate == resolved_target_path:
                continue
            if direct_candidate.exists() and not is_ignored_search_path(direct_candidate):
                return direct_candidate

    for candidate_name in preferred_names:
        for search_root in search_roots:
            try:
                for candidate_path in search_root.rglob(candidate_name):
                    if is_ignored_search_path(candidate_path):
                        continue
                    resolved_candidate_path = candidate_path.resolve()
                    if resolved_candidate_path == resolved_target_path:
                        continue
                    return resolved_candidate_path
            except OSError:
                continue

    for candidate_name in fallback_names:
        for search_root in search_roots:
            direct_candidate = (search_root / candidate_name).resolve()
            if direct_candidate == resolved_target_path:
                continue
            if direct_candidate.exists() and not is_ignored_search_path(direct_candidate):
                return direct_candidate

    for candidate_name in fallback_names:
        for search_root in search_roots:
            try:
                for candidate_path in search_root.rglob(candidate_name):
                    if is_ignored_search_path(candidate_path):
                        continue
                    resolved_candidate_path = candidate_path.resolve()
                    if resolved_candidate_path == resolved_target_path:
                        continue
                    return resolved_candidate_path
            except OSError:
                continue

    if fallback_source.exists():
        return fallback_source.resolve()
    return fallback_source.resolve()


def to_fullwidth_digits(value: int) -> str:
    return str(value).translate(str.maketrans("0123456789", "０１２３４５６７８９"))


def normalize_workbook_name_hint(name: str) -> str:
    hint = Path(name).stem
    hint = re.sub(r"[0-9０-９]+月", "", hint, count=1)
    hint = re.sub(r"（.*?）|\(.*?\)", "", hint)
    hint = hint.replace("変更バージョン", "")
    hint = hint.replace("_temp", "")
    return re.sub(r"[\s　]+", "", hint)


def build_previous_month_candidate_names(original_name: str, previous_month: int, suffix: str) -> list[str]:
    month_match = re.search(r"[0-9０-９]+月", original_name)
    if month_match is None:
        return []

    month_tokens = [f"{previous_month}月", f"{to_fullwidth_digits(previous_month)}月"]
    candidate_names: list[str] = []
    for month_token in month_tokens:
        replaced_name = f"{original_name[:month_match.start()]}{month_token}{original_name[month_match.end():]}"
        if replaced_name not in candidate_names:
            candidate_names.append(replaced_name)
        if replaced_name.endswith(f"_temp{suffix}"):
            non_temp_name = replaced_name.replace(f"_temp{suffix}", suffix)
            if non_temp_name not in candidate_names:
                candidate_names.append(non_temp_name)
    return candidate_names


def is_ignored_search_path(candidate_path: Path) -> bool:
    normalized_parts = {part.lower() for part in candidate_path.parts}
    return any(part in normalized_parts for part in ("old", "exe", "dist", "build"))


def resolve_previous_month_source(base_dir: Path, target_path: Path, reference_source: Path, year: int, month: int) -> Path | None:
    previous_year = year - 1 if month == 1 else year
    previous_month = 12 if month == 1 else month - 1

    candidate_names: list[str] = []
    for original_name in (target_path.name, reference_source.name):
        for candidate_name in build_previous_month_candidate_names(original_name, previous_month, target_path.suffix):
            if candidate_name not in candidate_names:
                candidate_names.append(candidate_name)

    seen_paths: set[Path] = set()
    for directory in (target_path.parent, reference_source.parent, base_dir):
        for candidate_name in candidate_names:
            candidate_path = (directory / candidate_name).resolve()
            if candidate_path in seen_paths:
                continue
            seen_paths.add(candidate_path)
            if candidate_path.exists():
                detected_year, detected_month, _ = detect_template_period(candidate_path)
                if detected_year == previous_year and detected_month == previous_month:
                    return candidate_path

    search_roots = []
    for directory in (target_path.parent, reference_source.parent, base_dir):
        resolved_directory = directory.resolve()
        if resolved_directory not in search_roots:
            search_roots.append(resolved_directory)

    for search_root in search_roots:
        try:
            for candidate_name in candidate_names:
                for candidate_path in search_root.rglob(candidate_name):
                    if is_ignored_search_path(candidate_path):
                        continue
                    resolved_path = candidate_path.resolve()
                    if resolved_path in seen_paths:
                        continue
                    seen_paths.add(resolved_path)
                    detected_year, detected_month, _ = detect_template_period(resolved_path)
                    if detected_year == previous_year and detected_month == previous_month:
                        return resolved_path
        except OSError:
            continue

    recursive_match_paths: set[Path] = set()
    for month_token in (f"{previous_month}月", f"{to_fullwidth_digits(previous_month)}月"):
        recursive_match_paths.update(base_dir.rglob(f"*{month_token}*.xls"))

    target_name_hint = normalize_workbook_name_hint(target_path.name)
    preferred_matches: list[Path] = []
    fallback_matches: list[Path] = []
    for candidate_path in sorted(recursive_match_paths):
        if is_ignored_search_path(candidate_path):
            continue
        if candidate_path.name.endswith("_temp.xls"):
            continue
        try:
            detected_year, detected_month, _ = detect_template_period(candidate_path)
        except Exception:
            continue
        if detected_year == previous_year and detected_month == previous_month:
            resolved_path = candidate_path.resolve()
            if normalize_workbook_name_hint(candidate_path.name) == target_name_hint:
                preferred_matches.append(resolved_path)
            else:
                fallback_matches.append(resolved_path)
    if preferred_matches:
        return preferred_matches[0]
    if fallback_matches:
        return fallback_matches[0]
    return None


def read_title_from_workbook(workbook_path: Path, sheet_index: int) -> str | None:
    try:
        worksheet = xlrd.open_workbook(str(workbook_path)).sheet_by_index(sheet_index - 1)
    except Exception:
        return None
    if worksheet.nrows < 1 or worksheet.ncols < 1:
        return None
    value = worksheet.cell_value(0, 0)
    return None if value is None else str(value)


def compare_workbooks_xlrd(source_path: Path, target_path: Path, sheet_index: int = 1) -> list[tuple[int, int, str, str]]:
    source_sheet = xlrd.open_workbook(str(source_path)).sheet_by_index(sheet_index - 1)
    target_sheet = xlrd.open_workbook(str(target_path)).sheet_by_index(sheet_index - 1)
    diffs: list[tuple[int, int, str, str]] = []
    max_row = max(source_sheet.nrows, target_sheet.nrows)
    max_col = max(source_sheet.ncols, target_sheet.ncols)
    for row in range(max_row):
        for col in range(max_col):
            source_value = normalize_cell_text(source_sheet.cell_value(row, col)) if row < source_sheet.nrows and col < source_sheet.ncols else ""
            target_value = normalize_cell_text(target_sheet.cell_value(row, col)) if row < target_sheet.nrows and col < target_sheet.ncols else ""
            if source_value != target_value:
                diffs.append((row + 1, col + 1, source_value, target_value))
    return diffs


def collect_assignment_diff_rows_xlrd(
    source_path: Path,
    target_path: Path,
    sheet_index: int,
    employee_rows: list[int],
    max_days: int = 31,
) -> list[dict[str, object]]:
    source_sheet = xlrd.open_workbook(str(source_path)).sheet_by_index(sheet_index - 1)
    target_sheet = xlrd.open_workbook(str(target_path)).sheet_by_index(sheet_index - 1)
    results: list[dict[str, object]] = []
    for row in employee_rows:
        row_index = row - 1
        source_name = normalize_cell_text(source_sheet.cell_value(row_index, 3)) if row_index < source_sheet.nrows and 3 < source_sheet.ncols else ""
        target_name = normalize_cell_text(target_sheet.cell_value(row_index, 3)) if row_index < target_sheet.nrows and 3 < target_sheet.ncols else ""
        diffs: list[dict[str, object]] = []
        for day in range(1, max_days + 1):
            column_index = day + 3
            source_value = normalize_cell_text(source_sheet.cell_value(row_index, column_index)) if row_index < source_sheet.nrows and column_index < source_sheet.ncols else ""
            target_value = normalize_cell_text(target_sheet.cell_value(row_index, column_index)) if row_index < target_sheet.nrows and column_index < target_sheet.ncols else ""
            if source_value != target_value:
                diffs.append({"day": day, "manual": source_value, "generated": target_value})
        results.append(
            {
                "row": row,
                "manual_name": source_name,
                "generated_name": target_name,
                "diff_count": len(diffs),
                "diffs": diffs,
            }
        )
    return results


def compare_worksheets(source_sheet, target_sheet) -> list[tuple[int, int, str, str]]:
    diffs: list[tuple[int, int, str, str]] = []
    max_row = max(source_sheet.UsedRange.Rows.Count, target_sheet.UsedRange.Rows.Count)
    max_col = max(source_sheet.UsedRange.Columns.Count, target_sheet.UsedRange.Columns.Count)
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            source_value = source_sheet.Cells(row, col).Text or ""
            target_value = target_sheet.Cells(row, col).Text or ""
            if str(source_value) != str(target_value):
                diffs.append((row, col, str(source_value), str(target_value)))
    return diffs


def merge_config_values(base: object, override: object) -> object:
    if isinstance(base, dict) and isinstance(override, dict):
        merged = dict(base)
        for key, value in override.items():
            merged[key] = merge_config_values(merged.get(key), value)
        return merged
    return override


def load_raw_config(config_path: Path, visited: set[Path] | None = None) -> dict[str, object]:
    resolved_path = config_path.resolve()
    visited = set() if visited is None else set(visited)
    if resolved_path in visited:
        chain = " -> ".join(str(path) for path in [*visited, resolved_path])
        raise ValueError(f"設定ファイルの base_config が循環しています: {chain}")
    if not resolved_path.exists():
        return {}

    visited.add(resolved_path)
    loaded = json.loads(resolved_path.read_text(encoding="utf-8"))
    if not isinstance(loaded, dict):
        raise ValueError(f"設定ファイルは JSON オブジェクトである必要があります: {resolved_path}")

    merged: dict[str, object] = {}
    base_config = loaded.get("base_config")
    if base_config:
        merged = load_raw_config(resolve_path(resolved_path.parent, str(base_config)), visited)

    return merge_config_values(merged, {key: value for key, value in loaded.items() if key != "base_config"})


def apply_period_overrides(raw: dict[str, object], year: int, month: int) -> dict[str, object]:
    period_overrides = raw.get("period_overrides", {})
    if not isinstance(period_overrides, dict):
        raise ValueError("period_overrides は JSON オブジェクトで指定してください。")

    merged = {key: value for key, value in raw.items() if key != "period_overrides"}
    for key in (f"{year:04d}-{month:02d}", f"{year}-{month}", f"{month:02d}", str(month)):
        override = period_overrides.get(key)
        if override is None:
            continue
        if not isinstance(override, dict):
            raise ValueError(f"period_overrides[{key}] は JSON オブジェクトで指定してください。")
        merged = merge_config_values(merged, override)
    return merged


def load_config(config_path: Path, year: int | None = None, month: int | None = None) -> SchedulerConfig:
    base_dir = config_path.resolve().parent
    raw = merge_config_values(default_config_dict(), load_raw_config(config_path))

    effective_year = int(raw["year"] if year is None else year)
    effective_month = int(raw["month"] if month is None else month)
    raw = apply_period_overrides(raw, effective_year, effective_month)
    raw["year"] = effective_year
    raw["month"] = effective_month

    year = int(raw["year"])
    month = int(raw["month"])
    days_in_month = normalize_days_in_month(year, month, raw.get("days_in_month"), "設定ファイル")
    shift_kinds = {str(key): str(value) for key, value in raw.get("shift_kinds", DEFAULT_SHIFT_KINDS).items()}
    if "夜休" not in shift_kinds:
        shift_kinds["夜休"] = "night_rest"
    count_symbols = {str(key): str(value) for key, value in raw.get("count_symbols", DEFAULT_COUNT_SYMBOLS).items()}
    rules = raw["rules"]
    night_symbols = set(symbol_names_by_kind(shift_kinds, "night"))
    night_rest_symbol = first_symbol_by_kind(shift_kinds, "night_rest")

    employees: list[EmployeeConfig] = []
    for employee_raw in raw["employees"]:
        display_name = normalize_cell_text(employee_raw.get("display_name")) or normalize_cell_text(employee_raw.get("name"))
        if not display_name:
            raise ValueError("employees[].display_name または employees[].name を指定してください。")
        employee_id = normalize_cell_text(employee_raw.get("employee_id")) or fallback_employee_id(display_name, len(employees))
        aliases = normalize_employee_aliases(employee_raw, display_name, employee_id)
        allowed_shifts = [str(shift) for shift in employee_raw["allowed_shifts"]]
        if night_rest_symbol and any(shift in night_symbols for shift in allowed_shifts) and night_rest_symbol not in allowed_shifts:
            allowed_shifts.append(night_rest_symbol)
        min_counts = {str(key): int(value) for key, value in employee_raw.get("min_counts", {}).items()}
        if str(employee_raw.get("employment", "")) == "full" and "日" in allowed_shifts and int(min_counts.get("日", 0)) < 1:
            min_counts["日"] = 1
        allow_night_rest = night_rest_symbol in allowed_shifts
        weekday_allowed_shifts = {
            parse_weekday_key(weekday): normalize_allowed_shift_rule(
                shifts,
                shift_kinds,
                display_name,
                f"weekday_allowed_shifts[{weekday}]",
                allow_night_rest,
            )
            for weekday, shifts in employee_raw.get("weekday_allowed_shifts", {}).items()
        }
        date_allowed_shift_overrides = {
            int(day): normalize_allowed_shift_rule(
                shifts,
                shift_kinds,
                display_name,
                f"date_allowed_shift_overrides[{day}]",
                allow_night_rest,
            )
            for day, shifts in employee_raw.get("date_allowed_shift_overrides", {}).items()
        }
        previous_tail = tuple(normalize_night_rest_sequence(employee_raw.get("previous_tail", []), shift_kinds))
        fixed_assignments = normalize_night_rest_assignments(
            {int(day): str(symbol) for day, symbol in employee_raw.get("fixed_assignments", {}).items()},
            shift_kinds,
            days_in_month,
            previous_shift=(previous_tail[-1] if previous_tail else None),
        )
        exact_rest_days, min_rest_days, max_rest_days = normalize_employee_rest_day_targets(
            employee_raw,
            display_name,
            days_in_month,
        )
        employees.append(
            EmployeeConfig(
                employee_id=employee_id,
                display_name=display_name,
                unit=str(employee_raw["unit"]),
                employment=str(employee_raw["employment"]),
                row=int(employee_raw["row"]),
                allowed_shifts=tuple(allowed_shifts),
                aliases=aliases,
                weekday_allowed_shifts=weekday_allowed_shifts,
                date_allowed_shift_overrides=date_allowed_shift_overrides,
                require_weekend_pair_rest=bool(employee_raw.get("require_weekend_pair_rest", False)),
                night_fairness_target=bool(employee_raw.get("night_fairness_target", False)),
                required_double_night_min_count=(
                    None
                    if employee_raw.get("required_double_night_min_count") is None
                    else int(employee_raw["required_double_night_min_count"])
                ),
                weekend_fairness_target=bool(employee_raw.get("weekend_fairness_target", False)),
                min_counts=min_counts,
                max_counts={str(key): int(value) for key, value in employee_raw.get("max_counts", {}).items()},
                max_consecutive_work_limit=(
                    None
                    if employee_raw.get("max_consecutive_work_limit") is None
                    else int(employee_raw["max_consecutive_work_limit"])
                ),
                max_four_day_streak_count=(
                    None
                    if employee_raw.get("max_four_day_streak_count") is None
                    else int(employee_raw["max_four_day_streak_count"])
                ),
                exact_rest_days=exact_rest_days,
                min_rest_days=min_rest_days,
                max_rest_days=max_rest_days,
                specified_holidays=tuple(int(day) for day in employee_raw.get("specified_holidays", [])),
                fixed_assignments=fixed_assignments,
                previous_tail=previous_tail,
            )
        )

    validate_employee_identity_constraints(employees)

    day_requirements = {
        int(day): {
            section: {str(symbol): int(count) for symbol, count in values.items()}
            for section, values in requirement.items()
        }
        for day, requirement in rules.get("day_requirements", {}).items()
    }

    return SchedulerConfig(
        config_path=config_path.resolve(),
        target_path=resolve_path(base_dir, raw["target_path"]),
        manual_source=resolve_path(base_dir, raw["manual_source"]),
        sheet_index=int(raw.get("sheet_index", 1)),
        year=year,
        month=month,
        days_in_month=days_in_month,
        unit_name=str(raw["unit_name"]),
        shift_kinds=shift_kinds,
        count_symbols=count_symbols,
        employees=tuple(employees),
        required_per_day={
            str(key): {str(inner_key): int(inner_value) for inner_key, inner_value in value.items()}
            for key, value in rules["required_per_day"].items()
            if key in {"南", "北"}
        },
        night_total_per_day=int(rules["required_per_day"].get("night_total", 1)),
        day_requirements=day_requirements,
        max_consecutive_work=int(rules.get("max_consecutive_work", 5)),
        max_consecutive_night=int(rules.get("max_consecutive_night", 2)),
        max_consecutive_rest=int(rules.get("max_consecutive_rest", 3)),
        max_consecutive_rest_with_special=int(rules.get("max_consecutive_rest_with_special", 5)),
        fairness_night_spread=(None if rules.get("fairness_night_spread") is None else int(rules["fairness_night_spread"])),
        fairness_weekend_spread=(None if rules.get("fairness_weekend_spread") is None else int(rules["fairness_weekend_spread"])),
        weekend_rest_count_mode=normalize_weekend_rest_count_mode(rules.get("weekend_rest_count_mode", "rest_special_night_rest")),
        require_weekend_pair_rest=bool(rules.get("require_weekend_pair_rest", True)),
        prefer_double_night=bool(rules.get("prefer_double_night", True)),
    )


def add_window_constraint(model: cp_model.CpModel, flags: list[object], window_size: int, max_allowed: int) -> None:
    if window_size <= 0 or len(flags) < window_size:
        return
    for start in range(len(flags) - window_size + 1):
        model.Add(sum(flags[start : start + window_size]) <= max_allowed)


def build_schedule_model(config: SchedulerConfig) -> tuple[cp_model.CpModel, dict[tuple[str, int, str], cp_model.IntVar]]:
    model = cp_model.CpModel()
    decision_vars: dict[tuple[str, int, str], cp_model.IntVar] = {}
    shift_order = list(config.shift_kinds.keys())
    early_symbols = symbol_names_by_kind(config.shift_kinds, "early")
    late_symbols = symbol_names_by_kind(config.shift_kinds, "late")
    night_symbols = symbol_names_by_kind(config.shift_kinds, "night")
    night_rest_symbols = symbol_names_by_kind(config.shift_kinds, "night_rest")
    rest_symbols = symbol_names_by_kind(config.shift_kinds, "rest")
    rest_like_symbols = rest_symbols + night_rest_symbols
    work_symbols = [symbol for symbol in shift_order if symbol not in rest_like_symbols]
    regular_rest_limit_symbols = (["休"] if "休" in shift_order else []) + night_rest_symbols
    weekend_pairs = weekend_pair_day_indexes(config.year, config.month, config.days_in_month)

    for employee in config.employees:
        employee_id = employee.employee_id
        for day in range(config.days_in_month):
            effective_allowed_shifts = effective_allowed_shifts_for_day(config, employee, day + 1)
            for shift in shift_order:
                decision_vars[employee_id, day, shift] = model.NewBoolVar(f"shift_{employee_id}_{day}_{shift}")
                if shift not in employee.allowed_shifts:
                    model.Add(decision_vars[employee_id, day, shift] == 0)
                if effective_allowed_shifts is not None and shift not in effective_allowed_shifts:
                    model.Add(decision_vars[employee_id, day, shift] == 0)
            model.Add(sum(decision_vars[employee_id, day, shift] for shift in shift_order) == 1)

    for day in range(config.days_in_month):
        for unit, requirements in config.required_per_day.items():
            unit_employee_ids = [employee.employee_id for employee in config.employees if employee.unit == unit]
            model.Add(sum(decision_vars[employee_id, day, shift] for employee_id in unit_employee_ids for shift in early_symbols) == requirements.get("early", 0))
            model.Add(sum(decision_vars[employee_id, day, shift] for employee_id in unit_employee_ids for shift in late_symbols) == requirements.get("late", 0))
        model.Add(sum(decision_vars[employee.employee_id, day, shift] for employee in config.employees for shift in night_symbols) == config.night_total_per_day)

    for day, requirement in config.day_requirements.items():
        day_index = day - 1
        if not 0 <= day_index < config.days_in_month:
            continue
        for shift, count in requirement.get("min", {}).items():
            model.Add(sum(decision_vars[employee.employee_id, day_index, shift] for employee in config.employees) >= count)
        for shift, count in requirement.get("max", {}).items():
            model.Add(sum(decision_vars[employee.employee_id, day_index, shift] for employee in config.employees) <= count)

    pair_vars: list[cp_model.IntVar] = []
    pair_vars_by_employee: dict[str, list[cp_model.IntVar]] = {}
    night_eligible_ids = selected_night_fairness_employee_ids(config, night_symbols)
    weekend_rest_symbols = weekend_rest_symbols_for_mode(config.shift_kinds, config.weekend_rest_count_mode)
    for employee in config.employees:
        employee_id = employee.employee_id
        pair_vars_by_employee[employee_id] = []

        if config.require_weekend_pair_rest and employee.require_weekend_pair_rest:
            for saturday_day, sunday_day in weekend_pairs:
                model.Add(
                    sum(decision_vars[employee_id, saturday_day, shift] for shift in rest_like_symbols)
                    + sum(decision_vars[employee_id, sunday_day, shift] for shift in rest_like_symbols)
                    >= 1
                )

        for day, shift in employee.fixed_assignments.items():
            if 1 <= day <= config.days_in_month:
                model.Add(decision_vars[employee_id, day - 1, shift] == 1)

        for holiday_day in employee.specified_holidays:
            if 2 <= holiday_day <= config.days_in_month:
                model.Add(sum(decision_vars[employee_id, holiday_day - 2, shift] for shift in night_symbols) == 0)

        tail = list(employee.previous_tail)
        if tail:
            last_shift = tail[-1]
            if last_shift in late_symbols:
                model.Add(sum(decision_vars[employee_id, 0, shift] for shift in early_symbols) == 0)
            if last_shift in night_symbols:
                model.Add(sum(decision_vars[employee_id, 0, shift] for shift in night_symbols + night_rest_symbols) == 1)

        if night_rest_symbols:
            previous_night = 1 if tail and tail[-1] in night_symbols else 0
            model.Add(sum(decision_vars[employee_id, 0, shift] for shift in night_rest_symbols) <= previous_night)
            for day in range(1, config.days_in_month):
                model.Add(
                    sum(decision_vars[employee_id, day, shift] for shift in night_rest_symbols)
                    <= sum(decision_vars[employee_id, day - 1, shift] for shift in night_symbols)
                )

        for day in range(config.days_in_month - 1):
            model.Add(
                sum(decision_vars[employee_id, day, shift] for shift in late_symbols)
                + sum(decision_vars[employee_id, day + 1, shift] for shift in early_symbols)
                <= 1
            )
            model.Add(
                sum(decision_vars[employee_id, day, shift] for shift in night_symbols)
                <= sum(decision_vars[employee_id, day + 1, shift] for shift in night_symbols + night_rest_symbols)
            )

        if config.prefer_double_night or employee.required_double_night_min_count is not None:
            for day in range(config.days_in_month - 1):
                pair_var = model.NewBoolVar(f"night_pair_{employee_id}_{day}")
                night_today = sum(decision_vars[employee_id, day, shift] for shift in night_symbols)
                night_tomorrow = sum(decision_vars[employee_id, day + 1, shift] for shift in night_symbols)
                model.Add(pair_var <= night_today)
                model.Add(pair_var <= night_tomorrow)
                model.Add(pair_var >= night_today + night_tomorrow - 1)
                pair_vars.append(pair_var)
                pair_vars_by_employee[employee_id].append(pair_var)

        if employee.required_double_night_min_count is not None:
            model.Add(sum(pair_vars_by_employee[employee_id]) >= employee.required_double_night_min_count)

        max_consecutive_work_limit = employee_max_consecutive_work(employee, config)
        work_flags = [1 if shift in work_symbols else 0 for shift in tail[-max_consecutive_work_limit :]]
        work_flags.extend(sum(decision_vars[employee_id, day, shift] for shift in work_symbols) for day in range(config.days_in_month))
        add_window_constraint(model, work_flags, max_consecutive_work_limit + 1, max_consecutive_work_limit)

        if employee.max_four_day_streak_count is not None and config.days_in_month >= 4:
            four_day_vars: list[cp_model.IntVar] = []
            for start_day in range(config.days_in_month - 3):
                window_var = model.NewBoolVar(f"four_day_window_{employee_id}_{start_day}")
                window_flags = [sum(decision_vars[employee_id, day, shift] for shift in work_symbols) for day in range(start_day, start_day + 4)]
                for flag in window_flags:
                    model.Add(window_var <= flag)
                model.Add(window_var >= sum(window_flags) - 3)
                four_day_vars.append(window_var)
            model.Add(sum(four_day_vars) <= employee.max_four_day_streak_count)

        regular_rest_flags = [1 if shift in regular_rest_limit_symbols else 0 for shift in tail[-config.max_consecutive_rest :]]
        regular_rest_flags.extend(
            sum(decision_vars[employee_id, day, shift] for shift in regular_rest_limit_symbols)
            for day in range(config.days_in_month)
        )
        add_window_constraint(model, regular_rest_flags, config.max_consecutive_rest + 1, config.max_consecutive_rest)

        all_rest_flags = [1 if shift in rest_like_symbols else 0 for shift in tail[-config.max_consecutive_rest_with_special :]]
        all_rest_flags.extend(sum(decision_vars[employee_id, day, shift] for shift in rest_like_symbols) for day in range(config.days_in_month))
        add_window_constraint(model, all_rest_flags, config.max_consecutive_rest_with_special + 1, config.max_consecutive_rest_with_special)

        night_flags = [1 if shift in night_symbols else 0 for shift in tail[-config.max_consecutive_night :]]
        night_flags.extend(sum(decision_vars[employee_id, day, shift] for shift in night_symbols) for day in range(config.days_in_month))
        add_window_constraint(model, night_flags, config.max_consecutive_night + 1, config.max_consecutive_night)

        for shift, minimum in employee.min_counts.items():
            model.Add(sum(decision_vars[employee_id, day, shift] for day in range(config.days_in_month)) >= minimum)
        for shift, maximum in employee.max_counts.items():
            model.Add(sum(decision_vars[employee_id, day, shift] for day in range(config.days_in_month)) <= maximum)

        rest_count = sum(decision_vars[employee_id, day, shift] for day in range(config.days_in_month) for shift in rest_like_symbols)
        if employee.exact_rest_days is not None:
            model.Add(rest_count == employee.exact_rest_days)
        else:
            if employee.min_rest_days is not None:
                model.Add(rest_count >= employee.min_rest_days)
            if employee.max_rest_days is not None:
                model.Add(rest_count <= employee.max_rest_days)

    objective_terms: list[object] = []
    weekend_days = weekend_day_indexes(config.year, config.month, config.days_in_month)
    if weekend_days and config.fairness_weekend_spread is not None:
        weekend_eligible = selected_weekend_fairness_employee_ids(config)
        if weekend_eligible:
            weekend_counts = {
                employee_id: sum(decision_vars[employee_id, day, shift] for day in weekend_days for shift in weekend_rest_symbols)
                for employee_id in weekend_eligible
            }
            max_weekend = model.NewIntVar(0, len(weekend_days), "max_weekend_rest")
            min_weekend = model.NewIntVar(0, len(weekend_days), "min_weekend_rest")
            for employee_id in weekend_eligible:
                model.Add(weekend_counts[employee_id] <= max_weekend)
                model.Add(weekend_counts[employee_id] >= min_weekend)
            model.Add(max_weekend - min_weekend <= config.fairness_weekend_spread)
            objective_terms.append((max_weekend - min_weekend) * 100)

    if config.fairness_night_spread is not None and night_eligible_ids:
        night_counts = {
            employee_id: sum(decision_vars[employee_id, day, shift] for day in range(config.days_in_month) for shift in night_symbols)
            for employee_id in night_eligible_ids
        }
        max_night = model.NewIntVar(0, config.days_in_month, "max_night")
        min_night = model.NewIntVar(0, config.days_in_month, "min_night")
        for employee_id in night_eligible_ids:
            model.Add(night_counts[employee_id] <= max_night)
            model.Add(night_counts[employee_id] >= min_night)
        model.Add(max_night - min_night <= config.fairness_night_spread)
        objective_terms.append((max_night - min_night) * 100)

    for unit in {employee.unit for employee in config.employees}:
        unit_employee_ids = [employee.employee_id for employee in config.employees if employee.unit == unit and employee.employment == "full"]
        if not unit_employee_ids:
            continue
        for label, symbols in (("early", early_symbols), ("late", late_symbols)):
            max_count = model.NewIntVar(0, config.days_in_month, f"max_{unit}_{label}")
            min_count = model.NewIntVar(0, config.days_in_month, f"min_{unit}_{label}")
            for employee_id in unit_employee_ids:
                count = sum(decision_vars[employee_id, day, shift] for day in range(config.days_in_month) for shift in symbols)
                model.Add(count <= max_count)
                model.Add(count >= min_count)
            objective_terms.append(max_count - min_count)

    if config.prefer_double_night and pair_vars:
        objective_terms.append(-sum(pair_vars))

    if objective_terms:
        model.Minimize(sum(objective_terms))

    return model, decision_vars


def build_schedule_from_solver(
    config: SchedulerConfig,
    decision_vars: dict[tuple[str, int, str], cp_model.IntVar],
    solver: cp_model.CpSolver,
) -> dict[str, list[str]]:
    shift_order = list(config.shift_kinds.keys())
    schedule: dict[str, list[str]] = {}
    for employee in config.employees:
        employee_id = employee.employee_id
        shifts: list[str] = []
        for day in range(config.days_in_month):
            selected_shift = next(
                shift for shift in shift_order if solver.Value(decision_vars[employee_id, day, shift]) == 1
            )
            shifts.append(selected_shift)
        schedule[employee_id] = shifts
    return schedule


def build_relaxed_schedule_model(config: SchedulerConfig) -> tuple[cp_model.CpModel, dict[tuple[str, int, str], cp_model.IntVar]]:
    model = cp_model.CpModel()
    decision_vars: dict[tuple[str, int, str], cp_model.IntVar] = {}
    shift_order = list(config.shift_kinds.keys())
    early_symbols = symbol_names_by_kind(config.shift_kinds, "early")
    late_symbols = symbol_names_by_kind(config.shift_kinds, "late")
    night_symbols = symbol_names_by_kind(config.shift_kinds, "night")
    night_rest_symbols = symbol_names_by_kind(config.shift_kinds, "night_rest")
    rest_symbols = symbol_names_by_kind(config.shift_kinds, "rest")
    rest_like_symbols = rest_symbols + night_rest_symbols
    work_symbols = [symbol for symbol in shift_order if symbol not in rest_like_symbols]
    weekend_pairs = weekend_pair_day_indexes(config.year, config.month, config.days_in_month)
    objective_terms: list[object] = []

    for employee in config.employees:
        employee_id = employee.employee_id
        for day in range(config.days_in_month):
            effective_allowed_shifts = effective_allowed_shifts_for_day(config, employee, day + 1)
            for shift in shift_order:
                decision_vars[employee_id, day, shift] = model.NewBoolVar(f"relaxed_shift_{employee_id}_{day}_{shift}")
                if shift not in employee.allowed_shifts:
                    model.Add(decision_vars[employee_id, day, shift] == 0)
                if effective_allowed_shifts is not None and shift not in effective_allowed_shifts:
                    model.Add(decision_vars[employee_id, day, shift] == 0)
            model.Add(sum(decision_vars[employee_id, day, shift] for shift in shift_order) == 1)

    for employee in config.employees:
        employee_id = employee.employee_id
        for day, shift in employee.fixed_assignments.items():
            if 1 <= day <= config.days_in_month:
                model.Add(decision_vars[employee_id, day - 1, shift] == 1)

        for holiday_day in employee.specified_holidays:
            if 2 <= holiday_day <= config.days_in_month:
                previous_night = sum(decision_vars[employee_id, holiday_day - 2, shift] for shift in night_symbols)
                objective_terms.append(previous_night * 900)

        tail = list(employee.previous_tail)
        max_consecutive_work_limit = employee_max_consecutive_work(employee, config)
        if tail:
            last_shift = tail[-1]
            if last_shift in late_symbols:
                objective_terms.append(sum(decision_vars[employee_id, 0, shift] for shift in early_symbols) * 600)
            if last_shift in night_symbols:
                objective_terms.append(
                    sum(decision_vars[employee_id, 0, shift] for shift in shift_order if shift not in night_symbols + night_rest_symbols) * 900
                )

        for day in range(config.days_in_month - 1):
            late_today = sum(decision_vars[employee_id, day, shift] for shift in late_symbols)
            early_tomorrow = sum(decision_vars[employee_id, day + 1, shift] for shift in early_symbols)
            late_early_violation = model.NewBoolVar(f"relaxed_late_early_{employee_id}_{day}")
            model.Add(late_early_violation >= late_today + early_tomorrow - 1)
            model.Add(late_early_violation <= late_today)
            model.Add(late_early_violation <= early_tomorrow)
            objective_terms.append(late_early_violation * 700)

            night_today = sum(decision_vars[employee_id, day, shift] for shift in night_symbols)
            invalid_follow = sum(
                decision_vars[employee_id, day + 1, shift]
                for shift in shift_order
                if shift not in night_symbols + night_rest_symbols
            )
            night_follow_violation = model.NewBoolVar(f"relaxed_night_follow_{employee_id}_{day}")
            model.Add(night_follow_violation >= night_today + invalid_follow - 1)
            model.Add(night_follow_violation <= night_today)
            model.Add(night_follow_violation <= invalid_follow)
            objective_terms.append(night_follow_violation * 900)

            if night_rest_symbols:
                invalid_night_rest = sum(decision_vars[employee_id, day + 1, shift] for shift in night_rest_symbols)
                previous_not_night = sum(
                    decision_vars[employee_id, day, shift] for shift in shift_order if shift not in night_symbols
                )
                night_rest_violation = model.NewBoolVar(f"relaxed_night_rest_{employee_id}_{day}")
                model.Add(night_rest_violation >= invalid_night_rest + previous_not_night - 1)
                model.Add(night_rest_violation <= invalid_night_rest)
                model.Add(night_rest_violation <= previous_not_night)
                objective_terms.append(night_rest_violation * 400)

        if config.require_weekend_pair_rest and employee.require_weekend_pair_rest:
            for saturday_day, sunday_day in weekend_pairs:
                saturday_work = sum(decision_vars[employee_id, saturday_day, shift] for shift in work_symbols)
                sunday_work = sum(decision_vars[employee_id, sunday_day, shift] for shift in work_symbols)
                weekend_violation = model.NewBoolVar(f"relaxed_weekend_pair_{employee_id}_{saturday_day}")
                model.Add(weekend_violation >= saturday_work + sunday_work - 1)
                model.Add(weekend_violation <= saturday_work)
                model.Add(weekend_violation <= sunday_work)
                objective_terms.append(weekend_violation * 650)

        for shift, minimum in employee.min_counts.items():
            count = sum(decision_vars[employee_id, day, shift] for day in range(config.days_in_month))
            deficit = model.NewIntVar(0, minimum, f"relaxed_min_{employee_id}_{shift}")
            model.Add(count + deficit >= minimum)
            objective_terms.append(deficit * 120)
        for shift, maximum in employee.max_counts.items():
            count = sum(decision_vars[employee_id, day, shift] for day in range(config.days_in_month))
            excess = model.NewIntVar(0, config.days_in_month, f"relaxed_max_{employee_id}_{shift}")
            model.Add(count - excess <= maximum)
            objective_terms.append(excess * 120)

        if employee.required_double_night_min_count is not None:
            pair_vars_for_employee: list[cp_model.IntVar] = []
            for day in range(config.days_in_month - 1):
                pair_var = model.NewBoolVar(f"relaxed_night_pair_{employee_id}_{day}")
                night_today = sum(decision_vars[employee_id, day, shift] for shift in night_symbols)
                night_tomorrow = sum(decision_vars[employee_id, day + 1, shift] for shift in night_symbols)
                model.Add(pair_var <= night_today)
                model.Add(pair_var <= night_tomorrow)
                model.Add(pair_var >= night_today + night_tomorrow - 1)
                pair_vars_for_employee.append(pair_var)
            double_night_deficit = model.NewIntVar(0, employee.required_double_night_min_count, f"relaxed_double_night_deficit_{employee_id}")
            model.Add(sum(pair_vars_for_employee) + double_night_deficit >= employee.required_double_night_min_count)
            objective_terms.append(double_night_deficit * 180)

        work_flags = [sum(decision_vars[employee_id, day, shift] for shift in work_symbols) for day in range(config.days_in_month)]
        if max_consecutive_work_limit < config.days_in_month:
            for start_day in range(config.days_in_month - max_consecutive_work_limit):
                window_sum = sum(work_flags[start_day : start_day + max_consecutive_work_limit + 1])
                excess = model.NewIntVar(0, max_consecutive_work_limit + 1, f"relaxed_work_window_excess_{employee_id}_{start_day}")
                model.Add(window_sum - excess <= max_consecutive_work_limit)
                objective_terms.append(excess * 200)

        if employee.max_four_day_streak_count is not None and config.days_in_month >= 4:
            four_day_window_vars: list[cp_model.IntVar] = []
            for start_day in range(config.days_in_month - 3):
                window_var = model.NewBoolVar(f"relaxed_four_day_window_{employee_id}_{start_day}")
                window_flags = work_flags[start_day : start_day + 4]
                for flag in window_flags:
                    model.Add(window_var <= flag)
                model.Add(window_var >= sum(window_flags) - 3)
                four_day_window_vars.append(window_var)
            four_day_excess = model.NewIntVar(0, config.days_in_month, f"relaxed_four_day_excess_{employee_id}")
            model.Add(sum(four_day_window_vars) - four_day_excess <= employee.max_four_day_streak_count)
            objective_terms.append(four_day_excess * 120)

        rest_count = sum(decision_vars[employee_id, day, shift] for day in range(config.days_in_month) for shift in rest_like_symbols)
        if employee.exact_rest_days is not None:
            rest_deficit = model.NewIntVar(0, employee.exact_rest_days, f"relaxed_exact_rest_deficit_{employee_id}")
            rest_excess = model.NewIntVar(0, config.days_in_month, f"relaxed_exact_rest_excess_{employee_id}")
            model.Add(rest_count + rest_deficit - rest_excess == employee.exact_rest_days)
            objective_terms.append((rest_deficit + rest_excess) * 120)
        else:
            if employee.min_rest_days is not None:
                rest_deficit = model.NewIntVar(0, employee.min_rest_days, f"relaxed_min_rest_{employee_id}")
                model.Add(rest_count + rest_deficit >= employee.min_rest_days)
                objective_terms.append(rest_deficit * 120)
            if employee.max_rest_days is not None:
                rest_excess = model.NewIntVar(0, config.days_in_month, f"relaxed_max_rest_{employee_id}")
                model.Add(rest_count - rest_excess <= employee.max_rest_days)
                objective_terms.append(rest_excess * 120)

    for day in range(config.days_in_month):
        for unit, requirements in config.required_per_day.items():
            unit_employee_ids = [employee.employee_id for employee in config.employees if employee.unit == unit]
            for label, symbols in (("early", early_symbols), ("late", late_symbols)):
                required = requirements.get(label, 0)
                actual = sum(decision_vars[employee_id, day, shift] for employee_id in unit_employee_ids for shift in symbols)
                deficit = model.NewIntVar(0, len(unit_employee_ids), f"relaxed_{unit}_{label}_deficit_{day}")
                excess = model.NewIntVar(0, len(unit_employee_ids), f"relaxed_{unit}_{label}_excess_{day}")
                model.Add(actual + deficit - excess == required)
                objective_terms.append((deficit + excess) * 1000)

        night_required = config.night_total_per_day
        night_actual = sum(decision_vars[employee.employee_id, day, shift] for employee in config.employees for shift in night_symbols)
        night_deficit = model.NewIntVar(0, len(config.employees), f"relaxed_night_deficit_{day}")
        night_excess = model.NewIntVar(0, len(config.employees), f"relaxed_night_excess_{day}")
        model.Add(night_actual + night_deficit - night_excess == night_required)
        objective_terms.append((night_deficit + night_excess) * 1000)

    if objective_terms:
        model.Minimize(sum(objective_terms))

    return model, decision_vars


def summarize_partial_validation(validation: dict[str, object]) -> list[str]:
    summary_items = [
        ("人数条件未充足", len(validation.get("staffing_issues", []))),
        ("曜日制限違反", len(validation.get("weekday_restriction_violations", []))),
        ("土日ペア条件違反", len(validation.get("weekend_pair_rest_violations", []))),
        ("夜勤後遷移違反", len(validation.get("night_follow_violations", []))),
        ("遅→早違反", len(validation.get("late_to_early_violations", []))),
        ("連勤違反", len(validation.get("work_streak_violations", []))),
        ("夜勤連続違反", len(validation.get("night_streak_violations", []))),
        ("連休違反", len(validation.get("regular_rest_violations", [])) + len(validation.get("all_rest_violations", []))),
    ]
    return [f"{label}: {count}件" for label, count in summary_items if count > 0]


def solve_schedule(config: SchedulerConfig) -> ScheduleSolveResult:
    model, decision_vars = build_schedule_model(config)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return ScheduleSolveResult(schedule=build_schedule_from_solver(config, decision_vars, solver))

    relaxed_model, relaxed_decision_vars = build_relaxed_schedule_model(config)
    relaxed_solver = cp_model.CpSolver()
    relaxed_solver.parameters.max_time_in_seconds = 30
    relaxed_solver.parameters.num_search_workers = 8
    relaxed_status = relaxed_solver.Solve(relaxed_model)
    if relaxed_status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("勤務表を作成できませんでした。設定ファイルの制約を見直してください。")

    partial_schedule = build_schedule_from_solver(config, relaxed_decision_vars, relaxed_solver)
    partial_validation = validate_schedule(config, partial_schedule)
    diagnostics = {
        "strict_status": solver.StatusName(status),
        "relaxed_status": relaxed_solver.StatusName(relaxed_status),
        "summary_lines": summarize_partial_validation(partial_validation),
    }
    return ScheduleSolveResult(
        schedule=partial_schedule,
        is_partial=True,
        message="厳密解を作成できなかったため、途中案と違反理由を出力しました。",
        diagnostics=diagnostics,
    )


def update_calendar_headers(worksheet, year: int, month: int, days_in_month: int) -> None:
    for day in range(1, 32):
        column = day + 4
        worksheet.Cells(4, column).Value = day if day <= days_in_month else ""
        worksheet.Cells(5, column).Value = JAPANESE_WEEKDAYS[calendar.weekday(year, month, day)] if day <= days_in_month else ""


def validate_calendar_headers_not_blank(worksheet, target_path: Path, sheet_index: int, days_in_month: int) -> None:
    for day in range(1, days_in_month + 1):
        column = day + 4
        date_text = normalize_cell_text(worksheet.Cells(4, column).Text)
        if not date_text:
            raise ValueError(
                "日付チェックで空欄を検出しました。"
                f" 勤怠記入前に {day}日の日付を入力してください。"
                f"\n対象ファイル: {target_path}"
                f"\nシート番号: {sheet_index}"
                f"\nセル位置: {day + 4}列目の4行目"
            )

    for day in range(1, days_in_month + 1):
        column = day + 4
        weekday_text = normalize_cell_text(worksheet.Cells(5, column).Text)
        if not weekday_text:
            raise ValueError(
                "曜日チェックで空欄を検出しました。"
                f" 勤怠記入前に {day}日の日付に対応する曜日を入力してください。"
                f"\n対象ファイル: {target_path}"
                f"\nシート番号: {sheet_index}"
                f"\nセル位置: {day + 4}列目の5行目"
            )


def should_write_employee_row(employee: EmployeeConfig) -> bool:
    return employee.employment != "part"


def write_schedule_to_excel(config: SchedulerConfig, schedule: dict[str, list[str]]) -> None:
    excel = create_excel_application()
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = open_workbook(excel, config.target_path)
    try:
        worksheet = workbook.Worksheets(config.sheet_index)
        validate_calendar_headers_not_blank(worksheet, config.target_path, config.sheet_index, config.days_in_month)
        update_calendar_headers(worksheet, config.year, config.month, config.days_in_month)

        reference_title = None
        try:
            reference_year, reference_month, _ = detect_template_period(config.manual_source, config.sheet_index)
            if reference_year == config.year and reference_month == config.month and config.manual_source.exists():
                reference_title = read_title_from_workbook(config.manual_source, config.sheet_index)
        except Exception:
            reference_title = None

        count_column_map = {"夜": 36, "早": 37, "遅": 38, "休": 39}
        night_rest_symbols = set(symbol_names_by_kind(config.shift_kinds, "night_rest"))
        for employee in config.employees:
            employee_id = employee.employee_id
            row = employee.row
            for day_offset in range(31):
                column = day_offset + 5
                if should_write_employee_row(employee) and day_offset < config.days_in_month:
                    value = schedule[employee_id][day_offset]
                else:
                    value = ""
                worksheet.Cells(row, column).Value = value

            for label, symbol in config.count_symbols.items():
                if should_write_employee_row(employee):
                    counted_symbols = {symbol}
                    if label == "休":
                        counted_symbols.update(night_rest_symbols)
                    count_value = sum(shift in counted_symbols for shift in schedule[employee_id])
                else:
                    count_value = ""
                worksheet.Cells(row, count_column_map[label]).Value = count_value

        worksheet.Cells(1, 1).Value = reference_title if reference_title is not None else f"勤務表　　【R{config.year - 2018}年　{config.month}　月】                             "
        worksheet.Cells(2, 3).Value = config.unit_name
        workbook.Save()
    finally:
        workbook.Close(True)
        excel.Quit()


def validate_schedule(config: SchedulerConfig, schedule: dict[str, list[str]]) -> dict[str, object]:
    early_symbols = set(symbol_names_by_kind(config.shift_kinds, "early"))
    late_symbols = set(symbol_names_by_kind(config.shift_kinds, "late"))
    day_symbols = set(symbol_names_by_kind(config.shift_kinds, "day"))
    night_symbols = set(symbol_names_by_kind(config.shift_kinds, "night"))
    night_rest_symbols = set(symbol_names_by_kind(config.shift_kinds, "night_rest"))
    rest_symbols = set(symbol_names_by_kind(config.shift_kinds, "rest"))
    rest_like_symbols = rest_symbols | night_rest_symbols
    work_symbols = set(symbol for symbol in config.shift_kinds if symbol not in rest_like_symbols)
    regular_rest_limit_symbols = ({"休"} if "休" in config.shift_kinds else set()) | night_rest_symbols
    weekend_days = weekend_day_indexes(config.year, config.month, config.days_in_month)
    weekend_pairs = weekend_pair_day_indexes(config.year, config.month, config.days_in_month)
    issues: list[str] = []
    staffing_issues: list[str] = []
    late_to_early_violations: list[str] = []
    night_follow_violations: list[str] = []
    weekday_restriction_violations: list[str] = []
    weekend_pair_rest_violations: list[str] = []
    work_streak_violations: list[str] = []
    night_streak_violations: list[str] = []
    regular_rest_violations: list[str] = []
    all_rest_violations: list[str] = []
    night_rest_chain_violations: list[str] = []
    used_special_leave = sorted(
        employee.display_name
        for employee in config.employees
        if "特" in schedule[employee.employee_id]
    )
    night_rest_chain_max: dict[str, int] = {}
    double_night_requirement_violations: list[str] = []
    four_day_streak_violations: list[str] = []

    for day in range(config.days_in_month):
        for unit, requirements in config.required_per_day.items():
            unit_employees = [employee for employee in config.employees if employee.unit == unit]
            early_count = sum(schedule[employee.employee_id][day] in early_symbols for employee in unit_employees)
            late_count = sum(schedule[employee.employee_id][day] in late_symbols for employee in unit_employees)
            if early_count != requirements.get("early", 0):
                message = f"{day + 1}日: {unit}の早番人数異常"
                issues.append(message)
                staffing_issues.append(message)
            if late_count != requirements.get("late", 0):
                message = f"{day + 1}日: {unit}の遅番人数異常"
                issues.append(message)
                staffing_issues.append(message)
        night_count = sum(schedule[employee.employee_id][day] in night_symbols for employee in config.employees)
        if night_count != config.night_total_per_day:
            message = f"{day + 1}日: 夜勤人数異常"
            issues.append(message)
            staffing_issues.append(message)

    for employee in config.employees:
        employee_id = employee.employee_id
        shifts = list(employee.previous_tail) + schedule[employee_id]
        max_consecutive_work_limit = employee_max_consecutive_work(employee, config)
        if config.require_weekend_pair_rest and employee.require_weekend_pair_rest:
            for saturday_day, sunday_day in weekend_pairs:
                saturday_shift = schedule[employee_id][saturday_day]
                sunday_shift = schedule[employee_id][sunday_day]
                if saturday_shift not in rest_like_symbols and sunday_shift not in rest_like_symbols:
                    message = f"{employee.display_name}: {saturday_day + 1}日(土)-{sunday_day + 1}日(日) の両方が勤務"
                    issues.append(message)
                    weekend_pair_rest_violations.append(message)
        for day in range(1, config.days_in_month + 1):
            effective_allowed_shifts = effective_allowed_shifts_for_day(config, employee, day)
            if effective_allowed_shifts is None:
                continue
            shift = schedule[employee_id][day - 1]
            if shift not in effective_allowed_shifts:
                message = f"{employee.display_name}: {day}日の勤務 {shift} が曜日別勤務制限に違反"
                issues.append(message)
                weekday_restriction_violations.append(message)
        work_streak = 0
        night_streak = 0
        all_rest_streak = 0
        regular_rest_streak = 0
        night_rest_chain_streak = 0
        max_night_rest_chain_streak = 0
        for index, shift in enumerate(shifts):
            if index < len(shifts) - 1:
                next_shift = shifts[index + 1]
                if shift in late_symbols and next_shift in early_symbols:
                    message = f"{employee.display_name}: 遅→早 違反"
                    issues.append(message)
                    late_to_early_violations.append(message)
                if shift in night_symbols and next_shift not in night_symbols | night_rest_symbols:
                    message = f"{employee.display_name}: 夜→{next_shift} 違反"
                    issues.append(message)
                    night_follow_violations.append(message)

            if shift in work_symbols:
                work_streak += 1
                all_rest_streak = 0
                regular_rest_streak = 0
            else:
                work_streak = 0
                all_rest_streak += 1
                regular_rest_streak = regular_rest_streak + 1 if shift in regular_rest_limit_symbols else 0

            night_streak = night_streak + 1 if shift in night_symbols else 0

            if work_streak > max_consecutive_work_limit:
                message = f"{employee.display_name}: {max_consecutive_work_limit}連勤超過"
                issues.append(message)
                work_streak_violations.append(message)
            if night_streak > config.max_consecutive_night:
                message = f"{employee.display_name}: 夜勤{config.max_consecutive_night}連続超過"
                issues.append(message)
                night_streak_violations.append(message)
            if regular_rest_streak > config.max_consecutive_rest:
                message = f"{employee.display_name}: 休{config.max_consecutive_rest}連続超過"
                issues.append(message)
                regular_rest_violations.append(message)
            if all_rest_streak > config.max_consecutive_rest_with_special:
                message = f"{employee.display_name}: 休/特{config.max_consecutive_rest_with_special}連続超過"
                issues.append(message)
                all_rest_violations.append(message)

            previous_shift = shifts[index - 1] if index > 0 else None
            if shift in night_rest_symbols:
                if previous_shift in night_symbols:
                    night_rest_chain_streak += 1
                    max_night_rest_chain_streak = max(max_night_rest_chain_streak, night_rest_chain_streak)
                    if night_rest_chain_streak == 10:
                        day_number = index - len(employee.previous_tail) + 1
                        message = f"{employee.display_name}: 夜勤明け休みが10回以上連続 ({day_number}日目で到達)"
                        night_rest_chain_violations.append(message)
                else:
                    night_rest_chain_streak = 0
            elif shift in rest_like_symbols:
                night_rest_chain_streak = 0

        night_rest_chain_max[employee.display_name] = max_night_rest_chain_streak

        if employee.max_four_day_streak_count is not None:
            four_day_window_count = count_consecutive_work_windows(schedule[employee_id], work_symbols, 4)
            if four_day_window_count > employee.max_four_day_streak_count:
                message = f"{employee.display_name}: 4連勤窓が {four_day_window_count} 回で、許容 {employee.max_four_day_streak_count} 回を超過"
                issues.append(message)
                four_day_streak_violations.append(message)

    specified_holiday_night_violations: list[str] = []
    specified_holiday_unchecked: list[str] = []
    specified_holiday_count = sum(len(employee.specified_holidays) for employee in config.employees)
    standard_day_shift_symbols = standard_day_symbols(config.shift_kinds)
    rest_day_counts = {
        employee.display_name: sum(shift in rest_like_symbols for shift in schedule[employee.employee_id])
        for employee in config.employees
    }
    exact_rest_day_targets = {
        employee.display_name: employee.exact_rest_days
        for employee in config.employees
        if employee.exact_rest_days is not None
    }
    exact_rest_day_violations: list[str] = []
    for employee in config.employees:
        if employee.exact_rest_days is None:
            continue
        actual_rest_days = rest_day_counts[employee.display_name]
        if actual_rest_days != employee.exact_rest_days:
            message = f"{employee.display_name}: 休み総数 {actual_rest_days} 回 (指定 {employee.exact_rest_days} 回)"
            issues.append(message)
            exact_rest_day_violations.append(message)

    for employee in config.employees:
        employee_id = employee.employee_id
        shifts = schedule[employee_id]
        for holiday_day in employee.specified_holidays:
            if not 1 <= holiday_day <= config.days_in_month:
                continue
            if holiday_day == 1:
                if employee.previous_tail:
                    if employee.previous_tail[-1] in night_symbols:
                        specified_holiday_night_violations.append(f"{employee.display_name}: 指定休日 {holiday_day} 日の前日が夜")
                else:
                    specified_holiday_unchecked.append(f"{employee.display_name}: 指定休日 {holiday_day} 日は前月末情報がないため未検証")
                continue
            if shifts[holiday_day - 2] in night_symbols:
                specified_holiday_night_violations.append(f"{employee.display_name}: 指定休日 {holiday_day} 日の前日が夜")

    night_eligible = [employee for employee in config.employees if any(config.shift_kinds.get(shift) == "night" for shift in employee.allowed_shifts)]
    night_counts = {employee.display_name: sum(shift in night_symbols for shift in schedule[employee.employee_id]) for employee in night_eligible}
    day_shift_counts = {employee.display_name: sum(shift in day_symbols for shift in schedule[employee.employee_id]) for employee in config.employees}
    standard_day_shift_counts = {
        employee.display_name: sum(shift in standard_day_shift_symbols for shift in schedule[employee.employee_id])
        for employee in config.employees
    }
    standard_day_shift_target_counts = {
        employee.display_name: standard_day_shift_counts[employee.display_name]
        for employee in config.employees
        if int(employee.min_counts.get("日", 0)) > 0
    }
    standard_day_shift_violations = [
        f"{employee.display_name}: 通常の日が {standard_day_shift_counts[employee.display_name]} 回で、最低 {employee.min_counts['日']} 回に未達"
        for employee in config.employees
        if int(employee.min_counts.get("日", 0)) > standard_day_shift_counts[employee.display_name]
    ]
    double_night_counts = {
        employee.display_name: sum(
            1
            for day in range(config.days_in_month - 1)
            if schedule[employee.employee_id][day] in night_symbols and schedule[employee.employee_id][day + 1] in night_symbols
        )
        for employee in config.employees
        if any(config.shift_kinds.get(shift) == "night" for shift in employee.allowed_shifts)
    }
    required_double_night_target_counts = {
        employee.display_name: int(employee.required_double_night_min_count)
        for employee in config.employees
        if employee.required_double_night_min_count is not None
    }
    for employee in config.employees:
        if employee.required_double_night_min_count is None:
            continue
        actual_count = double_night_counts.get(employee.display_name, 0)
        if actual_count < employee.required_double_night_min_count:
            message = f"{employee.display_name}: 夜夜 {actual_count} 回で、最低 {employee.required_double_night_min_count} 回に未達"
            issues.append(message)
            double_night_requirement_violations.append(message)

    night_fairness_names = [employee.display_name for employee in config.employees if employee.employee_id in selected_night_fairness_employee_ids(config, list(night_symbols))]
    night_fairness_counts = {name: night_counts[name] for name in night_fairness_names if name in night_counts}
    weekend_rest_symbols = set(weekend_rest_symbols_for_mode(config.shift_kinds, config.weekend_rest_count_mode))
    weekend_rest = {
        employee.display_name: sum(schedule[employee.employee_id][day] in weekend_rest_symbols for day in weekend_days)
        for employee in config.employees
        if employee.employee_id in selected_weekend_fairness_employee_ids(config)
    }
    weekend_four_day_target_counts = {
        employee.display_name: int(employee.max_four_day_streak_count)
        for employee in config.employees
        if employee.max_four_day_streak_count is not None
    }
    return {
        "issues": issues,
        "night_counts": night_counts,
        "night_spread": (0 if not night_counts else max(night_counts.values()) - min(night_counts.values())),
        "night_fairness_counts": night_fairness_counts,
        "night_fairness_spread": (0 if not night_fairness_counts else max(night_fairness_counts.values()) - min(night_fairness_counts.values())),
        "day_shift_counts": day_shift_counts,
        "rest_day_counts": rest_day_counts,
        "exact_rest_day_targets": exact_rest_day_targets,
        "exact_rest_day_violations": exact_rest_day_violations,
        "standard_day_shift_counts": standard_day_shift_counts,
        "standard_day_shift_target_counts": standard_day_shift_target_counts,
        "standard_day_shift_violations": standard_day_shift_violations,
        "double_night_counts": double_night_counts,
        "required_double_night_target_counts": required_double_night_target_counts,
        "double_night_requirement_violations": double_night_requirement_violations,
        "weekend_rest": weekend_rest,
        "weekend_rest_spread": (0 if not weekend_rest else max(weekend_rest.values()) - min(weekend_rest.values())),
        "weekend_rest_count_mode": config.weekend_rest_count_mode,
        "weekend_four_day_target_counts": weekend_four_day_target_counts,
        "used_special_leave": used_special_leave,
        "staffing_issues": staffing_issues,
        "late_to_early_violations": late_to_early_violations,
        "night_follow_violations": night_follow_violations,
        "weekday_restriction_violations": weekday_restriction_violations,
        "weekend_pair_rest_violations": weekend_pair_rest_violations,
        "work_streak_violations": work_streak_violations,
        "four_day_streak_violations": four_day_streak_violations,
        "night_streak_violations": night_streak_violations,
        "regular_rest_violations": regular_rest_violations,
        "all_rest_violations": all_rest_violations,
        "night_rest_chain_violations": night_rest_chain_violations,
        "night_rest_chain_max": night_rest_chain_max,
        "specified_holiday_count": specified_holiday_count,
        "specified_holiday_night_violations": specified_holiday_night_violations,
        "specified_holiday_unchecked": specified_holiday_unchecked,
        "previous_tail_configured": any(employee.previous_tail for employee in config.employees),
    }


def compare_workbooks(source_path: Path, target_path: Path, sheet_index: int = 1) -> list[tuple[int, int, str, str]]:
    excel = create_excel_application()
    excel.Visible = False
    excel.DisplayAlerts = False
    source_book = None
    target_book = None
    try:
        source_book = open_workbook(excel, source_path)
        target_book = open_workbook(excel, target_path)
        if source_book is None or target_book is None:
            raise RuntimeError("Excel COM で勤務表を開けませんでした。")
        return compare_worksheets(source_book.Worksheets(sheet_index), target_book.Worksheets(sheet_index))
    except Exception:
        return compare_workbooks_xlrd(source_path, target_path, sheet_index)
    finally:
        if source_book is not None:
            try:
                source_book.Close(False)
            except Exception:
                pass
        if target_book is not None:
            try:
                target_book.Close(False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass


def sync_workbook(source_path: Path, target_path: Path, sheet_index: int = 1) -> int:
    excel = create_excel_application()
    excel.Visible = False
    excel.DisplayAlerts = False
    source_book = open_workbook(excel, source_path)
    target_book = open_workbook(excel, target_path)
    try:
        source_sheet = source_book.Worksheets(sheet_index)
        target_sheet = target_book.Worksheets(sheet_index)
        max_row = max(source_sheet.UsedRange.Rows.Count, target_sheet.UsedRange.Rows.Count)
        max_col = max(source_sheet.UsedRange.Columns.Count, target_sheet.UsedRange.Columns.Count)
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                target_sheet.Cells(row, col).Formula = source_sheet.Cells(row, col).Formula
        target_book.Save()
        return len(compare_worksheets(source_sheet, target_sheet))
    finally:
        source_book.Close(False)
        target_book.Close(True)
        excel.Quit()


def collect_assignment_diff_rows(
    source_path: Path,
    target_path: Path,
    sheet_index: int,
    employee_rows: list[int],
    max_days: int = 31,
) -> list[dict[str, object]]:
    excel = create_excel_application()
    excel.Visible = False
    excel.DisplayAlerts = False
    source_book = None
    target_book = None
    try:
        source_book = open_workbook(excel, source_path)
        target_book = open_workbook(excel, target_path)
        if source_book is None or target_book is None:
            raise RuntimeError("Excel COM で勤務表を開けませんでした。")
        source_sheet = source_book.Worksheets(sheet_index)
        target_sheet = target_book.Worksheets(sheet_index)
        results: list[dict[str, object]] = []
        for row in employee_rows:
            source_name = str(source_sheet.Cells(row, 4).Text or "")
            target_name = str(target_sheet.Cells(row, 4).Text or "")
            diffs: list[dict[str, object]] = []
            for day in range(1, max_days + 1):
                column = day + 4
                source_value = str(source_sheet.Cells(row, column).Text or "")
                target_value = str(target_sheet.Cells(row, column).Text or "")
                if source_value != target_value:
                    diffs.append({"day": day, "manual": source_value, "generated": target_value})
            results.append(
                {
                    "row": row,
                    "manual_name": source_name,
                    "generated_name": target_name,
                    "diff_count": len(diffs),
                    "diffs": diffs,
                }
            )
        return results
    except Exception:
        return collect_assignment_diff_rows_xlrd(source_path, target_path, sheet_index, employee_rows, max_days)
    finally:
        if source_book is not None:
            try:
                source_book.Close(False)
            except Exception:
                pass
        if target_book is not None:
            try:
                target_book.Close(False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass


def build_validation_results(config: SchedulerConfig, validation: dict[str, object]) -> list[dict[str, object]]:
    full_time_employees = [employee.display_name for employee in config.employees if employee.employment == "full"]
    special_leave_names = list(validation.get("used_special_leave", []))
    part_time_entries = [
        f"{employee.display_name}: {next((shift for shift in employee.allowed_shifts if shift in {'5.5', '6.0'}), '')}"
        for employee in config.employees
        if employee.employment == "part"
    ]
    fixed_assignment_entries: list[str] = []
    fixed_assignment_count = 0
    specified_holiday_entries: list[str] = []
    for employee in config.employees:
        specified_holiday_days = tuple(sorted(set(employee.specified_holidays)))
        non_holiday_fixed_assignments = [
            (day, shift)
            for day, shift in sorted(employee.fixed_assignments.items())
            if not (shift == "休" and day in specified_holiday_days)
        ]
        fixed_assignment_count += len(non_holiday_fixed_assignments)
        if non_holiday_fixed_assignments:
            fixed_assignment_entries.append(
                f"{employee.display_name}: {', '.join(f'{day}日={shift}' for day, shift in non_holiday_fixed_assignments[:5])}"
            )
        if specified_holiday_days:
            displayed_days = ", ".join(f"{day}日" for day in specified_holiday_days[:5])
            if len(specified_holiday_days) > 5:
                displayed_days += ", ..."
            specified_holiday_entries.append(f"{employee.display_name}: {displayed_days}")

    weekend_days = ", ".join(str(day + 1) for day in weekend_day_indexes(config.year, config.month, config.days_in_month))
    weekend_pair_count = len(weekend_pair_day_indexes(config.year, config.month, config.days_in_month))
    has_required_symbols = all(symbol in config.shift_kinds for symbol in ["早", "遅", "日", "夜", "休", "特"]) and bool(symbol_names_by_kind(config.shift_kinds, "night_rest"))
    double_night_total = sum(int(count) for count in validation.get("double_night_counts", {}).values())
    exact_rest_day_targets = validation.get("exact_rest_day_targets", {})
    exact_rest_day_violations = validation.get("exact_rest_day_violations", [])
    required_double_night_target_counts = validation.get("required_double_night_target_counts", {})
    double_night_requirement_violations = validation.get("double_night_requirement_violations", [])
    standard_day_shift_target_counts = validation.get("standard_day_shift_target_counts", {})
    standard_day_shift_violations = validation.get("standard_day_shift_violations", [])
    weekday_rule_employee_count = sum(1 for employee in config.employees if employee.weekday_allowed_shifts)
    weekday_override_count = sum(len(employee.date_allowed_shift_overrides) for employee in config.employees)
    weekend_pair_target_count = sum(1 for employee in config.employees if employee.require_weekend_pair_rest)
    four_day_streak_violations = validation.get("four_day_streak_violations", [])
    four_day_target_counts = validation.get("weekend_four_day_target_counts", {})
    night_fairness_counts = validation.get("night_fairness_counts", validation.get("night_counts", {}))
    night_fairness_spread = validation.get("night_fairness_spread", validation.get("night_spread", 0))
    weekend_rest_mode = validation.get("weekend_rest_count_mode", config.weekend_rest_count_mode)

    results: list[dict[str, object]] = []
    if validation.get("partial_mode"):
        results.append(
            {
                "category": "途中経過出力",
                "title": "暫定案の出力",
                "status": "要注意",
                "detail": str(validation.get("partial_reason") or "厳密解を作成できなかったため、途中案を出力しました。"),
                "evidence": list(validation.get("partial_summary_lines", [])),
            }
        )

    results.extend([
        {
            "category": "ルール適合性",
            "title": "記号体系",
            "status": "適合" if has_required_symbols else "不適合",
            "detail": "ルールで定義された 早・遅・日・夜・夜休・休・特 を設定済みです。" if has_required_symbols else "ルールで定義された勤務記号が不足しています。",
            "evidence": [f"設定済み記号: {', '.join(display_symbol(symbol) for symbol in config.shift_kinds.keys())}"],
        },
        {
            "category": "ルール適合性",
            "title": "南北の早番・遅番配置",
            "status": "適合" if not validation["staffing_issues"] else "不適合",
            "detail": "各ユニットの早番・遅番人数条件を満たしています。" if not validation["staffing_issues"] else "人数条件を満たさない日があります。",
            "evidence": validation["staffing_issues"][:20],
        },
        {
            "category": "ルール適合性",
            "title": "遅番翌日の早番禁止",
            "status": "適合" if not validation["late_to_early_violations"] else "不適合",
            "detail": "遅→早 の禁止条件を満たしています。" if not validation["late_to_early_violations"] else "遅→早 の違反があります。",
            "evidence": validation["late_to_early_violations"][:20],
        },
        {
            "category": "ルール適合性",
            "title": "曜日別勤務制限",
            "status": "適合" if not validation["weekday_restriction_violations"] else "不適合",
            "detail": (
                f"曜日ルール {weekday_rule_employee_count} 人分、日付上書き {weekday_override_count} 件を反映しています。"
                if not validation["weekday_restriction_violations"]
                else "曜日ルールまたは日付上書きに違反があります。"
            ),
            "evidence": validation["weekday_restriction_violations"][:20],
        },
        {
            "category": "ルール適合性",
            "title": "指定勤務日の反映",
            "status": "要注意" if fixed_assignment_count == 0 else "適合",
            "detail": (
                "指定勤務は未設定です。"
                if fixed_assignment_count == 0
                else f"指定勤務 {fixed_assignment_count} 件を最優先で反映します。元勤務表を読み込み、設定ファイルの指定勤務で必要時だけ上書きします。"
            ),
            "evidence": fixed_assignment_entries[:20],
        },
        {
            "category": "ルール適合性",
            "title": "土日のどちらか休み",
            "status": (
                "要注意"
                if weekend_pair_target_count == 0
                else "適合" if not validation["weekend_pair_rest_violations"] else "不適合"
            ),
            "detail": (
                "対象職員が未設定のため未検証です。"
                if weekend_pair_target_count == 0
                else f"対象 {weekend_pair_target_count} 人について、{weekend_pair_count} 組の土日ペアで土曜か日曜のどちらかが休みまたは夜休になる条件を満たしています。"
                if not validation["weekend_pair_rest_violations"]
                else "土日ペアの両方が勤務になっている箇所があります。"
            ),
            "evidence": validation["weekend_pair_rest_violations"][:20],
        },
        {
            "category": "ルール適合性",
            "title": "夜勤翌日の勤務制限と夜勤2連続上限",
            "status": "適合" if not validation["night_follow_violations"] and not validation["night_streak_violations"] else "不適合",
            "detail": "夜勤後の遷移制限と夜夜の上限を満たしています。" if not validation["night_follow_violations"] and not validation["night_streak_violations"] else "夜勤後の遷移または夜勤連続数に違反があります。",
            "evidence": [*validation["night_follow_violations"][:10], *validation["night_streak_violations"][:10]],
        },
        {
            "category": "ルール適合性",
            "title": "連勤上限",
            "status": "適合" if not validation["work_streak_violations"] else "不適合",
            "detail": "設定された連勤上限を超える勤務はありません。" if not validation["work_streak_violations"] else "設定された連勤上限を超える勤務があります。",
            "evidence": validation["work_streak_violations"][:20],
        },
        {
            "category": "ルール適合性",
            "title": "4連勤許容回数",
            "status": (
                "要注意"
                if not four_day_target_counts
                else "適合" if not four_day_streak_violations
                else "不適合"
            ),
            "detail": (
                "4連勤許容回数は未設定です。"
                if not four_day_target_counts
                else f"対象 {len(four_day_target_counts)} 人について、4連勤窓の回数上限を満たしています。"
                if not four_day_streak_violations
                else "4連勤窓の回数上限を超える対象者がいます。"
            ),
            "evidence": four_day_streak_violations[:20],
        },
        {
            "category": "ルール適合性",
            "title": "連休上限",
            "status": "適合" if not validation["regular_rest_violations"] and not validation["all_rest_violations"] else "不適合",
            "detail": "休3連休まで、特を含めて5連休までの条件を満たしています。" if not validation["regular_rest_violations"] and not validation["all_rest_violations"] else "連休上限を超える箇所があります。",
            "evidence": [*validation["regular_rest_violations"][:10], *validation["all_rest_violations"][:10]],
        },
        {
            "category": "ルール適合性",
            "title": "夜勤回数の平等化",
            "status": "適合" if config.fairness_night_spread is None or night_fairness_spread <= config.fairness_night_spread else "不適合",
            "detail": f"夜勤回数のばらつきは {night_fairness_spread} 回です。",
            "evidence": [f"{name}: {count}回" for name, count in night_fairness_counts.items()],
        },
        {
            "category": "ルール適合性",
            "title": "夜夜の2連続を1回以上入れる配慮",
            "status": (
                (
                    "適合" if not double_night_requirement_violations else "不適合"
                )
                if required_double_night_target_counts
                else ("適合" if double_night_total > 0 else "要注意")
            ),
            "detail": (
                (
                    f"対象 {len(required_double_night_target_counts)} 人について、夜夜の最低回数条件を満たしています。"
                    if not double_night_requirement_violations
                    else "夜夜の最低回数条件に未達の対象者がいます。"
                )
                if required_double_night_target_counts
                else ("夜夜の2連続が少なくとも1回あります。" if double_night_total > 0 else "夜夜の2連続は確認できませんでした。")
            ),
            "evidence": (
                [
                    *double_night_requirement_violations,
                    *[
                        f"{name}: 実績{validation['double_night_counts'].get(name, 0)}回 / 最低{count}回"
                        for name, count in required_double_night_target_counts.items()
                    ],
                ][:20]
                if required_double_night_target_counts
                else [f"{name}: {count}回" for name, count in validation.get("double_night_counts", {}).items()]
            ),
        },
        {
            "category": "ルール適合性",
            "title": "休回数の個別指定",
            "status": (
                "要注意"
                if not exact_rest_day_targets
                else "適合" if not exact_rest_day_violations
                else "不適合"
            ),
            "detail": (
                "exact_rest_days が未設定のため未検証です。"
                if not exact_rest_day_targets
                else f"対象 {len(exact_rest_day_targets)} 人について、休み総数を指定回数に一致させています。"
                if not exact_rest_day_violations
                else "休み総数が指定回数と一致しない対象者がいます。"
            ),
            "evidence": [
                *exact_rest_day_violations,
                *[
                    f"{name}: 設定{target}回 / 実績{validation['rest_day_counts'][name]}回"
                    for name, target in exact_rest_day_targets.items()
                ],
            ][:20],
        },
        {
            "category": "ルール適合性",
            "title": "土日休みの平等化",
            "status": "適合" if config.fairness_weekend_spread is None or validation["weekend_rest_spread"] <= config.fairness_weekend_spread else "要注意",
            "detail": f"土日休みのばらつきは {validation['weekend_rest_spread']} 回です。カウント方式は {weekend_rest_mode} です。",
            "evidence": [f"{name}: {count}回" for name, count in validation["weekend_rest"].items()],
        },
        {
            "category": "ルール適合性",
            "title": "通常の『日』確保",
            "status": (
                "要注意"
                if not standard_day_shift_target_counts
                else "適合" if not standard_day_shift_violations
                else "不適合"
            ),
            "detail": (
                "対象者が設定されていないため未検証です。"
                if not standard_day_shift_target_counts
                else f"対象 {len(standard_day_shift_target_counts)} 人全員に、通常の『日』を月1回以上配置しています。"
                if not standard_day_shift_violations
                else "通常の『日』が月1回未満の対象者がいます。"
            ),
            "evidence": [
                *standard_day_shift_violations,
                *[f"{name}: {count}回" for name, count in standard_day_shift_target_counts.items()],
            ][:20],
        },
        {
            "category": "ルール適合性",
            "title": "指定休日前日の夜勤禁止",
            "status": (
                "要注意"
                if validation["specified_holiday_count"] == 0 or validation["specified_holiday_unchecked"]
                else "不適合" if validation["specified_holiday_night_violations"]
                else "適合"
            ),
            "detail": (
                "指定休日データが設定ファイルまたは対象勤務表に無いため自動判定できません。"
                if validation["specified_holiday_count"] == 0
                else "前月末情報が不足する指定休日があり、一部未検証です。"
                if validation["specified_holiday_unchecked"]
                else f"指定休日 {validation['specified_holiday_count']} 件について、前日に夜勤は入っていません。"
                if not validation["specified_holiday_night_violations"]
                else "指定休日の前日に夜勤が入っている箇所があります。"
            ),
            "evidence": [
                *validation["specified_holiday_night_violations"],
                *validation["specified_holiday_unchecked"],
                *specified_holiday_entries,
            ][:20],
        },
        {
            "category": "ルール適合性",
            "title": "月初への前月末勤務引継ぎ",
            "status": "適合" if validation.get("previous_tail_configured") else "要注意",
            "detail": "前月末勤務の持ち越し設定を使って検証しています。" if validation.get("previous_tail_configured") else "previous_tail が未設定のため月初引継ぎは未検証です。",
            "evidence": [],
        },
        {
            "category": "ルール適合性",
            "title": "夜勤明け休み（夜休）が10日以上続かない",
            "status": "適合" if not validation["night_rest_chain_violations"] else "不適合",
            "detail": "夜勤明け休みが10回以上続く並びはありません。" if not validation["night_rest_chain_violations"] else "夜勤明け休みが10回以上続く並びがあります。",
            "evidence": validation["night_rest_chain_violations"] or [
                f"{name}: 最大{count}回"
                for name, count in validation.get("night_rest_chain_max", {}).items()
            ],
        },
        {
            "category": "必須",
            "title": "総合検証",
            "status": "適合" if not validation["issues"] else "不適合",
            "detail": "実装済みの必須制約では違反は見つかりませんでした。" if not validation["issues"] else "ルール違反が見つかりました。",
            "evidence": validation["issues"],
        },
        {
            "category": "必須",
            "title": "夜勤回数の平等化",
            "status": "適合",
            "detail": f"夜勤回数のばらつきは {validation.get('night_spread', 0)} 回です。",
            "evidence": [f"{name}: {count}回" for name, count in validation.get("night_counts", {}).items()],
        },
        {
            "category": "配慮",
            "title": "土日休みの平等化",
            "status": "適合",
            "detail": f"土日休みのばらつきは {validation['weekend_rest_spread']} 回です。",
            "evidence": [f"{name}: {validation['weekend_rest'][name]}回" for name in full_time_employees if name in validation["weekend_rest"]],
        },
        {
            "category": "必須",
            "title": "特休の使用",
            "status": "参考",
            "detail": "特休が使われています。" if special_leave_names else "特休は使われていません。",
            "evidence": special_leave_names,
        },
        {
            "category": "運用",
            "title": "非常勤の時間表記",
            "status": "参考",
            "detail": "非常勤は時間表記を使う設定です。",
            "evidence": [entry for entry in part_time_entries if entry.strip().split(": ")[-1]],
        },
        {
            "category": "参考",
            "title": "対象月の土日",
            "status": "参考",
            "detail": f"{config.year}年{config.month}月の土日は {len(weekend_day_indexes(config.year, config.month, config.days_in_month))} 日です。",
            "evidence": [weekend_days] if weekend_days else [],
        },
    ])
    return results


def render_validation_report(
    config: SchedulerConfig,
    validation: dict[str, object],
    full_diffs: list[tuple[int, int, str, str]],
    assignment_rows: list[dict[str, object]],
) -> str:
    partial_note_html = ""
    if validation.get("partial_mode"):
        partial_note_html = (
            '<div class="note">'
            f"<p><strong>途中案を出力しました。</strong> {html.escape(str(validation.get('partial_reason', '')))}</p>"
            "</div>"
        )
    status_counts = {
        "適合": sum(1 for item in validation["results"] if item["status"] == "適合"),
        "不適合": sum(1 for item in validation["results"] if item["status"] == "不適合"),
        "要注意": sum(1 for item in validation["results"] if item["status"] == "要注意"),
        "参考": sum(1 for item in validation["results"] if item["status"] == "参考"),
    }
    assignment_diff_count = sum(int(item["diff_count"]) for item in assignment_rows)
    status_class = {"適合": "ok", "不適合": "ng", "要注意": "warn", "参考": "info"}

    validation_rows = []
    for item in validation["results"]:
        evidence_html = ""
        if item["evidence"]:
            evidence_html = "<ul>" + "".join(f"<li>{html.escape(str(line))}</li>" for line in item["evidence"]) + "</ul>"
        validation_rows.append(
            "<tr>"
            f"<td>{html.escape(str(item['category']))}</td>"
            f"<td>{html.escape(str(item['title']))}</td>"
            f"<td class=\"{status_class.get(str(item['status']), 'info')}\">{html.escape(str(item['status']))}</td>"
            f"<td>{html.escape(str(item['detail']))}{evidence_html}</td>"
            "</tr>"
        )

    assignment_diff_rows = []
    for item in sorted(assignment_rows, key=lambda row: int(row["diff_count"]), reverse=True):
        assignment_diff_rows.append(
            "<tr>"
            f"<td>{item['row']}</td>"
            f"<td>{html.escape(str(item['generated_name']))}</td>"
            f"<td>{html.escape(str(item['manual_name']))}</td>"
            f"<td>{item['diff_count']}</td>"
            "</tr>"
        )

    full_diff_rows = []
    for row, col, source_value, target_value in full_diffs[:20]:
        full_diff_rows.append(
            "<tr>"
            f"<td>R{row}C{col}</td>"
            f"<td>{html.escape(source_value)}</td>"
            f"<td>{html.escape(target_value)}</td>"
            "</tr>"
        )

    return f'''<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>あかねっこ自動生成検証と差分</title>
<style>
:root {{
  --bg: #f3efe7;
  --paper: #fffdf9;
  --ink: #1f2933;
  --muted: #5b6470;
  --line: #d9d1c3;
  --head: #ece4d8;
  --accent: #9a3412;
  --accent-soft: #fff1e8;
  --ok-bg: #edf8f0;
  --ok-fg: #166534;
  --ng-bg: #fff1f2;
  --ng-fg: #b91c1c;
  --warn-bg: #fff7e8;
  --warn-fg: #b45309;
  --info-bg: #ecfeff;
  --info-fg: #0f766e;
}}
* {{ box-sizing: border-box; }}
body {{
  margin: 0;
  color: var(--ink);
  background:
    radial-gradient(circle at top left, rgba(154, 52, 18, .08), transparent 28%),
    linear-gradient(180deg, #f7f2ea 0%, var(--bg) 100%);
  font-family: "Yu Gothic UI", "Meiryo", sans-serif;
}}
main {{
  max-width: 1240px;
  margin: 32px auto;
  background: var(--paper);
  padding: 32px;
  border-radius: 24px;
  border: 1px solid rgba(154, 52, 18, .08);
  box-shadow: 0 20px 50px rgba(74, 49, 24, .10);
}}
h1, h2, p {{ margin-top: 0; }}
h1 {{ margin-bottom: 10px; font-size: 2rem; line-height: 1.2; letter-spacing: .02em; }}
h2 {{ margin-bottom: 12px; font-size: 1.2rem; }}
.hero {{ padding: 20px 22px; border-radius: 18px; background: linear-gradient(135deg, #fff7f2 0%, #fffdf8 100%); border: 1px solid #f1dfd1; }}
.eyebrow {{ display: inline-block; margin-bottom: 10px; padding: 6px 10px; border-radius: 999px; color: var(--accent); background: var(--accent-soft); font-size: .84rem; font-weight: 700; }}
.subtext {{ color: var(--muted); margin-bottom: 0; }}
.summary {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 14px; margin: 22px 0 26px; }}
.card {{ padding: 16px 18px; border-radius: 16px; background: #f7f3ec; border: 1px solid #e7ddd0; }}
.card .label {{ display: block; margin-bottom: 6px; color: var(--muted); font-size: .92rem; }}
.card strong {{ display: block; font-size: 1.8rem; line-height: 1; }}
.card.ok-card {{ background: var(--ok-bg); }}
.card.ng-card {{ background: var(--ng-bg); }}
.card.warn-card {{ background: var(--warn-bg); }}
.card.info-card {{ background: var(--info-bg); }}
.card.accent-card {{ background: #f8efe7; }}
.section {{ margin-top: 30px; padding: 22px; border-radius: 18px; background: #fff; border: 1px solid #efe7dc; }}
.table-wrap {{ overflow-x: auto; border: 1px solid var(--line); border-radius: 14px; }}
table {{ width: 100%; border-collapse: collapse; margin-top: 0; background: #fff; }}
th, td {{ border-bottom: 1px solid var(--line); padding: 11px 12px; vertical-align: top; text-align: left; }}
th {{ position: sticky; top: 0; z-index: 1; background: var(--head); white-space: nowrap; }}
tbody tr:nth-child(even) {{ background: #fffcf7; }}
tbody tr:hover {{ background: #fff7ed; }}
.ok, .ng, .warn, .info {{ display: inline-block; min-width: 68px; padding: 4px 10px; border-radius: 999px; font-weight: 700; text-align: center; }}
.ok {{ color: var(--ok-fg); background: var(--ok-bg); }}
.ng {{ color: var(--ng-fg); background: var(--ng-bg); }}
.warn {{ color: var(--warn-fg); background: var(--warn-bg); }}
.info {{ color: var(--info-fg); background: var(--info-bg); }}
ul {{ margin: 8px 0 0 18px; padding-left: 10px; }}
li + li {{ margin-top: 4px; }}
.note {{ margin-top: 24px; padding: 16px 18px; color: #21415f; background: #eef6ff; border: 1px solid #c9e1ff; border-radius: 14px; }}
@media (max-width: 720px) {{
  body {{ padding: 12px; }}
  main {{ margin: 0; padding: 18px; border-radius: 18px; }}
  .section {{ padding: 16px; }}
  h1 {{ font-size: 1.5rem; }}
}}
</style>
</head>
<body>
<main>
<section class="hero">
<div class="eyebrow">自動生成レポート</div>
<h1>【白紙】あかねっこ.xls 自動生成結果と差分</h1>
<p class="subtext">対象ファイル: {html.escape(config.target_path.name)}</p>
</section>
{partial_note_html}
<div class="summary">
  <div class="card ok-card"><span class="label">適合</span><strong>{status_counts['適合']}</strong></div>
  <div class="card ng-card"><span class="label">不適合</span><strong>{status_counts['不適合']}</strong></div>
  <div class="card warn-card"><span class="label">要注意</span><strong>{status_counts['要注意']}</strong></div>
  <div class="card info-card"><span class="label">参考</span><strong>{status_counts['参考']}</strong></div>
  <div class="card accent-card"><span class="label">総差分</span><strong>{len(full_diffs)}</strong></div>
  <div class="card accent-card"><span class="label">勤務割当差分</span><strong>{assignment_diff_count}</strong></div>
</div>
<div class="section">
<h2>検証結果</h2>
<div class="table-wrap">
<table>
<thead><tr><th>区分</th><th>項目</th><th>判定</th><th>内容</th></tr></thead>
<tbody>{''.join(validation_rows)}</tbody>
</table>
</div>
</div>
<div class="section">
<h2>職員別 勤務割当差分件数</h2>
<div class="table-wrap">
<table>
<thead><tr><th>行</th><th>生成結果の氏名</th><th>手入力版の氏名</th><th>差分件数</th></tr></thead>
<tbody>{''.join(assignment_diff_rows)}</tbody>
</table>
</div>
</div>
<div class="section">
<h2>先頭20件のセル差分</h2>
<div class="table-wrap">
<table>
<thead><tr><th>セル</th><th>手入力版</th><th>生成結果</th></tr></thead>
<tbody>{''.join(full_diff_rows)}</tbody>
</table>
</div>
</div>
<div class="note">
<p>総差分は見出し、曜日、集計欄を含みます。勤務割当差分は職員行の日別勤務欄のみを集計しています。</p>
</div>
</main>
</body>
</html>'''


def write_validation_report(
    config: SchedulerConfig,
    validation_summary: dict[str, object],
    report_path: Path = DEFAULT_REPORT_PATH,
) -> Path:
    excluded_rows = {employee.row for employee in config.employees if not should_write_employee_row(employee)}
    full_diffs = [
        diff
        for diff in compare_workbooks(config.manual_source, config.target_path, config.sheet_index)
        if diff[0] not in excluded_rows
    ]
    assignment_rows = collect_assignment_diff_rows(
        config.manual_source,
        config.target_path,
        config.sheet_index,
        [employee.row for employee in config.employees if should_write_employee_row(employee)],
    )
    report_payload = {
        **validation_summary,
        "results": build_validation_results(config, validation_summary),
    }
    report_text = render_validation_report(config, report_payload, full_diffs, assignment_rows)
    report_path.write_text(report_text, encoding="utf-8")
    return report_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="あかねっこの勤務表を生成・同期・比較します。")
    subparsers = parser.add_subparsers(dest="command")

    generate_parser = subparsers.add_parser("generate", help="JSON設定から勤務表を自動作成")
    generate_parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH, help="設定JSONファイル")
    generate_parser.add_argument("--target", type=Path, default=None, help="書き込み先のExcelファイル")
    generate_parser.add_argument("--year", type=int, default=None, help="西暦年")
    generate_parser.add_argument("--month", type=int, default=None, help="月")
    generate_parser.add_argument("--unit-name", default=None, help="ユニット名")
    generate_parser.add_argument("--days", type=int, default=None, help="作成する日数")

    sync_parser = subparsers.add_parser("sync", help="手入力版の勤務表を対象Excelへ反映")
    sync_parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH, help="設定JSONファイル")
    sync_parser.add_argument("--source", type=Path, default=None, help="反映元のExcelファイル")
    sync_parser.add_argument("--target", type=Path, default=None, help="反映先のExcelファイル")
    sync_parser.add_argument("--sheet-index", type=int, default=None, help="反映対象のシート番号")

    compare_parser = subparsers.add_parser("compare", help="2つの勤務表の差分件数を確認")
    compare_parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH, help="設定JSONファイル")
    compare_parser.add_argument("--source", type=Path, default=None, help="比較元のExcelファイル")
    compare_parser.add_argument("--target", type=Path, default=None, help="比較先のExcelファイル")
    compare_parser.add_argument("--sheet-index", type=int, default=None, help="比較対象のシート番号")
    compare_parser.add_argument("--show-limit", type=int, default=20, help="表示する差分件数の上限")
    return parser


def parse_args() -> argparse.Namespace:
    parser = build_parser()
    raw_args = sys.argv[1:]
    if not raw_args or raw_args[0].startswith("-"):
        raw_args = ["generate", *raw_args]
    return parser.parse_args(raw_args)


def with_generate_overrides(config: SchedulerConfig, args: argparse.Namespace) -> SchedulerConfig:
    target_path = args.target.resolve() if args.target else config.target_path
    detected_year, detected_month, detected_days = detect_template_period(target_path, config.sheet_index)

    if detected_year is None and detected_month is None:
        raise ValueError(
            "対象の勤怠表から年と月を読み取れません。"
            "1行目のタイトルに「R8年4月」のような年と月を入力してください。"
            f"\n対象ファイル: {target_path}"
        )
    if detected_year is None:
        raise ValueError(
            "対象の勤怠表から年を読み取れません。"
            "1行目のタイトルに「R8年4月」のような年を入力してください。"
            f"\n対象ファイル: {target_path}"
        )
    if detected_month is None:
        raise ValueError(
            "対象の勤怠表から月を読み取れません。"
            "1行目のタイトルに「R8年4月」のような月を入力してください。"
            f"\n対象ファイル: {target_path}"
        )

    year = args.year if args.year is not None else detected_year
    month = args.month if args.month is not None else detected_month
    config = load_config(config.config_path, year=year, month=month)
    reference_source = resolve_reference_source(target_path, config.manual_source)
    requested_days = args.days if args.days is not None else detected_days
    days_in_month = normalize_days_in_month(year, month, requested_days, "テンプレートまたは実行引数")

    manual_year, manual_month, _ = detect_template_period(reference_source, config.sheet_index)
    manual_fixed_assignments: dict[str, dict[int, str]] = {}
    if manual_year == year and manual_month == month:
        candidate_manual_fixed_assignments = read_fixed_assignments_from_workbook(
            reference_source,
            config.sheet_index,
            config.employees,
            config.shift_kinds,
            days_in_month,
        )
        if not is_completed_schedule_like_fixed_assignments(candidate_manual_fixed_assignments, days_in_month):
            manual_fixed_assignments = candidate_manual_fixed_assignments

    target_specified_holidays = read_specified_holidays_from_workbook(
        target_path,
        config.sheet_index,
        config.employees,
        days_in_month,
    )
    workbook_employee_settings = read_workbook_employee_settings(target_path, config.sheet_index, config.employees)
    workbook_monthly_settings = read_workbook_monthly_settings(target_path, config.sheet_index)

    previous_tail_length = max(
        config.max_consecutive_work,
        config.max_consecutive_night,
        config.max_consecutive_rest,
        config.max_consecutive_rest_with_special,
    )
    previous_source = resolve_previous_month_source(config.config_path.parent, target_path, reference_source, year, month)
    previous_tails: dict[str, tuple[str, ...]] = {}
    if previous_source is not None and previous_tail_length > 0:
        previous_tails = read_previous_tail_from_workbook(
            previous_source,
            config.sheet_index,
            config.employees,
            config.shift_kinds,
            previous_tail_length,
        )

    merged_employees: list[EmployeeConfig] = []
    for employee in config.employees:
        fixed_assignments = dict(manual_fixed_assignments.get(employee.employee_id, {}))
        fixed_assignments.update(employee.fixed_assignments)
        specified_holidays = tuple(sorted({*employee.specified_holidays, *target_specified_holidays.get(employee.employee_id, ())}))
        for holiday_day in target_specified_holidays.get(employee.employee_id, ()): 
            fixed_assignments[holiday_day] = "休"
        previous_tail = previous_tails.get(employee.employee_id, employee.previous_tail)
        fixed_assignments = normalize_night_rest_assignments(
            fixed_assignments,
            config.shift_kinds,
            days_in_month,
            previous_shift=(previous_tail[-1] if previous_tail else None),
        )
        merged_employees.append(
            EmployeeConfig(
                employee_id=employee.employee_id,
                display_name=employee.display_name,
                unit=employee.unit,
                employment=employee.employment,
                row=employee.row,
                allowed_shifts=employee.allowed_shifts,
                aliases=employee.aliases,
                weekday_allowed_shifts=employee.weekday_allowed_shifts,
                date_allowed_shift_overrides=employee.date_allowed_shift_overrides,
                require_weekend_pair_rest=employee.require_weekend_pair_rest,
                night_fairness_target=bool(workbook_employee_settings.get(employee.employee_id, {}).get("night_fairness_target", employee.night_fairness_target)),
                required_double_night_min_count=(
                    workbook_employee_settings.get(employee.employee_id, {}).get("required_double_night_min_count", employee.required_double_night_min_count)
                ),
                weekend_fairness_target=bool(workbook_employee_settings.get(employee.employee_id, {}).get("weekend_fairness_target", employee.weekend_fairness_target)),
                min_counts=employee.min_counts,
                max_counts=employee.max_counts,
                max_consecutive_work_limit=workbook_employee_settings.get(employee.employee_id, {}).get("max_consecutive_work_limit", employee.max_consecutive_work_limit),
                max_four_day_streak_count=workbook_employee_settings.get(employee.employee_id, {}).get("max_four_day_streak_count", employee.max_four_day_streak_count),
                exact_rest_days=employee.exact_rest_days,
                min_rest_days=employee.min_rest_days,
                max_rest_days=employee.max_rest_days,
                specified_holidays=specified_holidays,
                fixed_assignments=fixed_assignments,
                previous_tail=previous_tail,
            )
        )

    return SchedulerConfig(
        config_path=config.config_path,
        target_path=target_path,
        manual_source=reference_source,
        sheet_index=config.sheet_index,
        year=year,
        month=month,
        days_in_month=days_in_month,
        unit_name=(args.unit_name if args.unit_name is not None else config.unit_name),
        shift_kinds=config.shift_kinds,
        count_symbols=config.count_symbols,
        employees=tuple(merged_employees),
        required_per_day=config.required_per_day,
        night_total_per_day=config.night_total_per_day,
        day_requirements=config.day_requirements,
        max_consecutive_work=config.max_consecutive_work,
        max_consecutive_night=config.max_consecutive_night,
        max_consecutive_rest=config.max_consecutive_rest,
        max_consecutive_rest_with_special=config.max_consecutive_rest_with_special,
        fairness_night_spread=workbook_monthly_settings.get("fairness_night_spread", config.fairness_night_spread),
        fairness_weekend_spread=workbook_monthly_settings.get("fairness_weekend_spread", config.fairness_weekend_spread),
        weekend_rest_count_mode=str(workbook_monthly_settings.get("weekend_rest_count_mode", config.weekend_rest_count_mode)),
        require_weekend_pair_rest=config.require_weekend_pair_rest,
        prefer_double_night=config.prefer_double_night,
    )


def main() -> None:
    args = parse_args()
    config = load_config(args.config)

    if args.command == "generate":
        config = with_generate_overrides(config, args)
        solve_result = solve_schedule(config)
        schedule = solve_result.schedule
        validation = validate_schedule(config, schedule)
        validation["partial_mode"] = solve_result.is_partial
        validation["partial_reason"] = solve_result.message
        validation["partial_summary_lines"] = solve_result.diagnostics.get("summary_lines", [])
        if validation["issues"] and not solve_result.is_partial:
            raise RuntimeError("ルール検証で問題が見つかりました: " + json.dumps(validation, ensure_ascii=False))
        write_schedule_to_excel(config, schedule)
        report_path = write_validation_report(config, validation)
        print(json.dumps({"report": str(report_path), **validation}, ensure_ascii=False, indent=2))
        return

    if args.command == "sync":
        target_path = args.target.resolve() if args.target else config.target_path
        detected_year, detected_month, _ = detect_template_period(target_path, config.sheet_index)
        if detected_year is not None and detected_month is not None:
            config = load_config(args.config, year=detected_year, month=detected_month)
        source_path = args.source.resolve() if args.source else config.manual_source
        sheet_index = args.sheet_index if args.sheet_index is not None else config.sheet_index
        print(f"remaining_diff_count={sync_workbook(source_path, target_path, sheet_index)}")
        return

    if args.command == "compare":
        target_path = args.target.resolve() if args.target else config.target_path
        detected_year, detected_month, _ = detect_template_period(target_path, config.sheet_index)
        if detected_year is not None and detected_month is not None:
            config = load_config(args.config, year=detected_year, month=detected_month)
        source_path = args.source.resolve() if args.source else config.manual_source
        sheet_index = args.sheet_index if args.sheet_index is not None else config.sheet_index
        diffs = compare_workbooks(source_path, target_path, sheet_index)
        print(f"diff_count={len(diffs)}")
        for row, col, source_value, target_value in diffs[: args.show_limit]:
            print(f"R{row}C{col}: [{source_value}] != [{target_value}]")
        return

    raise RuntimeError(f"未対応のコマンドです: {args.command}")


if __name__ == "__main__":
    main()