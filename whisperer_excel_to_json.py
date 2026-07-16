#!/usr/bin/env python3
"""
Excel/CSV to Incident Whisperer JSON converter.

v5 highlights:
- Better logging and diagnostics.
- Logs input workbook, available sheets, selected sheet, rows, columns.
- Column matching is case-insensitive and ignores spaces/newlines/punctuation.
- Sheet name is optional and can be overridden with --sheet.
- Config is split into required_fields and optional_fields.
- Mappings can be:
    null
    "Column Name"
    ["Primary Column", "Fallback Column"]
    {"columns": ["Primary", "Fallback"], "fallback": "row_number", "prefix": "ROW-"}
- Multiple text fields can be combined into inc_comments_and_work_notes.
- Text fields can have timestamps or no timestamps.
- Date normalization is configurable.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import logging
import re
import sys
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    import pandas as pd
except ImportError as exc:
    raise SystemExit("Missing dependency: pandas. Install with: pip install pandas openpyxl") from exc


LOGGER = logging.getLogger("whisperer_converter")


def normalize_name(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\ufeff", "")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = text.strip().casefold()
    return re.sub(r"[^a-z0-9]+", "", text)


def is_nullish(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return isinstance(value, str) and value.strip() == ""


def clean_text(value: Any) -> str:
    if is_nullish(value):
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()


def coerce_runtime_static(value: Any, input_path: Path, batch_id: str) -> Optional[str]:
    if value is None:
        return None
    if not isinstance(value, str):
        return clean_text(value)

    now = dt.datetime.now()
    utc_now = dt.datetime.now(dt.timezone.utc)
    replacements = {
        "__NOW__": now.strftime("%Y-%m-%d %H:%M:%S"),
        "__UTC_NOW__": utc_now.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "__INPUT_FILE__": input_path.name,
        "__INPUT_PATH__": str(input_path.resolve()),
        "__BATCH_ID__": batch_id,
    }
    return replacements.get(value.strip(), value)


@dataclass
class ConverterConfig:
    sheet_name: Optional[Any]
    required_fields: Dict[str, Any]
    optional_fields: Dict[str, Any]
    comments_worknotes: Dict[str, Any]
    static_fields: Dict[str, Any]
    date_fields: Dict[str, Any]
    deduplicate_by_inc_number: bool
    deduplicate_keep_latest_by: Optional[str]
    output_top_level_key: str
    skip_empty_rows: bool
    empty_row_required_any_of: List[str]


def load_config(path: Path) -> ConverterConfig:
    raw = json.loads(path.read_text(encoding="utf-8"))

    if "field_mapping" in raw and ("required_fields" not in raw and "optional_fields" not in raw):
        LOGGER.warning("Legacy 'field_mapping' detected. Prefer required_fields/optional_fields.")
        required_default = [
            "inc_number", "inc_state", "inc_sys_updated_on",
            "inc_sys_created_on", "inc_resolved_at"
        ]
        fm = raw.get("field_mapping") or {}
        required_fields = {k: fm.get(k) for k in required_default if k in fm}
        optional_fields = {k: v for k, v in fm.items() if k not in required_fields}
    else:
        required_fields = raw.get("required_fields") or {}
        optional_fields = raw.get("optional_fields") or {}

    return ConverterConfig(
        sheet_name=raw.get("sheet_name", None),
        required_fields=required_fields,
        optional_fields=optional_fields,
        comments_worknotes=raw.get("comments_worknotes") or {},
        static_fields=raw.get("static_fields") or {},
        date_fields=raw.get("date_fields") or {},
        deduplicate_by_inc_number=bool(raw.get("deduplicate_by_inc_number", True)),
        deduplicate_keep_latest_by=raw.get("deduplicate_keep_latest_by", "inc_sys_updated_on"),
        output_top_level_key=raw.get("output_top_level_key", "result"),
        skip_empty_rows=bool(raw.get("skip_empty_rows", True)),
        empty_row_required_any_of=list(raw.get("empty_row_required_any_of") or [
            "inc_number", "inc_short_description", "inc_description", "inc_comments_and_work_notes"
        ]),
    )


def list_sheets(input_path: Path) -> List[str]:
    if input_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return list(pd.ExcelFile(input_path).sheet_names)
    return []


def read_spreadsheet(input_path: Path, sheet_name: Optional[Any]) -> Tuple[pd.DataFrame, Optional[str]]:
    suffix = input_path.suffix.lower()

    if suffix == ".csv":
        LOGGER.info("Reading CSV: %s", input_path)
        return pd.read_csv(input_path, dtype=object, keep_default_na=False), None

    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError(f"Unsupported file extension: {suffix}")

    sheets = list_sheets(input_path)
    LOGGER.info("Input workbook: %s", input_path)
    LOGGER.info("Available sheets: %s", ", ".join(sheets))

    selected = 0 if sheet_name is None else sheet_name
    if sheet_name is None:
        LOGGER.info("No sheet specified. Reading first sheet by default.")

    df = pd.read_excel(input_path, sheet_name=selected, dtype=object, keep_default_na=False)

    actual = str(selected)
    if isinstance(selected, int) and 0 <= selected < len(sheets):
        actual = sheets[selected]

    LOGGER.info("Reading sheet: %s", actual)
    LOGGER.info("Read %s rows and %s columns.", len(df), len(df.columns))
    return df, actual


def build_column_lookup(columns: Iterable[Any]) -> Dict[str, str]:
    lookup: Dict[str, str] = {}
    duplicates: Dict[str, List[str]] = {}

    for col in columns:
        original = str(col)
        norm = normalize_name(original)
        if not norm:
            continue
        if norm in lookup:
            duplicates.setdefault(norm, [lookup[norm]]).append(original)
        else:
            lookup[norm] = original

    for norm, cols in duplicates.items():
        LOGGER.warning("Duplicate/ambiguous normalized column '%s': %s. First used: %s", norm, cols, cols[0])

    return lookup


def mapping_to_columns(mapping: Any) -> Tuple[List[str], Optional[Dict[str, Any]]]:
    """
    Returns configured source columns and optional mapping object.

    Accepted:
      null
      "Column"
      ["Column A", "Column B"]
      {"columns": ["Column A", "Column B"], "fallback": "row_number", "prefix": "ROW-"}
      {"column": "Column A"}
    """
    if mapping is None:
        return [], None

    if isinstance(mapping, str):
        return [mapping], None

    if isinstance(mapping, list):
        return [str(x) for x in mapping if x is not None and str(x).strip()], None

    if isinstance(mapping, dict):
        if "columns" in mapping and isinstance(mapping["columns"], list):
            return [str(x) for x in mapping["columns"] if x is not None and str(x).strip()], mapping
        if "column" in mapping and mapping["column"]:
            return [str(mapping["column"])], mapping
        return [], mapping

    return [str(mapping)], None


def resolve_columns(mapping: Any, lookup: Dict[str, str], field_name: str, required: bool) -> Tuple[List[str], Optional[Dict[str, Any]]]:
    requested_columns, mapping_obj = mapping_to_columns(mapping)

    if not requested_columns:
        if required:
            LOGGER.error("Required field '%s' has no mapping.", field_name)
        return [], mapping_obj

    resolved = []
    for requested in requested_columns:
        actual = lookup.get(normalize_name(requested))
        if actual:
            resolved.append(actual)
            LOGGER.debug("Resolved %s: '%s' -> '%s'", field_name, requested, actual)
        else:
            message = f"Column for field '{field_name}' not found: '{requested}'"
            if required:
                LOGGER.error(message)
            else:
                LOGGER.warning(message)

    return resolved, mapping_obj


def parse_date_from_string(value: str, field: str, config: ConverterConfig) -> Optional[dt.datetime]:
    value = value.strip()
    if not value:
        return None

    date_cfg = config.date_fields or {}
    default_order = str(date_cfg.get("input_order", "dmy")).lower()
    ambiguous_order = str(date_cfg.get("ambiguous_date_order", default_order)).lower()
    field_cfg = date_cfg.get("field_overrides", {}).get(field, {})
    order = str(field_cfg.get("input_order", default_order)).lower()
    field_ambiguous_order = str(field_cfg.get("ambiguous_date_order", ambiguous_order)).lower()

    # Allow explicit formats first.
    explicit_formats = field_cfg.get("input_formats") or date_cfg.get("input_formats") or []
    default_month_name_formats = [
        "%d-%b-%Y", "%d-%B-%Y", "%d-%b-%y", "%d-%B-%y",
        "%b-%d-%Y", "%B-%d-%Y", "%b-%d-%y", "%B-%d-%y",
        "%d %b %Y", "%d %B %Y", "%b %d %Y", "%B %d %Y",
        "%d-%b-%Y %H:%M:%S", "%d-%B-%Y %H:%M:%S",
        "%b-%d-%Y %H:%M:%S", "%B-%d-%Y %H:%M:%S"
    ]
    for fmt in list(explicit_formats) + default_month_name_formats:
        try:
            return dt.datetime.strptime(value, fmt)
        except ValueError:
            pass

    # Common numeric separators and optional time.
    # input_order:
    #   dmy  -> 10/04/2024 means 10 Apr
    #   mdy  -> 10/04/2024 means Oct 4
    #   auto -> infer when possible; if ambiguous, use ambiguous_date_order
    date_time_patterns = [
        r"^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$",
        r"^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$",
    ]

    m = re.match(date_time_patterns[0], value)
    if m:
        a_raw, b_raw, y, hh, mm, ss = m.groups()
        a, b = int(a_raw), int(b_raw)
        y = int(y)
        if y < 100:
            y += 2000

        effective_order = order
        if order == "auto":
            if a > 12 and b <= 12:
                effective_order = "dmy"
            elif b > 12 and a <= 12:
                effective_order = "mdy"
            else:
                effective_order = field_ambiguous_order

        if effective_order == "mdy":
            month, day = a, b
        elif effective_order == "dmy":
            day, month = a, b
        else:
            LOGGER.warning("Unsupported date input_order=%r for field '%s'. Use dmy, mdy, or auto.", order, field)
            return None

        try:
            return dt.datetime(y, month, day, int(hh or 0), int(mm or 0), int(ss or 0))
        except ValueError:
            LOGGER.warning("Invalid date for field '%s': %r using order=%s", field, value, effective_order)
            return None

    m = re.match(date_time_patterns[1], value)
    if m:
        y, month, day, hh, mm, ss = m.groups()
        try:
            return dt.datetime(int(y), int(month), int(day), int(hh or 0), int(mm or 0), int(ss or 0))
        except ValueError:
            LOGGER.warning("Invalid ISO-like date for field '%s': %r", field, value)
            return None

    return None


def normalize_date_value(value: Any, field: str, config: ConverterConfig) -> str:
    if is_nullish(value):
        return ""

    date_cfg = config.date_fields or {}
    output_format = date_cfg.get("output_format", "%d/%m/%Y %H:%M:%S")
    on_error = str(date_cfg.get("on_parse_error", "fail")).lower()
    date_field_names = set(date_cfg.get("fields") or [
        "inc_sys_created_on", "inc_sys_updated_on", "inc_resolved_at"
    ])

    if field not in date_field_names:
        return clean_text(value)

    parsed: Optional[dt.datetime] = None

    if isinstance(value, dt.datetime):
        parsed = value
    elif isinstance(value, dt.date):
        parsed = dt.datetime.combine(value, dt.time.min)
    elif hasattr(value, "to_pydatetime"):
        try:
            parsed = value.to_pydatetime()
        except Exception:
            parsed = None
    else:
        parsed = parse_date_from_string(clean_text(value), field, config)

    if not parsed:
        message = (
            f"Could not parse date for field '{field}'. "
            f"Value: {value!r}. "
            f"Expected to output format: {output_format}. "
            f"Set date_fields.input_order / ambiguous_date_order / input_formats in config."
        )
        if on_error == "empty":
            LOGGER.warning("%s Returning empty string.", message)
            return ""
        if on_error == "keep_original":
            LOGGER.warning("%s Passing original value.", message)
            return clean_text(value)
        raise ValueError(message)

    return parsed.strftime(output_format)


def get_first_non_empty_value(row: pd.Series, actual_columns: List[str]) -> Tuple[str, Optional[str]]:
    for col in actual_columns:
        val = row.get(col, "")
        if not is_nullish(val):
            return val, col
    return "", None


def build_field_value(row: pd.Series, row_number: int, field: str, mapping_obj: Optional[Dict[str, Any]], actual_columns: List[str], config: ConverterConfig) -> str:
    raw, used_col = get_first_non_empty_value(row, actual_columns)

    if is_nullish(raw) and mapping_obj:
        fallback = mapping_obj.get("fallback")
        if fallback == "row_number":
            prefix = mapping_obj.get("prefix", "ROW-")
            return f"{prefix}{row_number}"
        if fallback == "uuid":
            prefix = mapping_obj.get("prefix", "")
            return f"{prefix}{uuid.uuid4()}"

    try:
        return normalize_date_value(raw, field, config)
    except ValueError as exc:
        column_info = used_col or ", ".join(actual_columns) or "(no column)"
        raise ValueError(
            f"{exc} Source spreadsheet row: {row_number}; source column: {column_info}."
        ) from exc


TIMESTAMP_PATTERNS = [
    re.compile(
        r"(?P<ts>\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?)\s*-\s*(?P<rest>.*?)(?=\n\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*-|\Z)",
        re.S,
    ),
    re.compile(
        r"(?P<ts>\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?)\s*-\s*(?P<rest>.*?)(?=\n\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?\s*-|\Z)",
        re.S,
    ),
]


def parse_text_timestamp(ts_text: str, output_format: str, input_order: str) -> Tuple[Optional[dt.datetime], str]:
    dummy_cfg = ConverterConfig(
        sheet_name=None, required_fields={}, optional_fields={}, comments_worknotes={},
        static_fields={}, date_fields={"input_order": input_order, "output_format": output_format},
        deduplicate_by_inc_number=False, deduplicate_keep_latest_by=None,
        output_top_level_key="result", skip_empty_rows=True, empty_row_required_any_of=[]
    )
    parsed = parse_date_from_string(ts_text, "inc_comments_and_work_notes", dummy_cfg)
    if parsed:
        return parsed, parsed.strftime(output_format)
    return None, ts_text


def split_text_entries(text: str, label: str, date_output_format: str, input_order: str) -> List[Dict[str, Any]]:
    text = clean_text(text)
    if not text:
        return []

    for pattern in TIMESTAMP_PATTERNS:
        matches = list(pattern.finditer(text))
        if matches:
            entries = []
            for i, match in enumerate(matches):
                parsed, ts_out = parse_text_timestamp(match.group("ts"), date_output_format, input_order)
                body = f"{ts_out} - {match.group('rest').strip()}"
                entries.append({"timestamp": parsed, "text": f"[{label}]\n{body}".strip(), "raw_order": i})
            return entries

    return [{"timestamp": None, "text": f"[{label}]\n{text}".strip(), "raw_order": 0}]


def get_text_field_specs(config: ConverterConfig) -> List[Dict[str, Any]]:
    cw = config.comments_worknotes
    specs = cw.get("text_fields")

    if isinstance(specs, list):
        normalized = []
        for item in specs:
            if isinstance(item, str):
                normalized.append({"column": item, "label": item})
            elif isinstance(item, dict) and item.get("column"):
                normalized.append({"column": item["column"], "label": item.get("label") or item["column"]})
        return normalized

    # legacy support
    fallback = []
    if cw.get("work_notes_column"):
        fallback.append({"column": cw["work_notes_column"], "label": "Work notes"})
    if cw.get("comments_column"):
        fallback.append({"column": cw["comments_column"], "label": "Comments"})
    return fallback


def build_comments_worknotes(row: pd.Series, config: ConverterConfig, lookup: Dict[str, str]) -> str:
    cw = config.comments_worknotes
    output_format = cw.get("date_output_format", "%d/%m/%Y %H:%M:%S")
    input_order = cw.get("timestamp_input_order", (config.date_fields or {}).get("input_order", "dmy"))
    sort_chronologically = bool(cw.get("sort_chronologically", True))

    all_entries = []
    global_order = 0

    for spec in get_text_field_specs(config):
        columns, _ = resolve_columns(spec["column"], lookup, f"comments_worknotes:{spec['column']}", required=False)
        if not columns:
            continue
        raw, used_col = get_first_non_empty_value(row, columns)
        if is_nullish(raw):
            continue
        entries = split_text_entries(clean_text(raw), spec.get("label") or used_col or spec["column"], output_format, input_order)
        for e in entries:
            e["global_order"] = global_order
            global_order += 1
            all_entries.append(e)

    if not all_entries:
        return ""

    if sort_chronologically:
        all_entries.sort(key=lambda e: (e["timestamp"] is None, e["timestamp"] or dt.datetime.max, e["global_order"]))

    return "\n\n".join(e["text"] for e in all_entries)


def validate_mapping(config: ConverterConfig, lookup: Dict[str, str]) -> Tuple[Dict[str, Tuple[List[str], Optional[Dict[str, Any]]]], Dict[str, Tuple[List[str], Optional[Dict[str, Any]]]]]:
    LOGGER.info("Validating column mapping...")

    required = {}
    optional = {}

    for field, mapping in config.required_fields.items():
        cols, obj = resolve_columns(mapping, lookup, field, required=True)
        if cols or (obj and obj.get("fallback")):
            required[field] = (cols, obj)

    for field, mapping in config.optional_fields.items():
        cols, obj = resolve_columns(mapping, lookup, field, required=False)
        if cols or (obj and obj.get("fallback")):
            optional[field] = (cols, obj)

    missing_required = [f for f in config.required_fields if f not in required]
    if missing_required:
        raise ValueError(
            "Missing required mapped columns/fallbacks for: "
            + ", ".join(missing_required)
            + "\nAvailable columns:\n  "
            + "\n  ".join(lookup.values())
        )

    LOGGER.info("Required fields configured: %s", ", ".join(required.keys()) or "(none)")
    LOGGER.info("Optional fields configured: %s", ", ".join(optional.keys()) or "(none)")
    return required, optional


def row_is_empty(ticket: Dict[str, Any], config: ConverterConfig) -> bool:
    if not config.skip_empty_rows:
        return False
    for key in config.empty_row_required_any_of:
        if clean_text(ticket.get(key, "")):
            return False
    return True


def convert_dataframe(df: pd.DataFrame, config: ConverterConfig, input_path: Path) -> Dict[str, Any]:
    lookup = build_column_lookup(df.columns)

    LOGGER.debug("Spreadsheet columns:")
    for c in df.columns:
        LOGGER.debug("  %r -> %s", c, normalize_name(c))

    required, optional = validate_mapping(config, lookup)

    batch_id = str(uuid.uuid4())
    records = []
    skipped_empty = 0
    empty_required_counts = {field: 0 for field in required.keys()}

    for idx, row in df.iterrows():
        excel_row_number = idx + 2  # assuming row 1 is header
        ticket = {}

        for field, (cols, obj) in required.items():
            value = build_field_value(row, excel_row_number, field, obj, cols, config)
            ticket[field] = value
            if value == "":
                empty_required_counts[field] += 1

        for field, (cols, obj) in optional.items():
            value = build_field_value(row, excel_row_number, field, obj, cols, config)
            if value != "":
                ticket[field] = value

        comments_field = config.comments_worknotes.get("output_field", "inc_comments_and_work_notes")
        comments = build_comments_worknotes(row, config, lookup)
        if comments:
            ticket[comments_field] = comments

        for field, static_value in config.static_fields.items():
            val = coerce_runtime_static(static_value, input_path, batch_id)
            if val is not None and clean_text(val) != "":
                ticket[field] = val

        if row_is_empty(ticket, config):
            skipped_empty += 1
            continue

        records.append(ticket)

    for field, count in empty_required_counts.items():
        if count:
            LOGGER.warning("Required field '%s' is empty in %s converted row(s). Consider fallback columns or fallback='row_number'.", field, count)

    LOGGER.info("Converted records: %s", len(records))
    if skipped_empty:
        LOGGER.info("Skipped empty rows: %s", skipped_empty)

    if config.deduplicate_by_inc_number:
        before = len(records)
        records = deduplicate_records(records, config)
        after = len(records)
        if before != after:
            LOGGER.info("Deduplicated by inc_number: %s -> %s", before, after)

    return {config.output_top_level_key: records}


def parse_sort_date(value: str, config: ConverterConfig) -> dt.datetime:
    parsed = parse_date_from_string(clean_text(value), "inc_sys_updated_on", config)
    return parsed or dt.datetime.min


def deduplicate_records(records: List[Dict[str, Any]], config: ConverterConfig) -> List[Dict[str, Any]]:
    key_field = "inc_number"
    latest_field = config.deduplicate_keep_latest_by
    deduped = {}
    passthrough = []

    for record in records:
        key = clean_text(record.get(key_field, ""))
        if not key:
            passthrough.append(record)
            continue

        if key not in deduped:
            deduped[key] = record
            continue

        if latest_field:
            old_dt = parse_sort_date(deduped[key].get(latest_field, ""), config)
            new_dt = parse_sort_date(record.get(latest_field, ""), config)
            if new_dt >= old_dt:
                deduped[key] = record
        else:
            deduped[key] = record

    return passthrough + list(deduped.values())


def write_json(data: Dict[str, Any], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    count = len(next(iter(data.values()))) if data else 0
    LOGGER.info("Output written: %s", output_path)
    LOGGER.info("Output records: %s", count)


def configure_logging(level: str) -> None:
    logging.basicConfig(level=getattr(logging, level.upper(), logging.INFO), format="%(levelname)s | %(message)s")


def parse_sheet_arg(value: Optional[str]) -> Optional[Any]:
    if value is None:
        return None
    value = value.strip()
    return int(value) if re.fullmatch(r"\d+", value) else value


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Convert Excel/CSV incident exports to Incident Whisperer JSON.")
    p.add_argument("--input", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--config", required=True)
    p.add_argument("--sheet", required=False)
    p.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    configure_logging(args.log_level)

    input_path = Path(args.input)
    output_path = Path(args.output)
    config_path = Path(args.config)

    try:
        LOGGER.info("Starting conversion.")
        LOGGER.info("Input: %s", input_path)
        LOGGER.info("Config: %s", config_path)
        LOGGER.info("Output: %s", output_path)

        config = load_config(config_path)
        selected_sheet = parse_sheet_arg(args.sheet) if args.sheet is not None else config.sheet_name

        df, _ = read_spreadsheet(input_path, selected_sheet)
        result = convert_dataframe(df, config, input_path)
        write_json(result, output_path)

        LOGGER.info("Conversion completed successfully.")
        return 0

    except Exception as exc:
        LOGGER.error("Conversion failed: %s", exc)
        if args.log_level == "DEBUG":
            LOGGER.exception("Detailed traceback:")
        else:
            LOGGER.error("Run with --log-level DEBUG for full traceback.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
