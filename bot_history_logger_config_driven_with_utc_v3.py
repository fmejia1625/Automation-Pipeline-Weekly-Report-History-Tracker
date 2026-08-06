"""
Config-driven bot history snapshot logger with local + UTC timestamps.

Purpose:
- Reads the current bot pipeline data from Data_Set_Bots.
- Compares each bot against the latest prior snapshot in Bot_History_Data.
- Appends a new snapshot row for every current bot into Excel table "Table2".
- Highlights newly appended rows where a change was detected.
- Stores both local Eastern timestamp and UTC timestamp for distributed teams.

How to add a new tracked field:
1. Add the column to "columns" in bot_history_config_with_utc.json.
2. Add it to "compare_columns" if changes should be tracked.
3. Add it to "date_columns" only if it should be normalized as YYYY-MM.
4. Add a source-to-history mapping in "rename_map" if the source column name is different.

Notes:
- xlwings requires Excel/COM, so run this locally on your Windows machine.
- Keep bot_history_config_with_utc.json in the same folder as this script, or update CONFIG_PATH.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json
import re

import pandas as pd
import pytz
import xlwings as xw


# =========================================================
# CONFIG LOADING
# =========================================================
CONFIG_PATH = Path(__file__).with_name("bot_history_config_with_utc.json")


def load_config(config_path: Path) -> dict:
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    with config_path.open("r", encoding="utf-8") as f:
        cfg = json.load(f)

    required_keys = [
        "file_path",
        "source_sheet",
        "history_sheet",
        "history_table",
        "timezone",
        "columns",
        "compare_columns",
        "date_columns",
        "rename_map",
        "fallback_columns",
        "highlight_color_changed",
    ]

    missing = [key for key in required_keys if key not in cfg]
    if missing:
        raise KeyError(f"Missing required config keys: {missing}")

    return cfg


CFG = load_config(CONFIG_PATH)

FILE_PATH = Path(CFG["file_path"])
SOURCE_SHEET = CFG["source_sheet"]
HISTORY_SHEET = CFG["history_sheet"]
HISTORY_TABLE = CFG["history_table"]
LOCAL_TZ = pytz.timezone(CFG["timezone"])

COLUMNS = CFG["columns"]
COMPARE_COLS = CFG["compare_columns"]
DATE_COLS = set(CFG["date_columns"])
RENAME_MAP = CFG["rename_map"]
FALLBACK_COLUMNS = CFG["fallback_columns"]


def hex_to_excel_bgr(hex_color: str) -> int:
    """Convert #RRGGBB or RRGGBB to Excel COM BGR integer."""
    hex_color = hex_color.strip().lstrip("#")
    if len(hex_color) != 6:
        raise ValueError(f"Invalid hex color: {hex_color}")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return r + (g << 8) + (b << 16)


HIGHLIGHT_COLOR_CHANGED = hex_to_excel_bgr(CFG["highlight_color_changed"])


# =========================================================
# NORMALIZATION HELPERS
# =========================================================
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip headers and remove duplicate columns."""
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df.loc[:, ~df.columns.duplicated()]


def normalize_text(value) -> str:
    """Normalize text values for stable comparison."""
    if pd.isna(value):
        return ""

    text = str(value)

    # Normalize common invisible characters from Excel/SharePoint copy paths.
    text = (
        text.replace("\u00a0", " ")  # non-breaking space
        .replace("\u200b", "")      # zero-width space
        .replace("\ufeff", "")      # BOM marker
    )

    # Collapse repeated whitespace/newlines/tabs into single spaces.
    text = re.sub(r"\s+", " ", text).strip()

    if text.lower() in {"", "none", "nan", "null"}:
        return ""

    return text.lower()


def normalize_date(value) -> str:
    """
    Normalize date-like values to YYYY-MM for comparison.
    Non-date values fall back to normalized text.
    """
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if not text:
        return ""

    dt = pd.to_datetime(text, errors="coerce")
    if pd.isna(dt):
        return normalize_text(text)

    return dt.strftime("%Y-%m")


def normalize_bot_id(series: pd.Series) -> pd.Series:
    """Normalize bot IDs and remove Excel/pandas artifacts."""
    cleaned = (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .replace({"nan": "", "none": "", "null": ""})
    )

    # Align numeric-looking IDs read as floats (e.g., 20.0) with plain IDs (20).
    cleaned = cleaned.str.replace(r"\.0+$", "", regex=True)

    return cleaned


def normalize_dependency(series: pd.Series) -> pd.Series:
    """Standardize blank dependency values as 'None' for reporting."""
    cleaned = series.fillna("").astype(str).str.strip()
    return cleaned.mask(cleaned.eq("") | cleaned.str.lower().isin(["nan", "null", "none"]), "None")


# =========================================================
# DATA LOADING
# =========================================================
def load_data(file_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not file_path.exists():
        raise FileNotFoundError(f"Workbook not found: {file_path}")

    current = pd.read_excel(file_path, sheet_name=SOURCE_SHEET)
    history = pd.read_excel(file_path, sheet_name=HISTORY_SHEET)

    current = clean_columns(current).rename(columns=RENAME_MAP)
    history = clean_columns(history)

    # Apply fallback source columns, e.g. Dependency <- Dependency Display Pivot.
    for target_col, fallback_col in FALLBACK_COLUMNS.items():
        if target_col not in current.columns and fallback_col in current.columns:
            current[target_col] = current[fallback_col]

    for col in COLUMNS:
        if col not in current.columns:
            current[col] = ""
        if col not in history.columns:
            history[col] = ""

    current = current.loc[:, ~current.columns.duplicated()]
    history = history.loc[:, ~history.columns.duplicated()]

    current["Bot ID (S.No.)"] = normalize_bot_id(current["Bot ID (S.No.)"])
    history["Bot ID (S.No.)"] = normalize_bot_id(history["Bot ID (S.No.)"])

    # Prevent blank Excel rows from becoming fake history records.
    current = current[current["Bot ID (S.No.)"].ne("")].copy()
    history = history[history["Bot ID (S.No.)"].ne("")].copy()

    if "Dependency" in current.columns:
        current["Dependency"] = normalize_dependency(current["Dependency"])

    return current, history[COLUMNS]


# =========================================================
# TIMESTAMP PARSING
# =========================================================
def parse_history_timestamps(history: pd.DataFrame) -> pd.DataFrame:
    """
    Parse history timestamps consistently.

    Important:
    Everything is converted to timezone-aware UTC.
    This prevents pandas errors caused by mixing tz-aware UTC values
    with older tz-naive local timestamps.
    """
    history = history.copy()

    if history.empty:
        history["Snapshot Timestamp Parsed"] = pd.Series(dtype="datetime64[ns, UTC]")
        return history

    # Initialize as timezone-aware UTC so every later assignment is compatible.
    history["Snapshot Timestamp Parsed"] = pd.NaT
    history["Snapshot Timestamp Parsed"] = pd.to_datetime(
        history["Snapshot Timestamp Parsed"],
        utc=True,
        errors="coerce",
    )

    # Preferred option for distributed teams: parse UTC audit column.
    if "Snapshot Timestamp UTC" in history.columns:
        utc_text = history["Snapshot Timestamp UTC"].astype(str).str.strip()

        parsed_utc = pd.to_datetime(
            utc_text,
            format="%Y-%m-%d %H:%M:%S UTC",
            errors="coerce",
            utc=True,
        )

        missing_utc_mask = parsed_utc.isna()

        # Support ISO-like/manual UTC values.
        if missing_utc_mask.any():
            parsed_utc.loc[missing_utc_mask] = pd.to_datetime(
                utc_text[missing_utc_mask],
                errors="coerce",
                utc=True,
            )

        valid_utc_mask = parsed_utc.notna()
        history.loc[valid_utc_mask, "Snapshot Timestamp Parsed"] = parsed_utc.loc[valid_utc_mask]

    # Fallback for legacy history without usable UTC.
    missing_mask = history["Snapshot Timestamp Parsed"].isna()

    if missing_mask.any():
        timestamp_text = (
            history.loc[missing_mask, "Snapshot Date"].astype(str).str.strip()
            + " "
            + history.loc[missing_mask, "Snapshot Timestamp"].astype(str).str.strip()
        )

        # Parse legacy local Eastern timestamps.
        # localize explicitly to avoid tz-naive / tz-aware mixing.
        parsed_local = pd.to_datetime(
            timestamp_text,
            format="%m/%d/%Y %I:%M:%S %p %Z",
            errors="coerce",
        )

        # Fallback: older rows that used fixed EST text.
        still_missing = parsed_local.isna()
        if still_missing.any():
            legacy_timestamp_text = timestamp_text[still_missing].str.replace(
                " EDT",
                " EST",
                regex=False,
            )

            parsed_local.loc[still_missing] = pd.to_datetime(
                legacy_timestamp_text,
                format="%m/%d/%Y %I:%M:%S %p EST",
                errors="coerce",
            )

        # Final fallback for unusual legacy records.
        still_missing = parsed_local.isna()
        if still_missing.any():
            parsed_local.loc[still_missing] = pd.to_datetime(
                timestamp_text[still_missing],
                format="mixed",
                errors="coerce",
            )

        # Convert legacy parsed values to UTC.
        # If pandas returned tz-naive values, treat them as local Eastern time.
        if parsed_local.notna().any():
            if getattr(parsed_local.dt, "tz", None) is None:
                parsed_local = parsed_local.dt.tz_localize(
                    LOCAL_TZ,
                    ambiguous="NaT",
                    nonexistent="shift_forward",
                )
            parsed_local_utc = parsed_local.dt.tz_convert("UTC")

            history.loc[missing_mask, "Snapshot Timestamp Parsed"] = parsed_local_utc

    return history


# =========================================================
# SNAPSHOT BUILDING
# =========================================================
def build_snapshot_rows(current: pd.DataFrame, history: pd.DataFrame) -> list[list]:
    local_now = datetime.now(LOCAL_TZ)
    utc_now = local_now.astimezone(pytz.utc)

    snapshot_week = local_now.strftime("%G-W%V")
    snapshot_date = local_now.strftime("%m/%d/%Y")
    snapshot_timestamp = local_now.strftime("%I:%M:%S %p %Z")
    snapshot_timezone = str(LOCAL_TZ)
    snapshot_timestamp_utc = utc_now.strftime("%Y-%m-%d %H:%M:%S UTC")

    if not history.empty:
        history = parse_history_timestamps(history)

        # Deterministic ordering for "latest" selection:
        # 1) parsed timestamp, 2) original row order as tie-breaker.
        # This avoids picking the wrong prior row when many rows share the same second.
        history = history.reset_index(drop=True)
        history["_row_order"] = history.index
        history = history.sort_values(
            ["Snapshot Timestamp Parsed", "_row_order"],
            kind="mergesort",
            # NaT must be treated as oldest, otherwise stale/unparseable rows can be
            # selected as the "latest" and cause repeated false positives.
            na_position="first",
        )

    latest = history.drop_duplicates("Bot ID (S.No.)", keep="last")
    if "_row_order" in latest.columns:
        latest = latest.drop(columns=["_row_order"])
    history_is_empty = latest.empty
    prior_by_bot = latest.set_index("Bot ID (S.No.)").to_dict("index") if not latest.empty else {}

    output_rows: list[list] = []

    for _, row in current.iterrows():
        bot_id = row.get("Bot ID (S.No.)", "")
        old = prior_by_bot.get(bot_id, {})
        diffs: list[str] = []

        for col in COMPARE_COLS:
            new_value = row.get(col, "")
            old_value = old.get(col, "")

            if col in DATE_COLS:
                new_norm = normalize_date(new_value)
                old_norm = normalize_date(old_value)
            else:
                new_norm = normalize_text(new_value)
                old_norm = normalize_text(old_value)

            if old and new_norm != old_norm:
                diffs.append(f"{col}: '{old_norm}' → '{new_norm}'")

        if not old and history_is_empty:
            change_detected = "No"
            change_details = "Baseline snapshot"
        elif not old:
            change_detected = "Yes"
            change_details = "New bot added"
        elif diffs:
            change_detected = "Yes"
            change_details = " | ".join(diffs)
        else:
            change_detected = "No"
            change_details = "No Change Details Available"

        row_dict = {col: row.get(col, "") for col in COLUMNS}
        row_dict.update({
            "Snapshot Week": snapshot_week,
            "Snapshot Date": snapshot_date,
            "Snapshot Timestamp": snapshot_timestamp,
            "Snapshot Timezone": snapshot_timezone,
            "Snapshot Timestamp UTC": snapshot_timestamp_utc,
            "Change Detected": change_detected,
            "Change Details": change_details,
        })

        output_rows.append([row_dict.get(col, "") for col in COLUMNS])

    return output_rows


# =========================================================
# EXCEL WRITEBACK
# =========================================================
def append_rows_to_history(file_path: Path, output_rows: list[list]) -> None:
    if not output_rows:
        print("🟡 No current bots found to snapshot.")
        return

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    wb = None

    try:
        wb = app.books.open(str(file_path), update_links=False, read_only=False)

        # Retry workbook/sheet resolution once if Excel COM proxies are transiently unavailable.
        ws = wb.sheets[HISTORY_SHEET]
        if ws.api is None:
            wb.close()
            wb = app.books.open(str(file_path), update_links=False, read_only=False)
            ws = wb.sheets[HISTORY_SHEET]

        if ws.api is None:
            raise RuntimeError(
                f"Could not access worksheet COM API for '{HISTORY_SHEET}'. "
                "Close the workbook in other Excel instances and try again."
            )

        # Resolve table by name with a clear error if missing.
        table = None
        for i in range(1, ws.api.ListObjects.Count + 1):
            candidate = ws.api.ListObjects(i)
            if str(candidate.Name).strip().lower() == HISTORY_TABLE.strip().lower():
                table = candidate
                break

        if table is None:
            available = [
                str(ws.api.ListObjects(i).Name).strip()
                for i in range(1, ws.api.ListObjects.Count + 1)
            ]
            raise KeyError(
                f"History table '{HISTORY_TABLE}' was not found on sheet '{HISTORY_SHEET}'. "
                f"Available tables: {available}"
            )

        table_column_names = [
            str(table.ListColumns(i).Name).strip()
            for i in range(1, table.ListColumns.Count + 1)
        ]

        # Add any missing history columns to the Excel table automatically.
        for col in COLUMNS:
            if col not in table_column_names:
                new_col = table.ListColumns.Add()
                new_col.Name = col
                table_column_names.append(col)

        row_dicts = [dict(zip(COLUMNS, row)) for row in output_rows]
        values = [[row.get(col, "") for col in table_column_names] for row in row_dicts]

        existing_rows = table.ListRows.Count
        new_row_count = len(values)

        # Add required rows, then write values in one batch.
        for _ in range(new_row_count):
            table.ListRows.Add()

        start_row = existing_rows + 1
        end_row = existing_rows + new_row_count
        start_cell = table.ListRows(start_row).Range.Cells(1, 1)
        end_cell = table.ListRows(end_row).Range.Cells(1, len(table_column_names))
        write_range = ws.range((start_cell.Row, start_cell.Column), (end_cell.Row, end_cell.Column))
        write_range.value = values

        bot_id_col = table_column_names.index("Bot ID (S.No.)") + 1

        # Format only newly appended rows.
        for offset, row_data in enumerate(row_dicts):
            excel_row = start_row + offset
            row_range = table.ListRows(excel_row).Range

            if str(row_data.get("Change Detected", "")).strip().lower() == "yes":
                row_range.Interior.Color = HIGHLIGHT_COLOR_CHANGED
            else:
                row_range.Interior.Pattern = -4142

            row_range.Cells(1, bot_id_col).Font.Bold = True
            row_range.Cells(1, bot_id_col).HorizontalAlignment = -4108

        wb.save()

    finally:
        if wb is not None:
            wb.close()
        app.quit()


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    current, history = load_data(FILE_PATH)
    output_rows = build_snapshot_rows(current, history)

    changed_count = sum(
        1 for row in output_rows
        if str(row[COLUMNS.index("Change Detected")]).strip().lower() == "yes"
    )

    append_rows_to_history(FILE_PATH, output_rows)

    if history.empty:
        print("🟡 Baseline snapshot logged — no prior history to compare")
    elif changed_count == 0:
        print("🟡 Snapshot logged — NO CHANGES DETECTED")
    else:
        print(f"🟢 Snapshot logged — {changed_count} bots changed")


if __name__ == "__main__":
    main()
