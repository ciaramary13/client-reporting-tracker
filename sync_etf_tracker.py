"""
ETF Client Reporting Tracker - Auto Sync Script
------------------------------------------------
- Reads the latest export from BNY ETFMetricsUI
- Merges it into the master tracker
- Preserves the "API Identifier" column (manually maintained)
- Adds new clients automatically with a blank API Identifier
- Saves a timestamped backup before overwriting

Paths (update for each machine):
  EXPORT_PATH  = where the downloaded export lands
  TRACKER_PATH = the master ETF Client Reporting Tracker
"""

import pandas as pd
import openpyxl
import shutil
import os
import sys
import traceback
from datetime import datetime

# ─── CONFIG ────────────────────────────────────────────────────────────────────

EXPORT_PATH = r"C:\Users\ciara\Downloads\UserExecutionDetailsByProduct.xlsx"

TRACKER_PATH = (
    r"C:\Users\ciara\OneDrive - Fordham University\internship-projects"
    r"\client-reporting-tracker\ETF_Client_Reporting_Tracker_DUMMY.xlsx"
)

# Where to write the run log (one line per run, kept indefinitely)
LOG_PATH = r"C:\ETF_Sync\sync_log.txt"

# Whether to show a Windows toast notification on completion / failure.
# Requires: pip install win10toast-persist  (falls back to log-only if missing)
SHOW_NOTIFICATIONS = True

# Sheet names (adjust if different)
EXPORT_SHEET  = "Sheet1"       # sheet in the BNY export
TRACKER_SHEET = "Execution by Product"  # sheet in the master tracker

# Columns used together to uniquely identify a row (composite key).
# Client Name alone isn't unique -- the same client can have multiple rows
# for different products AND different reports (e.g. Voya/SSE-WDA/"Basket
# View Enhanced" vs Voya/SSE-WDA/"Order Detail Standard..." vs
# Voya/SSE-Standard/<blank>). Using all three together is the only way to
# uniquely identify a row.
MERGE_KEY = "Client Name"            # kept for messages/display purposes
KEY_COLS  = ["Client Name", "Product Type", "Report Name"]

# Columns that live ONLY in the tracker and must never be overwritten
PROTECTED_COLS = ["API Identifier"]

# Month abbreviations used to detect which export columns are "monthly data"
# columns (as opposed to static columns like Total Count or Product Type).
# These are ADDITIVE ONLY — a month column already in the tracker is never
# removed just because it dropped off the site's rolling export window.
MONTH_NAMES = {
    "jan", "feb", "mar", "apr", "may", "jun",
    "jul", "aug", "sep", "oct", "nov", "dec",
}

def is_month_col(col_name: str) -> bool:
    """A column counts as a 'month' column if its name starts with a
    3-letter month abbreviation (handles 'Jun', 'Jun 2026', 'Jun-26', etc.)."""
    first_word = str(col_name).strip().lower()[:3]
    return first_word in MONTH_NAMES


def has_year_suffix(col_name: str) -> bool:
    """True if the column name already ends with a 4-digit year, e.g.
    'Jun 2026' -- used to avoid double-stamping a year that's already there."""
    parts = str(col_name).strip().split()
    return len(parts) > 1 and parts[-1].isdigit() and len(parts[-1]) == 4


_MONTH_ORDER = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"]
)}

def month_sort_key(col_name: str):
    """Sort key for chronological ordering, newest first. Handles disambiguated
    names like 'May 2026' (year present) and plain 'May' (year unknown --
    sorts after same-month-with-year, oldest-last as a safe fallback)."""
    text = str(col_name).strip()
    abbr = text[:3].lower()
    month_num = _MONTH_ORDER.get(abbr, -1)
    parts = text.split()
    year = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
    return (-year, -month_num)


def infer_export_month_years(month_cols: list, as_of: datetime = None) -> dict:
    if as_of is None:
        as_of = datetime.now()

    year_map = {}
    year, month = as_of.year, as_of.month
    y, m = year, month
    for col in month_cols:
        abbr = str(col).strip().lower()[:3]
        target_month = _MONTH_ORDER.get(abbr)
        if target_month is None:
            continue
        # Walk backwards from current position until the abbreviation matches.
        # Starting the walk from wherever the previous column left off (rather
        # than resetting to "now" each time) keeps duplicate-named columns in
        # the export (which legitimately shouldn't happen, but just in case)
        # from collapsing onto the same year.
        for _ in range(13):
            if m - 1 == target_month:  # _MONTH_ORDER is 0-indexed (jan=0)
                year_map[col] = y
                m -= 1
                if m == 0:
                    m = 12
                    y -= 1
                break
            m -= 1
            if m == 0:
                m = 12
                y -= 1
    return year_map

# ─── NOTIFICATIONS / LOGGING ──────────────────────────────────────────────────

def log_line(message: str):
    """Append a single timestamped line to the log file (creates folder if needed)."""
    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {message}\n")


def notify(title: str, message: str):
    """Show a Windows toast notification. Silently does nothing if the
    optional package isn't installed — logging still happens regardless."""
    if not SHOW_NOTIFICATIONS:
        return
    try:
        from win10toast_persist import ToastNotifier
        toaster = ToastNotifier()
        toaster.show_toast(title, message, duration=10, threaded=True)
    except ImportError:
        # Notifications are a nice-to-have, not a hard requirement
        pass
    except Exception:
        # Never let a notification failure crash the actual sync
        pass

# ─── HELPERS ───────────────────────────────────────────────────────────────────

def backup_tracker(tracker_path: str) -> str:
    """Save a timestamped copy of the tracker before modifying it."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = os.path.join(os.path.dirname(tracker_path), "backups")
    os.makedirs(backup_dir, exist_ok=True)
    backup_path = os.path.join(backup_dir, f"ETF_Tracker_backup_{ts}.xlsx")
    shutil.copy2(tracker_path, backup_path)
    print(f"  ✅ Backup saved: {backup_path}")
    return backup_path


def load_sheet(path: str, sheet: str):
    """Load an Excel sheet, stripping leading/trailing whitespace from headers.

    Handles a tracker layout like:
        Row 1: title text (e.g. "List of User Executions ... as of May 2026")
        Row 2: year banner -- years like 2026 / 2025 spanning groups of month columns
        Row 3: real column headers (No., Client Name, ..., May, Apr, ..., Jun, May)
        Row 4+: data

    Because the same month abbreviation (e.g. "May") can appear twice for two
    different years, this reads the year-banner row directly above the header
    row and stamps each ambiguous month column with its year, producing
    unambiguous labels like "May 2026" / "May 2025".

    Returns (DataFrame, header_row_index) -- header_row_index is the
    0-indexed row number (pandas-style) where the real column headers were
    found, so the caller can write data back starting at the right place
    without disturbing the title/banner rows above it.
    """
    raw = pd.read_excel(path, sheet_name=sheet, header=None)

    # Find the real header row (first row that contains the Client Name column)
    header_row = None
    for i, row in raw.iterrows():
        if any(str(v).strip() == MERGE_KEY for v in row.values):
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            f"Could not find a '{MERGE_KEY}' column in {path} / {sheet}. "
            "Check the sheet name and column header."
        )

    raw_headers = raw.iloc[header_row].astype(str).str.strip()

    # If there's a year-banner row directly above the header row, use it to
    # disambiguate repeated month names. Forward-fill handles merged-cell-style
    # banners where the year only appears in the first column of its group.
    # Columns that already carry a 4-digit year (e.g. tracker headers already
    # saved as "Jun 2026") are left untouched to avoid double-stamping
    # ("Jun 2026 2026").
    final_headers = list(raw_headers)
    if header_row > 0:
        year_row = raw.iloc[header_row - 1]
        year_ffilled = year_row.ffill()
        for idx, col_name in enumerate(raw_headers):
            if is_month_col(col_name) and not has_year_suffix(col_name):
                year_val = year_ffilled.iloc[idx] if idx < len(year_ffilled) else None
                if pd.notna(year_val) and str(year_val).strip().isdigit():
                    final_headers[idx] = f"{col_name} {str(year_val).strip()}"

    df = raw.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = final_headers
    df.columns.name = None
    return df, header_row


# ─── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    print("\n🔄  ETF Tracker Sync — starting...\n")

    # 1. Sanity checks
    for label, path in [("Export", EXPORT_PATH), ("Tracker", TRACKER_PATH)]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"{label} file not found:\n  {path}")

    # 2. Backup tracker
    print("📦  Backing up tracker...")
    backup_tracker(TRACKER_PATH)

    # 3. Load both files
    print("📂  Loading export...")
    export_df, export_header_row = load_sheet(EXPORT_PATH, EXPORT_SHEET)

    print("📂  Loading tracker...")
    tracker_df, tracker_header_row = load_sheet(TRACKER_PATH, TRACKER_SHEET)

    # 3b. Fallback: if the export has NO year banner of its own (flat rolling
    # window with bare month names like "Jun", "May"...), infer each one's
    # year by counting backwards from today. If load_sheet already
    # disambiguated these using the export's own banner row, they'll already
    # have a year suffix and are skipped here.
    export_month_cols_raw = [c for c in export_df.columns if is_month_col(c) and not has_year_suffix(c)]
    if export_month_cols_raw:
        export_year_map = infer_export_month_years(export_month_cols_raw)
        export_df = export_df.rename(columns={
            c: f"{c} {export_year_map[c]}" for c in export_month_cols_raw if c in export_year_map
        })

    # 4. Identify site columns vs protected columns
    #    Site columns = everything in export except the key columns
    site_cols  = [c for c in export_df.columns if c not in KEY_COLS and c.strip()]
    month_cols_in_export = [c for c in site_cols if is_month_col(c)]
    static_cols_in_export = [c for c in site_cols if not is_month_col(c)]

    # Month columns the tracker already has but export no longer shows
    # (e.g. rolling 10-month export window) -- these must be preserved untouched
    tracker_month_cols = [c for c in tracker_df.columns if is_month_col(c)]
    months_to_preserve = [c for c in tracker_month_cols if c not in month_cols_in_export]
    if months_to_preserve:
        print(f"  🗓️  Preserving {len(months_to_preserve)} month column(s) no longer in export: "
              f"{', '.join(months_to_preserve)}")

    new_month_cols = [c for c in month_cols_in_export if c not in tracker_df.columns]
    if new_month_cols:
        print(f"  🗓️  Adding new month column(s): {', '.join(new_month_cols)}")

    # Make sure protected cols exist in tracker (add if missing)
    for col in PROTECTED_COLS:
        if col not in tracker_df.columns:
            tracker_df[col] = ""
            print(f"  ➕ Added missing protected column: '{col}'")

    # 4b. Make sure both files actually have the composite key columns
    for col in KEY_COLS:
        if col not in export_df.columns:
            raise ValueError(f"Export is missing expected column '{col}'")
        if col not in tracker_df.columns:
            raise ValueError(f"Tracker is missing expected column '{col}'")

    # Build a composite key (Client Name + Product Type + Report Name) so rows
    # for the same client but different products/reports are treated as distinct.
    # NOTE: we use a proper pandas MultiIndex (not a plain Index of tuples) --
    # indexing a plain Index with a tuple makes pandas treat it as
    # (row_selector, col_selector) instead of a single key, which raises a
    # confusing KeyError. A MultiIndex avoids that ambiguity entirely.
    def make_key(row):
        return tuple(str(row[c]).strip() for c in KEY_COLS)

    tracker_df["_key"] = tracker_df.apply(make_key, axis=1)

    # 5. True duplicates now mean same client AND same report name twice —
    #    a real data entry mistake worth flagging, not an expected pattern
    dup_mask = tracker_df["_key"].duplicated(keep=False)
    duplicate_keys = sorted(set(tracker_df.loc[dup_mask, "_key"]))
    if duplicate_keys:
        print(f"  ⚠️  Found exact duplicate (Client Name + Product Type + Report Name) rows in tracker — "
              f"keeping the FIRST occurrence of each, extras left untouched at the bottom:")
        for k in duplicate_keys:
            print(f"       – {k[0]} / {k[1]} / {k[2]}")

    first_occurrence = ~tracker_df["_key"].duplicated(keep="first")
    tracker_subset = tracker_df[first_occurrence]
    tracker_lookup = tracker_subset.set_index(
        pd.MultiIndex.from_tuples(tracker_subset["_key"], names=KEY_COLS)
    ).drop(columns=["_key"])

    # Keep any leftover true-duplicate rows as-is so no data is silently dropped
    leftover_dupe_rows = tracker_df[~first_occurrence].drop(columns=["_key"])

    # 6. Merge: update existing rows, append new ones
    updated_rows = []
    new_clients  = []
    matched_keys = set()

    for _, exp_row in export_df.iterrows():
        client = str(exp_row[MERGE_KEY]).strip()
        if not client or client.lower() == "nan":
            continue

        key = make_key(exp_row)

        if key in tracker_lookup.index:
            # Existing client+report — update site columns, keep protected cols
            row = tracker_lookup.loc[key].copy()
            for col in site_cols:
                if col in exp_row.index:
                    row[col] = exp_row[col]
            updated_rows.append(row)
            matched_keys.add(key)
        else:
            # New client+report combo — add with blank protected columns
            # and blank values for any preserved month columns the export doesn't cover
            new_row = exp_row.copy()
            for col in PROTECTED_COLS:
                new_row[col] = ""
            for col in months_to_preserve:
                new_row[col] = ""
            updated_rows.append(new_row)
            new_clients.append(f"{key[0]} / {key[1]} / {key[2]}")

    # Any tracker rows NOT present in today's export are kept as-is (not deleted)
    untouched_keys = [k for k in tracker_lookup.index if k not in matched_keys]
    for k in untouched_keys:
        updated_rows.append(tracker_lookup.loc[k])

    # 7. Rebuild dataframe
    result_df = pd.DataFrame(updated_rows).reset_index(drop=True)

    # Re-attach any leftover duplicate rows so nothing is silently dropped
    if not leftover_dupe_rows.empty:
        result_df = pd.concat([result_df, leftover_dupe_rows], ignore_index=True)

    # Re-order columns: preserve the TRACKER's original column order exactly
    # as-is (so API Identifier, or any other custom column, stays wherever the
    # human placed it) -- we only need to insert any brand-new month columns
    # that didn't exist in the tracker before today's sync.
    original_tracker_order = [c for c in tracker_df.columns if c != "_key"]
    all_cols = list(result_df.columns)

    if new_month_cols:
        # Insert new months right after the last existing month column (or
        # after Total Count / static cols if the tracker had no months yet),
        # sorted so the newest month ends up first among the months.
        existing_months_sorted = sorted(tracker_month_cols, key=month_sort_key)
        new_months_sorted = sorted(new_month_cols, key=month_sort_key)

        if existing_months_sorted:
            insert_after = existing_months_sorted[0]  # newest existing month
            insert_at = original_tracker_order.index(insert_after)
        else:
            # No months yet -- insert right after Total Count if present, else at the end
            insert_at = (
                original_tracker_order.index("Total Count") + 1
                if "Total Count" in original_tracker_order
                else len(original_tracker_order)
            )

        ordered = (
            original_tracker_order[:insert_at]
            + new_months_sorted
            + original_tracker_order[insert_at:]
        )
    else:
        ordered = original_tracker_order

    # Anything in result_df not yet accounted for (shouldn't normally happen,
    # but keeps us safe rather than silently dropping a column)
    ordered += [c for c in all_cols if c not in ordered]
    ordered = list(dict.fromkeys(ordered))   # deduplicate, preserve order
    result_df = result_df[[c for c in ordered if c in result_df.columns]]

    # 8. Write back to tracker IN PLACE -- preserves the title row, year
    # banner, merged cells, fonts, and fills. We deliberately do NOT use
    # pandas.ExcelWriter(if_sheet_exists="replace") here: that approach
    # deletes the whole sheet and rewrites it from a bare DataFrame, which
    # silently destroys any formatting/title rows/merged cells above the
    # header. Instead we open the workbook with openpyxl, clear only the
    # data rows below the existing header, and write fresh values into that
    # same region -- header row and everything above it is never touched.
    print("💾  Writing updated tracker...")
    wb = openpyxl.load_workbook(TRACKER_PATH)
    ws = wb[TRACKER_SHEET]

    header_excel_row = tracker_header_row + 1          # openpyxl is 1-indexed
    first_data_row    = header_excel_row + 1
    n_new_rows         = len(result_df)
    n_cols             = len(result_df.columns)

    # 8a. If new month columns were inserted, the year-banner row (directly
    # above the header row -- e.g. row 1 with "2026" / "2025" sitting over
    # the FIRST month column of each year) needs its merge range widened so
    # the new column is still visually grouped under the correct year.
    # Without this, inserting a column shifts everything right but the old
    # merge boundary stays put, silently mislabeling a month's year.
    if new_month_cols and header_excel_row > 1:
        year_banner_row = header_excel_row - 1

        # Map: which year does each column (by its NEW position) belong to,
        # based on the column's own header text (e.g. "Jan 2026" -> 2026)
        col_year = {}
        for col_idx, col_name in enumerate(result_df.columns, start=1):
            if is_month_col(col_name):
                parts = str(col_name).split()
                if len(parts) > 1 and parts[-1].isdigit():
                    col_year[col_idx] = int(parts[-1])

        if col_year:
            # Find each existing merge range in the year-banner row and
            # widen/shift it to exactly span the columns matching its year
            existing_year_merges = [
                rng for rng in list(ws.merged_cells.ranges)
                if rng.min_row == year_banner_row and rng.max_row == year_banner_row
            ]
            for rng in existing_year_merges:
                year_value = ws.cell(row=year_banner_row, column=rng.min_col).value
                if not isinstance(year_value, (int, float)):
                    continue
                year_value = int(year_value)
                matching_cols = [c for c, y in col_year.items() if y == year_value]
                if not matching_cols:
                    continue
                new_min, new_max = min(matching_cols), max(matching_cols)
                if (new_min, new_max) != (rng.min_col, rng.max_col):
                    ws.unmerge_cells(start_row=year_banner_row, start_column=rng.min_col,
                                      end_row=year_banner_row, end_column=rng.max_col)
                    # Clear the old anchor cell value if it's no longer the merge start
                    if new_min != rng.min_col:
                        old_val = ws.cell(row=year_banner_row, column=rng.min_col).value
                        ws.cell(row=year_banner_row, column=rng.min_col).value = None
                        ws.cell(row=year_banner_row, column=new_min).value = old_val
                    ws.merge_cells(start_row=year_banner_row, start_column=new_min,
                                    end_row=year_banner_row, end_column=new_max)
                    print(f"  🔧 Adjusted {year_value} year-banner merge to columns {new_min}-{new_max}")

    # Clear any existing data rows below the header (in case the new data
    # has fewer rows than before, so no stale rows linger at the bottom)
    if ws.max_row >= first_data_row:
        for row in ws.iter_rows(min_row=first_data_row, max_row=ws.max_row,
                                 min_col=1, max_col=max(ws.max_column, n_cols)):
            for cell in row:
                cell.value = None

    # Write headers (in case a new month column was added) and data.
    # Month-column headers are written WITHOUT their year suffix (the
    # tracker's convention keeps bare month names like "Jun" in row 2 and
    # relies on the year-banner row above for disambiguation).
    for col_idx, col_name in enumerate(result_df.columns, start=1):
        header_to_write = col_name
        if is_month_col(col_name) and has_year_suffix(col_name):
            header_to_write = " ".join(str(col_name).split()[:-1])
        ws.cell(row=header_excel_row, column=col_idx, value=header_to_write)

    for row_idx, (_, data_row) in enumerate(result_df.iterrows(), start=first_data_row):
        for col_idx, col_name in enumerate(result_df.columns, start=1):
            val = data_row[col_name]
            ws.cell(row=row_idx, column=col_idx, value=None if pd.isna(val) else val)

    wb.save(TRACKER_PATH)

    # 9. Summary
    summary = {
        "updated": len(matched_keys),
        "untouched": len(untouched_keys),
        "new": new_clients,
        "duplicates": duplicate_keys,
        "new_months": new_month_cols,
        "preserved_months": months_to_preserve,
    }

    print(f"\n✅  Sync complete!")
    print(f"   • {summary['updated']} existing rows updated")
    print(f"   • {summary['untouched']} rows left untouched (not in today's export)")
    if summary["new_months"]:
        print(f"   • {len(summary['new_months'])} new month column(s) added: {', '.join(summary['new_months'])}")
    if summary["preserved_months"]:
        print(f"   • {len(summary['preserved_months'])} older month column(s) preserved (not in latest export): "
              f"{', '.join(summary['preserved_months'])}")
    if summary["new"]:
        print(f"   • {len(summary['new'])} new row(s) added:")
        for c in summary["new"]:
            print(f"       – {c}  ← fill in API Identifier manually")
    else:
        print("   • No new client/report rows detected")
    if summary["duplicates"]:
        print(f"   • ⚠️  {len(summary['duplicates'])} exact duplicate row(s) detected — please clean up tracker when convenient")
    print()

    return summary


if __name__ == "__main__":
    try:
        result = main()

        # Build a short, human-readable status for the log + notification
        parts = [f"{result['updated']} updated"]
        if result["new"]:
            parts.append(f"{len(result['new'])} new client(s)")
        if result["duplicates"]:
            parts.append(f"{len(result['duplicates'])} duplicate(s) found")
        status_msg = ", ".join(parts)

        log_line(f"SUCCESS — {status_msg}")
        notify("ETF Tracker Sync ✅", status_msg)

    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"
        log_line(f"FAILED — {error_msg}")
        log_line(traceback.format_exc())
        notify("ETF Tracker Sync ❌ FAILED", f"{error_msg}\nCheck {LOG_PATH}")
        print(f"\n❌  Error: {e}")
        sys.exit(1)
