"""
Core processing logic for Duty Chart Consolidator
Supports:
  - Summary format (Faculty + Total Duties + date lists)  ← your current files
  - Detailed format (one row per duty)
"""

import re
from pathlib import Path
from typing import List, Dict, Optional, Tuple
import pandas as pd

from config import (
    SUMMARY_COLUMNS,
    DETAILED_COLUMNS,
    FILENAME_PATTERN,
    DEFAULT_PERIOD,
    DEFAULT_EXAM_TYPE,
    SHEET_NAME,
    IGNORE_NAMES,
    NORMALIZE_NAMES,
    DUTY_HOURS,
)


def parse_filename(filename: str) -> Tuple[str, str]:
    """Extract Period (e.g. SP-2026 / Sem3) and Exam Type (Mid / End) from filename."""
    name = Path(filename).stem
    match = re.search(FILENAME_PATTERN, name)
    if match:
        period = (match.group(1) or match.group(2) or "").replace(" ", "").upper()
        # Normalize SP-2026 / MO-2026 style
        period = re.sub(r"(SP|MO)[-\s]?(\d{4})", r"\1-\2", period, flags=re.I)
        period = re.sub(r"SEM[-\s]?(\d+|[IVX]+)", r"Sem\1", period, flags=re.I)

        exam_raw = match.group(3).lower()
        if exam_raw in ("final", "end", "endsem"):
            exam_type = "End"
        elif exam_raw == "mid":
            exam_type = "Mid"
        else:
            exam_type = exam_raw
        return period, exam_type
    return DEFAULT_PERIOD, DEFAULT_EXAM_TYPE


def clean_teacher_name(name) -> Optional[str]:
    if pd.isna(name):
        return None
    name = str(name).strip()
    if not name or name.lower() in IGNORE_NAMES:
        return None
    if NORMALIZE_NAMES:
        # Keep existing capitalization mostly, just clean extra spaces
        name = " ".join(name.split())
    return name


def _split_dates(date_str) -> List[str]:
    """Split comma-separated date string into clean list."""
    if pd.isna(date_str) or str(date_str).strip().lower() in ("", "none", "nan"):
        return []
    parts = re.split(r"[,;]\s*", str(date_str).strip())
    return [p.strip() for p in parts if p.strip()]


def detect_format(df: pd.DataFrame) -> str:
    """Return 'summary' or 'detailed' based on columns present."""
    cols = {c.strip().lower() for c in df.columns}
    if "faculty" in cols and ("total duties" in cols or "first half duties" in cols):
        return "summary"
    if any(k in cols for k in ("teacher name", "teacher", "name", "faculty")):
        return "detailed"
    return "unknown"


def read_summary_format(filepath: Path, period: str, exam_type: str) -> pd.DataFrame:
    """
    Handle files like SP-2026-END-SEM.xlsx
    Expand each faculty's date lists into individual duty records.
    """
    df = pd.read_excel(filepath, sheet_name=SHEET_NAME or 0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    teacher_col = SUMMARY_COLUMNS["teacher"]
    if teacher_col not in df.columns:
        # fallback
        for c in df.columns:
            if "faculty" in c.lower() or "teacher" in c.lower() or "name" in c.lower():
                teacher_col = c
                break

    hours = DUTY_HOURS.get(exam_type, DUTY_HOURS["Unknown"])
    records = []

    for idx, row in df.iterrows():
        teacher = clean_teacher_name(row.get(teacher_col))
        if not teacher:
            continue

        # Collect dates from both halves
        fh_dates = _split_dates(row.get(SUMMARY_COLUMNS.get("first_half_dates")))
        sh_dates = _split_dates(row.get(SUMMARY_COLUMNS.get("second_half_dates")))

        # Prefer expanding dates (most accurate). Fallback to Total Duties count.
        if fh_dates or sh_dates:
            for d in fh_dates:
                records.append({
                    "Teacher": teacher,
                    "Period": period,
                    "ExamType": exam_type,
                    "Hours": hours,
                    "Date": d,
                    "Session": "First Half",
                    "SourceFile": filepath.name,
                    "Phone": row.get(SUMMARY_COLUMNS.get("phone")),
                    "Email": row.get(SUMMARY_COLUMNS.get("email")),
                })
            for d in sh_dates:
                records.append({
                    "Teacher": teacher,
                    "Period": period,
                    "ExamType": exam_type,
                    "Hours": hours,
                    "Date": d,
                    "Session": "Second Half",
                    "SourceFile": filepath.name,
                    "Phone": row.get(SUMMARY_COLUMNS.get("phone")),
                    "Email": row.get(SUMMARY_COLUMNS.get("email")),
                })
        else:
            # No dates → use Total Duties (or First+Second)
            total = row.get(SUMMARY_COLUMNS.get("total_duties"))
            if pd.isna(total):
                fh = row.get(SUMMARY_COLUMNS.get("first_half_duties"), 0) or 0
                sh = row.get(SUMMARY_COLUMNS.get("second_half_duties"), 0) or 0
                total = int(fh) + int(sh)
            total = int(total) if not pd.isna(total) else 0
            for i in range(total):
                records.append({
                    "Teacher": teacher,
                    "Period": period,
                    "ExamType": exam_type,
                    "Hours": hours,
                    "Date": None,
                    "Session": None,
                    "SourceFile": filepath.name,
                    "Phone": row.get(SUMMARY_COLUMNS.get("phone")),
                    "Email": row.get(SUMMARY_COLUMNS.get("email")),
                })

    return pd.DataFrame(records)


def read_detailed_format(filepath: Path, period: str, exam_type: str) -> pd.DataFrame:
    """Handle classic one-row-per-duty sheets."""
    df = pd.read_excel(filepath, sheet_name=SHEET_NAME or 0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    teacher_col = DETAILED_COLUMNS.get("teacher", "Teacher Name")
    if teacher_col not in df.columns:
        for c in df.columns:
            if any(k in c.lower() for k in ("teacher", "faculty", "name")):
                teacher_col = c
                break
        else:
            raise ValueError(f"No teacher column found in {filepath.name}. Columns: {list(df.columns)}")

    hours = DUTY_HOURS.get(exam_type, DUTY_HOURS["Unknown"])
    records = []

    for idx, row in df.iterrows():
        teacher = clean_teacher_name(row.get(teacher_col))
        if not teacher:
            continue
        records.append({
            "Teacher": teacher,
            "Period": period,
            "ExamType": exam_type,
            "Hours": hours,
            "Date": row.get(DETAILED_COLUMNS.get("date")),
            "Session": row.get(DETAILED_COLUMNS.get("session")),
            "Room": row.get(DETAILED_COLUMNS.get("room")),
            "Subject": row.get(DETAILED_COLUMNS.get("subject")),
            "Role": row.get(DETAILED_COLUMNS.get("role"), "Invigilator"),
            "SourceFile": filepath.name,
        })

    return pd.DataFrame(records)


def read_one_file(filepath: Path) -> pd.DataFrame:
    period, exam_type = parse_filename(filepath.name)

    # Peek at columns to detect format
    preview = pd.read_excel(filepath, sheet_name=SHEET_NAME or 0, nrows=2, engine="openpyxl")
    preview.columns = [str(c).strip() for c in preview.columns]
    fmt = detect_format(preview)

    if fmt == "summary":
        return read_summary_format(filepath, period, exam_type)
    elif fmt == "detailed":
        return read_detailed_format(filepath, period, exam_type)
    else:
        raise ValueError(
            f"Cannot detect format of {filepath.name}.\n"
            f"Columns found: {list(preview.columns)}"
        )


def process_folder(folder_path: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Process all Excel files in a folder.
    Returns: detailed, summary, ranking
    """
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(f"Folder not found: {folder_path}")

    excel_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
    excel_files = [
        f for f in excel_files
        if not f.name.startswith("~$")
        and "Consolidation_Report" not in f.name
        and "Duty_Consolidation" not in f.name
    ]

    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in {folder_path}")

    all_records = []
    errors = []

    for f in sorted(excel_files):
        try:
            df = read_one_file(f)
            all_records.append(df)
            print(f"✓ {f.name}: {len(df)} duty slots  ({df['ExamType'].iloc[0] if len(df) else '?'} @ {DUTY_HOURS.get(df['ExamType'].iloc[0], '?')} hrs)")
        except Exception as e:
            errors.append(f"✗ {f.name}: {e}")
            print(f"✗ {f.name}: {e}")

    if not all_records:
        raise ValueError("No valid duty records extracted.\n" + "\n".join(errors))

    detailed = pd.concat(all_records, ignore_index=True)

    # Combined key
    detailed["Period_Exam"] = detailed["Period"] + " - " + detailed["ExamType"]

    # ---- Summary pivot (counts) ----
    pivot_count = (
        detailed.groupby(["Teacher", "Period_Exam"])
        .size()
        .unstack(fill_value=0)
    )
    pivot_count["Total_Duties"] = pivot_count.sum(axis=1)

    # ---- Hours pivot ----
    pivot_hours = (
        detailed.groupby(["Teacher", "Period_Exam"])["Hours"]
        .sum()
        .unstack(fill_value=0)
    )
    pivot_hours["Total_Hours"] = pivot_hours.sum(axis=1)

    summary = pivot_count.copy()
    summary["Total_Hours"] = pivot_hours["Total_Hours"]
    summary = summary.sort_values("Total_Hours", ascending=False).reset_index()

    # ---- Ranking ----
    ranking = (
        detailed.groupby("Teacher")
        .agg(
            Total_Hours=("Hours", "sum"),
            Total_Duties=("Teacher", "count"),
            Mid_Duties=("ExamType", lambda x: (x == "Mid").sum()),
            End_Duties=("ExamType", lambda x: (x == "End").sum()),
            Files=("SourceFile", "nunique"),
            Periods=("Period", lambda x: ", ".join(sorted(x.unique()))),
        )
        .sort_values("Total_Hours", ascending=False)
        .reset_index()
    )

    return detailed, summary, ranking


def export_reports(
    detailed: pd.DataFrame,
    summary: pd.DataFrame,
    ranking: pd.DataFrame,
    output_path: str = "Duty_Consolidation_Report.xlsx",
):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        ranking.to_excel(writer, sheet_name="Teacher Load Ranking", index=False)
        summary.to_excel(writer, sheet_name="Breakdown by Period", index=False)
        detailed.to_excel(writer, sheet_name="Detailed Duty List", index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except Exception:
                        pass
                ws.column_dimensions[column_letter].width = min(max_length + 2, 45)

    print(f"\n✅ Report saved to: {output_path}")
    return output_path
