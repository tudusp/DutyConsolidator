"""
Configuration for Duty Chart Consolidator
-----------------------------------------
Supports two formats:
  A) Summary format (like SP-2026-END-SEM.xlsx)  → one row per faculty with Total Duties + Dates
  B) Detailed format (one row per duty assignment)
"""

# ============================================================
# 1. COLUMN MAPPING
# ============================================================

# --- Format A: Summary-style (your current files) ---
SUMMARY_COLUMNS = {
    "teacher": "Faculty",
    "total_duties": "Total Duties",
    "first_half_duties": "First Half Duties",
    "second_half_duties": "Second Half Duties",
    "first_half_dates": "First Half Dates",
    "second_half_dates": "Second Half Dates",
    "phone": "Phone No",
    "email": "Email ID",
}

# --- Format B: Detailed (one row per duty) ---
DETAILED_COLUMNS = {
    "teacher": "Teacher Name",
    "date": "Date",
    "session": "Session",
    "room": "Room",
    "subject": "Subject",
    "role": "Duty Type",
}

# ============================================================
# 2. FILE NAME PARSING
# ============================================================
# Examples that work:
#   SP-2026-END-SEM.xlsx   → Period=SP-2026, ExamType=End
#   SP-2026-MID-SEM.xlsx   → Period=SP-2026, ExamType=Mid
#   Sem3_Mid_2025.xlsx     → Period=Sem3,    ExamType=Mid
#   Sem5_End_Oct2025.xlsx  → Period=Sem5,    ExamType=End

FILENAME_PATTERN = r"(?i)(?:((?:SP|MO)[-\s]?\d{4})|(Sem[-\s]?(?:\d+|[IVX]+))).*?(ENDSEM|MID|END|FINAL)"

DEFAULT_PERIOD = "Unknown"
DEFAULT_EXAM_TYPE = "Unknown"

# ============================================================
# 3. DUTY HOURS
# ============================================================
DUTY_HOURS = {
    "Mid": 2,
    "End": 3,
    "Unknown": 2,
}

# ============================================================
# 4. CLEANING
# ============================================================
IGNORE_NAMES = ["", "nan", "none", "n/a", "tbd", "vacant", "total", "grand total"]
NORMALIZE_NAMES = True

SHEET_NAME = None   # None = first sheet
