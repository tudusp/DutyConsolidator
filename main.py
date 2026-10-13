"""
Duty Chart Consolidator – Streamlit Web App
Run:  streamlit run main.py
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import shutil
from processor import process_folder, export_reports
import config

# ── Persistent file store ─────────────────────────────────────────────────────
SAVED_DIR = Path("saved_files")
SAVED_DIR.mkdir(exist_ok=True)


def saved_files() -> list[Path]:
    """Return sorted list of .xlsx / .xls files in the saved store."""
    return sorted(SAVED_DIR.glob("*.xls*"))


def save_uploaded(file) -> Path:
    dest = SAVED_DIR / file.name
    dest.write_bytes(file.getbuffer())
    return dest


def delete_saved(name: str):
    target = SAVED_DIR / name
    if target.exists():
        target.unlink()


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Duty Chart Consolidator",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Teacher Duty Chart Consolidator")
st.markdown(
    """
Upload all your duty Excel files (Mid / End exams).  
The system calculates **total load in hours** (Mid = 2 hrs, End = 3 hrs),  
gives a **breakdown by period**, and a full **detailed duty list**.
"""
)

# ── Sidebar settings ──────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    st.markdown("**Hours per duty**")
    mid_hrs = st.number_input("Mid-semester hours", min_value=1, max_value=8, value=config.DUTY_HOURS["Mid"])
    end_hrs = st.number_input("End-semester hours", min_value=1, max_value=8, value=config.DUTY_HOURS["End"])
    config.DUTY_HOURS["Mid"] = mid_hrs
    config.DUTY_HOURS["End"] = end_hrs

    st.markdown("---")
    st.markdown("**Supported file names**")
    st.code("SP-2026-END-SEM.xlsx\nSP-2026-MID-SEM.xlsx\nSem3_Mid_2025.xlsx")

# ── File manager ──────────────────────────────────────────────────────────────
current_files = saved_files()
file_count = len(current_files)
expander_label = f"📁 Input Files — {file_count} file(s) ready" if file_count else "📁 Input Files — no files loaded"

with st.expander(expander_label, expanded=file_count == 0):

    # -- Upload area (uses a resettable key so the widget clears after saving) -
    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0

    new_uploads = st.file_uploader(
        "Add / replace duty chart Excel files (.xlsx)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state['uploader_key']}",
    )

    if new_uploads:
        for f in new_uploads:
            save_uploaded(f)
        st.session_state["uploader_key"] += 1   # clears the widget on next render
        st.rerun()

    # -- Saved file list with delete buttons -----------------------------------
    if current_files:
        st.markdown(f"**{file_count} file(s) ready to process:**")

        header = st.columns([5, 1.5, 1])
        header[0].caption("File name")
        header[1].caption("Size")
        header[2].caption("Remove")

        for fp in current_files:
            col_name, col_size, col_del = st.columns([5, 1.5, 1])
            size_kb = fp.stat().st_size / 1024
            col_name.write(f"📄 {fp.name}")
            col_size.caption(f"{size_kb:.1f} KB")
            if col_del.button(
                "🗑️",
                key=f"del_{fp.name}",
                help=f"Remove {fp.name}",
                on_click=delete_saved,
                args=(fp.name,),
            ):
                st.rerun()
    else:
        st.info("👆 Upload your duty Excel files above to get started.")

# ── Process button ────────────────────────────────────────────────────────────
st.markdown("")
can_process = len(saved_files()) > 0
if st.button("🚀 Process All Files", type="primary", disabled=not can_process):
    with st.spinner("Reading and consolidating duties..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            for fp in saved_files():
                shutil.copy2(fp, tmp_path / fp.name)
            try:
                detailed, summary, ranking = process_folder(str(tmp_path))
                st.session_state["detailed"] = detailed
                st.session_state["summary"] = summary
                st.session_state["ranking"] = ranking
                st.session_state["processed"] = True
            except Exception as e:
                st.error(f"Error: {e}")
                st.stop()

# ── Results ───────────────────────────────────────────────────────────────────
if st.session_state.get("processed"):
    detailed = st.session_state["detailed"]
    summary  = st.session_state["summary"]
    ranking  = st.session_state["ranking"]

    st.markdown("---")
    st.header("📊 Results")

    total_hours = int(detailed["Hours"].sum())
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Duty Hours", f"{total_hours} hrs")
    col2.metric("Total Duties", len(detailed))
    col3.metric("Unique Teachers", detailed["Teacher"].nunique())
    col4.metric("Files Processed", detailed["SourceFile"].nunique())

    st.caption(f"⏱ Mid = {config.DUTY_HOURS['Mid']} hrs  |  End = {config.DUTY_HOURS['End']} hrs")

    tab1, tab2, tab3 = st.tabs(
        ["🏆 Teacher Load Ranking", "📈 Breakdown by Period", "📝 Detailed Duty List"]
    )

    with tab1:
        st.subheader("Teachers sorted by total duty hours")
        st.dataframe(ranking, use_container_width=True, height=500)
        st.bar_chart(ranking.set_index("Teacher")["Total_Hours"].head(25))

    with tab2:
        st.subheader("Duty count + Total Hours by Period / Exam Type")
        st.dataframe(summary, use_container_width=True, height=500)

    with tab3:
        st.subheader("Complete list of every duty slot")
        c1, c2, c3 = st.columns(3)
        teachers = ["All"] + sorted(detailed["Teacher"].unique().tolist())
        selected_teacher = c1.selectbox("Teacher", teachers)
        all_periods = sorted(detailed["Period"].unique().tolist())
        selected_periods = c2.multiselect("Period", all_periods, placeholder="All periods")
        exams = ["All"] + sorted(detailed["ExamType"].unique().tolist())
        selected_exam = c3.selectbox("Exam Type", exams)

        filtered = detailed.copy()
        if selected_teacher != "All":
            filtered = filtered[filtered["Teacher"] == selected_teacher]
        if selected_periods:
            filtered = filtered[filtered["Period"].isin(selected_periods)]
        if selected_exam != "All":
            filtered = filtered[filtered["ExamType"] == selected_exam]

        m1, m2, m3, m4 = st.columns(4)
        first_half  = filtered[filtered["Session"] == "First Half"]
        second_half = filtered[filtered["Session"] == "Second Half"]
        m1.metric("Total Duties", len(filtered))
        m2.metric("Total Hours", f"{int(filtered['Hours'].sum())} hrs")
        m3.metric("First Half Duties", len(first_half))
        m4.metric("Second Half Duties", len(second_half))
        st.caption(f"Showing {len(filtered)} of {len(detailed)} records")

        st.dataframe(filtered, use_container_width=True, height=500)

    st.markdown("---")
    output_file = "Duty_Consolidation_Report.xlsx"
    export_reports(detailed, summary, ranking, output_file)

    with open(output_file, "rb") as f:
        st.download_button(
            label="⬇️ Download Full Excel Report",
            data=f,
            file_name="Duty_Consolidation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ── Supported formats help ────────────────────────────────────────────────────
with st.expander("ℹ️ Supported formats"):
    st.markdown(
        """
### Format A – Summary (your current style)
| Faculty | First Half Duties | Second Half Duties | Total Duties | First Half Dates | Second Half Dates |
|---------|-------------------|--------------------|--------------|------------------|-------------------|
| Dr. X   | 2                 | 1                  | 3            | 28-04-2026, ...  | 29-04-2026        |

The system expands the date lists into individual duty slots.

### Format B – Detailed (one row per duty)
| Teacher Name | Date       | Session | Room |
|--------------|------------|---------|------|
| Dr. X        | 2025-03-10 | Morning | 101  |

### File naming
- `SP-2026-END-SEM.xlsx` → Period = SP-2026, Exam = End (3 hrs)
- `SP-2026-MID-SEM.xlsx` → Period = SP-2026, Exam = Mid (2 hrs)
"""
    )
