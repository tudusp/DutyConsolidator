"""
Duty Chart Consolidator – Streamlit Web App
Run:  streamlit run main.py
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
from processor import process_folder, export_reports
import config

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

uploaded_files = st.file_uploader(
    "Upload duty chart Excel files (.xlsx)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} file(s) uploaded")

    if st.button("🚀 Process All Files", type="primary"):
        with st.spinner("Reading and consolidating duties..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_path = Path(tmpdir)
                for f in uploaded_files:
                    (tmp_path / f.name).write_bytes(f.getbuffer())

                try:
                    detailed, summary, ranking = process_folder(str(tmp_path))
                    st.session_state["detailed"] = detailed
                    st.session_state["summary"] = summary
                    st.session_state["ranking"] = ranking
                    st.session_state["processed"] = True
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.stop()

if st.session_state.get("processed"):
    detailed = st.session_state["detailed"]
    summary = st.session_state["summary"]
    ranking = st.session_state["ranking"]

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

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Duties (filtered)", len(filtered))
        m2.metric("Total Hours (filtered)", f"{int(filtered['Hours'].sum())} hrs")
        m3.caption(f"Showing {len(filtered)} of {len(detailed)} records")

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
else:
    st.info("👆 Upload your duty Excel files and click **Process All Files**.")

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
