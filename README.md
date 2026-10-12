# Teacher Duty Chart Consolidator

A simple Python system that consolidates multiple exam duty Excel sheets and calculates the **total load** of each teacher across all mid-term and end-term exams.

## Features

- ✅ Upload / process many Excel files at once
- ✅ Automatic extraction of **Semester** and **Exam Type** (Mid / End) from filename
- ✅ **Duty load in hours** (Mid = 2 hrs, End = 3 hrs) + count of duties
- ✅ **Breakdown** by Semester + Exam Type + Total Hours
- ✅ **Detailed list** of every duty assignment
- ✅ Clean Excel report with three sheets
- ✅ Works as a **web app** (Streamlit) or **command-line tool**

---

## Quick Start

### 1. Install dependencies

```bash
cd duty_consolidator
pip install -r requirements.txt
```

### 2A. Use the Web App (Recommended)

```bash
streamlit run main.py
```

Then open the browser, upload your Excel files, and click **Process All Files**.

### 2B. Use the Command Line

```bash
python cli.py /path/to/folder/with/excel/files
```

Example with the included sample data:

```bash
python cli.py ./sample_data
```

---

## How to prepare your files

### File naming (important)

Name your files like this so the system can detect Semester & Exam Type automatically:

| Filename example          | Detected Semester | Detected Exam Type |
|---------------------------|-------------------|--------------------|
| `Sem3_Mid_2025.xlsx`      | Sem3              | Mid                |
| `Sem5_End_Oct2025.xlsx`   | Sem5              | End                |
| `Semester-4_Midterm.xlsx` | Semester4         | Mid                |

### Excel content

Your Excel should have at least a **Teacher Name** column.

Typical columns:

| Teacher Name       | Date       | Session   | Room | Duty Type    |
|--------------------|------------|-----------|------|--------------|
| Dr. Ananya Sharma  | 2025-03-10 | Morning   | 101  | Invigilator  |
| Prof. Rajesh Kumar | 2025-03-10 | Afternoon | 102  | Reliever     |

You can change the expected column names in `config.py` or in the web app sidebar.

---

## Output

The system generates one Excel file with three sheets:

1. **Teacher Load Ranking** – sorted by total number of duties
2. **Breakdown by Semester** – counts for each Semester-Exam combination + Total
3. **Detailed Duty List** – every single duty record with source file

---

## Customization

Open `config.py` to:

- Change column names (`COLUMN_MAP`)
- Add extra teacher columns (e.g. Reliever, Additional Invigilator)
- Adjust filename parsing regex
- Ignore certain names (Total, Vacant, etc.)

---

## Sample Data

The folder `sample_data/` contains three example duty charts so you can test immediately:

```bash
python cli.py ./sample_data
```

---

## Requirements

- Python 3.9+
- pandas, openpyxl, streamlit
