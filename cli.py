"""
Command-line version of Duty Chart Consolidator
Usage:
    python cli.py /path/to/folder/containing/excel/files
"""

import sys
from pathlib import Path
from processor import process_folder, export_reports


def main():
    if len(sys.argv) < 2:
        print("Usage: python cli.py <folder_with_excel_files>")
        print("Example: python cli.py ./sample_data")
        sys.exit(1)

    folder = sys.argv[1]
    print(f"\n🔍 Scanning folder: {folder}\n")

    try:
        detailed, summary, ranking = process_folder(folder)

        print("\n" + "=" * 70)
        print("🏆 TEACHER LOAD RANKING (sorted by Total Hours)")
        print("   Mid duty = 2 hrs  |  End duty = 3 hrs")
        print("=" * 70)
        print(ranking.head(20).to_string(index=False))

        print("\n" + "=" * 70)
        print("📈 BREAKDOWN BY PERIOD / EXAM TYPE  (+ Total Hours)")
        print("=" * 70)
        print(summary.head(20).to_string(index=False))

        output = Path(folder) / "Duty_Consolidation_Report.xlsx"
        export_reports(detailed, summary, ranking, str(output))

        print(f"\n✅ Done! Full report saved to:\n   {output}")

    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
