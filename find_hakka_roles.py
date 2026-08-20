"""
Find 專屬客語堂 service role assignments from worship schedule files.

Extracts for Hakka (客語堂) service only:
  - 司會 (MC/Host)
  - 投影同工 / 投影製作放映 (Projection)
  - 司琴 (Pianist)
"""

import os
import sys
import glob
import datetime
from collections import defaultdict

# Force UTF-8 output on Windows (avoids cp950 encoding errors)
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.run(["pip", "install", "openpyxl"], check=True)
    import openpyxl

try:
    from docx import Document
except ImportError:
    import subprocess
    subprocess.run(["pip", "install", "python-docx"], check=True)
    from docx import Document


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TARGET_ROLES = {
    "司會": "司會 (MC/Host)",
    "司琴": "司琴 (Pianist)",
    "投影": "投影同工 (Projection)",
}


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def cell_text(cell):
    """Return stripped string value of a cell, or empty string."""
    if cell is None or cell.value is None:
        return ""
    return str(cell.value).strip()


def date_label(value):
    """Convert Excel datetime or serial number to YYYY-MM-DD string."""
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)) and 40000 < value < 60000:
        base = datetime.datetime(1899, 12, 30)
        return (base + datetime.timedelta(days=int(value))).strftime("%Y-%m-%d")
    return str(value).strip()


# ─────────────────────────────────────────────
# XLSX parsing
# ─────────────────────────────────────────────

def find_hakka_row_range(ws):
    """
    Identify the row range for 客語堂 section in the worksheet.

    The section header '客語堂' is written one character per row in column A
    (e.g. A4='客', A5='語', A6='堂'). The section ends where column B contains
    '講員' again (start of the next section, usually 國語堂).

    Returns (start_row, end_row) — the rows containing Hakka service roles,
    i.e. the rows between (and including) the first role row after '客' and
    the row just before the next '講員' row.
    """
    col_a_values = {}  # row -> value in col A
    col_b_values = {}  # row -> value in col B

    for row in ws.iter_rows(min_col=1, max_col=2):
        a_cell, b_cell = row[0], row[1]
        a_val = cell_text(a_cell)
        b_val = cell_text(b_cell)
        if a_val:
            col_a_values[a_cell.row] = a_val
        if b_val:
            col_b_values[b_cell.row] = b_val

    # Find rows where col A contains a single character that is part of '客語堂'
    hakka_chars = set("客語堂")
    hakka_a_rows = sorted(r for r, v in col_a_values.items() if v in hakka_chars)

    if not hakka_a_rows:
        return None

    # The section starts at the row with '客' in col A (or first hakka char row)
    section_start = hakka_a_rows[0]

    # Row before section_start may have 講員 (preacher) — include it if it does
    # (row 3 is typically the preacher row shared or separate)
    # We want actual service roles: find the role rows after section_start
    # The section ends just before the next '講員' row
    role_start = section_start  # roles start here

    # Find next '講員' row after section_start (signals next section)
    next_section_row = None
    for row_idx in sorted(col_b_values.keys()):
        val = col_b_values[row_idx]
        if row_idx > section_start and ("講員" in val or "讲员" in val):
            next_section_row = row_idx
            break

    if next_section_row:
        section_end = next_section_row - 1
    else:
        section_end = ws.max_row

    return role_start, section_end


def parse_xlsx(filepath):
    """
    Parse xlsx and return { role_display_name: { date_str: [names] } }
    for the 客語堂 section only.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    results = defaultdict(lambda: defaultdict(list))

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        hakka_range = find_hakka_row_range(ws)
        if not hakka_range:
            continue
        hakka_start, hakka_end = hakka_range

        # Date row: row 2 always contains the Sunday dates in columns C onward
        date_cols = {}  # col_index -> date_str
        for col_idx in range(3, ws.max_column + 1):
            cell = ws.cell(row=2, column=col_idx)
            if cell.value:
                date_cols[col_idx] = date_label(cell.value)

        if not date_cols:
            continue

        # Find each target role within the Hakka section
        for row_idx in range(hakka_start, hakka_end + 1):
            b_cell = ws.cell(row=row_idx, column=2)
            b_val = cell_text(b_cell)

            for role_key, role_display in TARGET_ROLES.items():
                if role_key in b_val:
                    for col_idx, date_str in date_cols.items():
                        name_cell = ws.cell(row=row_idx, column=col_idx)
                        name = cell_text(name_cell)
                        if name:
                            results[role_display][date_str].append(name)
                    break  # found the role for this row, move on

    return results


# ─────────────────────────────────────────────
# DOCX parsing
# ─────────────────────────────────────────────

def is_hakka_row(section_label):
    """
    Return True if the section label cell (column 0) belongs to 客語堂.
    In the docx the cell text is '客\n語\n堂' (one char per line).
    We accept any cell whose text contains 客 and 語 but NOT 國/蓉/粵/華,
    which would indicate the Mandarin section.
    """
    label = section_label.replace("\n", "").replace(" ", "")
    if not label:
        return False  # blank = continuation of previous section (招待 row etc.)
    mandarin_markers = {"國", "蓉", "粵", "華"}
    return "客" in label and "語" in label and not any(m in label for m in mandarin_markers)


def parse_docx(filepath):
    """
    Parse docx and return { role_display_name: { date_str: [names] } }
    for the 客語堂 section only.

    Table structure (one row per role):
      Col 0: section label ('客\\n語\\n堂' or '蓉元\\n國\\n語\\n堂' …)
      Col 1: role name ('司會', '司琴', '投影製作放映' …)
      Col 2+: person names, one per Sunday date
    Row 0 is the header: col 0 empty, col 1 = '類別  日期', col 2+ = dates ('1/5', '1/12' …)
    """
    doc = Document(filepath)
    results = defaultdict(lambda: defaultdict(list))

    for table in doc.tables:
        if len(table.rows) < 3:
            continue

        # Build deduplicated grid: list of lists of cell texts
        grid = []
        for row in table.rows:
            seen = set()
            row_texts = []
            for cell in row.cells:
                if id(cell._tc) not in seen:
                    seen.add(id(cell._tc))
                    row_texts.append(cell.text.strip())
            grid.append(row_texts)

        # Row 0 is the header row; extract date labels from col 2 onward
        header = grid[0]
        date_col_map = {}  # col_idx -> date_str
        for col_idx in range(2, len(header)):
            t = header[col_idx]
            if t and ("/" in t or "-" in t) and any(c.isdigit() for c in t):
                date_col_map[col_idx] = t

        if not date_col_map:
            continue

        # Process data rows (skip header row 0)
        for row_texts in grid[1:]:
            if len(row_texts) < 2:
                continue

            section_label = row_texts[0]
            role_label = row_texts[1]

            # Only process rows that belong to 客語堂
            if not is_hakka_row(section_label):
                continue

            for role_key, role_display in TARGET_ROLES.items():
                if role_key in role_label:
                    for col_idx, date_str in date_col_map.items():
                        if col_idx < len(row_texts):
                            name = row_texts[col_idx].strip()
                            if name:
                                results[role_display][date_str].append(name)
                    break

    return results


# ─────────────────────────────────────────────
# Aggregate and display
# ─────────────────────────────────────────────

def collect_unique_names(all_results):
    """Return { role_display: sorted list of unique names } across all files."""
    role_names = defaultdict(set)
    for _file, role_data in all_results.items():
        for role, date_data in role_data.items():
            for date, names in date_data.items():
                for name in names:
                    role_names[role].add(name)
    return {role: sorted(names) for role, names in role_names.items()}


def print_results(all_results, unique_names):
    role_order = list(TARGET_ROLES.values())
    separator = "=" * 60

    print(separator)
    print("  專屬客語堂 服事名單（彙整）")
    print(separator)

    for role in role_order:
        names = unique_names.get(role, [])
        print(f"\n【{role}】")
        if names:
            for name in names:
                print(f"  • {name}")
        else:
            print("  （未找到資料）")

    print(f"\n{separator}")
    print("  詳細排班（依檔案 / 日期）")
    print(separator)

    for filename, role_data in sorted(all_results.items()):
        if not any(role_data.values()):
            continue
        print(f"\n▶ {filename}")
        for role in role_order:
            date_data = role_data.get(role, {})
            if not date_data:
                continue
            print(f"  【{role}】")
            for date in sorted(date_data.keys()):
                names = ", ".join(date_data[date])
                print(f"    {date}: {names}")

    print()


def main():
    all_results = {}

    xlsx_files = sorted(glob.glob(os.path.join(BASE_DIR, "*.xlsx")))
    docx_files = sorted(glob.glob(os.path.join(BASE_DIR, "*.docx")))

    for filepath in xlsx_files:
        filename = os.path.basename(filepath)
        print(f"Parsing XLSX: {filename} ...")
        try:
            result = parse_xlsx(filepath)
            all_results[filename] = result
        except Exception as e:
            print(f"  ERROR: {e}")

    for filepath in docx_files:
        filename = os.path.basename(filepath)
        print(f"Parsing DOCX: {filename} ...")
        try:
            result = parse_docx(filepath)
            all_results[filename] = result
        except Exception as e:
            print(f"  ERROR: {e}")

    print()
    unique_names = collect_unique_names(all_results)
    print_results(all_results, unique_names)


if __name__ == "__main__":
    main()
