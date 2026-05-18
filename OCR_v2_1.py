import fitz  # PyMuPDF
import os
import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import sys
import threading

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass


# ─────────────────────────────────────────────
#  COMPILED REGEX CONSTANTS  (compiled once at import, reused everywhere)
# ─────────────────────────────────────────────

# Semester header: "Fall 2023 - Urbana-Champaign"
_RE_SEM = re.compile(
    r"^(Fall|Spring|Summer|Winter)\s+(\d{4})\s*[-–]\s*(.+)", re.IGNORECASE
)
# Course line: SUBJ NNN  Title  credit  grade  pts  [R]
_RE_COURSE = re.compile(
    r"^([A-Z]{2,5})\s{1,5}(\d{3}[A-Z]?)\s{2,}(.+?)\s{2,}"
    r"(\d+\.\d{2})\s+([A-F][+-]?|[A-F]W?|W|ABS|CR|NC)\s+(\d+\.\d{2})"
)
# In-progress course: no grade or points
_RE_INPROG = re.compile(
    r"^([A-Z]{2,5})\s{1,5}(\d{3}[A-Z]?)\s{2,}(.+?)\s{2,}(\d+\.\d{2})\s+IN PROGRESS"
)
# Lines to skip (summaries, footers, headers)
_RE_SKIP = re.compile(
    r"^\s*(Ehrs:|QPts:|GPA:|TRANSFER CREDIT|INSTITUTION CREDIT|END OF TRANSCRIPT|"
    r"CONTINUED ON|Page \d|TRANSCRIPT TOTALS|Earned Hrs|TOTAL|OVERALL|"
    r"Grainger|Nondegree|College\s*:|Major\s*:|SUBJ|_{5,}|\*{5,})",
    re.IGNORECASE,
)
# Grade-scale header detector: "CRED GRD" → captures "GRD"
_RE_GRD_HEADER = re.compile(r"\bCRED\s+(\S+)", re.IGNORECASE)


# ─────────────────────────────────────────────
#  PARSING
# ─────────────────────────────────────────────

def parse_student_info(text):
    """
    Extract name, university ID, and major from a page.
    Uses the original line-scan logic that was confirmed working,
    plus major extraction retained from the new version.
    """
    lines = text.splitlines()
    name = "Unknown"
    uid = "Unknown"
    major = "Unknown"

    # Original working logic: scan lines 2-20 for name (has comma) and 9-digit ID
    for line in lines[2:20]:
        line = line.strip()
        if ',' in line and all(x.isalpha() or x in " ,-." for x in line.replace(',', '')) and name == "Unknown":
            name = line
        if re.match(r"^6\d{8}$", line) and uid == "Unknown":
            uid = line
        if name != "Unknown" and uid != "Unknown":
            break

    # Major extraction
    major = next((line.split(":")[1].strip() for line in lines if "Major :" in line), "Unknown")

    return name, uid, major


def parse_courses(text):
    """
    Parse all semester blocks from a block of text.
    Returns dict: { semester_label -> [course_dict, ...] }

    Uses module-level pre-compiled regex constants (_RE_SEM, _RE_COURSE, etc.)
    to avoid recompilation on every call.

    Also detects the grade-scale label (评分基准) from the UIUC column header:
        SUBJ NO.   COURSE TITLE   CRED GRD   PTS R
    """
    semesters = {}
    current_semester = None

    # Detect grade scale from the transcript column header line (compiled at module level)
    grd_scale = ""
    for raw in text.splitlines():
        hm = _RE_GRD_HEADER.search(raw)
        if hm and "SUBJ" in raw.upper():
            grd_scale = hm.group(1).strip()
            break

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if _RE_SKIP.match(line):
            continue

        sm = _RE_SEM.match(line)
        if sm:
            season, year, _ = sm.group(1), sm.group(2), sm.group(3)
            current_semester = f"{season.capitalize()} {year}"
            semesters.setdefault(current_semester, [])
            continue

        if current_semester is None:
            continue

        cm = _RE_COURSE.match(line)
        if cm:
            semesters[current_semester].append({
                "subj":      cm.group(1),
                "number":    cm.group(2),
                "title":     cm.group(3).strip(),
                "credits":   float(cm.group(4)),
                "grade":     cm.group(5),
                "points":    float(cm.group(6)),
                "grd_scale": grd_scale,
            })
            continue

        ip = _RE_INPROG.match(line)
        if ip:
            semesters[current_semester].append({
                "subj":      ip.group(1),
                "number":    ip.group(2),
                "title":     ip.group(3).strip(),
                "credits":   float(ip.group(4)),
                "grade":     "IN PROGRESS",
                "points":    None,
                "grd_scale": grd_scale,
            })

    return semesters


def sort_semesters(semesters_dict):
    order = {"Spring": 1, "Summer": 2, "Fall": 3, "Winter": 4}
    return sorted(
        semesters_dict.items(),
        key=lambda x: (int(x[0].split()[1]), order.get(x[0].split()[0], 0))
    )


def process_pdf(pdf_path):
    """
    Open a multi-page PDF, carry student info across all pages,
    merge semester data from every page, return unified transcript dict.
    """
    doc = fitz.open(pdf_path)
    name, uid, major = None, None, None
    all_courses = {}

    for page_num, page in enumerate(doc, start=1):
        text = page.get_text()

        # Always try to pull student info (repeated on every page in UIUC format)
        n, u, mj = parse_student_info(text)
        if n and n != "Unknown" and (not name or name == "Unknown"):
            name = n
        if u and u != "Unknown" and (not uid or uid == "Unknown"):
            uid = u
        if mj and mj != "Unknown" and (not major or major == "Unknown"):
            major = mj

        # Parse courses from this page and merge
        page_sems = parse_courses(text)
        for sem, courses in page_sems.items():
            all_courses.setdefault(sem, [])
            # Avoid duplicates: check by (subj+number+grade)
            existing_keys = {
                (c["subj"], c["number"], c["grade"])
                for c in all_courses[sem]
            }
            for c in courses:
                key = (c["subj"], c["number"], c["grade"])
                if key not in existing_keys:
                    all_courses[sem].append(c)
                    existing_keys.add(key)

    doc.close()

    return {
        "name":     name or "Unknown",
        "uid":      uid or "Unknown",
        "major":    major or "Unknown",
        "semesters": sort_semesters(all_courses),
    }


# ─────────────────────────────────────────────
#  EXCEL EXPORT  – matches ZJU import template
# ─────────────────────────────────────────────
#
# Columns (from screenshot):
#  A  学号(ZJU ID)
#  B  UIN
#  C  Name
#  D  学历
#  E  学术机构
#  F  转校模型编号
#  G  学科设定
#  H  专业设定(Major)
#  I  来源机构
#  J  包含在平均成绩
#  K  接续学期
#  L  抵免同等组
#  M  学年
#  N  外部学期
#  O  学校学科
#  P  学校课程编号
#  Q  描述
#  R  已获分(Credit)
#  S  成绩输入(Grade)
#  T  课程ID(Course ID)

HEADERS = [
    "学号(ZJU ID)",              # A
    "UIN",                        # B
    "Name",                       # C
    "学历",                       # D
    "学术机构",                   # E
    "转校\n模型\n编号",           # F
    "学科设定",                   # G
    "专业设定(Major)",            # H
    "来源机构",                   # I
    "包含在\n平均成绩",           # J
    "接续学期",                   # K
    "抵免同等\n组",               # L
    "学年",                       # M
    "外部学期",                   # N
    "学校学科",                   # O
    "学校课程\n编号",             # P
    "描述",                       # Q
    "已获分(Credit)",             # R
    "成绩输入(Grade)",            # S
    "课程ID(Course ID)",          # T
    "评分计划",                   # U
    "评分基准",                   # V
    "已抵免",                     # W
    "正式成绩",                   # X
    "包括在\n学分 FA WI\n统计中", # Y
]

# Column widths (characters)
COL_WIDTHS = [14, 12, 16, 8, 12, 8, 10, 22, 8, 10, 10, 10, 8, 10, 10, 14, 30, 14, 16, 18,
              10, 10, 10, 10, 14]

# Header background: yellow (matches screenshot)
HEADER_FILL   = PatternFill("solid", fgColor="FFFF00")
HEADER_FONT   = Font(bold=True, name="Arial", size=10)
CELL_FONT     = Font(name="Arial", size=10)
THIN_BORDER   = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

# Season → external term full names for 外部学期
SEASON_TO_EXT = {"Spring": "SPRING", "Summer": "SUMMER", "Fall": "FALL", "Winter": "WINTER"}


def _semester_to_year_term(semester_label):
    """'Fall 2023' → year='2023', ext_term='FALL'"""
    parts = semester_label.split()
    season, year = parts[0], parts[1]
    ext = SEASON_TO_EXT.get(season, season.upper())
    return year, ext


def load_uin_zjuid_map(ref_path):
    """
    Load a UIN → ZJU ID mapping from a reference Excel or CSV file.

    Expected columns (any order, case-insensitive):
        UIN  (or 'University ID', 'uin')
        ZJU ID  (or 'zjuid', '学号', '学号(ZJU ID)')

    Returns a dict:  { uin_string -> zjuid_string }
    Rows with missing values in either column are skipped.
    """
    mapping = {}
    if not ref_path:
        return mapping

    ext = os.path.splitext(ref_path)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm", ".xls"):
            wb = openpyxl.load_workbook(ref_path, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
        elif ext == ".csv":
            import csv
            with open(ref_path, newline="", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                rows = list(reader)
        else:
            print(f"  ✗ Reference file format not supported: {ext}")
            return mapping

        if not rows:
            return mapping

        # Find column indices from header row
        header = [str(h).strip().lower() if h else "" for h in rows[0]]
        uin_aliases  = {"uin", "university id", "universityid"}
        zjuid_aliases = {"zju id", "zjuid", "学号", "学号(zju id)"}

        uin_col   = next((i for i, h in enumerate(header) if h in uin_aliases),   None)
        zjuid_col = next((i for i, h in enumerate(header) if h in zjuid_aliases), None)

        if uin_col is None or zjuid_col is None:
            print(f"  ✗ Reference file: could not find UIN or ZJU ID columns.")
            print(f"    Headers found: {rows[0]}")
            return mapping

        for row in rows[1:]:
            if len(row) <= max(uin_col, zjuid_col):
                continue
            uin_val   = str(row[uin_col]).strip()   if row[uin_col]   else ""
            zjuid_val = str(row[zjuid_col]).strip() if row[zjuid_col] else ""
            if uin_val and zjuid_val and uin_val.lower() != "none":
                mapping[uin_val] = zjuid_val

        print(f"  ✓ Loaded {len(mapping)} UIN→ZJU ID mapping(s) from reference file.")
    except Exception as e:
        print(f"  ✗ Failed to load reference file: {e}")

    return mapping


def load_course_map(ref_path):
    """
    Load a course-code lookup from a reference Excel or CSV file.

    Expected columns (any order, case-insensitive):
        course code  (or 'course', 'code', '学校课程编号', 'subj no', 'subj_no')
        course id    (or 'courseid', 'course_id', '课程ID', 'zju course id')

    Optionally also:
        评分计划, 评分基准, 已抵免, 正式成绩, 包括在学分FA WI统计中

    Returns a dict: { "ECE 329" -> { "course_id": "...", "评分计划": ..., ... } }
    """
    mapping = {}
    if not ref_path:
        return mapping

    ext = os.path.splitext(ref_path)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm", ".xls"):
            wb = openpyxl.load_workbook(ref_path, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
        elif ext == ".csv":
            import csv
            with open(ref_path, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.reader(f))
        else:
            print(f"  ✗ Course reference file format not supported: {ext}")
            return mapping

        if not rows:
            return mapping

        header = [str(h).strip().lower() if h else "" for h in rows[0]]

        code_aliases = {"course code", "course", "code", "学校课程编号", "subj no", "subj_no", "course_code"}
        id_aliases   = {"course id", "courseid", "course_id", "课程id", "zju course id", "课程id(course id)"}

        code_col = next((i for i, h in enumerate(header) if h in code_aliases), None)
        id_col   = next((i for i, h in enumerate(header) if h in id_aliases),   None)

        # Optional extra columns — map Chinese header → dict key
        extra_aliases = {
            "评分计划": "评分计划",
            "评分基准": "评分基准",
            "已抵免":   "已抵免",
            "正式成绩": "正式成绩",
            "包括在学分fa wi统计中": "包括在学分FAWI统计中",
            "包括在\n学分 fa wi\n统计中": "包括在学分FAWI统计中",
        }
        extra_cols = {}
        for i, h in enumerate(header):
            if h in extra_aliases:
                extra_cols[extra_aliases[h]] = i

        if code_col is None:
            print(f"  ✗ Course reference file: could not find course code column.")
            print(f"    Headers found: {list(rows[0])}")
            return mapping

        for row in rows[1:]:
            if not row or len(row) <= code_col:
                continue
            code_val = str(row[code_col]).strip() if row[code_col] else ""
            if not code_val:
                continue
            entry = {}
            if id_col is not None and len(row) > id_col:
                entry["course_id"] = str(row[id_col]).strip() if row[id_col] else ""
            for key, col_idx in extra_cols.items():
                entry[key] = str(row[col_idx]).strip() if len(row) > col_idx and row[col_idx] else ""
            mapping[code_val.upper()] = entry

        print(f"  ✓ Loaded {len(mapping)} course mapping(s) from course reference file.")
    except Exception as e:
        print(f"  ✗ Failed to load course reference file: {e}")

    return mapping


def load_term_map(ref_path):
    """
    Load a reference file for 接续学期 lookup.
    Logic TBD — returns empty dict until implemented.
    """
    mapping = {}
    if not ref_path:
        return mapping
    # TODO: implement lookup logic once the column structure is determined
    print(f"  ⚠ 接续学期 reference loaded but lookup logic is not yet implemented.")
    return mapping


def safe_save(wb, output_path, retries=3):
    """
    Save a workbook, handling Windows file-lock (PermissionError / errno 13).
    If the file is open in Excel the user is prompted to close it, then retried.
    """
    import time
    for attempt in range(1, retries + 1):
        try:
            wb.save(output_path)
            return True
        except PermissionError:
            filename = os.path.basename(output_path)
            if attempt < retries:
                answer = messagebox.askretrycancel(
                    "File Locked",
                    f'Cannot write to:\n  {filename}\n\n'
                    f'The file is open in another program (e.g. Excel).\n'
                    f'Please close it, then click Retry.',
                )
                if not answer:
                    print(f"  ✗ Skipped (file locked): {filename}")
                    return False
                time.sleep(0.5)
            else:
                messagebox.showerror(
                    "Permission Denied",
                    f'Could not save:\n  {filename}\n\n'
                    f'Please close the file in Excel and run again.',
                )
                print(f"  ✗ Failed after {retries} attempts (file locked): {filename}")
                return False
    return False


def export_to_excel(transcript, output_path, zjuid="", course_map=None, term_map=None):
    """
    Write transcript data to Excel using append-only row writes.

    openpyxl 3.1+ removed write_only streaming, so we use normal mode with
    ws.append() for each data row — rows are serialised incrementally and the
    sheet never holds more than the current row's Cell objects in RAM at once.
    Column widths, freeze pane, and header styling are applied up front.
    """
    if course_map is None:
        course_map = {}
    if term_map is None:
        term_map = {}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transcript"

    # ── Column widths + header row height (set before data so xlsx is valid) ──
    ws.row_dimensions[1].height = 45
    for col_idx, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A2"

    # ── Header row ──
    ws.append(HEADERS)
    for col_idx in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.border    = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ── Data rows — built as plain lists and appended one at a time ──
    for semester_label, courses in transcript["semesters"]:
        year, ext_term = _semester_to_year_term(semester_label)
        for course in courses:
            course_code = f"{course['subj']} {course['number']}"
            cm          = course_map.get(course_code.upper(), {})
            grd_scale   = course.get("grd_scale", "") or cm.get("评分基准", "")

            ws.append([
                zjuid,                                    # A  ZJU ID
                transcript["uid"],                        # B  UIN
                transcript["name"],                       # C  Name
                "UGRD",                                   # D  学历
                "ZJUNV",                                  # E  学术机构
                "1",                                      # F  转校模型编号
                "UC001",                                  # G  学科设定
                transcript["major"],                      # H  专业设定
                "UIUC",                                   # I  来源机构
                "Y",                                      # J  包含在平均成绩
                term_map.get(transcript["uid"], ""),      # K  接续学期
                "1",                                      # L  抵免同等组
                year,                                     # M  学年
                ext_term,                                 # N  外部学期
                course["subj"],                           # O  学校学科
                course["number"],                         # P  学校课程编号
                f"{course['subj']} {course['number']}",  # Q  描述
                course["credits"],                        # R  已获分
                course["grade"],                          # S  成绩输入
                cm.get("course_id", ""),                  # T  课程ID
                "UGS",                                    # U  评分计划
                grd_scale,                                # V  评分基准
                course["credits"],                        # W  已抵免
                course["grade"],                          # X  正式成绩
                "Y",                                      # Y  包括在学分FA WI统计中
            ])

            # Style the row just appended
            row_idx = ws.max_row
            for col_idx in range(1, len(HEADERS) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font      = CELL_FONT
                cell.border    = THIN_BORDER
                cell.alignment = Alignment(vertical="center", wrap_text=False)

    if safe_save(wb, output_path):
        print(f"  ✓ Saved: {os.path.basename(output_path)}")


def combine_all_excels(excel_files, output_folder, combined_filename="combined_transcripts.xlsx"):
    combined_wb = openpyxl.Workbook()
    combined_ws = combined_wb.active
    combined_ws.title = "Combined"

    headers_written = False
    seen = set()

    for file in excel_files:
        try:
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            if not headers_written:
                combined_ws.append(list(rows[0]))
                # Style header
                for col_idx in range(1, len(rows[0]) + 1):
                    cell = combined_ws.cell(row=1, column=col_idx)
                    cell.font = HEADER_FONT
                    cell.fill = HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    combined_ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS[col_idx - 1]
                combined_ws.row_dimensions[1].height = 45
                combined_ws.freeze_panes = "A2"
                headers_written = True

            for row in rows[1:]:
                key = (row[1], row[15], row[12], row[13])  # UIN, course_code, year, term
                if key not in seen:
                    combined_ws.append(list(row))
                    seen.add(key)
        except Exception as e:
            print(f"  ✗ Could not read {os.path.basename(file)}: {e}")

    combined_path = os.path.join(output_folder, combined_filename)
    if safe_save(combined_wb, combined_path):
        print(f"\n✓ Combined file saved: {combined_path}")


# ─────────────────────────────────────────────
#  PROCESSING PIPELINE
# ─────────────────────────────────────────────

def process_files(files, output_folder, progress_var, progress_label, root, uin_map=None, course_map=None, term_map=None):
    excel_files = []
    total = len(files)
    if uin_map is None:
        uin_map = {}
    if course_map is None:
        course_map = {}
    if term_map is None:
        term_map = {}

    for i, pdf_path in enumerate(files, start=1):
        filename = os.path.basename(pdf_path)
        base = os.path.splitext(filename)[0]
        pct = int((i - 1) / total * 100)

        root.after(0, lambda p=pct, f=filename: (
            progress_var.set(p),
            progress_label.config(text=f"Processing: {f}")
        ))

        try:
            print(f"\n[{i}/{total}] {filename}")
            transcript = process_pdf(pdf_path)
            print(f"  Student : {transcript['name']} ({transcript['uid']})")
            print(f"  Major   : {transcript['major']}")
            sem_count = len(transcript['semesters'])
            course_count = sum(len(c) for _, c in transcript['semesters'])
            print(f"  Semesters: {sem_count}  |  Courses: {course_count}")

            zjuid = uin_map.get(transcript['uid'], "")
            if zjuid:
                print(f"  ZJU ID  : {zjuid}")
            else:
                print(f"  ZJU ID  : (not found in reference file)")

            out_xlsx = os.path.join(output_folder, f"{base}.xlsx")
            export_to_excel(transcript, out_xlsx, zjuid=zjuid, course_map=course_map, term_map=term_map)
            excel_files.append(out_xlsx)

        except PermissionError:
            print(f"  ✗ Skipped due to file lock: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process {filename}:\n{e}")
            print(f"  ✗ Error: {e}")

    if len(excel_files) > 1:
        print("\nCombining all files...")
        combine_all_excels(excel_files, output_folder)
    elif len(excel_files) == 1:
        print("\n(Only one file – no combined output needed)")

    root.after(0, lambda: (
        progress_var.set(100),
        progress_label.config(text=f"Done! Processed {len(excel_files)} file(s).")
    ))
    print("\n✓ All done.")


def run_in_thread(files, output_folder, progress_var, progress_label, root, btn_files, btn_folder, uin_map=None, course_map=None, term_map=None):
    btn_files.config(state=tk.DISABLED)
    btn_folder.config(state=tk.DISABLED)
    t = threading.Thread(
        target=lambda: process_files(files, output_folder, progress_var, progress_label, root, uin_map, course_map, term_map),
        daemon=True
    )
    t.start()

    def check():
        if t.is_alive():
            root.after(200, check)
        else:
            btn_files.config(state=tk.NORMAL)
            btn_folder.config(state=tk.NORMAL)

    root.after(200, check)


# ─────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────

def main():
    global output_folder_path

    # Use TkinterDnD.Tk() as the root window if available — required for DND to work.
    # Fall back silently to plain tk.Tk() if the package isn't installed.
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        _dnd_available = True
        root = TkinterDnD.Tk()
    except ImportError:
        _dnd_available = False
        root = tk.Tk()

    root.title("Transcript Reader")
    root.geometry("860x820")
    root.resizable(False, False)

    try:
        icon_path = resource_path("icon.png")
        if os.path.exists(icon_path):
            root.iconphoto(False, tk.PhotoImage(file=icon_path))
    except Exception:
        pass

    BG     = "#FAFAFA"
    FG     = "#222222"
    ACCENT = "#000000"
    BTN_BG = "#E0E0E0"
    BTN_HV = "#C0C0C0"
    TXT_BG = "#FFFFFF"

    root.configure(bg=BG)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("Accent.TButton", background=BTN_BG, foreground=FG,
                    font=("Segoe UI", 11), padding=8, borderwidth=0, focusthickness=0)
    style.map("Accent.TButton",
              background=[("active", BTN_HV), ("pressed", BTN_HV)],
              relief=[("pressed", "flat"), ("!pressed", "flat")])
    style.configure("TLabel", background=BG, foreground=FG, font=("Segoe UI", 10))
    style.configure("Header.TLabel", font=("Segoe UI", 22, "bold"),
                    background=BG, foreground=ACCENT)
    style.configure("TFrame", background=BG)
    style.configure("Prog.TFrame", background=BG)

    output_folder_path = tk.StringVar()
    ref_file_path      = tk.StringVar()
    term_ref_file_path = tk.StringVar()
    progress_var = tk.IntVar(value=0)

    # ── Header ──
    ttk.Label(root, text="Transcript Reader", style="Header.TLabel").pack(pady=(20, 15))

    # ── Output folder row ──
    top_frame = ttk.Frame(root, padding=5)
    top_frame.pack(fill=tk.X, padx=20, pady=(0, 6))
    top_frame.config(borderwidth=1, relief="solid")

    lbl_output = ttk.Label(top_frame, text="Output Folder: (not selected)", anchor="w")

    def browse_output():
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            output_folder_path.set(folder)
            lbl_output.config(text=f"Output Folder: {folder}")
            _refresh_action_buttons()

    ttk.Button(top_frame, text="Select Output Folder",
               command=browse_output, style="Accent.TButton").pack(side=tk.LEFT)
    lbl_output.pack(side=tk.LEFT, padx=15, fill=tk.X, expand=True)

    # ── Reference file row ──
    ref_frame = ttk.Frame(root, padding=5)
    ref_frame.pack(fill=tk.X, padx=20, pady=(0, 6))
    ref_frame.config(borderwidth=1, relief="solid")

    lbl_ref = ttk.Label(ref_frame, text="Reference File: (optional — links UIN to ZJU ID)", anchor="w")

    def browse_ref():
        path = filedialog.askopenfilename(
            title="Select UIN→ZJU ID Reference File",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("All Files", "*.*")]
        )
        if path:
            ref_file_path.set(path)
            lbl_ref.config(text=f"Reference File: {os.path.basename(path)}")

    def clear_ref():
        ref_file_path.set("")
        lbl_ref.config(text="Reference File: (optional — links UIN to ZJU ID)")

    ttk.Button(ref_frame, text="Select Reference File",
               command=browse_ref, style="Accent.TButton").pack(side=tk.LEFT)
    ttk.Button(ref_frame, text="Clear",
               command=clear_ref, style="Accent.TButton").pack(side=tk.LEFT, padx=(4, 0))
    lbl_ref.pack(side=tk.LEFT, padx=15, fill=tk.X, expand=True)

    # ── Course reference file row ──
    course_ref_file_path = tk.StringVar()
    course_ref_frame = ttk.Frame(root, padding=5)
    course_ref_frame.pack(fill=tk.X, padx=20, pady=(0, 6))
    course_ref_frame.config(borderwidth=1, relief="solid")

    lbl_course_ref = ttk.Label(course_ref_frame,
                               text="Course Reference: (optional — looks up course ID & extra fields)",
                               anchor="w")

    def browse_course_ref():
        path = filedialog.askopenfilename(
            title="Select Course Code Reference File",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("All Files", "*.*")]
        )
        if path:
            course_ref_file_path.set(path)
            lbl_course_ref.config(text=f"Course Reference: {os.path.basename(path)}")

    def clear_course_ref():
        course_ref_file_path.set("")
        lbl_course_ref.config(text="Course Reference: (optional — looks up course ID & extra fields)")

    ttk.Button(course_ref_frame, text="Select Course Reference",
               command=browse_course_ref, style="Accent.TButton").pack(side=tk.LEFT)
    ttk.Button(course_ref_frame, text="Clear",
               command=clear_course_ref, style="Accent.TButton").pack(side=tk.LEFT, padx=(4, 0))
    lbl_course_ref.pack(side=tk.LEFT, padx=15, fill=tk.X, expand=True)

    # ── 接续学期 reference file row ──
    term_ref_frame = ttk.Frame(root, padding=5)
    term_ref_frame.pack(fill=tk.X, padx=20, pady=(0, 6))
    term_ref_frame.config(borderwidth=1, relief="solid")

    lbl_term_ref = ttk.Label(term_ref_frame,
                             text="接续学期 Reference: (optional — logic TBD)",
                             anchor="w")

    def browse_term_ref():
        path = filedialog.askopenfilename(
            title="Select 接续学期 Reference File",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("All Files", "*.*")]
        )
        if path:
            term_ref_file_path.set(path)
            lbl_term_ref.config(text=f"接续学期 Reference: {os.path.basename(path)}")

    def clear_term_ref():
        term_ref_file_path.set("")
        lbl_term_ref.config(text="接续学期 Reference: (optional — logic TBD)")

    ttk.Button(term_ref_frame, text="Select 接续学期 Reference",
               command=browse_term_ref, style="Accent.TButton").pack(side=tk.LEFT)
    ttk.Button(term_ref_frame, text="Clear",
               command=clear_term_ref, style="Accent.TButton").pack(side=tk.LEFT, padx=(4, 0))
    lbl_term_ref.pack(side=tk.LEFT, padx=15, fill=tk.X, expand=True)

    # ── Action buttons ──
    btn_frame = ttk.Frame(root, padding=5)
    btn_frame.pack(pady=(0, 6), padx=20)
    btn_frame.config(borderwidth=1, relief="solid")

    def _get_uin_map():
        ref = ref_file_path.get()
        if ref:
            print(f"\nLoading UIN reference: {os.path.basename(ref)}")
            return load_uin_zjuid_map(ref)
        return {}

    def _get_course_map():
        ref = course_ref_file_path.get()
        if ref:
            print(f"\nLoading course reference: {os.path.basename(ref)}")
            return load_course_map(ref)
        return {}

    def _get_term_map():
        ref = term_ref_file_path.get()
        if ref:
            print(f"\nLoading 接续学期 reference: {os.path.basename(ref)}")
            return load_term_map(ref)
        return {}

    def pick_files():
        files = filedialog.askopenfilenames(
            title="Select PDF files", filetypes=[("PDF Files", "*.pdf")]
        )
        if files:
            progress_var.set(0)
            progress_label.config(text="Starting…")
            run_in_thread(list(files), output_folder_path.get(),
                          progress_var, progress_label, root, btn_files, btn_folder,
                          uin_map=_get_uin_map(), course_map=_get_course_map(),
                          term_map=_get_term_map())

    def pick_folder():
        folder = filedialog.askdirectory(title="Select PDF folder")
        if folder:
            files = [
                os.path.join(folder, f)
                for f in sorted(os.listdir(folder))
                if f.lower().endswith(".pdf")
            ]
            if not files:
                messagebox.showwarning("No PDFs", "No PDF files found in the selected folder.")
                return
            progress_var.set(0)
            progress_label.config(text="Starting…")
            run_in_thread(files, output_folder_path.get(),
                          progress_var, progress_label, root, btn_files, btn_folder,
                          uin_map=_get_uin_map(), course_map=_get_course_map(),
                          term_map=_get_term_map())

    btn_files  = ttk.Button(btn_frame, text="Select Multiple PDFs",
                            command=pick_files, style="Accent.TButton", state=tk.DISABLED)
    btn_folder = ttk.Button(btn_frame, text="Select Folder",
                            command=pick_folder, style="Accent.TButton", state=tk.DISABLED)
    btn_files.grid(row=0, column=0, padx=5)
    btn_folder.grid(row=0, column=1, padx=5)

    def _refresh_action_buttons():
        state = tk.NORMAL if output_folder_path.get() else tk.DISABLED
        btn_files.config(state=state)
        btn_folder.config(state=state)

    # ── Progress bar ──
    prog_frame = ttk.Frame(root, style="Prog.TFrame")
    prog_frame.pack(fill=tk.X, padx=20, pady=(0, 6))

    progress_label = ttk.Label(prog_frame, text="Idle", anchor="w", style="TLabel")
    progress_label.pack(fill=tk.X, pady=(0, 2))

    pbar = ttk.Progressbar(prog_frame, variable=progress_var,
                           maximum=100, mode="determinate", length=820)
    pbar.pack(fill=tk.X)

    # ── Drag-and-drop zone ──
    drop_frame = tk.Frame(root, bg="#E8E8E8", relief="groove", bd=2,
                          cursor="hand2")
    drop_frame.pack(fill=tk.X, padx=20, pady=(6, 4))

    drop_label = tk.Label(
        drop_frame,
        text="⬇  Drag & Drop PDF files or folders here",
        font=("Segoe UI", 10, "italic"),
        fg="#666666", bg="#E8E8E8", pady=8
    )
    drop_label.pack()

    def _on_drop(event):
        if not output_folder_path.get():
            messagebox.showwarning("No output folder",
                                   "Please select an output folder first.")
            return
        raw = event.data
        # Paths with spaces are wrapped in braces on Windows; handle both cases
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        paths = [a or b for a, b in paths]
        pdf_files = []
        for p in paths:
            if os.path.isdir(p):
                pdf_files += [
                    os.path.join(p, f)
                    for f in sorted(os.listdir(p))
                    if f.lower().endswith(".pdf")
                ]
            elif p.lower().endswith(".pdf"):
                pdf_files.append(p)
        if not pdf_files:
            messagebox.showwarning("No PDFs", "No PDF files found in the dropped items.")
            return
        progress_var.set(0)
        progress_label.config(text="Starting…")
        run_in_thread(pdf_files, output_folder_path.get(),
                      progress_var, progress_label, root, btn_files, btn_folder,
                      uin_map=_get_uin_map(), course_map=_get_course_map(),
                      term_map=_get_term_map())

    if _dnd_available:
        from tkinterdnd2 import DND_FILES
        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind("<<Drop>>", _on_drop)
        drop_label.dnd_bind("<<Drop>>", _on_drop)  # label covers most of the frame area
        drop_label.config(text="⬇  Drag & Drop PDF files or folders here")
    else:
        drop_label.config(
            text="⬇  Drag & Drop — run: pip install tkinterdnd2  |  Use buttons above",
            fg="#999999"
        )

    # ── Log text box ──
    txt_box = ScrolledText(root, width=100, height=22, font=("Consolas", 10),
                           state="disabled", bg=TXT_BG, fg=FG,
                           relief="sunken", borderwidth=1)
    txt_box.pack(padx=20, pady=(4, 4), fill=tk.BOTH, expand=True)

    sys.stdout = StdoutRedirector(txt_box)
    sys.stderr = StdoutRedirector(txt_box)

    # ── Copyright ──
    ttk.Label(root, text="© 2025 Jaden Peterson Wen   © 2026 Junway Lin",
              font=("Segoe UI", 8, "italic"),
              foreground="#777777", background=BG).place(relx=1.0, rely=1.0,
                                                         anchor="se", x=-10, y=-3)

    root.mainloop()


if __name__ == "__main__":
    main()