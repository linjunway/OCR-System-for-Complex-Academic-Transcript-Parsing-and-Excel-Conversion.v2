import fitz  # PyMuPDF
import os
import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import xlsxwriter                          # pip install xlsxwriter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import sys
import threading
import concurrent.futures
import multiprocessing

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
# Continuation markers — signal end of a column/page section
_RE_CONTINUED  = re.compile(
    r"(CONTINUED\s+ON\s+(NEXT\s+COLUMN|PAGE\s+\d+))", re.IGNORECASE
)


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


'''def _split_page_into_segments(text):
    """
    UIUC transcripts use a two-column layout per page.
    fitz.page.get_text() returns text left-column-first then right-column.
    A "CONTINUED ON NEXT COLUMN" or "CONTINUED ON PAGE N" marker signals
    the boundary where the left column ends and the right column begins.

    This function splits the full page text into ordered segments so that
    each segment is parsed independently.  The key effect is that
    current_semester context is NOT carried from the left column into the
    right column — instead each segment starts fresh and re-establishes its
    own current_semester from its own header lines.

    Returns a list of text strings: [left_segment, right_segment, ...]
    If no continuation marker is found, returns [text] (single segment).
    """
    raw_lines  = text.splitlines()
    segments   = []
    current    = []

    for line in raw_lines:
        if _RE_CONTINUED.search(line):
            # End this segment at the continuation marker (don't include marker)
            if current:
                segments.append("\n".join(current))
            current = []
        else:
            current.append(line)

    if current:
        segments.append("\n".join(current))

    return segments if segments else [text]'''


def parse_courses(text, current_semester=None):
    """
    Parse all semester blocks from a block of text.
    Returns tuple: (semesters_dict, updated_current_semester)
    """
    semesters = {}

    # Detect grade scale once from the full page text
    grd_scale = ""
    for raw in text.splitlines():
        hm = _RE_GRD_HEADER.search(raw)
        if hm and "SUBJ" in raw.upper():
            grd_scale = hm.group(1).strip()
            break

    # Process all lines sequentially in the fixed spatial reading order
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
            
        # Ignore continuation markers and skip lines.
        # This naturally lets the current_semester persist to the next line.
        if _RE_SKIP.match(line) or _RE_CONTINUED.search(line):
            continue

        # Check if we encounter a new semester header
        sm = _RE_SEM.match(line)
        if sm:
            season, year, _ = sm.group(1), sm.group(2), sm.group(3)
            current_semester = f"{season.capitalize()} {year}"
            semesters.setdefault(current_semester, [])
            continue

        # If no semester context exists yet, skip lines until one is established
        if current_semester is None:
            continue

        # Match regular completed courses
        cm = _RE_COURSE.match(line)
        if cm:
            semesters.setdefault(current_semester, [])
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

        # Match in-progress courses
        ip = _RE_INPROG.match(line)
        if ip:
            semesters.setdefault(current_semester, [])
            semesters[current_semester].append({
                "subj":      ip.group(1),
                "number":    ip.group(2),
                "title":     ip.group(3).strip(),
                "credits":   float(ip.group(4)),
                "grade":     "IN PROGRESS",
                "points":    None,
                "grd_scale": grd_scale,
            })

    return semesters, current_semester


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
    Ensures strict Column 1 -> Column 2 spatial reading order.
    """
    doc = fitz.open(pdf_path)
    name, uid, major = None, None, None
    all_courses = {}

    current_semester = None  # Global context to persist across columns/pages

    for page_num, page in enumerate(doc, start=1):
        # ── 1. Enforce Pg X C1 -> Pg X C2 spatial extraction ──
        blocks = page.get_text("blocks")
        page_center = page.rect.width / 2

        left_col, right_col = [], []
        for b in blocks:
            # b[6] == 0 confirms it is a text block, not an image
            if b[6] == 0:  
                # b[0] is the left x-coordinate of the block
                if b[0] < page_center:
                    left_col.append(b)
                else:
                    right_col.append(b)

        # Sort blocks vertically top-to-bottom within each column
        left_col.sort(key=lambda b: b[1])
        right_col.sort(key=lambda b: b[1])

        # Reconstruct strict sequential text
        text = ""
        for b in left_col:
            text += b[4] + "\n"
        for b in right_col:
            text += b[4] + "\n"

        # ── 2. Parse Student Info ──
        n, u, mj = parse_student_info(text)
        if n and n != "Unknown" and (not name or name == "Unknown"):
            name = n
        if u and u != "Unknown" and (not uid or uid == "Unknown"):
            uid = u
        if mj and mj != "Unknown" and (not major or major == "Unknown"):
            major = mj

        # ── 3. Parse Courses and Carry Over Context ──
        page_sems, current_semester = parse_courses(text, current_semester)

        for sem, courses in page_sems.items():
            all_courses.setdefault(sem, [])
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
#  VERIFICATION LAYER
# ─────────────────────────────────────────────
#
#  verify_transcript() runs a suite of checks on a parsed transcript dict
#  and returns a list of VerificationIssue named-tuples.  Each issue has:
#
#    severity   "ERROR"   — data is missing or clearly wrong; output may be
#                           unusable.  The file is still exported but flagged.
#               "WARNING" — data looks suspicious or incomplete; human review
#                           recommended.
#               "INFO"    — neutral observation (e.g. IN PROGRESS courses).
#
#    field      Which field / column the issue relates to.
#    message    Human-readable description.
#    course     The course code that triggered the issue, or "" for
#               transcript-level issues.
#
#  verify_batch() aggregates issues across all transcripts in a batch and
#  prints a consolidated fidelity report to stdout at the end of processing.
# ─────────────────────────────────────────────

from collections import namedtuple

VerificationIssue = namedtuple("VerificationIssue", ["severity", "field", "message", "course"])

# Valid UIUC letter grades (including special codes)
_VALID_GRADES = {
    "A+","A","A-","B+","B","B-","C+","C","C-","D+","D","D-","F",
    "W","AW","DFR","CR","NC","ABS","IN PROGRESS",
}

# Reasonable credit hour bounds per course
_MIN_CREDITS = 0.5
_MAX_CREDITS = 6.0


def verify_transcript(transcript, course_map=None, uin_map=None):
    """
    Run the full verification suite on a parsed transcript dict.

    Checks performed
    ────────────────
    T1  Student name present and not 'Unknown'
    T2  University ID present, not 'Unknown', matches 9-digit pattern
    T3  Major present and not 'Unknown'
    T4  At least one semester was parsed
    T5  At least one course was parsed

    Per-course checks (for every course in every semester):
    C1  Course code format — SUBJ NNN pattern
    C2  Credit hours in expected range (_MIN_CREDITS – _MAX_CREDITS)
    C3  Grade is a recognised UIUC grade code
    C4  Grade points consistent with letter grade (spot-check A=4.0, F=0.0)
    C5  Course ID resolved from course_map (WARNING if missing)
    C6  Duplicate course detection within same semester

    Cross-field checks:
    X1  UIN found in uin_map (WARNING if provided map but no match)
    X2  Semester label format valid ("Season YYYY")

    Returns list[VerificationIssue]  (empty = all clear)
    """
    course_map = course_map or {}
    uin_map    = uin_map    or {}
    issues     = []

    def _add(severity, field, message, course=""):
        issues.append(VerificationIssue(severity, field, message, course))

    # ── T1–T5: Transcript-level ───────────────────────────────────────────────
    if not transcript["name"] or transcript["name"] == "Unknown":
        _add("ERROR", "Name", "Student name could not be extracted.")
    if not transcript["uid"] or transcript["uid"] == "Unknown":
        _add("ERROR", "UIN", "University ID could not be extracted.")
    elif not re.match(r"^\d{9}$", transcript["uid"]):
        _add("WARNING", "UIN",
             f"University ID '{transcript['uid']}' does not match expected 9-digit format.")

    if not transcript["major"] or transcript["major"] == "Unknown":
        _add("WARNING", "Major", "Major could not be extracted.")

    if not transcript["semesters"]:
        _add("ERROR", "Semesters", "No semesters were parsed from this transcript.")
        return issues   # nothing more to check

    all_courses_count = sum(len(c) for _, c in transcript["semesters"])
    if all_courses_count == 0:
        _add("ERROR", "Courses", "No courses were parsed from this transcript.")
        return issues

    # ── X1: UIN map lookup ────────────────────────────────────────────────────
    if uin_map and transcript["uid"] not in ("Unknown", ""):
        if transcript["uid"] not in uin_map:
            _add("WARNING", "ZJU ID",
                 f"UIN '{transcript['uid']}' not found in UIN→ZJUID reference map.")

    # ── X2 + per-course checks ────────────────────────────────────────────────
    _grade_pts = {"A+": 4.0, "A": 4.0, "A-": 3.67,
                  "B+": 3.33, "B": 3.0, "B-": 2.67,
                  "C+": 2.33, "C": 2.0, "C-": 1.67,
                  "D+": 1.33, "D": 1.0, "D-": 0.67, "F": 0.0}

    for semester_label, courses in transcript["semesters"]:
        # X2: semester format
        if not re.match(r"^(Fall|Spring|Summer|Winter)\s+\d{4}$",
                        semester_label, re.IGNORECASE):
            _add("WARNING", "Semester",
                 f"Unexpected semester label format: '{semester_label}'")

        seen_in_sem = {}   # code → first index, for duplicate detection
        for idx, course in enumerate(courses):
            code = f"{course['subj']} {course['number']}"

            # C1: course code format
            if not re.match(r"^[A-Z]{2,5}\s+\d{3}[A-Z]?$", code):
                _add("WARNING", "CourseCode",
                     f"Unexpected course code format: '{code}'", code)

            # C2: credit hours
            cred = course.get("credits", 0)
            if not (_MIN_CREDITS <= cred <= _MAX_CREDITS):
                _add("WARNING", "Credits",
                     f"Unusual credit value {cred} (expected {_MIN_CREDITS}–{_MAX_CREDITS})",
                     code)

            # C3: grade
            grade = course.get("grade", "")
            if grade not in _VALID_GRADES:
                _add("ERROR", "Grade",
                     f"Unrecognised grade '{grade}'", code)

            # C4: grade-points consistency (only for completed courses with pts)
            if grade in _grade_pts and course.get("points") is not None:
                expected_pts = round(_grade_pts[grade] * cred, 2)
                actual_pts   = round(course["points"], 2)
                tolerance    = 0.05
                if abs(expected_pts - actual_pts) > tolerance:
                    _add("WARNING", "GradePoints",
                         f"Points mismatch for {code}: "
                         f"grade {grade} × {cred}cr = {expected_pts:.2f} "
                         f"but transcript shows {actual_pts:.2f}", code)

            # C5: course ID resolution
            if course_map and code.upper() not in course_map:
                _add("INFO", "CourseID",
                     f"No course ID mapping found for '{code}'", code)

            # C6: duplicate detection
            if code in seen_in_sem:
                _add("WARNING", "Duplicate",
                     f"Course '{code}' appears more than once in {semester_label}", code)
            else:
                seen_in_sem[code] = idx

    return issues


def _format_verification_report(filename, issues):
    """Format issues for a single file into printable lines."""
    if not issues:
        return [f"  ✓ Verification passed — no issues found."]
    lines = []
    counts = {"ERROR": 0, "WARNING": 0, "INFO": 0}
    for issue in issues:
        counts[issue.severity] += 1
        course_tag = f" [{issue.course}]" if issue.course else ""
        lines.append(f"  {issue.severity:<8} {issue.field}{course_tag}: {issue.message}")
    summary = (f"  Totals → "
               f"{counts['ERROR']} error(s), "
               f"{counts['WARNING']} warning(s), "
               f"{counts['INFO']} info")
    return lines + [summary]


def compute_fidelity_score(all_issues_by_file):
    """
    Compute a 0–100 fidelity score across all processed files.

    Scoring:
        Each file starts at 100.
        ERROR   deducts 20 points (capped at 0).
        WARNING deducts  5 points (capped at 0).
        INFO    deducts  0 points.
    Final score = mean across all files.
    """
    if not all_issues_by_file:
        return 100.0
    file_scores = []
    for issues in all_issues_by_file:
        score = 100.0
        for issue in issues:
            if issue.severity == "ERROR":
                score -= 20
            elif issue.severity == "WARNING":
                score -= 5
        file_scores.append(max(0.0, score))
    return round(sum(file_scores) / len(file_scores), 1)


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


def _xw_formats(wb):
    """
    Create and return xlsxwriter Format objects for a given Workbook.
    Must be called per-workbook since formats are bound to their workbook instance.
    Returns a dict with keys: 'header', 'cell'.
    """
    header_fmt = wb.add_format({
        "bold":        True,
        "font_name":   "Arial",
        "font_size":   10,
        "bg_color":    "#FFFF00",
        "border":      1,
        "align":       "center",
        "valign":      "vcenter",
        "text_wrap":   True,
    })
    cell_fmt = wb.add_format({
        "font_name":   "Arial",
        "font_size":   10,
        "border":      1,
        "valign":      "vcenter",
    })
    return {"header": header_fmt, "cell": cell_fmt}


def _semester_to_year_term(semester_label):
    """'Fall 2023' → year='2023', ext_term='FALL'"""
    parts = semester_label.split()
    season, year = parts[0], parts[1]
    ext = SEASON_TO_EXT.get(season, season.upper())
    return year, ext


def _compute_term_code(season, year_str):
    """
    Compute the ZJU 接续学期 term code from a season name and 4-digit year string.

    Rule (derived from ZJU semester numbering convention):
        20XX FALL    ->  XX10      e.g. 2023 FALL   -> 2310
        20XX WINTER  ->  XX10      e.g. 2023 WINTER -> 2310
        20XX SPRING  ->  (XX-1)20  e.g. 2024 SPRING -> 2320
        20XX SUMMER  ->  (XX-1)50  e.g. 2024 SUMMER -> 2350

    Spring and Summer belong to the academic year that started the previous Fall,
    so their code prefix uses (year - 1)'s last two digits, not the calendar year.
    """
    try:
        year = int(year_str)
        xx   = year % 100
        s    = season.capitalize()
        if s in ("Fall", "Winter"):
            prefix, suffix = xx,        10
        elif s == "Spring":
            prefix, suffix = (xx - 1) % 100, 20
        elif s == "Summer":
            prefix, suffix = (xx - 1) % 100, 50
        else:
            return ""
        return f"{prefix:02d}{suffix}"
    except (ValueError, TypeError):
        return ""


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
    Load a course-code → course-ID lookup from a reference Excel or CSV file.

    Expected format (col indices, 0-based):
        Col 0  —  course code   e.g. "ECE 110"
        Col 1  —  course ID     e.g. 100101  or  "100101"  (may be int or str)

    The file may optionally have a header row — detected automatically.
    Column-header-based lookup is tried first (case-insensitive aliases);
    if headers are not recognised, col 0 / col 1 are used directly.

    Course IDs are always stored as strings.  IDs that end in '0' are often
    saved as integers by Excel (e.g. 100110 → 100110) or as strings
    (e.g. "100110").  Both are handled: the value is converted via
    str(int(v)) if it looks numeric, otherwise str(v).strip() is used,
    ensuring trailing zeros are never lost.

    Returns: { "ECE 110": {"course_id": "100101"}, ... }
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
            wb.close()
        elif ext == ".csv":
            import csv
            with open(ref_path, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.reader(f))
        else:
            print(f"  ✗ Course Code Map format not supported: {ext}")
            return mapping

        if not rows:
            return mapping

        # ── Detect header row ─────────────────────────────────────────────────
        code_aliases = {"course code", "course", "code", "学校课程编号",
                        "subj no", "subj_no", "course_code"}
        id_aliases   = {"course id", "courseid", "course_id", "课程id",
                        "zju course id", "课程id(course id)"}

        header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
        code_col = next((i for i, h in enumerate(header) if h in code_aliases), None)
        id_col   = next((i for i, h in enumerate(header) if h in id_aliases),   None)

        if code_col is not None and id_col is not None:
            # Recognised header row — skip it
            data_rows = rows[1:]
            print(f"  Course Code Map: using header-detected columns "
                  f"(code=col {code_col}, id=col {id_col})")
        else:
            # No recognised headers — treat col 0 as code, col 1 as ID
            data_rows = rows if len(rows[0]) < 2 or rows[0][0] is None else rows
            code_col, id_col = 0, 1
            # Skip first row only if it looks like a text header (non-numeric code)
            if data_rows and data_rows[0][0] is not None:
                first = str(data_rows[0][0]).strip()
                if not re.match(r'^[A-Z]{2,5}\s*\d{3}', first.upper()):
                    data_rows = data_rows[1:]  # skip header
            print(f"  Course Code Map: headers not recognised, using col 0 (code) / col 1 (id)")

        # ── Build mapping ─────────────────────────────────────────────────────
        def _normalise_id(raw):
            """
            Convert a cell value to a clean course-ID string.
            Handles:
              - int/float from Excel  (e.g. 100110  or 100110.0)
              - string               (e.g. "100110" or " 100110 ")
            Trailing zeros are preserved by converting through int first
            when the value is clearly numeric.
            """
            if raw is None:
                return ""
            if isinstance(raw, float):
                # Excel stores integers as floats (e.g. 100110.0)
                return str(int(raw)) if raw == int(raw) else str(raw).strip()
            if isinstance(raw, int):
                return str(raw)
            # String — strip whitespace; if it looks purely numeric, normalise
            s = str(raw).strip()
            if re.match(r'^\d+$', s):
                return s   # already a clean digit string, keep as-is
            if re.match(r'^\d+\.0+$', s):
                return s.split(".")[0]  # "100110.0" → "100110"
            return s

        for row in data_rows:
            if not row or len(row) <= max(code_col, id_col):
                continue
            code_raw = row[code_col]
            id_raw   = row[id_col]
            if code_raw is None:
                continue
            code_val = str(code_raw).strip().upper()
            id_val   = _normalise_id(id_raw)
            if code_val:
                mapping[code_val] = {"course_id": id_val}

        print(f"  ✓ Loaded {len(mapping)} course code mapping(s).")
        # Debug sample
        sample = list(mapping.items())[:3]
        for k, v in sample:
            print(f"    {k!r:20} → {v['course_id']!r}")

    except Exception as e:
        print(f"  ✗ Failed to load Course Code Map: {e}")

    return mapping


def load_term_map(ref_path):
    """
    Load Term Code Map reference file for 接续学期 lookup.

    Expected format (Excel or CSV) — two columns, case-insensitive headers:
        Col 1: term label  e.g. "2025 FALL", "2024 SPRING"  (header: 'term', 'semester', '学期' or similar)
        Col 2: term code   e.g. "2510"                       (header: 'code', 'term code', '接续学期' or similar)

    Lookup key is normalised to uppercase with whitespace collapsed so that
    "2025 fall", "2025 Fall", "2025  FALL" all resolve correctly.

    Returns: { "2025 FALL": "2510", "2024 SPRING": "2490", ... }
    """
    mapping = {}
    if not ref_path:
        return mapping

    ext = os.path.splitext(ref_path)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm", ".xls"):
            wb = openpyxl.load_workbook(ref_path, read_only=True, data_only=True)
            ws = wb.active
            rows = [tuple(cell for cell in row) for row in ws.iter_rows(values_only=True)]
            wb.close()
        elif ext == ".csv":
            import csv
            with open(ref_path, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.reader(f))
        else:
            print(f"  ✗ Term Code Map format not supported: {ext}")
            return mapping

        if not rows:
            return mapping

        # Normalise headers
        header = [str(h).strip().lower() if h else "" for h in rows[0]]

        term_aliases = {"term", "semester", "学期", "term label", "year term",
                        "external term", "外部学期", "col1", "column1"}
        code_aliases = {"code", "term code", "接续学期", "接续", "value",
                        "col2", "column2"}

        term_col = next((i for i, h in enumerate(header) if h in term_aliases), None)
        code_col = next((i for i, h in enumerate(header) if h in code_aliases), None)

        # Fallback: if headers unrecognised, assume col 0 = term, col 1 = code
        if term_col is None or code_col is None:
            if len(rows[0]) >= 2:
                print(f"  ⚠ Term Code Map: headers not recognised, assuming col 0 = term, col 1 = code.")
                print(f"    Headers found: {list(rows[0])}")
                term_col, code_col = 0, 1
                rows = rows[1:]   # skip header row only if we recognised it before; here skip nothing
            else:
                print(f"  ✗ Term Code Map: need at least 2 columns. Found: {list(rows[0])}")
                return mapping
        else:
            rows = rows[1:]  # skip recognised header row

        for row in rows:
            if not row or len(row) <= max(term_col, code_col):
                continue
            term_val = str(row[term_col]).strip() if row[term_col] is not None else ""
            code_val = str(row[code_col]).strip() if row[code_col] is not None else ""
            if term_val and code_val:
                # Normalise key: uppercase, collapse internal whitespace
                key = " ".join(term_val.upper().split())
                mapping[key] = code_val

        print(f"  ✓ Loaded {len(mapping)} term code mapping(s).")
    except Exception as e:
        print(f"  ✗ Failed to load Term Code Map: {e}")

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
    Stream transcript data to disk using xlsxwriter.

    xlsxwriter writes rows directly to the ZIP stream on disk — the workbook
    is never fully held in RAM regardless of row count. All formatting
    (header style, column widths, freeze pane, cell borders) is applied in a
    single pass without a second file-open.

    Requires: pip install xlsxwriter
    """
    if course_map is None:
        course_map = {}
    if term_map is None:
        term_map = {}

    try:
        wb = xlsxwriter.Workbook(output_path, {"constant_memory": True})
        ws = wb.add_worksheet("Transcript")
        fmt = _xw_formats(wb)

        # ── Column widths + freeze pane ──────────────────────────────
        for col_idx, width in enumerate(COL_WIDTHS):
            ws.set_column(col_idx, col_idx, width)
        ws.freeze_panes(1, 0)           # freeze row 1
        ws.set_row(0, 45)               # header row height

        # ── Header row ───────────────────────────────────────────────
        for col_idx, header in enumerate(HEADERS):
            ws.write(0, col_idx, header, fmt["header"])

        # ── Data rows — streamed one at a time ───────────────────────
        row_idx = 1
        for semester_label, courses in transcript["semesters"]:
            year, ext_term = _semester_to_year_term(semester_label)
            for course in courses:
                course_code = f"{course['subj']} {course['number']}"
                cm          = course_map.get(course_code.upper(), {})
                grd_scale   = course.get("grd_scale", "") or cm.get("评分基准", "")

                # 接续学期: computed from season+year using ZJU numbering formula
                # Fall/Winter: XX10 | Spring: (XX-1)20 | Summer: (XX-1)50
                season_label = semester_label.split()[0]
                term_code = _compute_term_code(season_label, year)

                row = [
                    zjuid,
                    transcript["uid"],
                    transcript["name"],
                    "UGRD",
                    "ZJUNV",
                    "1",
                    "UC001",
                    transcript["major"],
                    "UIUC",
                    "Y",
                    term_code,                                # K  接续学期
                    "1",
                    year,
                    ext_term,
                    course["subj"],
                    course["number"],
                    f"{course['subj']} {course['number']}",
                    course["credits"],
                    course["grade"],
                    cm.get("course_id", ""),
                    "UGS",
                    grd_scale,
                    course["credits"],
                    course["grade"],
                    "Y",
                ]
                for col_idx, val in enumerate(row):
                    ws.write(row_idx, col_idx, val, fmt["cell"])
                row_idx += 1

        wb.close()
        print(f"  ✓ Saved: {os.path.basename(output_path)}")

    except PermissionError:
        # File locked — fall through to safe_save retry dialog via a dummy openpyxl save
        print(f"  ⚠ xlsxwriter: permission denied on first attempt, retrying via safe_save…")
        import tempfile, shutil
        tmp = output_path + ".tmp.xlsx"
        wb2 = xlsxwriter.Workbook(tmp, {"constant_memory": True})
        ws2 = wb2.add_worksheet("Transcript")
        fmt2 = _xw_formats(wb2)
        for col_idx, width in enumerate(COL_WIDTHS):
            ws2.set_column(col_idx, col_idx, width)
        ws2.freeze_panes(1, 0)
        ws2.set_row(0, 45)
        for col_idx, header in enumerate(HEADERS):
            ws2.write(0, col_idx, header, fmt2["header"])
        row_idx = 1
        for semester_label, courses in transcript["semesters"]:
            year, ext_term = _semester_to_year_term(semester_label)
            for course in courses:
                course_code = f"{course['subj']} {course['number']}"
                cm          = course_map.get(course_code.upper(), {})
                grd_scale   = course.get("grd_scale", "") or cm.get("评分基准", "")
                row = [zjuid, transcript["uid"], transcript["name"], "UGRD", "ZJUNV",
                       "1", "UC001", transcript["major"], "UIUC", "Y",
                       _compute_term_code(semester_label.split()[0], year),
                       "1", year, ext_term,
                       course["subj"], course["number"],
                       f"{course['subj']} {course['number']}",
                       course["credits"], course["grade"], cm.get("course_id", ""),
                       "UGS", grd_scale, course["credits"], course["grade"], "Y"]
                for col_idx, val in enumerate(row):
                    ws2.write(row_idx, col_idx, val, fmt2["cell"])
                row_idx += 1
        wb2.close()
        # Now use safe_save logic by loading the tmp and saving to target
        openwb = openpyxl.load_workbook(tmp)
        if safe_save(openwb, output_path):
            print(f"  ✓ Saved (retry): {os.path.basename(output_path)}")
        try:
            os.remove(tmp)
        except Exception:
            pass
    except Exception as e:
        print(f"  ✗ Failed to write {os.path.basename(output_path)}: {e}")
        raise


def combine_all_excels(excel_files, output_folder, combined_filename="combined_transcripts.xlsx"):
    """
    Merge individual transcript xlsx files into one combined file.

    Pass 1 — read each source file with openpyxl read_only=True (streams rows
    from disk without loading the entire workbook into RAM).
    Pass 2 — write combined output with xlsxwriter constant_memory=True
    (streams rows to disk, never holds full workbook in RAM).
    """
    combined_path = os.path.join(output_folder, combined_filename)

    try:
        wb_out = xlsxwriter.Workbook(combined_path, {"constant_memory": True})
        ws_out = wb_out.add_worksheet("Combined")
        fmt    = _xw_formats(wb_out)

        for col_idx, width in enumerate(COL_WIDTHS):
            ws_out.set_column(col_idx, col_idx, width)
        ws_out.freeze_panes(1, 0)
        ws_out.set_row(0, 45)

        headers_written = False
        seen    = set()
        out_row = 0

        for file in excel_files:
            try:
                # read_only=True: rows are yielded lazily, never fully in RAM
                wb_in = openpyxl.load_workbook(file, read_only=True, data_only=True)
                ws_in = wb_in.active

                first_row = True
                for row in ws_in.iter_rows(values_only=True):
                    if first_row:
                        first_row = False
                        if not headers_written:
                            for col_idx, val in enumerate(row):
                                ws_out.write(out_row, col_idx, val, fmt["header"])
                            out_row += 1
                            headers_written = True
                        continue  # skip header rows from subsequent files

                    key = (row[1], row[15], row[12], row[13])
                    if key not in seen:
                        for col_idx, val in enumerate(row):
                            ws_out.write(out_row, col_idx, val if val is not None else "", fmt["cell"])
                        out_row += 1
                        seen.add(key)

                wb_in.close()

            except Exception as e:
                print(f"  ✗ Could not read {os.path.basename(file)}: {e}")

        wb_out.close()
        print(f"\n✓ Combined file saved: {combined_path}")

    except PermissionError:
        messagebox.showerror(
            "File Locked",
            f'Cannot write combined file:\n  {combined_filename}\n\n'
            f'Please close it in Excel and run again.'
        )
        print(f"  ✗ Combined file locked: {combined_filename}")


# ─────────────────────────────────────────────
#  PROCESSING PIPELINE
# ─────────────────────────────────────────────

def _process_one_pdf(args):
    """
    Top-level worker function for parallel execution via ProcessPoolExecutor.

    Must be defined at module level (not nested) so it is picklable by the
    multiprocessing machinery on all platforms. Returns a result dict that is
    safe to pass back across process boundaries (plain Python types only —
    no tkinter objects, no open file handles).

    Args:
        args: tuple of (pdf_path, output_folder, uin_map, course_map, term_map)

    Returns:
        dict with keys: pdf_path, out_xlsx (str|None), log (list[str]),
                        error (str|None)
    """
    pdf_path, output_folder, uin_map, course_map, term_map = args
    filename = os.path.basename(pdf_path)
    base     = os.path.splitext(filename)[0]
    log      = []
    out_xlsx = None

    try:
        transcript = process_pdf(pdf_path)

        # ── Detect unrecognised format ────────────────────────────────────────
        # If all key fields are Unknown and no courses were found, the PDF layout
        # was not recognised by the parser. Flag it instead of writing an empty file.
        course_count = sum(len(c) for _, c in transcript['semesters'])
        if (transcript['name'] == "Unknown"
                and transcript['uid'] == "Unknown"
                and course_count == 0):
            log.append(f"  ⚠ Unrecognised format — no data extracted.")
            return {"pdf_path": pdf_path, "out_xlsx": None,
                    "log": log, "error": None, "skipped": True, "issues": []}

        log.append(f"  Student  : {transcript['name']} ({transcript['uid']})")
        log.append(f"  Major    : {transcript['major']}")
        sem_count = len(transcript['semesters'])
        log.append(f"  Semesters: {sem_count}  |  Courses: {course_count}")

        zjuid = uin_map.get(transcript['uid'], "")
        log.append(f"  ZJU ID   : {zjuid if zjuid else '(not found in reference file)'}")

        # ── Verification ──────────────────────────────────────────────────────
        issues = verify_transcript(transcript, course_map=course_map, uin_map=uin_map)
        log.append(f"  Verification:")
        log.extend(_format_verification_report(filename, issues))

        out_xlsx = os.path.join(output_folder, f"{base}.xlsx")
        export_to_excel(transcript, out_xlsx, zjuid=zjuid,
                        course_map=course_map, term_map=term_map)
        log.append(f"  ✓ Saved  : {os.path.basename(out_xlsx)}")

    except PermissionError:
        log.append(f"  ✗ Skipped (file locked): {filename}")
        out_xlsx = None
    except Exception as e:
        log.append(f"  ✗ Error  : {e}")
        return {"pdf_path": pdf_path, "out_xlsx": None,
                "log": log, "error": str(e), "skipped": False, "issues": []}

    return {"pdf_path": pdf_path, "out_xlsx": out_xlsx,
            "log": log, "error": None, "skipped": False, "issues": issues}


def _detect_worker_count():
    """
    Return the number of parallel workers to use.

    Caps at min(cpu_count, 4) — beyond 4 the disk I/O of xlsxwriter writes
    becomes the bottleneck rather than CPU, so more workers add overhead
    without meaningful speedup for typical batch sizes.
    """
    try:
        cpus = multiprocessing.cpu_count()
    except Exception:
        cpus = 1
    return max(1, min(cpus, 4))


def _parallel_available():
    """
    Return True if ProcessPoolExecutor can safely be used on this platform.

    On Windows, multiprocessing requires the entry point to be protected by
    `if __name__ == '__main__'`. When running as a frozen PyInstaller exe or
    from a tkinter GUI (no __main__ guard), we fall back to threading to avoid
    recursive process spawning.

    On macOS with Python 3.8+ the default start method changed from 'fork' to
    'spawn', which has the same pickling requirement — safe as long as the
    worker function is at module level (which it is).
    """
    # Always safe on Linux (fork). Safe on Windows/macOS if not frozen.
    if getattr(sys, 'frozen', False):
        # PyInstaller exe: multiprocessing can work but requires careful setup.
        # Default to threading for safety unless explicitly tested.
        return False
    return True


def process_files(files, output_folder, progress_var, progress_label, root,
                  uin_map=None, course_map=None, term_map=None):
    """
    Process a list of PDF transcripts, writing one xlsx per file then combining.

    Strategy:
    - Single file  → direct sequential processing (no parallelism overhead)
    - Multiple files, parallel available → ProcessPoolExecutor with up to
      min(cpu_count, 4) workers; each PDF parsed + written in its own process
    - Multiple files, parallel not available (frozen exe, etc.) → sequential
      fallback with the same logic

    Progress and log messages are pushed back to the Tkinter thread via
    root.after() throughout.
    """
    excel_files   = []
    skipped_files = []   # files that produced no data (unrecognised format)
    all_issues    = []   # verification issues per successfully parsed file
    total         = len(files)
    uin_map       = uin_map    or {}
    course_map    = course_map or {}
    term_map      = term_map   or {}

    use_parallel = _parallel_available() and total > 1
    workers      = _detect_worker_count() if use_parallel else 1

    print(f"  Mode     : {'parallel (' + str(workers) + ' workers)' if use_parallel else 'sequential'}")
    print(f"  Files    : {total}")

    # Build args list once — dicts are picklable, fitz/xlsxwriter objects are not
    args_list = [
        (pdf_path, output_folder, uin_map, course_map, term_map)
        for pdf_path in files
    ]

    completed = 0

    def _on_result(result, index):
        """Called in the main thread after each worker result arrives."""
        nonlocal completed, excel_files, skipped_files, all_issues
        completed += 1
        filename = os.path.basename(result["pdf_path"])

        print(f"\n[{index}/{total}] {filename}")
        for line in result["log"]:
            print(line)

        if result.get("skipped"):
            skipped_files.append(filename)
        elif result["error"]:
            messagebox.showerror("Error",
                                 f"Failed to process {filename}:\n{result['error']}")
        else:
            if result["out_xlsx"]:
                excel_files.append(result["out_xlsx"])
            if result.get("issues") is not None:
                all_issues.append(result["issues"])

        pct = int(completed / total * 100)
        root.after(0, lambda p=pct, f=filename: (
            progress_var.set(p),
            progress_label.config(text=f"Processed: {f}")
        ))

    if use_parallel:
        # ── Parallel: submit all jobs, collect futures in submission order ──
        futures_ordered = []
        with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
            for args in args_list:
                futures_ordered.append(executor.submit(_process_one_pdf, args))

            for idx, future in enumerate(
                concurrent.futures.as_completed(futures_ordered), start=1
            ):
                try:
                    result = future.result()
                except Exception as e:
                    fname = os.path.basename(args_list[idx - 1][0])
                    result = {"pdf_path": args_list[idx - 1][0],
                              "out_xlsx": None,
                              "log": [f"  ✗ Worker crash: {e}"],
                              "error": str(e)}
                _on_result(result, idx)

    else:
        # ── Sequential fallback ───────────────────────────────────────────────
        for idx, args in enumerate(args_list, start=1):
            filename = os.path.basename(args[0])
            root.after(0, lambda f=filename: progress_label.config(
                text=f"Processing: {f}"))
            result = _process_one_pdf(args)
            _on_result(result, idx)

    # ── Combine ───────────────────────────────────────────────────────────────
    # Sort to preserve original file order in the combined output
    order = {path: i for i, path in enumerate(files)}
    excel_files.sort(key=lambda p: order.get(
        os.path.join(output_folder,
                     os.path.splitext(os.path.basename(p.replace(".xlsx", "")))[0] + ".pdf"),
        999))

    if len(excel_files) > 1:
        print("\nCombining all files…")
        combine_all_excels(excel_files, output_folder)
    elif len(excel_files) == 1:
        print("\n(Single file — no combined output needed)")

    root.after(0, lambda: (
        progress_var.set(100),
        progress_label.config(text=f"Done! Processed {len(excel_files)}/{total} file(s).")
    ))
    print(f"\n✓ All done. {len(excel_files)}/{total} file(s) exported.")

    # ── Fidelity score ────────────────────────────────────────────────────────
    if all_issues:
        score = compute_fidelity_score(all_issues)
        total_errors   = sum(1 for f in all_issues for i in f if i.severity == "ERROR")
        total_warnings = sum(1 for f in all_issues for i in f if i.severity == "WARNING")
        print(f"\n{'═'*60}")
        print(f"  ACCURACY / FIDELITY REPORT")
        print(f"  Files verified : {len(all_issues)}")
        print(f"  Total ERRORs   : {total_errors}")
        print(f"  Total WARNINGs : {total_warnings}")
        print(f"  Fidelity score : {score:.1f} / 100")
        if score >= 95:
            print(f"  ✓ High confidence — output is likely accurate.")
        elif score >= 80:
            print(f"  ⚠ Medium confidence — review warnings before submitting.")
        else:
            print(f"  ✗ Low confidence — errors detected, manual review required.")
        print(f"{'═'*60}")
    if skipped_files:
        print(f"\n{'─'*60}")
        print(f"⚠  {len(skipped_files)} file(s) could not be processed")
        print(f"   (unrecognised layout — no data was extracted):")
        for fname in skipped_files:
            print(f"   • {fname}")
        print(f"{'─'*60}")
        root.after(0, lambda: messagebox.showwarning(
            "Unprocessed Files",
            f"{len(skipped_files)} file(s) produced no output due to an\n"
            f"unrecognised transcript layout:\n\n"
            + "\n".join(f"  • {f}" for f in skipped_files)
            + "\n\nCheck the log for details."
        ))


def run_in_thread(files, output_folder, progress_var, progress_label, root,
                  btn_primary, btn_secondary, uin_map=None, course_map=None, term_map=None):
    btn_primary.config(state=tk.DISABLED)
    btn_secondary.config(state=tk.DISABLED)
    t = threading.Thread(
        target=lambda: process_files(files, output_folder, progress_var, progress_label,
                                     root, uin_map, course_map, term_map),
        daemon=True
    )
    t.start()

    def check():
        if t.is_alive():
            root.after(200, check)
        else:
            btn_primary.config(state=tk.NORMAL)
            btn_secondary.config(state=tk.NORMAL)

    root.after(200, check)


# ═══════════════════════════════════════════════════════════════════════════════
#  GUI  —  ZJUI-UIUC Transcript Parser v2.1
#
#  Layout (1280 × 800):
#    • Title banner   — Illini Orange (#FF5F05), bold white title + logo
#    • Working frame  — White card on Illini Blue background, split left/right
#        Left panel   — Drag-and-drop zone (icon centred) + Browse / Clear /
#                       Parse buttons below
#        Right panel  — Select Output Folder / UIN to ZJUID / Course Code Map /
#                       Term Code Map  (gold buttons; status updates go to log)
#    • Dialogue box   — scrolled log (all status messages printed here)
#    • Progress bar   — teal fill on Illini Blue track at bottom
#
#  Flow:
#    1. Browse / Drag-and-drop → files held in dropped_files list (no parse yet)
#    2. Click "Parse" → validation → run_in_thread → process_files
#
#  Colors:
#    ILLINI_BLUE   = #13294B   window bg, progress track
#    ILLINI_ORANGE = #FF5F05   title banner
#    CHINESE_GOLD  = #CA9D08   primary action buttons
#    LIGHT_RED     = #E05555   clear / destructive buttons
#    WHITE         = #FFFFFF   card, log bg, button text
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    global output_folder_path

    # ── Root window ──────────────────────────────────────────────────────────
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        _dnd_available = True
        root = TkinterDnD.Tk()
    except ImportError:
        _dnd_available = False
        root = tk.Tk()

    root.title("ZJUI-UIUC Transcript Parser v2.1")
    root.geometry("1280x800")
    root.resizable(False, False)

    try:
        icon_path = resource_path("icon.png")
        if os.path.exists(icon_path):
            root.iconphoto(False, tk.PhotoImage(file=icon_path))
    except Exception:
        pass

    # ── Color palette ─────────────────────────────────────────────────────────
    ILLINI_BLUE   = "#13294B"
    ILLINI_ORANGE = "#FF5F05"
    CHINESE_GOLD  = "#CA9D08"
    LIGHT_RED     = "#E05555"
    WHITE         = "#FFFFFF"
    GOLD_HOVER    = "#A07D06"
    RED_HOVER     = "#B83E3E"
    TEAL          = "#00C5A1"

    root.configure(bg=ILLINI_BLUE)

    # ── ttk styles ────────────────────────────────────────────────────────────
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure("Gold.TButton",
                    background=CHINESE_GOLD, foreground=WHITE,
                    font=("Segoe UI", 10, "bold"), padding=(10, 7),
                    borderwidth=0, focusthickness=0, relief="flat")
    style.map("Gold.TButton",
              background=[("active", GOLD_HOVER), ("pressed", GOLD_HOVER),
                          ("disabled", "#888855")],
              foreground=[("disabled", "#CCCCCC")],
              relief=[("pressed", "flat"), ("!pressed", "flat")])

    style.configure("Red.TButton",
                    background=LIGHT_RED, foreground=WHITE,
                    font=("Segoe UI", 10, "bold"), padding=(10, 7),
                    borderwidth=0, focusthickness=0, relief="flat")
    style.map("Red.TButton",
              background=[("active", RED_HOVER), ("pressed", RED_HOVER)],
              relief=[("pressed", "flat"), ("!pressed", "flat")])

    style.configure("Blue.Horizontal.TProgressbar",
                    troughcolor=ILLINI_BLUE, background=TEAL,
                    thickness=18, borderwidth=0)

    # ── State variables ───────────────────────────────────────────────────────
    output_folder_path = tk.StringVar()
    ref_file_path      = tk.StringVar()   # UIN → ZJU ID
    course_ref_path    = tk.StringVar()   # Course Code Map
    term_ref_path      = tk.StringVar()   # Term Code Map
    progress_var       = tk.IntVar(value=0)
    dropped_files      = []               # held until Parse is clicked

    # ── Map loaders ───────────────────────────────────────────────────────────
    def _get_uin_map():
        ref = ref_file_path.get()
        if ref:
            print(f"Loading UIN→ZJUID map:   {os.path.basename(ref)}")
            return load_uin_zjuid_map(ref)
        return {}

    def _get_course_map():
        ref = course_ref_path.get()
        if ref:
            print(f"Loading Course Code Map: {os.path.basename(ref)}")
            return load_course_map(ref)
        return {}

    def _get_term_map():
        ref = term_ref_path.get()
        if ref:
            print(f"Loading Term Code Map:   {os.path.basename(ref)}")
            return load_term_map(ref)
        return {}

    # ─────────────────────────────────────────────────────────────────────────
    #  TITLE BANNER
    # ─────────────────────────────────────────────────────────────────────────
    banner = tk.Frame(root, bg=ILLINI_ORANGE, height=90)
    banner.pack(fill=tk.X)
    banner.pack_propagate(False)

    # ZJU-UIUC logo — file: zju_logo.png  (placed alongside the .py or in bundle)
    try:
        logo_path = resource_path("zju_logo.png")
        if not os.path.exists(logo_path):
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     "zju_logo.png")
        if os.path.exists(logo_path):
            raw_logo = tk.PhotoImage(file=logo_path)
            # Target ~50px wide (≈60% of previous ~80px target).
            # subsample(n) reduces each dimension by factor n.
            # We compute n so that width / n ≈ 50px.
            scale = max(1, raw_logo.width() // 100)
            logo_img = raw_logo.subsample(scale, scale)
            lbl_logo = tk.Label(banner, image=logo_img, bg=ILLINI_ORANGE, bd=0)
            lbl_logo.image = logo_img   # prevent GC
            lbl_logo.pack(side=tk.LEFT, padx=(14, 8), pady=5)
    except Exception:
        pass

    tk.Label(banner,
             text="ZJUI-UIUC Transcript Parser",
             font=("Segoe UI", 26, "bold"),
             fg=WHITE, bg=ILLINI_ORANGE).pack(side=tk.LEFT, padx=6)

    # ─────────────────────────────────────────────────────────────────────────
    #  MAIN WORKING AREA  (white card)
    # ─────────────────────────────────────────────────────────────────────────
    card = tk.Frame(root, bg=WHITE, bd=0)
    card.pack(fill=tk.BOTH, expand=True, padx=28, pady=18)

    # ── Left panel ───────────────────────────────────────────────────────────
    left = tk.Frame(card, bg=WHITE)
    left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(16, 8), pady=16)

    # Drag-and-drop zone
    DND_BG = "#D8D8D8"
    dnd_frame = tk.Frame(left, bg=DND_BG, bd=0)
    dnd_frame.pack(fill=tk.BOTH, expand=True)

    # Centred icon + label inside DND zone
    dnd_inner = tk.Frame(dnd_frame, bg=DND_BG)
    dnd_inner.place(relx=0.5, rely=0.45, anchor="center")

    # Load dnd_icon.png; scale to ~80px; fall back to Unicode symbol
    _dnd_icon_ref = None
    try:
        dnd_icon_path = resource_path("dnd_icon.png")
        if not os.path.exists(dnd_icon_path):
            dnd_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                         "dnd_icon.png")
        if os.path.exists(dnd_icon_path):
            raw_icon = tk.PhotoImage(file=dnd_icon_path)
            # Target ~80px wide; subsample to get close
            scale = max(1, raw_icon.width() // 80)
            _dnd_icon_ref = raw_icon.subsample(scale, scale)
            lbl_icon = tk.Label(dnd_inner, image=_dnd_icon_ref, bg=DND_BG, bd=0)
            lbl_icon.image = _dnd_icon_ref   # prevent GC
            lbl_icon.pack(pady=(0, 8))
        else:
            raise FileNotFoundError
    except Exception:
        tk.Label(dnd_inner, text="⤵",
                 font=("Segoe UI", 48), fg="#888888", bg=DND_BG).pack()

    tk.Label(dnd_inner, text="Drag and Drop",
             font=("Segoe UI", 13), fg="#555555", bg=DND_BG).pack()

    dnd_status = tk.Label(dnd_frame, text="No files selected",
                          font=("Segoe UI", 9, "italic"),
                          fg="#666666", bg=DND_BG)
    dnd_status.pack(side=tk.BOTTOM, pady=(0, 8))

    # ── Button row: Browse | Clear | (spacer) | Parse ─────────────────────
    btn_row = tk.Frame(left, bg=WHITE)
    btn_row.pack(fill=tk.X, pady=(10, 0))

    def _do_browse():
        files = filedialog.askopenfilenames(
            title="Select PDF files", filetypes=[("PDF Files", "*.pdf")]
        )
        if files:
            nonlocal dropped_files
            dropped_files = list(files)
            count = len(dropped_files)
            dnd_status.config(text=f"{count} file(s) selected — click Parse to begin",
                              fg="#333333")
            print(f"Selected {count} file(s):")
            for f in dropped_files:
                print(f"  {os.path.basename(f)}")

    def _do_clear():
        nonlocal dropped_files
        dropped_files = []
        dnd_status.config(text="No files selected", fg="#666666")
        progress_var.set(0)
        lbl_progress.config(text="Idle")
        print("File selection cleared.")

    def _do_parse():
        if not output_folder_path.get():
            messagebox.showwarning("No Output Folder",
                                   "Please select an output folder first.")
            return
        if not dropped_files:
            messagebox.showwarning("No Files",
                                   "No PDF files selected. Use Browse or drag-and-drop first.")
            return
        progress_var.set(0)
        lbl_progress.config(text="Starting…")
        print(f"\n{'─'*60}")
        print(f"Parsing {len(dropped_files)} file(s)…")
        run_in_thread(list(dropped_files), output_folder_path.get(),
                      progress_var, lbl_progress, root,
                      btn_parse, btn_browse,
                      uin_map=_get_uin_map(),
                      course_map=_get_course_map(),
                      term_map=_get_term_map())

    btn_browse = ttk.Button(btn_row, text="Browse File(s)",
                            command=_do_browse, style="Gold.TButton")
    btn_clear  = ttk.Button(btn_row, text="Clear",
                            command=_do_clear, style="Red.TButton")
    btn_parse  = ttk.Button(btn_row, text="▶  Parse",
                            command=_do_parse, style="Gold.TButton")

    btn_browse.pack(side=tk.LEFT, padx=(0, 6))
    btn_clear.pack(side=tk.LEFT)
    btn_parse.pack(side=tk.RIGHT)   # Parse on the right end of the row

    # Vertical separator
    tk.Frame(card, bg="#CCCCCC", width=1).pack(side=tk.LEFT, fill=tk.Y, pady=16)

    # ── Right panel: reference buttons (no sub-labels; status → log) ─────────
    right = tk.Frame(card, bg=WHITE, width=220)
    right.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 16), pady=16)
    right.pack_propagate(False)

    def _pick_output():
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            output_folder_path.set(folder)
            print(f"Output folder: {folder}")

    def _pick_ref(path_var, label, title):
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("All Files", "*.*")]
        )
        if path:
            path_var.set(path)
            print(f"{label}: {os.path.basename(path)}")

    ttk.Button(right, text="Select Output Folder", style="Gold.TButton",
               command=_pick_output).pack(fill=tk.X, pady=(0, 12))

    ttk.Button(right, text="UIN to ZJUID", style="Gold.TButton",
               command=lambda: _pick_ref(ref_file_path,
                                         "UIN→ZJUID map",
                                         "Select UIN→ZJUID Reference")
               ).pack(fill=tk.X, pady=(0, 12))

    ttk.Button(right, text="Course Code Map", style="Gold.TButton",
               command=lambda: _pick_ref(course_ref_path,
                                         "Course Code Map",
                                         "Select Course Code Map")
               ).pack(fill=tk.X, pady=(0, 12))

    ttk.Button(right, text="Term Code Map (Optional)", style="Gold.TButton",
               command=lambda: _pick_ref(term_ref_path,
                                         "Term Code Map",
                                         "Select Term Code Map")
               ).pack(fill=tk.X, pady=(0, 12))

    # ─────────────────────────────────────────────────────────────────────────
    #  DIALOGUE BOX  (all status messages printed here via StdoutRedirector)
    # ─────────────────────────────────────────────────────────────────────────
    log_frame = tk.Frame(root, bg=ILLINI_BLUE)
    log_frame.pack(fill=tk.BOTH, expand=True, padx=28, pady=(0, 6))

    txt_box = ScrolledText(log_frame, font=("Consolas", 10),
                           state="disabled", bg=WHITE, fg="#222222",
                           relief="flat", borderwidth=0, height=10)
    txt_box.pack(fill=tk.BOTH, expand=True)

    sys.stdout = StdoutRedirector(txt_box)
    sys.stderr = StdoutRedirector(txt_box)

    # ─────────────────────────────────────────────────────────────────────────
    #  PROGRESS BAR
    # ─────────────────────────────────────────────────────────────────────────
    prog_frame = tk.Frame(root, bg=ILLINI_BLUE)
    prog_frame.pack(fill=tk.X, padx=28, pady=(0, 6))

    lbl_progress = tk.Label(prog_frame, text="Idle",
                            font=("Segoe UI", 9), fg=WHITE, bg=ILLINI_BLUE, anchor="w")
    lbl_progress.pack(fill=tk.X)

    pbar = ttk.Progressbar(prog_frame, variable=progress_var,
                           maximum=100, mode="determinate",
                           style="Blue.Horizontal.TProgressbar")
    pbar.pack(fill=tk.X)

    # ─────────────────────────────────────────────────────────────────────────
    #  COPYRIGHT
    # ─────────────────────────────────────────────────────────────────────────
    tk.Label(root, text="© 2025 Jaden Peterson Wen   © 2026 Junway Lin",
             font=("Segoe UI", 8, "italic"),
             fg="#7799BB", bg=ILLINI_BLUE).pack(anchor="e", padx=10, pady=(0, 4))

    # ─────────────────────────────────────────────────────────────────────────
    #  DRAG-AND-DROP BINDING
    # ─────────────────────────────────────────────────────────────────────────
    def _on_drop(event):
        raw = event.data
        paths = re.findall(r'\{([^}]+)\}|(\S+)', raw)
        paths = [a or b for a, b in paths]
        pdf_files = []
        for p in paths:
            if os.path.isdir(p):
                pdf_files += [os.path.join(p, f)
                               for f in sorted(os.listdir(p))
                               if f.lower().endswith(".pdf")]
            elif p.lower().endswith(".pdf"):
                pdf_files.append(p)
        if not pdf_files:
            messagebox.showwarning("No PDFs", "No PDF files found in the dropped items.")
            return
        nonlocal dropped_files
        dropped_files = pdf_files
        count = len(dropped_files)
        dnd_status.config(text=f"{count} file(s) ready — click Parse to begin",
                          fg="#333333")
        print(f"Dropped {count} file(s):")
        for f in dropped_files:
            print(f"  {os.path.basename(f)}")

    if _dnd_available:
        from tkinterdnd2 import DND_FILES
        for widget in (dnd_frame, dnd_inner):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", _on_drop)

    # ─────────────────────────────────────────────────────────────────────────

    root.mainloop()


if __name__ == "__main__":
    # Required on Windows when using ProcessPoolExecutor in a frozen exe (PyInstaller).
    # On Linux/macOS this is a no-op.
    multiprocessing.freeze_support()
    main()