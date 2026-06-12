# ZJUI-UIUC Transcript Parser

**Version:** 2.3  
**Author:** Junway Lin (© 2026)

A desktop GUI tool that extracts course data from UIUC (University of Illinois Urbana-Champaign) PDF transcripts and converts it into structured Excel files formatted for import into the Zhejiang University International (ZJUI) student system. This program builds on the original First-Generation Parser developed by Jaden Wen Peterson.

---

## Features

- **PDF parsing** — Extracts student name, UIN, major, and full course history from multi-page UIUC transcripts.
- **Two-column layout support** — Correctly handles UIUC's two-column transcript format by enforcing left-column-first, right-column-second reading order.
- **Excel export** — Outputs one `.xlsx` file per transcript, formatted to match the ZJU course import template (25 columns, Chinese headers).
- **Batch processing** — Processes multiple PDFs in parallel (up to 4 workers) and merges all results into a single `combined_transcripts.xlsx`.
- **Reference file support** — Optionally load three reference files to enrich the output:
  - **UIN → ZJU ID map** — Maps UIUC student IDs to ZJU student numbers.
  - **Course Code Map** — Maps UIUC course codes (e.g. `ECE 110`) to ZJU internal course IDs.
  - **Term Code Map** — Maps semester labels (e.g. `2025 FALL`) to ZJU term codes.
- **Verification layer** — Automatically checks each parsed transcript for errors and warnings (missing fields, invalid grades, credit anomalies, duplicates) and produces a fidelity score out of 100.
- **Drag-and-drop interface** — Built with Tkinter; supports drag-and-drop via `tkinterdnd2` when available.
- **In-progress course handling** — Detects and records courses marked `IN PROGRESS` without a grade or points.

---

## Requirements

### Python version
Python 3.8 or higher.

### Dependencies

Install all required packages with:

```bash
pip install pymupdf openpyxl xlsxwriter
```

| Package | Purpose |
|---|---|
| `PyMuPDF` (`fitz`) | PDF reading and text block extraction |
| `openpyxl` | Reading reference Excel files; fallback save logic |
| `xlsxwriter` | Streaming Excel output (low memory usage) |
| `tkinterdnd2` *(optional)* | Drag-and-drop support in the GUI |

`tkinter` is included in most standard Python distributions. If missing, install it via your system package manager (e.g. `sudo apt install python3-tk` on Ubuntu).

To enable drag-and-drop:

```bash
pip install tkinterdnd2
```

---

## Installation

1. Clone or download this repository.
2. Install the required packages (see above).
3. Place the following optional image assets in the same directory as the script if you have them:
   - `icon.png` — application window icon
   - `zju_logo.png` — logo shown in the title banner
   - `dnd_icon.png` — icon shown in the drag-and-drop zone

---

## Usage

### Running the application

```bash
python OCR_v2_3.py
```

### Step-by-step workflow

1. **Select output folder** — Click **Select Output Folder** on the right panel to choose where `.xlsx` files will be saved.
2. **Load reference files** *(optional but recommended)*:
   - **UIN to ZJUID** — An Excel or CSV file mapping UIUC UINs to ZJU student IDs.
   - **Course Code Map** — An Excel or CSV file mapping course codes to ZJU course IDs.
   - **Term Code Map** — An Excel or CSV file mapping semester labels to ZJU term codes.
3. **Add PDF files** — Either drag-and-drop PDFs onto the drop zone, or click **Browse File(s)** to select them via a file dialog. You may also drop a folder; all PDFs inside it will be added.
4. **Parse** — Click **▶ Parse** to begin processing. Progress is shown in the progress bar and log window at the bottom.
5. **Check the output folder** — Each PDF produces one `.xlsx` file. If more than one PDF was processed, a `combined_transcripts.xlsx` is also created.

---

## Reference File Formats

### UIN → ZJU ID map

Excel (`.xlsx`, `.xlsm`, `.xls`) or CSV with at least two columns. Headers are case-insensitive.

| UIN | ZJU ID |
|---|---|
| 612345678 | 3220001234 |
| 623456789 | 3220005678 |

Accepted header aliases: `UIN`, `University ID` / `学号`, `学号(ZJU ID)`, `zjuid`.

### Course Code Map

Excel or CSV with two columns: course code and ZJU course ID.

| Course Code | Course ID |
|---|---|
| ECE 110 | 100101 |
| CS 225 | 100202 |

Accepted header aliases: `Course Code`, `Code`, `学校课程编号` / `Course ID`, `CourseID`, `课程ID`.

### Term Code Map *(optional)*

Excel or CSV mapping semester labels to ZJU term codes.

| Term | Code |
|---|---|
| 2025 FALL | 2510 |
| 2026 SPRING | 2520 |

Accepted header aliases: `Term`, `Semester`, `学期` / `Code`, `Term Code`, `接续学期`.

> **Note:** If no Term Code Map is provided, term codes are computed automatically using ZJU's semester numbering formula:  
> `Fall/Winter YYYY → YY10` | `Spring YYYY → (YY-1)20` | `Summer YYYY → (YY-1)50`

---

## Output Format

Each output `.xlsx` contains one row per course with the following 25 columns:

| Column | Header | Description |
|---|---|---|
| A | 学号(ZJU ID) | ZJU student number (from reference file) |
| B | UIN | UIUC University ID |
| C | Name | Student name |
| D | 学历 | Degree level (fixed: `UGRD`) |
| E | 学术机构 | Academic institution (fixed: `ZJUNV`) |
| F | 转校模型编号 | Transfer model number (fixed: `1`) |
| G | 学科设定 | Subject setting (fixed: `UC001`) |
| H | 专业设定(Major) | Major from transcript |
| I | 来源机构 | Source institution (fixed: `UIUC`) |
| J | 包含在平均成绩 | Include in GPA (fixed: `Y`) |
| K | 接续学期 | ZJU term code |
| L | 抵免同等组 | Credit equivalency group (fixed: `1`) |
| M | 学年 | Academic year |
| N | 外部学期 | External term (e.g. `FALL`, `SPRING`) |
| O | 学校学科 | Course subject (e.g. `ECE`) |
| P | 学校课程编号 | Course number (e.g. `110`) |
| Q | 描述 | Course description (`SUBJ NNN`) |
| R | 已获分(Credit) | Credit hours earned |
| S | 成绩输入(Grade) | Letter grade |
| T | 课程ID(Course ID) | ZJU internal course ID (from reference file) |
| U | 评分计划 | Grading plan (fixed: `UGS`) |
| V | 评分基准 | Grade scale |
| W | 已抵免 | Credits transferred |
| X | 正式成绩 | Official grade |
| Y | 包括在学分FA WI统计中 | Include in FA/WI credit count (fixed: `Y`) |

---

## Verification and Fidelity Score

After processing, the tool runs automated checks on every parsed transcript and prints a report to the log. Each issue has a severity level:

- **ERROR** — Critical problem (e.g. name or UIN not found, no courses parsed). Deducts 20 points from the fidelity score.
- **WARNING** — Suspicious data (e.g. unusual credits, grade/points mismatch, duplicate course). Deducts 5 points.
- **INFO** — Neutral observation (e.g. course not found in Course Code Map). No deduction.

The **fidelity score** (0–100) is the average per-file score across all processed transcripts:

| Score | Confidence level |
|---|---|
| ≥ 95 | ✓ High — output is likely accurate |
| 80–94 | ⚠ Medium — review warnings before submitting |
| < 80 | ✗ Low — errors detected, manual review required |

---

## Notes

- The parser is specifically designed for **UIUC official transcripts**. Other universities' transcript formats are not supported and will be flagged as unrecognised.
- If a PDF cannot be parsed (all fields `Unknown`, no courses found), it is skipped and listed in a warning dialog at the end.
- On Windows, if an output `.xlsx` file is open in Excel when the tool tries to save it, a retry dialog will appear asking you to close it first.
- When running as a frozen PyInstaller executable, parallel processing is automatically disabled and the tool falls back to sequential mode.

---

## License

© 2025 Jaden Peterson Wen · © 2026 Junway Lin. All rights reserved.
