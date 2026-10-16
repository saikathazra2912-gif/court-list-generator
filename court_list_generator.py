# -*- coding: utf-8 -*-
"""
NEXT-DAY COURT LIST GENERATOR
For West Bengal correctional-home Court List workbooks.

Put this file, COURT_TEMPLATE.xlsx, and a PrisonersList*.xls/.xlsx file
in the same folder. Double-click Run_Court_List_Generator.bat.

Default output date = TOMORROW.
The source .xls used by the user's office is an HTML table saved with .xls
extension, so this script reads that format without requiring xlrd.
"""

from __future__ import annotations

import copy
import html
import os
import re
import sys
from datetime import date, datetime, timedelta
from html.parser import HTMLParser
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

from openpyxl import load_workbook, Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# USER SETTINGS
# ============================================================

# Exact visual template. Keep COURT_TEMPLATE.xlsx beside this script.
TEMPLATE_NAME = "COURT_TEMPLATE.xlsx"

# "tomorrow" is the normal one-click behavior.
# Change to "ask" if you want a date dialog every time.
DATE_MODE = "tomorrow"

# The inspected reference workbook is PORTRAIT.
# Change to "landscape" if your office printout must be landscape.
PAGE_ORIENTATION = "landscape"

# A-Z sorting is applied inside every court.
SORT_BY = "Prisoner Name"

# True = duplicate JID inside the same court gets the same SL.
# False = every prisoner gets a new SL.
SAME_JID_SAME_SL_WITHIN_COURT = True

# If you literally want the same SL globally even when the same JID
# appears in different courts, set this to True.
# NOTE: that can create gaps in the serial numbers of a court section.
SAME_JID_SAME_SL_GLOBALLY = False

# Conditional formatting: duplicate values in JID, Prisoner Name,
# Father Name, FIR/Case No, and Court Name.
ADD_DUPLICATE_CONDITIONAL_FORMATTING = True

# Output name pattern.
OUTPUT_PATTERN = "COURT LIST ON {date}.xlsx"


# ============================================================
# COURT AREA ORDER REQUESTED BY USER
# ============================================================

AREA_ORDER = [
    "Alipore",
    "Calcutta",
    "Baruipur",
    "Diamond Harbour",
    "Kakdwip",
    "Howrah",
    "Chandannagar",
    "Hooghly",
    "Barasat",
    "Barrackpore",
    "Bidhannagar",
    "MPs & MLAs Court",
    "JJB Salt Lake",
    "Basirhat",
    "Basirahat",
    "Bongaon",
    "Kalyani",
    "Ranaghat",
    "Krishnanagar",
    "others",
]

# Aliases/normalization. The source file can use different spellings.
AREA_ALIASES = {
    "BARRACKPORE": "Barrackpore",
    "BARRACKPUR": "Barrackpore",
    "BASIRHAT": "Basirhat",
    "BASIRAHAT": "Basirahat",
    "BASIRAHAT": "Basirahat",
    "BASIRHAT": "Basirhat",
    "BONGAON": "Bongaon",
    "BONGAON": "Bongaon",
    "BARASAT": "Barasat",
    "BIDHANNAGAR": "Bidhannagar",
    "SALT LAKE": "JJB Salt Lake",
    "JJB SALT LAKE": "JJB Salt Lake",
    "KALYANI": "Kalyani",
    "RANAGHAT": "Ranaghat",
    "KRISHNANAGAR": "Krishnanagar",
    "ALIPORE": "Alipore",
    "CALCUTTA": "Calcutta",
    "KOLKATA": "Calcutta",
    "BARUIPUR": "Baruipur",
    "DIAMOND HARBOUR": "Diamond Harbour",
    "DIAMOND HARBOR": "Diamond Harbour",
    "KAKDWIP": "Kakdwip",
    "HOWRAH": "Howrah",
    "CHANDANNAGAR": "Chandannagar",
    "HOOGHLY": "Hooghly",
    "MP & MLA": "MPs & MLAs Court",
    "MPS & MLAS": "MPs & MLAs Court",
    "MPS & MLAS COURT": "MPs & MLAs Court",
}


# ============================================================
# COURT WITHIN-AREA ORDER
# ============================================================
# Specific courts can be listed here by exact/partial text.
# Lower number = earlier in that area.
#
# The requested Barasat sequence is explicitly encoded first.
# Similar keyword rules are used for the other areas.

BARASAT_ORDER = [
    ("CHIEF JUDICIAL MAGISTRATE", 10),
    ("ACJM", 20),
    ("J.M 1ST CLASS, 1ST COURT", 30),
    ("J.M., 1ST COURT", 31),
    ("J.M. 1ST COURT", 31),
    ("JM 1ST CLASS", 30),
    ("JM 1ST COURT", 31),
    ("JM 2ND COURT", 32),
    ("JM 3RD COURT", 33),
    ("JM 4TH COURT", 34),
    ("JM 5TH COURT", 35),
    ("POCSO", 40),
    ("ADJ., 1ST COURT", 50),
    ("ADJ., 2ND COURT", 51),
    ("ADJ., 3RD COURT", 52),
    ("ADJ., 4TH COURT", 53),
    ("ADJ., 5TH COURT", 54),
    ("ADJFTC NO. 1", 60),
    ("ADJFTC NO. 2", 61),
    ("ADJFTC NO. 3", 62),
    ("ADJFTC NO. 4", 63),
    ("ADJFTC NO. 5", 64),
    ("SESSION JUDGE", 70),
    ("ADJ., 6TH COURT", 71),
    ("ADJ., 7TH COUR", 71.5),
    ("ADDL.DISTRICT & SESSIONS JUDGE 8TH", 72),
    ("ADJ 8TH COURT", 72),
]

AREA_GENERIC_ORDER = {
    "Barasat": BARASAT_ORDER,
    "Barrackpore": [
        ("ACJM", 10),
        ("SPL EXEC MAGISTRATE", 20),
        ("SPECIAL JUDGE POCSO", 30),
        ("SPL JUD COURT", 40),
        ("ADJ., 1ST COURT", 50),
        ("ADJ., 2ND COURT", 51),
        ("ADJ., 3RD COURT", 52),
        ("ADJ., 4TH COURT", 53),
        ("ADJ., 5TH COURT", 54),
        ("ADSJ FTC NO. 1", 60),
        ("ADSJ FTC NO. 2", 61),
        ("ADSJ FTC NO. 3", 62),
        ("ADSJ FTC NO. 4", 63),
        ("ADSJ FTC NO. 5", 64),
    ],
    "Bidhannagar": [
        ("CHIEF JUDICIAL MAGISTRATE", 10),
        ("ACJM", 20),
        ("JM", 30),
        ("POCSO", 40),
        ("ADJ", 50),
        ("SESSIONS", 60),
    ],
    "Basirhat": [
        ("CHIEF JUDICIAL MAGISTRATE", 10),
        ("ACJM", 20),
        ("J.M 1ST CLASS, 1ST COURT", 30),
        ("J.M., 1ST COURT", 31),
        ("JM 1ST COURT", 31),
        ("J.M., 2ND COURT", 32),
        ("JM 2ND COURT", 32),
        ("POCSO", 40),
        ("ADJ", 50),
        ("FTC NO. 1", 60),
        ("FTC NO. 2", 61),
        ("FTC NO. 3", 62),
        ("FTC NO. 4", 63),
        ("FTC NO. 5", 64),
    ],
    "Basirahat": [
        ("CHIEF JUDICIAL MAGISTRATE", 10),
        ("ACJM", 20),
        ("J.M 1ST CLASS, 1ST COURT", 30),
        ("J.M., 1ST COURT", 31),
        ("JM 1ST COURT", 31),
        ("J.M., 2ND COURT", 32),
        ("JM 2ND COURT", 32),
        ("POCSO", 40),
        ("ADJ", 50),
        ("FTC NO. 1", 60),
        ("FTC NO. 2", 61),
        ("FTC NO. 3", 62),
        ("FTC NO. 4", 63),
        ("FTC NO. 5", 64),
    ],
    "Bongaon": [
        ("ACJM", 20),
        ("JM", 30),
        ("POCSO", 40),
        ("ADJ", 50),
        ("FTC", 60),
    ],
    "Krishnanagar": [
        ("ACJM", 20),
        ("JM", 30),
        ("POCSO", 40),
        ("ADJ", 50),
        ("NDPS", 55),
        ("SESSIONS", 60),
    ],
}


# ============================================================
# HELPERS
# ============================================================

def clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return html.unescape(str(value)).strip()


def norm(value) -> str:
    s = clean_text(value).upper()
    s = re.sub(r"\s+", " ", s)
    return s


def natural_key(value):
    s = norm(value)
    return [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", s)]


def find_area(court_name: str) -> str:
    s = norm(court_name)

    # Preserve the user's requested distinction between BASIRHAT and
    # BASIRAHAT when both spellings occur in the source workbook.
    if "BASIRAHAT" in s:
        return "Basirahat"
    if "BASIRHAT" in s:
        return "Basirhat"

    # Exact priority for the requested areas.
    # Longer strings first avoids partial collisions.
    aliases = sorted(AREA_ALIASES.items(), key=lambda x: len(x[0]), reverse=True)
    for alias, area in aliases:
        if alias in s:
            return area

    return "others"


def area_rank(court_name: str) -> int:
    area = find_area(court_name)
    return AREA_ORDER.index(area) if area in AREA_ORDER else len(AREA_ORDER) - 1


def court_rank_within_area(court_name: str) -> tuple:
    area = find_area(court_name)
    s = norm(court_name)

    # Required precedence within an area:
    # 1) Magistrate courts whose names contain JM / J.M. before POCSO
    # 2) POCSO before ADJ / A.D.J. courts
    # Normalise punctuation so J.M., J.M and JM are treated alike.
    compact = re.sub(r"[^A-Z0-9]+", "", s)
    is_jm_court = bool(re.search(r"(^|[^A-Z])J\.?\s*M\.?([^A-Z]|$)", s)) or "JM" in compact
    if is_jm_court:
        return (35, natural_key(court_name))
    if "POCSO" in s:
        return (40, natural_key(court_name))
    for keyword, rank in AREA_GENERIC_ORDER.get(area, []):
        if keyword in s:
            return (rank, natural_key(court_name))

    return (10000, natural_key(court_name))


def court_sort_key(court_name: str):
    return (area_rank(court_name), *court_rank_within_area(court_name))


def transformed_name(row) -> str:
    name = clean_text(row["Prisoner Name"])
    jid = clean_text(row["JID"])
    gender = norm(row["Gender"])

    # Avoid adding the suffix twice if a future source already contains it.
    if re.search(r"-RP\s*$", jid, flags=re.I) and not re.search(r"\(RP\)\s*$", name, flags=re.I):
        name += " (RP)"

    if gender == "H" and not re.search(r"\(TRANS\)\s*$", name, flags=re.I):
        name += " (TRANS)"

    return name


def transformed_location(value) -> str:
    s = clean_text(value)
    u = norm(s)

    replacements = {
        "FEMALE HOSPITAL": "FH",
        "FEMALE1": "F",
        "FEMALE 1": "F",
        "TB WARD": "TB",
        "TBWARD": "TB",
    }

    for old, new in replacements.items():
        if u == old:
            return new

    # Normal ward/barrack numbers are preserved.
    # The displayed column names W and B are handled in the workbook.
    return s


# ============================================================
# READ SOURCE FILE
# ============================================================

class SimpleTableParser(HTMLParser):
    """Small stdlib HTML table reader for office-generated .xls HTML files."""

    def __init__(self):
        super().__init__()
        self.rows = []
        self.current_row = None
        self.current_cell = None
        self.in_table = False

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "table":
            self.in_table = True
        elif self.in_table and tag == "tr":
            self.current_row = []
        elif self.in_table and tag in ("td", "th"):
            self.current_cell = []

    def handle_data(self, data):
        if self.current_cell is not None:
            self.current_cell.append(data)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("td", "th") and self.current_cell is not None:
            value = re.sub(r"\s+", " ", "".join(self.current_cell)).strip()
            self.current_row.append(value)
            self.current_cell = None
        elif tag == "tr" and self.current_row is not None:
            self.rows.append(self.current_row)
            self.current_row = None
        elif tag == "table":
            self.in_table = False


def read_source(path: Path):
    suffix = path.suffix.lower()

    # Normal modern Excel file.
    if suffix in (".xlsx", ".xlsm"):
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        return rows

    # The office PrisonersList .xls supplied with this task is HTML.
    raw = path.read_text(encoding="utf-8", errors="ignore")
    parser = SimpleTableParser()
    parser.feed(raw)

    if not parser.rows:
        raise ValueError(
            "Could not read the source .xls. It is neither a supported .xlsx "
            "file nor an HTML-table .xls file."
        )

    return parser.rows


def rows_to_records(rows):
    if not rows:
        raise ValueError("The source file is empty.")

    headers = [clean_text(x) for x in rows[0]]
    required = [
        "PID", "JID", "Prisoner Name", "Father Name", "Gender",
        "Age", "FIR/Case No", "PS Name", "Ward", "Barrack", "Court Name"
    ]

    header_map = {norm(h): i for i, h in enumerate(headers)}
    missing = [h for h in required if norm(h) not in header_map]
    if missing:
        raise ValueError("Missing source columns: " + ", ".join(missing))

    records = []
    for raw in rows[1:]:
        if not raw:
            continue

        rec = {}
        for h in required:
            idx = header_map[norm(h)]
            rec[h] = raw[idx] if idx < len(raw) else ""

        if not clean_text(rec["JID"]) and not clean_text(rec["Prisoner Name"]):
            continue

        # Preserve source values while normalizing display-only fields.
        rec["PID"] = clean_text(rec["PID"])
        rec["JID"] = clean_text(rec["JID"])
        rec["Prisoner Name"] = clean_text(rec["Prisoner Name"])
        rec["Father Name"] = clean_text(rec["Father Name"])
        rec["Gender"] = clean_text(rec["Gender"])
        rec["FIR/Case No"] = clean_text(rec["FIR/Case No"])
        rec["Ward"] = transformed_location(rec["Ward"])
        rec["Barrack"] = transformed_location(rec["Barrack"])
        rec["Court Name"] = clean_text(rec["Court Name"])
        rec["_display_name"] = transformed_name(rec)

        records.append(rec)

    return records


# ============================================================
# SERIAL NUMBER ASSIGNMENT
# ============================================================

def assign_serial_numbers(records):
    """
    Default behavior follows the practical meaning of the existing
    workbook: serial numbers restart at 1 for each court.
    Repeated JID inside that court receives the same SL.
    """

    if SAME_JID_SAME_SL_GLOBALLY:
        jid_to_sl = {}
        next_sl = 1
        for rec in records:
            jid = rec["JID"]
            if jid not in jid_to_sl:
                jid_to_sl[jid] = next_sl
                next_sl += 1
            rec["_SL"] = jid_to_sl[jid]
        return

    by_court = {}
    for rec in records:
        by_court.setdefault(rec["Court Name"], []).append(rec)

    for court_records in by_court.values():
        # IMPORTANT: sort first, then assign SL. This makes SL ascending in
        # the displayed A-Z name order, rather than preserving source order.
        court_records.sort(key=lambda r: (natural_key(r.get("_display_name", "")), natural_key(r.get("JID", ""))))
        jid_to_sl = {}
        next_sl = 1
        for rec in court_records:
            jid = clean_text(rec["JID"])
            # Empty JIDs are not treated as duplicate identifiers.
            if SAME_JID_SAME_SL_WITHIN_COURT and jid and jid in jid_to_sl:
                rec["_SL"] = jid_to_sl[jid]
            else:
                rec["_SL"] = next_sl
                if SAME_JID_SAME_SL_WITHIN_COURT and jid:
                    jid_to_sl[jid] = next_sl
                next_sl += 1


# ============================================================
# TEMPLATE / FORMATTING
# ============================================================

def copy_cell_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell._style = copy.copy(src_cell._style)
    if src_cell.number_format:
        dst_cell.number_format = src_cell.number_format
    if src_cell.alignment:
        dst_cell.alignment = copy.copy(src_cell.alignment)
    if src_cell.protection:
        dst_cell.protection = copy.copy(src_cell.protection)


def copy_row_style(ws, source_row, target_row, max_col=8):
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    ws.row_dimensions[target_row].hidden = ws.row_dimensions[source_row].hidden
    for col in range(1, max_col + 1):
        copy_cell_style(ws.cell(source_row, col), ws.cell(target_row, col))


def copy_column_dimensions(ws):
    # Already preserved by loading the template.
    pass


def unmerge_all(ws):
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))


def clear_sheet(ws):
    unmerge_all(ws)
    # Delete all rows after keeping the sheet object/template metadata.
    if ws.max_row > 1:
        ws.delete_rows(1, ws.max_row)
    else:
        ws["A1"] = None


def add_duplicate_cf(ws, ranges):
    if not ADD_DUPLICATE_CONDITIONAL_FORMATTING or not ranges:
        return

    # Compare duplicates down each entire logical column, including across
    # different court sections. Ignore blank cells. Use a formula rule because
    # Excel's built-in duplicateValues rule can behave inconsistently when the
    # report is split into many separate section ranges.
    fill = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
    font = Font(color="FF9C0006")
    for cell_range in ranges:
        first_cell = cell_range.split(":", 1)[0]
        col = re.match(r"[A-Z]+", first_cell).group(0)
        formula = f'AND({first_cell}<>"",COUNTIF(${col}:${col},{first_cell})>1)'
        rule = FormulaRule(formula=[formula], fill=fill, font=font, stopIfTrue=False)
        ws.conditional_formatting.add(cell_range, rule)


def set_page_setup(ws):
    ws.sheet_view.showGridLines = False

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    # Force landscape/fit-to-width settings into the workbook print metadata.
    ws.sheet_properties.pageSetUpPr.autoPageBreaks = False

    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.50
    ws.page_margins.bottom = 0.50
    ws.page_margins.header = 0.20
    ws.page_margins.footer = 0.20

    ws.print_options.horizontalCentered = False
    ws.print_options.verticalCentered = False


def rebuild_workbook(template_path: Path, records, output_path: Path, display_date: str):
    if template_path.exists():
        wb = load_workbook(template_path)
        ws = wb.active

        # Capture template styles BEFORE deleting rows.
        style_header = {}
        style_colheader = {}
        style_data = {}
        style_blank = {}

        # These row numbers are taken from the supplied reference workbook.
        for col in range(1, 9):
            style_header[col] = copy.copy(ws.cell(3, col)._style)
            style_colheader[col] = copy.copy(ws.cell(5, col)._style)
            style_data[col] = copy.copy(ws.cell(6, col)._style)
            style_blank[col] = copy.copy(ws.cell(7, col)._style)

        # Preserve special title styles.
        title_styles = {col: copy.copy(ws.cell(1, col)._style) for col in range(1, 9)}

        clear_sheet(ws)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "PrisonersList"

        # Fallback styles if the template is missing.
        thin = Side(style="thin", color="000000")
        normal_border = Border(left=thin, right=thin, top=thin, bottom=thin)
        title_font = Font(name="Algerian", size=16, bold=True)
        section_font = Font(name="Calibri", size=14, bold=True)
        header_font = Font(name="Calibri", size=11, bold=True)
        normal_font = Font(name="Calibri", size=11)

        style_header = {}
        style_colheader = {}
        style_data = {}
        style_blank = {}
        title_styles = {}

        for col in range(1, 9):
            style_header[col] = copy.copy(ws.cell(1, col)._style)
            style_colheader[col] = copy.copy(ws.cell(1, col)._style)
            style_data[col] = copy.copy(ws.cell(1, col)._style)
            style_blank[col] = copy.copy(ws.cell(1, col)._style)
            title_styles[col] = copy.copy(ws.cell(1, col)._style)

    # Sort courts by requested area sequence, then within-area sequence.
    courts = sorted(
        {r["Court Name"] for r in records},
        key=court_sort_key
    )

    # Sort records A-Z within each court.
    grouped = {}
    for court in courts:
        grouped[court] = sorted(
            [r for r in records if r["Court Name"] == court],
            key=lambda x: (
                natural_key(x.get("_display_name", "")),
                natural_key(x.get("JID", "")),
            )
        )

    # First title.
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=7)
    title = f"COURT LIST ON {display_date} ({len(records)})"
    ws.cell(1, 1).value = title

    for col in range(1, 9):
        ws.cell(1, col)._style = copy.copy(title_styles[col])

    # Ensure merged title style remains on A1.
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    current_row = 3
    duplicate_ranges = []

    column_headers = [
        "SL", "JID", "Prisoner Name", "Father Name",
        "FIR/Case No", "W", "B", "Court Name"
    ]

    for court in courts:
        rows = grouped[court]

        # Court heading: two merged rows A:G.
        header_row = current_row
        ws.merge_cells(
            start_row=header_row,
            start_column=1,
            end_row=header_row + 1,
            end_column=7,
        )
        ws.cell(header_row, 1).value = court
        ws.cell(header_row, 1).alignment = Alignment(
            horizontal="center", vertical="center"
        )

        for col in range(1, 9):
            ws.cell(header_row, col)._style = copy.copy(style_header[col])
            ws.cell(header_row + 1, col)._style = copy.copy(style_header[col])

        # Explicitly restore the court-subheading borders AFTER merging.
        # The reference layout has a rectangular outer border around A:G
        # across the two heading rows, plus a fully bordered H cell on each row.
        # Applying styles only before/after merge can cause openpyxl to drop
        # some merged-cell edge borders, so set the perimeter deliberately.
        thin_side = Side(style="thin", color="000000")
        for col in range(1, 8):
            left = thin_side if col == 1 else Side(style=None)
            right = thin_side if col == 7 else Side(style=None)
            ws.cell(header_row, col).border = Border(
                left=left, right=right, top=thin_side, bottom=Side(style=None)
            )
            ws.cell(header_row + 1, col).border = Border(
                left=left, right=right, top=Side(style=None), bottom=thin_side
            )
        for rr in (header_row, header_row + 1):
            ws.cell(rr, 8).border = Border(
                left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
            )

        # Column heading.
        hdr = header_row + 2
        for col, value in enumerate(column_headers, start=1):
            cell = ws.cell(hdr, col)
            cell.value = value
            cell._style = copy.copy(style_colheader[col])
            cell.alignment = Alignment(horizontal="center", vertical="center")
            # Every column heading cell must have all four visible borders.
            thin_side = Side(style="thin", color="000000")
            old = cell.border
            cell.border = Border(
                left=old.left if old.left and old.left.style else thin_side,
                right=old.right if old.right and old.right.style else thin_side,
                top=old.top if old.top and old.top.style else thin_side,
                bottom=old.bottom if old.bottom and old.bottom.style else thin_side,
            )

        # Data.
        data_start = hdr + 1
        for idx, rec in enumerate(rows):
            rr = data_start + idx
            values = [
                rec["_SL"],
                rec["JID"],
                rec["_display_name"],
                rec["Father Name"],
                rec["FIR/Case No"],
                rec["Ward"],
                rec["Barrack"],
                rec["Court Name"],
            ]

            for col, value in enumerate(values, start=1):
                cell = ws.cell(rr, col)
                cell.value = value
                cell._style = copy.copy(style_data[col])
                # Guarantee visible cell borders in every generated data row.
                # Keep any existing template side style; fill missing sides thin black.
                old_border = cell.border
                thin_side = Side(style="thin", color="000000")
                cell.border = Border(
                    left=old_border.left if old_border.left and old_border.left.style else thin_side,
                    right=old_border.right if old_border.right and old_border.right.style else thin_side,
                    top=old_border.top if old_border.top and old_border.top.style else thin_side,
                    bottom=old_border.bottom if old_border.bottom and old_border.bottom.style else thin_side,
                    diagonal=old_border.diagonal,
                    diagonal_direction=old_border.diagonal_direction,
                    diagonalUp=old_border.diagonalUp,
                    diagonalDown=old_border.diagonalDown,
                    outline=old_border.outline,
                    vertical=old_border.vertical,
                    horizontal=old_border.horizontal,
                )

            # W and B in the supplied workbook are centered.
            ws.cell(rr, 6).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(rr, 7).alignment = Alignment(horizontal="center", vertical="center")

        # Two blank rows after every court.
        after_data = data_start + len(rows)
        for rr in (after_data, after_data + 1):
            for col in range(1, 9):
                ws.cell(rr, col)._style = copy.copy(style_blank[col])

        # Duplicate-value CF on the five requested logical columns:
        # B (JID), C (Prisoner Name), D (Father Name), E (FIR/Case No), H (Court Name).
        if len(rows) > 0:
            last_data_row = data_start + len(rows) - 1
            # Separate rules per column are essential: duplicate detection
            # must compare JID with JID, names with names, etc., not across B:E.
            for col in (2, 3, 4, 5, 8):
                letter = get_column_letter(col)
                duplicate_ranges.append(f"{letter}{data_start}:{letter}{last_data_row}")

        current_row = after_data + 2

    # Fix title after all merges/styles.
    ws["A1"].value = title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # Reapply the merged-title border/style to merged placeholders if possible.
    for col in range(1, 8):
        ws.cell(1, col)._style = copy.copy(title_styles[col])

    # The template already contains the exact column widths, borders,
    # fonts, and alignment used by the reference workbook. Do not overwrite
    # them here; this is what keeps the generated printout visually stable.

    # Print setup.
    set_page_setup(ws)

    # Set print area to actual content.
    ws.print_area = f"A1:H{ws.max_row}"

    # Remove any inherited conditional formatting.
    ws.conditional_formatting._cf_rules.clear()

    # Add duplicate conditional formatting.
    add_duplicate_cf(ws, duplicate_ranges)

    # Active cell.
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"

    wb.save(output_path)


# ============================================================
# ONE-CLICK UI
# ============================================================

def choose_source(folder: Path) -> Path | None:
    candidates = sorted(
        [
            p for p in folder.iterdir()
            if p.is_file()
            and p.name.lower().startswith("prisonerslist")
            and p.suffix.lower() in (".xls", ".xlsx", ".xlsm")
        ],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    if candidates:
        # Most recently modified PrisonersList is the normal one-click input.
        return candidates[0]

    return Path(
        filedialog.askopenfilename(
            title="Select PrisonersList file",
            filetypes=[
                ("PrisonersList", "PrisonersList*.xls;PrisonersList*.xlsx"),
                ("Excel files", "*.xls;*.xlsx;*.xlsm"),
                ("All files", "*.*"),
            ],
        )
    ) if filedialog.askopenfilename else None


def get_output_date():
    if DATE_MODE == "tomorrow":
        return date.today() + timedelta(days=1)

    # Simple text dialog using tkinter.
    import tkinter.simpledialog as simpledialog
    root = Tk()
    root.withdraw()
    s = simpledialog.askstring(
        "Court List Date",
        "Enter output date as DD-MM-YYYY:",
    )
    root.destroy()

    if not s:
        raise ValueError("No output date entered.")

    return datetime.strptime(s.strip(), "%d-%m-%Y").date()


def main():
    base = Path(__file__).resolve().parent
    template = base / TEMPLATE_NAME

    root = Tk()
    root.withdraw()

    try:
        # Prefer newest PrisonersList in the same folder.
        candidates = sorted(
            [
                p for p in base.iterdir()
                if p.is_file()
                and p.name.lower().startswith("prisonerslist")
                and p.suffix.lower() in (".xls", ".xlsx", ".xlsm")
            ],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )

        if candidates:
            source = candidates[0]
        else:
            source = filedialog.askopenfilename(
                title="Select PrisonersList file",
                filetypes=[
                    ("PrisonersList", "*.xls;*.xlsx;*.xlsm"),
                    ("All Excel files", "*.xls;*.xlsx;*.xlsm"),
                    ("All files", "*.*"),
                ],
            )
            if not source:
                root.destroy()
                return
            source = Path(source)

        out_date = get_output_date()
        display_date = out_date.strftime("%d-%m-%Y")
        output = base / OUTPUT_PATTERN.format(date=display_date)

        rows = read_source(source)
        records = rows_to_records(rows)

        if not records:
            raise ValueError("No prisoner records were found.")

        assign_serial_numbers(records)

        rebuild_workbook(
            template_path=template,
            records=records,
            output_path=output,
            display_date=display_date,
        )

        messagebox.showinfo(
            "Court List Generator",
            f"Completed successfully.\n\n"
            f"Source: {source.name}\n"
            f"Prisoners: {len(records)}\n"
            f"Output date: {display_date}\n\n"
            f"Created:\n{output.name}",
        )

        # Open the output folder in Windows.
        if os.name == "nt":
            os.startfile(base)

    except Exception as exc:
        messagebox.showerror(
            "Court List Generator - Error",
            f"{type(exc).__name__}: {exc}\n\n"
            "The source file was not modified.",
        )
        raise
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
