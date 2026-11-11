from pathlib import Path

from datetime import datetime, date

from decimal import Decimal

from copy import copy

import io
import os
import hashlib
import shutil
import tempfile

import re

import tkinter as tk

import warnings

import zipfile

from tkinter import filedialog, messagebox

from pypdf import PdfReader

from openpyxl import load_workbook

from openpyxl.formula.translate import Translator

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from openpyxl.utils import get_column_letter, range_boundaries

APP_VERSION = "1.1.0"

APP_TITLE = f"ACH Returns Updater v{APP_VERSION}"

CURRENT_SHEET = "Current Returns"

CURRENT_TABLE = "CurrentReturns"

REPROCESSED_SHEET = "Reprocessed Returns"

REPROCESSED_TABLE = "ReprocessedReturns"

# Excel worksheet row limit. Managed rules expand with the tables.

RULES_LAST_ROW = 1048576

# The columns this app writes into (the workbook may also contain extra

# columns of your own; those are left alone and carried along with each row).

APP_COLUMNS = [

    "Bank Info Updated",

    "Date Reprocessed",

    "Employee Name",

    "Employee Email",

    "EmpID",

    "Last Contacted",

    "Receiving Bank ID",

    "Receiving Account",

    "Amount",

    "Reason",

    "Pay Date",

    "Report Date",

    "Notes",

]

# Look of each column in new rows: (number format, horizontal alignment)

DATE_FMT, TEXT_FMT, AMOUNT_FMT = "mm/dd/yyyy", "@", "#,##0.00"

COLUMN_STYLES = {

    "Bank Info Updated": (TEXT_FMT, "center"),

    "Date Reprocessed": (DATE_FMT, "center"),

    "Employee Name": (TEXT_FMT, None),

    "Employee Email": (TEXT_FMT, None),

    "EmpID": (TEXT_FMT, None),

    "Last Contacted": (TEXT_FMT, None),

    "Receiving Bank ID": (TEXT_FMT, None),

    "Receiving Account": (TEXT_FMT, None),

    "Amount": (AMOUNT_FMT, None),

    "Reason": (TEXT_FMT, None),

    "Pay Date": (DATE_FMT, "center"),

    "Report Date": (DATE_FMT, "center"),

    "Notes": (TEXT_FMT, None),

}

_THIN = Side(style="thin", color="D9D9D9")

CELL_FONT = Font(name="Arial", size=10)

CELL_FILL = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")

CELL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

# ---------------------------------------------------------------------------

# PDF patterns. Everything inside a single line uses [ \t] instead of \s so a

# match can never wander across a line break (or a page break).

# ---------------------------------------------------------------------------

# First line of every record:

#   [Recipient ID] time time-zone entry-class amount reason-code

# The recipient ID is optional; the amount is ".00" on Notifications of Change.

RECORD_PATTERN = re.compile(
    r"^[ \t]*(?:(?P<recipient_id>\S.*?)[ \t]+)?"
    r"\d{1,2}:\d{2}:\d{2}[ \t]+[AP]M[ \t]+\S+[ \t]+"
    r"[A-Z]{2,4}[ \t]+(?P<amount>(?:\d[\d,]*)?\.\d{2})[ \t]+"
    r"(?P<reason>[RC]\d{2})[ \t]*$", re.MULTILINE)
NAME_LINE_PATTERN = re.compile(
    r"(?P<name>.+?)[ \t]*(?P<date>\d{2}/\d{2}/\d{4})"
    r"[ \t]+[A-Z]{2,4}\b[ \t]*(?P<first>.*)$")
TRACE_PATTERN = re.compile(r"Trace\s+Number\s*:\s*(\d{15})[ \t]+(\d{15})(?!\d)")
BANK_PATTERN = re.compile(r"Receiving\s+Bank\s+ID\s*:[ \t]*(\d{9})(?!\d)")
ACCOUNT_PATTERN = re.compile(r"Receiving\s+Account\s*:[ \t]*([0-9A-Za-z*-]+)")
TRACE_LABEL_PATTERN = re.compile(r"Trace\s+Number\s*:")
PAGE_FOOTER_PATTERN = re.compile(r"^\s*Page:\s*\d+\s*$")
REPORT_DATE_PATTERN = re.compile(r"As\s+of\s+(\d{2}/\d{2}/\d{4})", re.I)

def clean_page(page_text, page_number):
    lines = page_text.replace("\xa0", " ").splitlines()
    meaningful = [x for x in lines if x.strip() and not PAGE_FOOTER_PATTERN.match(x)]
    if not meaningful:
        raise ValueError(f"Page {page_number} has no readable report text. Import stopped.")
    if not meaningful[0].strip().startswith("ACH Return/NOC Report"):
        raise ValueError(f"Page {page_number}: unsupported or missing report header.")
    end = next((i for i,x in enumerate(lines[:25]) if "Treasury Information Reporting" in x), None)
    if end is None:
        raise ValueError(f"Page {page_number}: could not identify the end of the header.")
    body = "\n".join(x for x in lines[end+1:] if not PAGE_FOOTER_PATTERN.match(x))
    if not body.strip():
        raise ValueError(f"Page {page_number} contains only a header. Import stopped.")
    return body

def read_description(lines, start):

    """Collect a reason-code description that may wrap onto several lines.

    It ends at the 'Credit'/'Debit' line (which also holds the Company Entry

    Description, e.g. 'PAYROLL Credit') or at 'Original'."""

    parts = []

    for line in lines[start:]:

        if re.search(r"(?:^|\s)(?:Credit|Debit)$", line) or line.startswith(

            "Original"

        ):

            break

        if line:

            parts.append(line)

    return " ".join(parts)

def parse_pdf_records(page_texts):

    """Turn the text of every PDF page into (report_date, records, skipped).

    Only records with an amount greater than 0 are returned; Notifications of

    Change ($0.00) are still read so the trace-number checks stay honest."""

    if not any(text.strip() for text in page_texts):

        raise ValueError(

            "No readable text was found. This PDF may need OCR first."

        )

    report_dates = set(REPORT_DATE_PATTERN.findall("\n".join(page_texts)))

    if len(report_dates) != 1:

        raise ValueError(

            "Expected one report date in the PDF; "

            f"found: {sorted(report_dates)}"

        )

    report_date = datetime.strptime(report_dates.pop(), "%m/%d/%Y")

    cleaned_pages = [

        clean_page(text, number)

        for number, text in enumerate(page_texts, start=1)

    ]

    for number, text in enumerate(cleaned_pages, start=1):

        if not text.strip():

            print(

                f"WARNING: page {number} has no readable text (it may be "

                "an image). Records on that page would be missing."

            )

    text = "\n".join(cleaned_pages)

    record_starts = list(RECORD_PATTERN.finditer(text))

    if not record_starts:

        raise ValueError(

            "No return or NOC records matched the expected report layout."

        )

    records = []

    for index, match in enumerate(record_starts):

        end = (

            record_starts[index + 1].start()

            if index + 1 < len(record_starts)

            else len(text)

        )

        block = text[match.end():end]

        lines = [line.strip() for line in block.split("\n")]

        employee = None

        for position, line in enumerate(lines):

            found = NAME_LINE_PATTERN.match(line)

            if found and found["name"]:

                continued = read_description(lines, position + 1)

                description = " ".join(

                    part for part in (found["first"], continued) if part

                )

                employee = {

                    "name": found["name"],

                    "date": found["date"],

                    "description": description,

                }

                break

        traces = TRACE_PATTERN.search(block)

        bank = BANK_PATTERN.search(block)

        account = ACCOUNT_PATTERN.search(block)

        missing = [

            label

            for label, value in [

                ("employee name / effective date / description", employee),

                ("original and returned trace numbers", traces),

                ("receiving bank ID", bank),

                ("receiving account", account),

            ]

            if value is None

        ]

        if missing:

            raise ValueError(

                f"Record {index + 1} could not be fully read. "

                f"Missing: {', '.join(missing)}. "

                "Nothing was changed in the Excel file."

            )

        # Every record must hold exactly one trace-number line.

        if len(TRACE_LABEL_PATTERN.findall(block)) != 1:

            raise ValueError(

                f"Record {index + 1} does not contain exactly one trace "

                "number. Nothing was changed in the Excel file."

            )

        records.append({

            # One space between words, so repeated names always match.

            "Employee Name": " ".join(employee["name"].split()),

            "Receiving Bank ID": bank.group(1),

            "Receiving Account": account.group(1),

            "Amount": float(Decimal(match["amount"].replace(",", "") or "0")),

            "Reason": employee["description"].strip(),

            "Pay Date": datetime.strptime(employee["date"], "%m/%d/%Y"),

            "Report Date": report_date,

        })

    # Check that no trace-number record was silently skipped.

    trace_count = len(TRACE_LABEL_PATTERN.findall(text))

    if trace_count != len(records):

        raise ValueError(

            f"Found {trace_count} trace-number records but parsed "

            f"{len(records)} rows. Nothing was changed in the Excel file."

        )

    payable = [record for record in records if record["Amount"] > 0]

    if not payable:

        raise ValueError(

            "No records with an amount greater than 0 were found."

        )

    return report_date, payable, len(records) - len(payable)

# ---------------------------------------------------------------------------

# Excel workbook helpers

# ---------------------------------------------------------------------------

class TableInfo:

    """One Excel table on one sheet, with its header names mapped to columns."""

    def __init__(self, ws, table):

        self.ws = ws

        self.table = table

        self.c1, self.r1, self.c2, self.r2 = range_boundaries(table.ref)

        self.cols = {}

        for column in range(self.c1, self.c2 + 1):

            value = ws.cell(self.r1, column).value

            if value is not None:

                self.cols[str(value).strip()] = column

def get_table_info(wb, sheet_name, table_name, required_columns):

    if sheet_name not in wb.sheetnames:

        raise ValueError(

            f'This workbook has no sheet named "{sheet_name}". '

            "Please pick the ACH Returns workbook."

        )

    ws = wb[sheet_name]

    if table_name in ws.tables:

        table = ws.tables[table_name]

    elif False:

        pass

    else:

        raise ValueError(

            f'The "{sheet_name}" sheet needs exactly one Excel table '

            f'(named {table_name}). Please pick the ACH Returns workbook.'

        )

    if table.totalsRowCount:

        raise ValueError(

            f'The table on "{sheet_name}" has a Total Row turned on. '

            "Turn it off (Table Design > Total Row) and run the app again."

        )

    info = TableInfo(ws, table)

    missing = [name for name in required_columns if name not in info.cols]

    if missing:

        raise ValueError(

            f'The "{sheet_name}" table is missing these columns: '

            f'{", ".join(missing)}.'

        )

    return info

def row_is_blank(ws, row, first_col, last_col):

    return all(

        ws.cell(row, column).value in (None, "")

        for column in range(first_col, last_col + 1)

    )

def count_filled_rows(info):

    return sum(

        1

        for row in range(info.r1 + 1, info.r2 + 1)

        if not row_is_blank(info.ws, row, info.c1, info.c2)

    )

def to_date(value):

    if isinstance(value, datetime):

        return value.date()

    if isinstance(value, date):

        return value

    if isinstance(value, str):

        for pattern in ("%m/%d/%Y", "%Y-%m-%d"):

            try:

                return datetime.strptime(value.strip(), pattern).date()

            except ValueError:

                pass

    return None

def count_rows_with_report_date(info, wanted):

    column = info.cols.get("Report Date")

    if column is None:

        return 0

    return sum(

        1

        for row in range(info.r1 + 1, info.r2 + 1)

        if to_date(info.ws.cell(row, column).value) == wanted

    )

def filter_is_active(info):

    """True if the table has a filter applied or any hidden rows."""

    auto_filter = info.table.autoFilter

    has_criteria = bool(auto_filter is not None and auto_filter.filterColumn)

    has_hidden = any(

        dim.hidden

        for row, dim in info.ws.row_dimensions.items()

        if info.r1 < row <= info.r2

    )

    return has_criteria or has_hidden

def insert_records(cur, records):

    """Insert the records at the top of the Current Returns table and push

    the existing rows down. Everything typed into the sheet stays with its row."""

    ws, table = cur.ws, cur.table

    r1, r2, c1, c2 = cur.r1, cur.r2, cur.c1, cur.c2

    n = len(records)

    empty_table = all(

        row_is_blank(ws, row, c1, c2) for row in range(r1 + 1, r2 + 1)

    )

    final_last_row = r1 + n if empty_table else r2 + n

    if final_last_row > RULES_LAST_ROW:

        raise ValueError(

            f"The Current Returns table would grow past row {RULES_LAST_ROW:,}, "

            "which exceeds Excel's worksheet capacity. "

            "Start a new year's workbook or ask for the limit to be raised."

        )

    # Refuse occupied cells below even when the table itself is empty.
    for cells in ws.iter_rows(min_row=r2+1, max_row=max(r2+1, final_last_row), min_col=c1, max_col=c2):
        if any(c.value is not None for c in cells):
            raise ValueError("Data below the table would be overwritten. Move it first.")
    source_row = None  # old first data row, used to carry along extra columns

    if not empty_table:

        last_col = max(ws.max_column, c2)

        # Nothing may be sitting directly under the table.

        for row in ws.iter_rows(min_row=r2 + 1, max_row=r2 + n, max_col=last_col):

            if any(cell.value not in (None, "") for cell in row):

                raise ValueError(

                    "There is data directly below the Current Returns table, "

                    "and the app would overwrite it. Move or delete it first."

                )

        # Remember custom row heights so they move with their rows.

        heights = {

            row: dim.height

            for row, dim in list(ws.row_dimensions.items())

            if r1 < row <= r2 and dim.height is not None

        }

        # Push the existing rows down (formulas in them are adjusted).

        ws.move_range(

            f"A{r1 + 1}:{get_column_letter(last_col)}{r2}",

            rows=n,

            translate=True,

        )

        # Email hyperlinks must be re-pointed at their new cells.

        for row in ws.iter_rows(

            min_row=r1 + 1 + n, max_row=r2 + n, max_col=last_col

        ):

            for cell in row:

                if cell.hyperlink is not None:

                    cell.hyperlink.ref = cell.coordinate

        for row in heights:

            ws.row_dimensions[row].height = None

        for row, height in heights.items():

            ws.row_dimensions[row + n].height = height

        source_row = r1 + 1 + n

    # Write the new rows.

    for offset, record in enumerate(records):

        row = r1 + 1 + offset

        for name, column in cur.cols.items():

            cell = ws.cell(row, column)

            if name in COLUMN_STYLES:

                number_format, horizontal = COLUMN_STYLES[name]

                cell.font = CELL_FONT

                cell.fill = CELL_FILL

                cell.border = CELL_BORDER

                cell.number_format = number_format

                cell.alignment = Alignment(

                    horizontal=horizontal, vertical="center"

                )

                cell.value = record.get(name)
                if isinstance(cell.value, str):
                    cell.data_type = "s"

            elif source_row is not None:

                # An extra column you added: copy its look, and carry

                # a formula down so it works on the new rows too.

                source = ws.cell(source_row, column)

                cell._style = copy(source._style)

                if isinstance(source.value, str) and source.value.startswith("="):

                    cell.value = Translator(

                        source.value, origin=source.coordinate

                    ).translate_formula(cell.coordinate)

    # Grow the table (and its filter buttons) to cover the new rows.

    new_ref = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{final_last_row}"

    table.ref = new_ref

    if table.autoFilter is not None:

        table.autoFilter.ref = new_ref

        table.autoFilter.filterColumn = []  # clear any active filter

    # Un-hide every row, so nothing stays hidden in the wrong place.

    for row, dim in ws.row_dimensions.items():

        if r1 < row <= final_last_row and dim.hidden:

            dim.hidden = False

def load_workbook_checked(path):

    """Open the workbook and return (workbook, warnings_about_lost_features)."""

    # Excel's newer checkboxes would be flattened by this app; refuse them.

    with zipfile.ZipFile(path) as archive:

        if "xl/featurePropertyBag/featurePropertyBag.xml" in archive.namelist():

            raise ValueError(

                "This workbook contains Excel checkboxes, which this app "

                "would remove when it saves. Please use the ACH Returns "

                "template that has the ✓ dropdowns."

            )

    with warnings.catch_warnings(record=True) as caught:

        warnings.simplefilter("always")

        wb = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")

    notes = sorted({

        str(item.message)

        for item in caught

        if "not supported" in str(item.message).lower()

        or "extension" in str(item.message).lower()

    })

    return wb, notes

def cell_signature(cell):
    link = cell.hyperlink
    return (cell.value, cell.data_type, cell.number_format,
            (link.target, link.location) if link else None,
            cell.comment.text if cell.comment else None)

def maintain_rules(wb):
    # Rebuild only the tracker's two known rules; retain unrelated rules.
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    for sheet, table in [(CURRENT_SHEET, CURRENT_TABLE), (REPROCESSED_SHEET, REPROCESSED_TABLE)]:
        info = get_table_info(wb, sheet, table, ["Bank Info Updated", "Date Reprocessed", "Employee Name"])
        ws = info.ws
        end = min(1048576, max(10000, info.r2 + 1000))
        start = info.r1 + 1
        a = get_column_letter(info.cols["Bank Info Updated"])
        d = get_column_letter(info.cols["Date Reprocessed"])
        name = get_column_letter(info.cols["Employee Name"])
        # These are the template's managed completion/repeated-name rules.
        # Preserve other rule types and unrelated expressions.
        for cf in list(ws.conditional_formatting):
            retained = [r for r in ws.conditional_formatting[cf]
                        if not (r.type == "expression" and any(
                            "COUNTIF(" in f.upper() or "✓" in f for f in (r.formula or [])))]
            del ws.conditional_formatting[str(cf.sqref)]
            for rule in retained:
                ws.conditional_formatting.add(str(cf.sqref), rule)
        ws.conditional_formatting.add(
            f"{get_column_letter(info.c1)}{start}:{get_column_letter(info.c2)}{end}",
            FormulaRule(formula=[f'AND(${a}{start}="✓",ISNUMBER(${d}{start}),${d}{start}>0)'],
                        fill=PatternFill("solid", fgColor="C6EFCE")))
        ws.conditional_formatting.add(f"{name}{start}:{name}{end}",
            FormulaRule(formula=[f'AND(${name}{start}<>"",COUNTIF(${name}${start}:${name}${end},${name}{start})>1)'],
                        font=Font(color="9C0006")))
        for col, kind in [(a,"list"),(d,"date")]:
            colnum = info.cols["Bank Info Updated" if kind=="list" else "Date Reprocessed"]
            for dv in list(ws.data_validations.dataValidation):
                kept = [str(r) for r in dv.sqref.ranges if not (r.min_col == r.max_col == colnum)]
                if kept: dv.sqref = " ".join(kept)
                else: ws.data_validations.dataValidation.remove(dv)
            dv = DataValidation(type=kind, allow_blank=True)
            if kind=="list": dv.formula1 = '"✓"'
            else:
                dv.operator="between"; dv.formula1="1"; dv.formula2="2958465"
            dv.showErrorMessage=True
            dv.error="Choose the checkmark or enter a valid Excel date."
            ws.add_data_validation(dv); dv.add(f"{col}{start}:{col}{end}")

def save_in_place(wb, path, stat_before, expected_filled_rows):
    buffer = io.BytesIO()
    wb.save(buffer)
    data = buffer.getvalue()
    check = load_workbook(io.BytesIO(data), keep_vba=path.suffix.lower()==".xlsm")
    try:
        check_cur = get_table_info(check, CURRENT_SHEET, CURRENT_TABLE, APP_COLUMNS)
        if count_filled_rows(check_cur) != expected_filled_rows:
            raise RuntimeError("Row-count verification failed. Original file was not changed.")
        for sheet in (CURRENT_SHEET, REPROCESSED_SHEET):
            before, after = wb[sheet], check[sheet]
            for row in before.iter_rows():
                for cell in row:
                    if cell.value is not None and cell_signature(cell) != cell_signature(after[cell.coordinate]):
                        raise RuntimeError(f"Saved-content verification failed at {sheet}!{cell.coordinate}. Original unchanged.")
    finally:
        check.close()
    original = path.read_bytes()
    current = path.stat()
    if (current.st_mtime_ns,current.st_size)!=(stat_before.st_mtime_ns,stat_before.st_size):
        raise RuntimeError("Workbook changed during import. Original was not changed by this app.")
    # Verify the macro payload survives before committing.
    with zipfile.ZipFile(io.BytesIO(original)) as old, zipfile.ZipFile(io.BytesIO(data)) as new:
        if "xl/vbaProject.bin" in old.namelist():
            if "xl/vbaProject.bin" not in new.namelist() or old.read("xl/vbaProject.bin")!=new.read("xl/vbaProject.bin"):
                raise RuntimeError("VBA preservation check failed. Original unchanged.")
    backup = path.with_name(path.stem + ".backup-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f") + path.suffix)
    with backup.open("xb") as stream:
        stream.write(original)
        stream.flush()
        os.fsync(stream.fileno())
    if path.read_bytes() != original:
        raise RuntimeError("Workbook changed before saving. Import canceled.")
    try:
        with path.open("r+b") as stream:
            stream.write(data)
            stream.truncate()
            stream.flush()
            os.fsync(stream.fileno())
        if path.read_bytes() != data:
            raise OSError("Saved bytes do not match generated workbook")
    except Exception as exc:
        raise RuntimeError(f"Saving failed; the workbook may be partially changed. Restore this backup: {backup}") from exc
    print(f"Recovery backup: {backup}")

# ---------------------------------------------------------------------------
# The app

# ---------------------------------------------------------------------------

def run(root):

    # 1) Pick the PDF.

    downloads = Path.home() / "Downloads"

    selected_pdf = filedialog.askopenfilename(

        parent=root,

        title="Select the ACH Return PDF",

        initialdir=str(downloads if downloads.exists() else Path.cwd()),

        filetypes=[("PDF files", "*.pdf")],

    )

    if not selected_pdf:

        print("Canceled. Nothing was changed.")

        return

    pdf_path = Path(selected_pdf)

    if not pdf_path.is_file():

        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    reader = PdfReader(str(pdf_path))

    page_texts = [page.extract_text() or "" for page in reader.pages]

    report_date, records, skipped = parse_pdf_records(page_texts)

    # 2) Pick the Excel file.

    box_folder = Path.home() / "Box"

    selected_xlsx = filedialog.askopenfilename(

        parent=root,

        title="Select the ACH Returns Excel file to update",

        initialdir=str(box_folder if box_folder.exists() else Path.home()),

        filetypes=[("Excel files", "*.xlsx *.xlsm")],

    )

    if not selected_xlsx:

        print("Canceled. Nothing was changed.")

        return

    excel_path = Path(selected_xlsx)

    stat_before = excel_path.stat()

    wb, load_notes = load_workbook_checked(excel_path)

    try:

        cur = get_table_info(wb, CURRENT_SHEET, CURRENT_TABLE, APP_COLUMNS)

        rep = get_table_info(

            wb, REPROCESSED_SHEET, REPROCESSED_TABLE, ["Report Date"]

        )

        # 3) Has this report already been imported?

        already = count_rows_with_report_date(

            cur, report_date.date()

        ) + count_rows_with_report_date(rep, report_date.date())

        if already:

            proceed = messagebox.askyesno(

                parent=root,

                title=APP_TITLE,

                icon="warning",

                default="no",

                message=(

                    f"This workbook already has {already} row(s) from a report "

                    f"dated {report_date:%m/%d/%Y}.\n\n"

                    "Importing it again will likely add duplicates. Add it anyway?"

                ),

            )

            if not proceed:

                print("Canceled. Nothing was changed.")

                return

        # 4) Confirm.

        lines = []

        if skipped:

            lines.append(f"Skipped {skipped} record(s) with a $0.00 amount.")

        if filter_is_active(cur):

            lines.append(

                "A filter (or hidden rows) was found on Current Returns. "

                "It will be cleared. Re-apply your filter after the import."

            )

        years = set(re.findall(r"(?<!\d)(20\d{2})(?!\d)", excel_path.name))

        if years and str(report_date.year) not in years:

            lines.append(

                f"Heads up: the file name mentions {', '.join(sorted(years))} "

                f"but this report is dated {report_date:%m/%d/%Y}."

            )

        for note in load_notes:

            lines.append(f"Excel feature warning: {note}")

        confirmed = messagebox.askyesno(

            parent=root,

            title=APP_TITLE,

            message=(

                f"Add {len(records)} return(s) dated {report_date:%m/%d/%Y} "

                f'to "{excel_path.name}"?'

            ),

            detail="\n\n".join(lines) if lines else None,

        )

        if not confirmed:

            print("Canceled. Nothing was changed.")

            return

        # 5) Update and save.

        filled_before = count_filled_rows(cur)
        old_rows = [[cell_signature(cur.ws.cell(r,c)) for c in range(cur.c1,cur.c2+1)]
                    for r in range(cur.r1+1,cur.r2+1)]

        insert_records(cur, records)
        if filled_before:
            for offset, old_row in enumerate(old_rows):
                for coloffset, signature in enumerate(old_row):
                    if signature[1] == "f": continue  # formulas are intentionally translated
                    cell = cur.ws.cell(cur.r1+1+len(records)+offset,cur.c1+coloffset)
                    if cell_signature(cell) != signature:
                        raise RuntimeError("Existing-row preservation check failed. Original unchanged.")
        maintain_rules(wb)

        save_in_place(wb, excel_path, stat_before, filled_before + len(records))

    finally:

        wb.close()

    summary = (

        f"Added {len(records)} return(s) to the top of Current Returns."

        + (f"\nSkipped {skipped} record(s) with a $0.00 amount." if skipped else "")

        + f"\n\nSaved: {excel_path}"

    )

    print(summary)

    messagebox.showinfo(parent=root, title=APP_TITLE, message=summary)

def main():

    root = tk.Tk()

    root.withdraw()

    root.attributes("-topmost", True)

    try:

        run(root)

    except PermissionError:

        message = (

            "The file could not be accessed or saved. Close the Excel file "

            "if it is open (here or on another computer), and make sure you "

            "have permission to the PDF and the Excel file."

        )

        print(message)

        messagebox.showerror(parent=root, title=APP_TITLE, message=message)

    except Exception as error:

        print(f"Update failed: {error}")

        messagebox.showerror(

            parent=root,

            title=APP_TITLE,

            message=str(error),

        )

    finally:

        root.destroy()

if __name__ == "__main__":

    main()