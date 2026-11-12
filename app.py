from pathlib import Path
from datetime import datetime, date
from decimal import Decimal
from copy import copy
import io
import os
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
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation


APP_VERSION = "1.1.1"
APP_TITLE = f"ACH Returns Updater v{APP_VERSION}"

CURRENT_SHEET = "Current Returns"
CURRENT_TABLE = "CurrentReturns"
REPROCESSED_SHEET = "Reprocessed Returns"
REPROCESSED_TABLE = "ReprocessedReturns"

# Excel's maximum worksheet row.
RULES_LAST_ROW = 1048576

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

DATE_FMT = "mm/dd/yyyy"
TEXT_FMT = "@"
AMOUNT_FMT = "#,##0.00"

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
CELL_FILL = PatternFill(
    "solid",
    start_color="FFFFFF",
    end_color="FFFFFF",
)
CELL_BORDER = Border(
    left=_THIN,
    right=_THIN,
    top=_THIN,
    bottom=_THIN,
)


# ---------------------------------------------------------------------------
# PDF parsing
# ---------------------------------------------------------------------------

RECORD_PATTERN = re.compile(
    r"^[ \t]*(?:(?P<recipient_id>\S.*?)[ \t]+)?"
    r"\d{1,2}:\d{2}:\d{2}[ \t]+[AP]M[ \t]+\S+[ \t]+"
    r"[A-Z]{2,4}[ \t]+"
    r"(?P<amount>(?:\d[\d,]*)?\.\d{2})[ \t]+"
    r"(?P<reason>[RC]\d{2})[ \t]*$",
    re.MULTILINE,
)

NAME_LINE_PATTERN = re.compile(
    r"(?P<name>.+?)[ \t]*"
    r"(?P<date>\d{2}/\d{2}/\d{4})"
    r"[ \t]+[A-Z]{2,4}\b[ \t]*(?P<first>.*)$"
)

TRACE_PATTERN = re.compile(
    r"Trace\s+Number\s*:\s*(\d{15})[ \t]+(\d{15})(?!\d)"
)

BANK_PATTERN = re.compile(
    r"Receiving\s+Bank\s+ID\s*:[ \t]*(\d{9})(?!\d)"
)

ACCOUNT_PATTERN = re.compile(
    r"Receiving\s+Account\s*:[ \t]*([0-9A-Za-z*-]+)"
)

TRACE_LABEL_PATTERN = re.compile(r"Trace\s+Number\s*:")
PAGE_FOOTER_PATTERN = re.compile(r"^\s*Page:\s*\d+\s*$")

REPORT_DATE_PATTERN = re.compile(
    r"As\s+of\s+(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE,
)


def clean_page(page_text, page_number):
    """Remove repeated report headers and page footers."""
    lines = page_text.replace("\xa0", " ").splitlines()

    meaningful = [
        line
        for line in lines
        if line.strip() and not PAGE_FOOTER_PATTERN.match(line)
    ]

    if not meaningful:
        raise ValueError(
            f"Page {page_number} has no readable report text. "
            "Import stopped."
        )

    if not meaningful[0].strip().startswith("ACH Return/NOC Report"):
        raise ValueError(
            f"Page {page_number}: unsupported or missing report header."
        )

    header_end = next(
        (
            index
            for index, line in enumerate(lines[:25])
            if "Treasury Information Reporting" in line
        ),
        None,
    )

    if header_end is None:
        raise ValueError(
            f"Page {page_number}: could not identify the end "
            "of the header."
        )

    body = "\n".join(
        line
        for line in lines[header_end + 1:]
        if not PAGE_FOOTER_PATTERN.match(line)
    )

    if not body.strip():
        raise ValueError(
            f"Page {page_number} contains only a header. Import stopped."
        )

    return body


def read_description(lines, start):
    """Join a description that wraps onto multiple lines."""
    parts = []

    for line in lines[start:]:
        if (
            re.search(r"(?:^|\s)(?:Credit|Debit)$", line)
            or line.startswith("Original")
        ):
            break

        if line:
            parts.append(line)

    return " ".join(parts)


def parse_pdf_records(page_texts):
    """
    Return (report_date, positive_amount_records, skipped_count).

    Zero-dollar records are parsed and validated, then excluded.
    Pages are joined after removing headers so records can span pages.
    """
    if not any(text.strip() for text in page_texts):
        raise ValueError(
            "No readable text was found. This PDF may need OCR first."
        )

    report_dates = set(
        REPORT_DATE_PATTERN.findall("\n".join(page_texts))
    )

    if len(report_dates) != 1:
        raise ValueError(
            "Expected one report date in the PDF; "
            f"found: {sorted(report_dates)}"
        )

    report_date = datetime.strptime(
        report_dates.pop(),
        "%m/%d/%Y",
    )

    cleaned_pages = [
        clean_page(text, number)
        for number, text in enumerate(page_texts, start=1)
    ]

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
                    part
                    for part in (found["first"], continued)
                    if part
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

        if len(TRACE_LABEL_PATTERN.findall(block)) != 1:
            raise ValueError(
                f"Record {index + 1} does not contain exactly one "
                "trace number. Nothing was changed in the Excel file."
            )

        records.append({
            "Employee Name": " ".join(employee["name"].split()),
            "Receiving Bank ID": bank.group(1),
            "Receiving Account": account.group(1),
            "Amount": float(
                Decimal(match["amount"].replace(",", "") or "0")
            ),
            "Reason": employee["description"].strip(),
            "Pay Date": datetime.strptime(
                employee["date"],
                "%m/%d/%Y",
            ),
            "Report Date": report_date,
        })

    trace_count = len(TRACE_LABEL_PATTERN.findall(text))

    if trace_count != len(records):
        raise ValueError(
            f"Found {trace_count} trace-number records but parsed "
            f"{len(records)} rows. Nothing was changed in the Excel file."
        )

    payable = [
        record
        for record in records
        if record["Amount"] > 0
    ]

    if not payable:
        raise ValueError(
            "No records with an amount greater than 0 were found."
        )

    return report_date, payable, len(records) - len(payable)


# ---------------------------------------------------------------------------
# Workbook helpers
# ---------------------------------------------------------------------------

class TableInfo:
    """One table, its boundaries, and its header-to-column mapping."""

    def __init__(self, ws, table):
        self.ws = ws
        self.table = table

        self.c1, self.r1, self.c2, self.r2 = range_boundaries(
            table.ref
        )

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

    if table_name not in ws.tables:
        raise ValueError(
            f'The "{sheet_name}" sheet needs an Excel table '
            f'named "{table_name}". '
            "Please pick the ACH Returns workbook."
        )

    table = ws.tables[table_name]

    if table.totalsRowCount:
        raise ValueError(
            f'The table on "{sheet_name}" has a Total Row turned on. '
            "Turn it off (Table Design > Total Row) and run the app again."
        )

    info = TableInfo(ws, table)

    missing = [
        name
        for name in required_columns
        if name not in info.cols
    ]

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
        if not row_is_blank(
            info.ws,
            row,
            info.c1,
            info.c2,
        )
    )


def to_date(value):
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):
        for pattern in ("%m/%d/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(
                    value.strip(),
                    pattern,
                ).date()
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
    """Check for active table filters or hidden table rows."""
    auto_filter = info.table.autoFilter

    has_criteria = bool(
        auto_filter is not None and auto_filter.filterColumn
    )

    has_hidden = any(
        dimension.hidden
        for row, dimension in info.ws.row_dimensions.items()
        if info.r1 < row <= info.r2
    )

    return has_criteria or has_hidden


def insert_records(cur, records):
    """Insert new records at the top and move existing rows down."""
    ws = cur.ws
    table = cur.table

    r1, r2, c1, c2 = cur.r1, cur.r2, cur.c1, cur.c2
    n = len(records)

    empty_table = all(
        row_is_blank(ws, row, c1, c2)
        for row in range(r1 + 1, r2 + 1)
    )

    final_last_row = r1 + n if empty_table else r2 + n

    if final_last_row > RULES_LAST_ROW:
        raise ValueError(
            f"The Current Returns table would grow past row "
            f"{RULES_LAST_ROW:,}, which exceeds Excel's worksheet "
            "capacity. Start another workbook."
        )

    # Check below the table, including when the table is empty.
    for cells in ws.iter_rows(
        min_row=r2 + 1,
        max_row=max(r2 + 1, final_last_row),
        min_col=c1,
        max_col=c2,
    ):
        if any(cell.value is not None for cell in cells):
            raise ValueError(
                "Data below the table would be overwritten. Move it first."
            )

    source_row = None

    if not empty_table:
        last_col = max(ws.max_column, c2)

        for row in ws.iter_rows(
            min_row=r2 + 1,
            max_row=r2 + n,
            max_col=last_col,
        ):
            if any(cell.value not in (None, "") for cell in row):
                raise ValueError(
                    "There is data directly below the Current Returns "
                    "table, and the app would overwrite it. "
                    "Move or delete it first."
                )

        heights = {
            row: dimension.height
            for row, dimension in list(ws.row_dimensions.items())
            if r1 < row <= r2 and dimension.height is not None
        }

        ws.move_range(
            f"A{r1 + 1}:{get_column_letter(last_col)}{r2}",
            rows=n,
            translate=True,
        )

        # Update hyperlink anchors after moving the cells.
        for row in ws.iter_rows(
            min_row=r1 + 1 + n,
            max_row=r2 + n,
            max_col=last_col,
        ):
            for cell in row:
                if cell.hyperlink is not None:
                    cell.hyperlink.ref = cell.coordinate

        for row in heights:
            ws.row_dimensions[row].height = None

        for row, height in heights.items():
            ws.row_dimensions[row + n].height = height

        source_row = r1 + 1 + n

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
                    horizontal=horizontal,
                    vertical="center",
                )

                cell.value = record.get(name)

                if isinstance(cell.value, str):
                    cell.data_type = "s"

            elif source_row is not None:
                source = ws.cell(source_row, column)
                cell._style = copy(source._style)

                if (
                    isinstance(source.value, str)
                    and source.value.startswith("=")
                ):
                    cell.value = Translator(
                        source.value,
                        origin=source.coordinate,
                    ).translate_formula(cell.coordinate)

    new_ref = (
        f"{get_column_letter(c1)}{r1}:"
        f"{get_column_letter(c2)}{final_last_row}"
    )

    table.ref = new_ref

    if table.autoFilter is not None:
        table.autoFilter.ref = new_ref
        table.autoFilter.filterColumn = []

    for row, dimension in ws.row_dimensions.items():
        if r1 < row <= final_last_row and dimension.hidden:
            dimension.hidden = False


def load_workbook_checked(path):
    """Load the workbook and collect unsupported-feature warnings."""
    with zipfile.ZipFile(path) as archive:
        if (
            "xl/featurePropertyBag/featurePropertyBag.xml"
            in archive.namelist()
        ):
            raise ValueError(
                "This workbook contains Excel checkboxes, which this "
                "app would remove when it saves. Please use the "
                "ACH Returns template that has the ✓ dropdowns."
            )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")

        wb = load_workbook(
            path,
            keep_vba=path.suffix.lower() == ".xlsm",
        )

    notes = sorted({
        str(item.message)
        for item in caught
        if (
            "not supported" in str(item.message).lower()
            or "extension" in str(item.message).lower()
        )
    })

    return wb, notes


def cell_signature(cell):
    link = cell.hyperlink

    return (
        cell.value,
        cell.data_type,
        cell.number_format,
        (link.target, link.location) if link else None,
        cell.comment.text if cell.comment else None,
    )


def maintain_rules(wb):
    """Restore the tracker's formatting and validation ranges."""
    table_pairs = [
        (CURRENT_SHEET, CURRENT_TABLE),
        (REPROCESSED_SHEET, REPROCESSED_TABLE),
    ]

    for sheet_name, table_name in table_pairs:
        info = get_table_info(
            wb,
            sheet_name,
            table_name,
            [
                "Bank Info Updated",
                "Date Reprocessed",
                "Employee Name",
            ],
        )

        ws = info.ws
        end = min(1048576, max(10000, info.r2 + 1000))
        start = info.r1 + 1

        updated_col = get_column_letter(
            info.cols["Bank Info Updated"]
        )
        date_col = get_column_letter(
            info.cols["Date Reprocessed"]
        )
        name_col = get_column_letter(
            info.cols["Employee Name"]
        )

        # Replace the template's checkmark and COUNTIF rules.
        # Other rule types and unrelated expressions are retained.
        for cf in list(ws.conditional_formatting):
            retained = [
                rule
                for rule in ws.conditional_formatting[cf]
                if not (
                    rule.type == "expression"
                    and any(
                        "COUNTIF(" in formula.upper()
                        or "✓" in formula
                        for formula in (rule.formula or [])
                    )
                )
            ]

            del ws.conditional_formatting[str(cf.sqref)]

            for rule in retained:
                ws.conditional_formatting.add(
                    str(cf.sqref),
                    rule,
                )

        full_range = (
            f"{get_column_letter(info.c1)}{start}:"
            f"{get_column_letter(info.c2)}{end}"
        )

        completed_formula = (
            f'AND(${updated_col}{start}="✓",'
            f'ISNUMBER(${date_col}{start}),'
            f'${date_col}{start}>0)'
        )

        ws.conditional_formatting.add(
            full_range,
            FormulaRule(
                formula=[completed_formula],
                fill=PatternFill(
                    "solid",
                    fgColor="C6EFCE",
                ),
            ),
        )

        repeated_name_formula = (
            f'AND(${name_col}{start}<>"",'
            f'COUNTIF(${name_col}${start}:${name_col}${end},'
            f'${name_col}{start})>1)'
        )

        ws.conditional_formatting.add(
            f"{name_col}{start}:{name_col}{end}",
            FormulaRule(
                formula=[repeated_name_formula],
                font=Font(color="9C0006"),
            ),
        )

        for column_letter, kind in [
            (updated_col, "list"),
            (date_col, "date"),
        ]:
            header = (
                "Bank Info Updated"
                if kind == "list"
                else "Date Reprocessed"
            )
            column_number = info.cols[header]

            for validation in list(
                ws.data_validations.dataValidation
            ):
                kept = [
                    str(cell_range)
                    for cell_range in validation.sqref.ranges
                    if not (
                        cell_range.min_col
                        == cell_range.max_col
                        == column_number
                    )
                ]

                if kept:
                    validation.sqref = " ".join(kept)
                else:
                    ws.data_validations.dataValidation.remove(
                        validation
                    )

            validation = DataValidation(
                type=kind,
                allow_blank=True,
            )

            if kind == "list":
                validation.formula1 = '"✓"'
            else:
                validation.operator = "between"
                validation.formula1 = "1"
                validation.formula2 = "2958465"

            validation.showErrorMessage = True
            validation.error = (
                "Choose the checkmark or enter a valid Excel date."
            )

            ws.add_data_validation(validation)
            validation.add(
                f"{column_letter}{start}:{column_letter}{end}"
            )


# ---------------------------------------------------------------------------
# Save without creating backup copies
# ---------------------------------------------------------------------------

def save_in_place(wb, path, stat_before, expected_filled_rows):
    """Validate in memory, then save to the selected workbook."""
    buffer = io.BytesIO()
    wb.save(buffer)
    data = buffer.getvalue()

    check = load_workbook(
        io.BytesIO(data),
        keep_vba=path.suffix.lower() == ".xlsm",
    )

    try:
        check_cur = get_table_info(
            check,
            CURRENT_SHEET,
            CURRENT_TABLE,
            APP_COLUMNS,
        )

        if count_filled_rows(check_cur) != expected_filled_rows:
            raise RuntimeError(
                "Row-count verification failed. "
                "Original file was not changed."
            )

        for sheet_name in (CURRENT_SHEET, REPROCESSED_SHEET):
            before = wb[sheet_name]
            after = check[sheet_name]

            for row in before.iter_rows():
                for cell in row:
                    if (
                        cell.value is not None
                        and cell_signature(cell)
                        != cell_signature(after[cell.coordinate])
                    ):
                        raise RuntimeError(
                            "Saved-content verification failed at "
                            f"{sheet_name}!{cell.coordinate}. "
                            "Original unchanged."
                        )
    finally:
        check.close()

    original = path.read_bytes()
    current = path.stat()

    if (
        current.st_mtime_ns,
        current.st_size,
    ) != (
        stat_before.st_mtime_ns,
        stat_before.st_size,
    ):
        raise RuntimeError(
            "Workbook changed during import. "
            "Original was not changed by this app."
        )

    # Verify that the embedded VBA project remains intact.
    with (
        zipfile.ZipFile(io.BytesIO(original)) as old_archive,
        zipfile.ZipFile(io.BytesIO(data)) as new_archive,
    ):
        vba_path = "xl/vbaProject.bin"

        if vba_path in old_archive.namelist():
            if (
                vba_path not in new_archive.namelist()
                or old_archive.read(vba_path)
                != new_archive.read(vba_path)
            ):
                raise RuntimeError(
                    "VBA preservation check failed. Original unchanged."
                )

    # Obtain write access before changing anything.
    # No backup files are created.
    try:
        stream = path.open("r+b")
    except PermissionError as exc:
        raise PermissionError(
            "Close this workbook in Excel and try again. "
            "If it is already closed, check your write permissions. "
            "Nothing was changed and no backup was created."
        ) from exc

    with stream:
        if stream.read() != original:
            raise RuntimeError(
                "Workbook changed before saving. "
                "Import canceled; nothing written."
            )

        try:
            stream.seek(0)
            stream.write(data)
            stream.truncate()
            stream.flush()
            os.fsync(stream.fileno())

            stream.seek(0)

            if stream.read() != data:
                raise OSError(
                    "Saved bytes do not match generated workbook"
                )

        except Exception as exc:
            # Attempt to restore the original bytes from memory.
            # This does not create a separate workbook or backup file.
            try:
                stream.seek(0)
                stream.write(original)
                stream.truncate()
                stream.flush()
                os.fsync(stream.fileno())

                stream.seek(0)

                if stream.read() != original:
                    raise OSError(
                        "Original-byte verification failed"
                    )

            except Exception as restore_error:
                raise RuntimeError(
                    "Saving failed and the original contents could "
                    "not be restored. Check the workbook and use your "
                    "existing version history if needed. "
                    f"Save error: {exc}; "
                    f"restoration error: {restore_error}"
                ) from restore_error

            raise RuntimeError(
                "Saving failed, but the original workbook contents "
                f"were restored. Details: {exc}"
            ) from exc


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------

def run(root):
    # 1. Select the PDF.
    downloads = Path.home() / "Downloads"

    selected_pdf = filedialog.askopenfilename(
        parent=root,
        title="Select the ACH Return PDF",
        initialdir=str(
            downloads if downloads.exists() else Path.cwd()
        ),
        filetypes=[("PDF files", "*.pdf")],
    )

    if not selected_pdf:
        print("Canceled. Nothing was changed.")
        return

    pdf_path = Path(selected_pdf)

    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    reader = PdfReader(str(pdf_path))
    page_texts = [
        page.extract_text() or ""
        for page in reader.pages
    ]

    report_date, records, skipped = parse_pdf_records(page_texts)

    # 2. Select the workbook.
    box_folder = Path.home() / "Box"

    selected_excel = filedialog.askopenfilename(
        parent=root,
        title="Select the ACH Returns Excel file to update",
        initialdir=str(
            box_folder if box_folder.exists() else Path.home()
        ),
        filetypes=[("Excel files", "*.xlsx *.xlsm")],
    )

    if not selected_excel:
        print("Canceled. Nothing was changed.")
        return

    excel_path = Path(selected_excel)

    # Check Excel's owner file before loading or editing.
    owner_file = excel_path.with_name(
        "~$" + excel_path.name
    )

    if owner_file.exists():
        raise PermissionError(
            "Excel reports this workbook is in use. "
            "Close it and try again."
        )

    # Test write access without writing anything.
    try:
        with excel_path.open("r+b"):
            pass
    except PermissionError as exc:
        raise PermissionError(
            "Close the workbook in Excel and try again. "
            "If already closed, check write permissions."
        ) from exc

    stat_before = excel_path.stat()
    wb, load_notes = load_workbook_checked(excel_path)

    try:
        cur = get_table_info(
            wb,
            CURRENT_SHEET,
            CURRENT_TABLE,
            APP_COLUMNS,
        )

        rep = get_table_info(
            wb,
            REPROCESSED_SHEET,
            REPROCESSED_TABLE,
            ["Report Date"],
        )

        # 3. Warn if this report date already appears in either sheet.
        already = (
            count_rows_with_report_date(
                cur,
                report_date.date(),
            )
            + count_rows_with_report_date(
                rep,
                report_date.date(),
            )
        )

        if already:
            proceed = messagebox.askyesno(
                parent=root,
                title=APP_TITLE,
                icon="warning",
                default="no",
                message=(
                    f"This workbook already has {already} row(s) "
                    "from a report "
                    f"dated {report_date:%m/%d/%Y}.\n\n"
                    "Importing it again will likely add duplicates. "
                    "Add it anyway?"
                ),
            )

            if not proceed:
                print("Canceled. Nothing was changed.")
                return

        # 4. Confirm the import.
        details = []

        if skipped:
            details.append(
                f"Skipped {skipped} record(s) with a $0.00 amount."
            )

        if filter_is_active(cur):
            details.append(
                "A filter (or hidden rows) was found on Current Returns. "
                "It will be cleared. Re-apply your filter after the import."
            )

        years = set(
            re.findall(
                r"(?<!\d)(20\d{2})(?!\d)",
                excel_path.name,
            )
        )

        if years and str(report_date.year) not in years:
            details.append(
                "Heads up: the file name mentions "
                f"{', '.join(sorted(years))} "
                f"but this report is dated {report_date:%m/%d/%Y}."
            )

        for note in load_notes:
            details.append(f"Excel feature warning: {note}")

        confirmed = messagebox.askyesno(
            parent=root,
            title=APP_TITLE,
            message=(
                f"Add {len(records)} return(s) "
                f"dated {report_date:%m/%d/%Y} "
                f'to "{excel_path.name}"?'
            ),
            detail="\n\n".join(details) if details else None,
        )

        if not confirmed:
            print("Canceled. Nothing was changed.")
            return

        # 5. Update the in-memory workbook.
        filled_before = count_filled_rows(cur)

        old_rows = [
            [
                cell_signature(cur.ws.cell(row, column))
                for column in range(cur.c1, cur.c2 + 1)
            ]
            for row in range(cur.r1 + 1, cur.r2 + 1)
        ]

        insert_records(cur, records)

        # Confirm existing nonformula values, links, and notes moved intact.
        if filled_before:
            for offset, old_row in enumerate(old_rows):
                for column_offset, signature in enumerate(old_row):
                    if signature[1] == "f":
                        # Formulas are intentionally translated.
                        continue

                    cell = cur.ws.cell(
                        cur.r1 + 1 + len(records) + offset,
                        cur.c1 + column_offset,
                    )

                    if cell_signature(cell) != signature:
                        raise RuntimeError(
                            "Existing-row preservation check failed. "
                            "Original unchanged."
                        )

        maintain_rules(wb)

        # 6. Validate and save without creating backup files.
        save_in_place(
            wb,
            excel_path,
            stat_before,
            filled_before + len(records),
        )

    finally:
        wb.close()

    summary = (
        f"Added {len(records)} return(s) "
        "to the top of Current Returns."
        + (
            f"\nSkipped {skipped} record(s) with a $0.00 amount."
            if skipped
            else ""
        )
        + f"\n\nSaved: {excel_path}"
    )

    print(summary)

    messagebox.showinfo(
        parent=root,
        title=APP_TITLE,
        message=summary,
    )


def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        run(root)

    except PermissionError as error:
        message = str(error)
        print(message)

        messagebox.showerror(
            parent=root,
            title=APP_TITLE,
            message=message,
        )

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