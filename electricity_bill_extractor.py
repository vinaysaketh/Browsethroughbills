"""
Electricity Bill HTML Extractor
Select a folder containing HTML electricity bills and export a consolidated Excel workbook.

The extraction rules are generic. Validate against sample bills from your DISCOM before
using the output for formal reporting.
"""

import re
import sys
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


FIELDS = {
    "CMD (kVA)": [
        r"\bCMD\b", r"Contracted\s+Maximum\s+Demand", r"Contract\s+Demand"
    ],
    "RMD (kVA)": [
        r"\bRMD\b", r"Recorded\s+Maximum\s+Demand",
        r"Maximum\s+Demand\s+Recorded", r"Recorded\s+Demand"
    ],
    "kWh": [
        r"Active\s+Energy", r"Energy\s+Consumed\s*\(?\s*kWh",
        r"Units\s+Consumed", r"\bkWh\b"
    ],
    "kVAh": [
        r"Apparent\s+Energy", r"Energy\s+Consumed\s*\(?\s*kVAh",
        r"\bkVAh\b"
    ],
    "PF": [r"Power\s+Factor", r"\bPF\b"],
    "Total Electricity Bill (INR)": [
        r"Total\s+Electricity\s+Bill", r"Total\s+Current\s+Bill",
        r"Net\s+Amount\s+Payable", r"Total\s+Amount\s+Payable",
        r"Current\s+Bill\s+Amount", r"Amount\s+Payable",
        r"Total\s+Bill"
    ],
}

# These phrases are deliberately excluded as bill-total labels where possible.
TOTAL_EXCLUSIONS = re.compile(
    r"arrears|previous\s+balance|security\s+deposit|"
    r"adjustment|late\s+payment|surcharge",
    re.IGNORECASE
)


def clean_number(value):
    """Convert a displayed number to float; return None if no number is found."""
    if value is None:
        return None
    value = str(value).replace(",", "").replace("₹", "")
    value = re.sub(r"\b(?:INR|Rs\.?)\b", "", value, flags=re.IGNORECASE)
    match = re.search(r"[-+]?\d+(?:\.\d+)?", value)
    return float(match.group()) if match else None


def label_matches(label, patterns):
    return any(re.search(p, label, re.IGNORECASE) for p in patterns)


def extract_from_rows(rows, field, patterns):
    """
    Find a matching label in an HTML table and read the closest following cell.
    Handles common layouts: Label | Value and Label | Unit | Value.
    """
    for row in rows:
        for idx, cell in enumerate(row):
            if not label_matches(cell, patterns):
                continue
            if field == "Total Electricity Bill (INR)" and TOTAL_EXCLUSIONS.search(cell):
                continue

            # Search the next few cells, not the whole row, to avoid unrelated values.
            for candidate in row[idx + 1:idx + 4]:
                number = clean_number(candidate)
                if number is not None:
                    return number
    return None


def extract_from_text(page_text, field, patterns):
    """Fallback for non-tabular HTML labels."""
    for pattern in patterns:
        regex = (
            pattern
            + r"\s*(?:[:\-]\s*|\s{1,})"
            + r"(?:₹|INR|Rs\.?)?\s*"
            + r"([-+]?\d[\d,]*(?:\.\d+)?)"
        )
        match = re.search(regex, page_text, re.IGNORECASE)
        if match:
            if field == "Total Electricity Bill (INR)":
                nearby = page_text[max(0, match.start() - 60):match.start()]
                if TOTAL_EXCLUSIONS.search(nearby):
                    continue
            return clean_number(match.group(1))
    return None


def extract_bill(html_path):
    raw = html_path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "lxml")

    rows = []
    for tr in soup.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
        if cells:
            rows.append(cells)

    page_text = soup.get_text(" ", strip=True)

    record = {
        "File Name": html_path.name,
        "CMD (kVA)": None,
        "RMD (kVA)": None,
        "kWh": None,
        "kVAh": None,
        "PF": None,
        "Total Electricity Bill (INR)": None,
    }

    for field, patterns in FIELDS.items():
        value = extract_from_rows(rows, field, patterns)
        if value is None:
            value = extract_from_text(page_text, field, patterns)
        record[field] = value

    return record


def format_workbook(output_path, dataframe):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dataframe.to_excel(writer, sheet_name="Bill Summary", index=False)
        ws = writer.sheets["Bill Summary"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        header_fill = PatternFill("solid", fgColor="1F4E78")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        for col_cells in ws.columns:
            width = max(len(str(c.value or "")) for c in col_cells) + 3
            ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(width, 42)

        # Numeric display formatting
        for row in ws.iter_rows(min_row=2):
            for cell in row[2:]:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'


def main():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    folder = filedialog.askdirectory(title="Select folder containing HTML electricity bills")
    if not folder:
        return

    input_folder = Path(folder)
    html_files = sorted(
        list(input_folder.glob("*.html")) + list(input_folder.glob("*.htm"))
    )

    if not html_files:
        messagebox.showwarning("No HTML files", "No .html or .htm files were found in the selected folder.")
        return

    records = []
    errors = []
    for html_file in html_files:
        try:
            records.append(extract_bill(html_file))
        except Exception as exc:
            errors.append(f"{html_file.name}: {exc}")
            records.append({
                "File Name": html_file.name,
                **{field: None for field in FIELDS}
            })

    df = pd.DataFrame(records)
    required = list(FIELDS.keys())
    df["Extraction Status"] = df[required].notna().all(axis=1).map(
        {True: "Complete", False: "Verify Missing Values"}
    )

    output_file = filedialog.asksaveasfilename(
        title="Save consolidated Excel workbook",
        defaultextension=".xlsx",
        initialfile="Consolidated_Electricity_Bills.xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")]
    )
    if not output_file:
        return

    try:
        format_workbook(output_file, df)
    except Exception as exc:
        messagebox.showerror("Export failed", f"Could not create Excel file:\n{exc}")
        return

    summary = (
        f"Processed HTML files: {len(html_files)}\n"
        f"Complete records: {(df['Extraction Status'] == 'Complete').sum()}\n"
        f"Records to verify: {(df['Extraction Status'] != 'Complete').sum()}\n\n"
        f"Excel saved to:\n{output_file}"
    )
    if errors:
        summary += "\n\nSome files had parsing errors. Check the console/log if run from source."

    messagebox.showinfo("Extraction completed", summary)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        # Keep a useful error message even in the no-console EXE.
        try:
            messagebox.showerror("Unexpected error", str(exc))
        except Exception:
            print(f"Unexpected error: {exc}", file=sys.stderr)
            raise
