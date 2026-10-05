#!/usr/bin/env python3
"""Generate internship_bank.xlsx following the user's "Internship bank" sheet structure.

Creates three sheets:
  1. "Indonesia Intern"
  2. "Indo Intern (Not in Website)"
  3. "Intern Abroad"

Each sheet has the same 10-column header. The "Indonesia Intern" sheet is
populated with one row per company; every data column other than No. and
Company is filled with the literal string "tidak tersedia" because no verified
data is available (do not fabricate company data).
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

HEADERS = [
    "No.",
    "Company",
    "Intern Role",
    "Application Date",
    "Internship Date",
    "Salary Range (Rp)",
    "Criteria/Eligibility",
    "Intern web",
    "About Company",
    "Office Location",
]

SHEET_NAMES = [
    "Indonesia Intern",
    "Indo Intern (Not in Website)",
    "Intern Abroad",
]

INDONESIA_INTERN_COMPANIES = [
    "DBS Bank",
    "Deloitte",
    "Huawei",
    "Accenture",
    "Tencent",
    "Alibaba",
    "Google Indonesia",
    "McKinsey",
    "JP Morgan",
    "Traveloka",
    "Grab",
    "Samsung",
    "Ticket.com",
    "BCA",
    "Citibank",
    "Microsoft Indonesia",
    "AWS Indonesia",
    "IBM Indonesia",
    "Oracle Indonesia",
    "SAP Indonesia",
    "Meta (Facebook) Indonesia",
    "UOB",
    "HSBC",
    "Garena",
    "Ericsson Indonesia",
    "Shopee",
    "CIMB Niaga",
    "PwC",
    "Sea Labs Indonesia",
    "Tokopedia",
    "Gojek",
]

PLACEHOLDER = "tidak tersedia"

# Output location for the generated workbook.
OUTPUT_PATH = Path(
    "/projects/sandbox/.kiro/workflow-output/internships/internship_bank.xlsx"
)


def write_header(sheet):
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)


def build_workbook():
    wb = Workbook()

    # First sheet: Indonesia Intern (with company data rows).
    indo = wb.active
    indo.title = SHEET_NAMES[0]
    write_header(indo)
    for idx, company in enumerate(INDONESIA_INTERN_COMPANIES, start=1):
        row = [idx, company] + [PLACEHOLDER] * (len(HEADERS) - 2)
        indo.append(row)

    # Remaining sheets: header only.
    for name in SHEET_NAMES[1:]:
        sheet = wb.create_sheet(title=name)
        write_header(sheet)

    return wb


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = build_workbook()
    wb.save(OUTPUT_PATH)
    print(f"Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
