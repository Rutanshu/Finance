"""
Builds a realistic sample month-end folder so you can try PaperTrail in 30 seconds
without pointing it at real company data.

    python make_sample.py
    python -m papertrail sample_books
"""
from pathlib import Path
import zipfile
import openpyxl

OUT = Path(__file__).parent / "sample_books"
OUT.mkdir(exist_ok=True)


def book(name, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sname, rows in sheets.items():
        ws = wb.create_sheet(sname[:31])
        for r in rows:
            ws.append(r)
    wb.save(OUT / name)
    print("  wrote", name)


def inject_connections(name, entries):
    """
    Stamp a synthetic xl/connections.xml into an already-saved workbook, so the sample folder
    demonstrates PaperTrail's database/web-query detection without needing real ERP/API access.
    Not meant to round-trip through Excel — just to exercise the same zip member PaperTrail reads.
    """
    xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    for i, e in enumerate(entries, start=1):
        if e["type"] == "db":
            xml.append(f'<connection id="{i}" name="{e["name"]}"><dbPr connection="{e["conn"]}" '
                       f'command="SELECT * FROM {e.get("table", "dbo.GL")}"/></connection>')
        else:
            xml.append(f'<connection id="{i}" name="{e["name"]}"><webPr sourceData="1" url="{e["url"]}"/></connection>')
    xml.append("</connections>")
    with zipfile.ZipFile(OUT / name, "a", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/connections.xml", "\n".join(xml))
    print("  + connections.xml ->", name)


def inject_query_table(name, sheet_part="sheet1.xml", connection_id=1):
    """
    Wire a connection to the specific worksheet it loads into — the same
    xl/worksheets/_rels/*.rels -> queryTables/*.xml -> connectionId chain real Excel/Power Query
    produces, so PaperTrail's sheet-attribution can actually be exercised against sample data.
    """
    qt = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          f'name="ExternalData_{connection_id}" connectionId="{connection_id}" autoFormatId="16"/>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rIdQT1" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" '
            f'Target="../queryTables/queryTable{connection_id}.xml"/></Relationships>')
    with zipfile.ZipFile(OUT / name, "a", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"xl/queryTables/queryTable{connection_id}.xml", qt)
        z.writestr(f"xl/worksheets/_rels/{sheet_part}.rels", rels)
    print("  + queryTable ->", name, sheet_part)


print("building sample workbooks in", OUT)

book("Sales_Export.xlsx", {"Data": [
    ["invoice_id", "customer", "amount_net", "tax", "invoice_date"],
    ["INV-1001", "Acme Ltd", 128400, 23112, "2026-07-03"],
    ["INV-1002", "Borealis", 94500, 17010, "2026-07-09"],
    ["INV-1003", "Cresta", 210000, 37800, "2026-07-21"],
]})

book("Payroll_Export.xlsx", {"Run": [
    ["employee_id", "gross_pay", "employer_cost", "cost_centre"],
    ["E-01", 82000, 9840, "CC-100"],
    ["E-02", 61000, 7320, "CC-200"],
]})

book("Bank_Statement.xlsx", {"Stmt": [
    ["value_date", "description", "debit", "credit"],
    ["2026-07-05", "ACME PAYMENT", 0, 151512],
    ["2026-07-18", "SUPPLIER XYZ", 44000, 0],
]})

book("Revenue_Register.xlsx", {
    "Data": [
        ["invoice_id", "customer", "amount_net", "gl_code"],
        ["INV-1001", "Acme Ltd", "='[Sales_Export.xlsx]Data'!$C$2", 4000],
        ["INV-1002", "Borealis", "='[Sales_Export.xlsx]Data'!$C$3", 4000],
        ["INV-1003", "Cresta", "='[Sales_Export.xlsx]Data'!$C$4", 4100],
    ],
    "Summary": [
        ["gl_code", "total"],
        [4000, "=SUMIF(Data!D2:D4,A2,Data!C2:C4)"],
        [4100, "=SUMIF(Data!D2:D4,A3,Data!C2:C4)"],
        ["with rounding", "=ROUND(B2*1.18,2)+12500"],
    ],
})

book("Payroll_Summary.xlsx", {"Summary": [
    ["cost_centre", "gross_pay", "gl_code"],
    ["CC-100", "='[Payroll_Export.xlsx]Run'!$B$2", 6000],
    ["CC-200", "='[Payroll_Export.xlsx]Run'!$B$3", 6000],
    ["total", "=SUM(B2:B3)", ""],
]})

book("AP_Ledger.xlsx", {"Invoices": [
    ["vendor", "invoice_no", "amount", "gl_code", "accrual"],
    ["Supplier XYZ", "S-771", 44000, 5200, "N"],
    ["Vendor Q", "VQ-19", 18750, 5300, "Y"],
    ["late fee", "", "=C2*0.025+1500", "", ""],
]})

book("Bank_Recon.xlsx", {"Recon": [
    ["value_date", "book_amount", "bank_amount", "variance"],
    ["2026-07-05", 151512, "='[Bank_Statement.xlsx]Stmt'!$D$2", "=B2-C2"],
    ["2026-07-18", 44000, "='[Bank_Statement.xlsx]Stmt'!$C$3", "=B3-C3"],
    ["unmatched", "", "", "=SUM(D2:D3)"],
]})

book("Trial_Balance.xlsx", {"TB": [
    ["gl_code", "account", "debit", "credit"],
    [4000, "Revenue - services", 0, "='[Revenue_Register.xlsx]Summary'!$B$2"],
    [4100, "Revenue - goods", 0, "='[Revenue_Register.xlsx]Summary'!$B$3"],
    [6000, "Payroll", "='[Payroll_Summary.xlsx]Summary'!$B$4", 0],
    [5200, "Cost of sales", "='[AP_Ledger.xlsx]Invoices'!$C$2", 0],
    [1000, "Cash at bank", "='[Bank_Recon.xlsx]Recon'!$B$2", 0],
    ["check", "debits less credits", "=SUM(C2:C6)-SUM(D2:D6)", ""],
]})

book("Journals.xlsx", {"Journals": [
    ["journal_no", "gl_code", "amount", "narrative", "approver"],
    ["J-01", 5300, 18750, "Accrue Vendor Q", "Controller"],
    ["J-02", 6000, "=OFFSET(C2,0,0)*0.5", "Bonus accrual", "Controller"],
]})

book("Management_Accounts.xlsx", {
    "PL": [
        ["line", "amount"],
        ["Revenue", "='[Trial_Balance.xlsx]TB'!$D$2+'[Trial_Balance.xlsx]TB'!$D$3"],
        ["Cost of sales", "='[Trial_Balance.xlsx]TB'!$C$5"],
        ["Payroll", "='[Trial_Balance.xlsx]TB'!$C$4"],
        ["Adjustments", "='[Journals.xlsx]Journals'!$C$2"],
        ["EBITDA", "=B2-B3-B4-B5"],
    ],
    "BS": [
        ["line", "amount"],
        ["Cash", "='[Trial_Balance.xlsx]TB'!$C$6"],
        ["Receivables", 245000],
    ],
})

book("Board_Pack_Figures.xlsx", {"KPI": [
    ["kpi", "value"],
    ["EBITDA", "='[Management_Accounts.xlsx]PL'!$B$6"],
    ["Cash", "='[Management_Accounts.xlsx]BS'!$B$2"],
    ["Budget variance", "=INDIRECT(\"B2\")-1850000"],
]})

# a deliberately broken link — points at a file that is not in the folder
book("Segment_Report.xlsx", {"Seg": [
    ["segment", "revenue"],
    ["North", "='[Regional_Split_v7_FINAL.xlsx]Data'!$B$2"],
    ["South", "='[Regional_Split_v7_FINAL.xlsx]Data'!$B$3"],
]})

# ---- external data sources: database, web query, and a live web formula -----------------------
# ERP_Extract.xlsx *looks* like a plain workbook, but it is refreshed from the finance ERP over
# a database connection — the giveaway lives in xl/connections.xml, not in any formula.
book("ERP_Extract.xlsx", {"Extract": [
    ["gl_code", "period", "amount"],
    [4000, "2026-07", 338800],
    [4100, "2026-07", 210000],
]})
inject_connections("ERP_Extract.xlsx", [
    {"type": "db", "name": "FinanceERP",
     "conn": "Provider=SQLOLEDB;Data Source=FINANCE-SQL01;Initial Catalog=ERP;UID=svc_finance;PWD=hunter2;",
     "table": "dbo.GL_Extract"},
])
inject_query_table("ERP_Extract.xlsx", sheet_part="sheet1.xml", connection_id=1)  # loads into "Extract"

# Market_Rates.xlsx is refreshed from a legacy web query — again, invisible to a formula scan.
book("Market_Rates.xlsx", {"Rates": [
    ["symbol", "rate"],
    ["GBPUSD", 1.27],
    ["EURUSD", 1.08],
]})
inject_connections("Market_Rates.xlsx", [
    {"type": "web", "name": "ECB reference rates", "url": "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"},
])
inject_query_table("Market_Rates.xlsx", sheet_part="sheet1.xml", connection_id=1)  # loads into "Rates"

# Treasury_Dashboard.xlsx converges all three external-source patterns into one file:
# a cross-file link to a DB-backed extract, a cross-file link to a web-query-backed file,
# and a formula that reaches the internet directly every time it recalculates.
book("Treasury_Dashboard.xlsx", {
    "FX": [
        ["currency", "spot_rate", "source"],
        ["EUR", '=WEBSERVICE("https://api.exchangerate.host/latest?base=USD&symbols=EUR")', "live API"],
        ["GBP", "='[Market_Rates.xlsx]Rates'!$B$2", "Power Query"],
    ],
    "GL": [
        ["gl_code", "amount"],
        [4000, "='[ERP_Extract.xlsx]Extract'!$C$2"],
        [4100, "='[ERP_Extract.xlsx]Extract'!$C$3"],
    ],
})

print("\ndone. now run:  python -m papertrail sample_books")
