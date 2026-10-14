"""Builds sample accounting problems (plain cells: no bold, no borders)."""

import datetime as dt

from openpyxl import Workbook


def _wb():
    wb = Workbook()
    return wb, wb.active


def journal_problem(path):
    """Journal only, one credit amount missing. Expect TB, ledger and statements."""
    wb, ws = _wb()
    ws.title = "Journal"
    ws["A1"] = "Sunny Day Services"
    ws["A2"] = "General Journal"
    ws.append([])
    ws.append(["Date", "Account Titles and Explanation", "Ref", "Debit", "Credit"])
    d = dt.date
    rows = [
        (d(2025, 1, 1), "Cash", 20000, None), (None, "   J. Sunny, Capital", None, 20000),
        (None, "(Owner investment)", None, None),
        (d(2025, 1, 3), "Equipment", 8000, None), (None, "   Cash", None, 3000),
        (None, "   Notes Payable", None, None),            # missing: 5000
        (d(2025, 1, 10), "Accounts Receivable", 4500, None), (None, "   Service Revenue", None, 4500),
        (d(2025, 1, 15), "Rent Expense", 1200, None), (None, "   Cash", None, 1200),
        (d(2025, 1, 20), "Salaries Expense", 1800, None), (None, "   Cash", None, 1800),
        (d(2025, 1, 25), "Cash", 2500, None), (None, "   Accounts Receivable", None, 2500),
        (d(2025, 1, 28), "Supplies", 600, None), (None, "   Accounts Payable", None, 600),
        (d(2025, 1, 31), "J. Sunny, Drawing", 1000, None), (None, "   Cash", None, 1000),
        (d(2025, 1, 31), "Depreciation Expense", 150, None),
        (None, "   Accumulated Depreciation - Equipment", None, 150),
    ]
    for date, acct, dr, cr in rows:
        ws.append([date, acct, None, dr, cr])
    wb.save(path)


def tb_missing(path):
    wb, ws = _wb()
    ws.title = "TB"
    ws["A1"] = "Acme Corp"
    ws["A2"] = "Trial Balance"
    ws["A3"] = "December 31, 2025"
    ws.append([])
    ws.append(["Account", "Debit", "Credit"])
    for row in [("Cash", 15000, None), ("Accounts Receivable", 6000, None),
                ("Merchandise Inventory", 9000, None), ("Equipment", 30000, None),
                ("Accumulated Depreciation", None, 6000), ("Accounts Payable", None, 4000),
                ("Notes Payable", None, 10000), ("Common Stock", None, 20000),
                ("Retained Earnings", None, None),       # missing: 8000
                ("Dividends", 2000, None), ("Sales Revenue", None, 60000),
                ("Sales Returns and Allowances", 1000, None), ("Cost of Goods Sold", 35000, None),
                ("Salaries Expense", 8000, None), ("Rent Expense", 2000, None),
                ("Income Tax Expense", 1000, None)]:
        ws.append(list(row))
    ws.append(["Totals", None, None])
    wb.save(path)


def templates(path):
    """A trial balance plus blank statements the student must complete."""
    wb, tb = _wb()
    tb.title = "Trial Balance"
    tb.append(["Account", "Debit", "Credit"])
    for row in [("Cash", 12000, None), ("Accounts Receivable", 3000, None),
                ("Supplies", 500, None), ("Equipment", 20000, None),
                ("Accumulated Depreciation", None, 4000), ("Accounts Payable", None, 2500),
                ("Common Stock", None, 15000), ("Retained Earnings", None, 6000),
                ("Dividends", 1000, None), ("Service Revenue", None, 18000),
                ("Salaries Expense", 6000, None), ("Rent Expense", 2400, None),
                ("Depreciation Expense", 600, None)]:
        tb.append(list(row))
    tb.append(["Totals", 45500, 45500])

    ws = wb.create_sheet("Income Statement")
    for row in [["Income Statement"], ["For the Year Ended December 31, 2025"], [],
                ["Revenues"], ["Service Revenue", None, None], ["Expenses"],
                ["Salaries Expense", None], ["Rent Expense", None],
                ["Depreciation Expense", None], ["Total expenses", None, None],
                ["Net income", None, None]]:
        ws.append(row)
    ws["B5"] = "?"         # one placeholder so the item column is known
    ws["C10"] = "?"

    re_ws = wb.create_sheet("Retained Earnings")
    for row in [["Statement of Retained Earnings"], [],
                ["Retained earnings, January 1", None], ["Add: Net income", None],
                ["Less: Dividends", None], ["Retained earnings, December 31", None]]:
        re_ws.append(row)

    bs = wb.create_sheet("Balance Sheet")
    for row in [["Balance Sheet"], ["December 31, 2025"], [], ["Assets"], ["Cash", "?"],
                ["Accounts Receivable", None], ["Supplies", None], ["Equipment", None],
                ["Less: Accumulated Depreciation", None], ["Total assets", None, "?"], [],
                ["Liabilities and Stockholders' Equity"], ["Accounts Payable", None],
                ["Total liabilities", None, None], ["Common Stock", None],
                ["Retained Earnings", None], ["Total stockholders' equity", None, None],
                ["Total liabilities and stockholders' equity", None, None], [],
                ["Current ratio", None], ["Debt to assets", None]]:
        bs.append(row)
    wb.save(path)


def worksheet10(path):
    wb, ws = _wb()
    ws.title = "Worksheet"
    ws["A1"] = "Pioneer Advertising"
    ws["A2"] = "Worksheet"
    ws.append([])
    ws.append(["Account Titles", "Trial Balance", None, "Adjustments", None,
               "Adjusted Trial Balance", None, "Income Statement", None, "Balance Sheet"])
    ws.append([None, "Dr", "Cr", "Dr", "Cr", "Dr", "Cr", "Dr", "Cr", "Dr", "Cr"])
    rows = [("Cash", 15200, None, None, None), ("Supplies", 2500, None, None, 1500),
            ("Prepaid Insurance", 600, None, None, 50), ("Equipment", 5000, None, None, None),
            ("Notes Payable", None, 5000, None, None), ("Accounts Payable", None, 2500, None, None),
            ("Unearned Service Revenue", None, 1200, 400, None),
            ("Owner's Capital", None, 10000, None, None), ("Owner's Drawings", 500, None, None, None),
            ("Service Revenue", None, 10000, None, 400), ("Salaries Expense", 4000, None, None, None),
            ("Rent Expense", 900, None, None, None),
            ("Supplies Expense", None, None, 1500, None), ("Insurance Expense", None, None, 50, None)]
    for name, tdr, tcr, adr, acr in rows:
        ws.append([name, tdr, tcr, adr, acr])
    ws.append(["Totals"])
    ws.append(["Net Income"])
    ws.append(["Totals"])
    wb.save(path)


def depreciation(path):
    wb, ws = _wb()
    ws.title = "Problem"
    ws.append(["Equipment cost", 50000])
    ws.append(["Salvage value", 5000])
    ws.append(["Useful life (years)", 5])
    ws.append(["Method", "Double-declining-balance"])
    ws.append([])
    ws.append(["Year", "Depreciation expense", "Accumulated depreciation", "Book value"])
    for y in range(1, 6):
        ws.append([y])
    ws.append([])
    ws.append(["Depreciable cost", None])
    wb.save(path)


def loan(path):
    wb, ws = _wb()
    ws.append(["Loan amount", 10000])
    ws.append(["Annual interest rate", 0.06])
    ws.append(["Term (years)", 3])
    ws.append(["Payments are made annually"])
    ws.append(["Annual payment", None])
    wb.save(path)


def bond(path):
    wb, ws = _wb()
    ws.append(["Face value", 100000])
    ws.append(["Stated rate", 0.08])
    ws.append(["Market rate", 0.10])
    ws.append(["Term (years)", 5])
    ws.append(["Interest is paid semiannually"])
    ws.append(["Issue price", "?"])
    ws.append([])
    ws.append(["Period", "Cash paid", "Interest expense", "Discount amortization",
               "Carrying value"])
    ws.append(["Issue"])
    wb.save(path)


def inventory(path):
    wb, ws = _wb()
    ws.append(["Date", "Description", "Units", "Unit cost", "Total cost"])
    ws.append(["Jan 1", "Beginning inventory", 100, 10])
    ws.append(["Mar 5", "Purchase", 200, 12])
    ws.append(["Aug 9", "Purchase", 150, 13])
    ws.append([])
    ws.append(["Units sold", 300])
    ws.append([])
    ws.append([None, "FIFO", "LIFO", "Weighted average"])
    ws.append(["Cost of goods sold"])
    ws.append(["Ending inventory"])
    wb.save(path)


ALL = {"journal": journal_problem, "tb_missing": tb_missing, "templates": templates,
       "worksheet10": worksheet10, "depreciation": depreciation, "loan": loan, "bond": bond,
       "inventory": inventory}


def journal_two_columns(path):
    """Credit account titles in their own column, no 'Account' header, no explanations."""
    wb, ws = _wb()
    ws.append(["Date", "Description", None, "Dr", "Cr"])
    ws.append(["May 1", "Cash", None, 9000, None])
    ws.append([None, None, "Common Stock", None, 9000])
    ws.append([])
    ws.append(["May 4", "Office Supplies", None, 400, None])
    ws.append([None, None, "Cash", None, 400])
    ws.append([])
    ws.append(["May 9", "Cash", None, 2500, None])
    ws.append([None, None, "Consulting Revenue", None, 2500])
    ws.append([])
    ws.append(["May 30", "Wages Expense", None, 800, None])
    ws.append([None, None, "Cash", None, 800])
    wb.save(path)


def tb_no_header(path):
    wb, ws = _wb()
    ws.append(["Bright Co."])
    ws.append(["Adjusted Trial Balance"])
    ws.append([None, "Debit", "Credit"])
    for row in [("Cash", 4000, None), ("Accounts receivable", 1500, None),
                ("Prepaid rent", 600, None), ("Salaries payable", None, 300),
                ("Owner, capital", None, 5000), ("Owner, withdrawals", 200, None),
                ("Fees earned", None, 2400), ("Salaries expense", 1100, None),
                ("Rent expense", 300, None)]:
        ws.append(list(row))
    wb.save(path)


def comparative(path):
    wb, ws = _wb()
    ws.append(["Comparative Balance Sheet"])
    ws.append([None, 2025, 2024])
    ws.append(["Current assets"])
    ws.append(["Cash", 5000, 4000])
    ws.append(["Accounts receivable", 3000, 2500])
    ws.append(["Total current assets", None, None])
    ws.append(["Equipment", 10000, 10000])
    ws.append(["Total assets", None, None])
    ws.append(["Current liabilities"])
    ws.append(["Accounts payable", 2000, 1500])
    ws.append(["Total liabilities", None, None])
    ws.append(["Stockholders' equity"])
    ws.append(["Common stock", 10000, 10000])
    ws.append(["Retained earnings", 6000, 5000])
    ws.append(["Total stockholders' equity", None, None])
    ws.append(["Total liabilities and stockholders' equity", None, None])
    wb.save(path)


def merchandising(path):
    """Numbers given, every subtotal left plain blank (no ?, no borders, no bold)."""
    wb, ws = _wb()
    for row in [["Income Statement"], ["For the Year Ended December 31, 2025"],
                ["Sales revenue", 100000], ["Less: Sales returns and allowances", 3000],
                ["Less: Sales discounts", 2000], ["Net sales"], ["Cost of goods sold", 60000],
                ["Gross profit"], ["Operating expenses"], ["Salaries expense", 15000],
                ["Rent expense", 6000], ["Total operating expenses"],
                ["Income from operations"], ["Interest expense", 1000],
                ["Income before income taxes"], ["Income tax expense", 3000], ["Net income"],
                [], ["Gross profit rate"], ["Profit margin"]]:
        ws.append(row)
    wb.save(path)


ALL.update({"journal_two_columns": journal_two_columns, "tb_no_header": tb_no_header,
            "comparative": comparative, "merchandising": merchandising})
