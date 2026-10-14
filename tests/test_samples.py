"""End-to-end: build a problem, solve it, let LibreOffice calculate, check answers."""

import pytest

from solver import accounts as A
from solver.totals import canonical


def find(ws, label, col):
    """Value in column `col` of the first row whose column-A text contains `label`."""
    for row in ws.iter_rows():
        v = row[0].value
        if isinstance(v, str) and label.lower() in v.lower():
            return ws.cell(row=row[0].row, column=col).value
    raise AssertionError(f"{label!r} not found on {ws.title}")


def test_journal_builds_everything(solved):
    result, wb = solved("journal")
    assert wb["Journal"]["E10"].value == 5000          # missing credit worked out
    tb = wb["Trial Balance"]
    assert find(tb, "Totals", 2) == find(tb, "Totals", 3) == 30250
    assert find(wb["Income Statement"], "Net income", 3) == 1350
    assert find(wb["Owner's Equity"], "Capital, ending", 3) == 20350
    bs = wb["Balance Sheet"]
    assert find(bs, "Total assets", 3) == 25950
    assert find(bs, "Check:", 3) == 0
    assert "General Ledger" in wb.sheetnames


def test_trial_balance_missing_amount(solved):
    result, wb = solved("tb_missing")
    ws = wb["TB"]
    assert find(ws, "Retained Earnings", 3) == 9000
    assert find(ws, "Totals", 2) == find(ws, "Totals", 3) == 109000
    assert find(wb["Income Statement"], "Gross profit", 3) == 24000
    assert find(wb["Income Statement"], "Net income", 3) == 13000
    assert find(wb["Balance Sheet"], "Check:", 3) == 0


def test_fill_existing_statement_templates(solved):
    result, wb = solved("templates")
    inc = wb["Income Statement"]
    assert find(inc, "Total expenses", 3) == 9000
    assert find(inc, "Net income", 3) == 9000
    re_ws = wb["Retained Earnings"]
    assert find(re_ws, "Add: Net income", 2) == 9000
    assert find(re_ws, "December 31", 2) == 14000
    bs = wb["Balance Sheet"]
    assert find(bs, "Total assets", 3) == 31500
    assert find(bs, "Total stockholders", 3) == 29000
    assert find(bs, "Total liabilities and", 3) == 31500
    assert find(bs, "Current ratio", 2) == pytest.approx(6.2)
    assert find(bs, "Debt to assets", 2) == pytest.approx(2500 / 31500)
    # no extra statements created when the workbook already has them
    assert not any(s.kind == "sheet" for s in result.steps)


def test_ten_column_worksheet(solved):
    result, wb = solved("worksheet10")
    ws = wb["Worksheet"]
    assert ws["F7"].value == 1000           # supplies after adjustment
    assert ws["H21"].value == 3950          # net income
    assert ws["H22"].value == ws["I22"].value == 10400
    assert ws["J22"].value == ws["K22"].value == 22250
    assert find(wb["Balance Sheet"], "Check:", 3) == 0
    assert find(wb["Owner's Equity"], "ending", 3) == 13450


def test_depreciation_ddb(solved):
    result, wb = solved("depreciation")
    ws = wb.active
    assert [ws.cell(row=r, column=2).value for r in range(7, 12)] == \
        [20000, 12000, 7200, 4320, 1480]
    assert ws["D11"].value == 5000
    assert find(ws, "Depreciable cost", 2) == 45000


def test_loan(solved):
    result, wb = solved("loan")
    assert find(wb.active, "Annual payment", 2) == pytest.approx(3741.10, abs=0.01)
    sched = wb["Loan Amortization"]
    assert sched["F6"].value == pytest.approx(0, abs=0.01)


def test_bond(solved):
    result, wb = solved("bond")
    ws = wb.active
    assert find(ws, "Issue price", 2) == pytest.approx(92278.27, abs=0.01)
    assert ws["E19"].value == pytest.approx(100000, abs=0.01)
    assert ws["C10"].value == pytest.approx(4613.91, abs=0.01)


def test_inventory(solved):
    result, wb = solved("inventory")
    ws = wb.active
    assert [ws.cell(row=9, column=c).value for c in (2, 3)] == [3400, 3750]
    assert [ws.cell(row=10, column=c).value for c in (2, 3)] == [1950, 1600]
    assert ws["D9"].value == pytest.approx(3566.67, abs=0.01)


def test_steps_are_one_cell_each(solved):
    result, wb = solved("templates")
    cells = [(s.sheet, s.cell) for s in result.steps if s.kind == "fill"]
    assert len(cells) == len(set(cells))
    assert all(s.note for s in result.steps)


@pytest.mark.parametrize("name,expected", [
    ("Cash", A.ASSET), ("Accumulated Depreciation - Equipment", A.CONTRA_ASSET),
    ("Prepaid Insurance", A.ASSET), ("Unearned Service Revenue", A.LIABILITY),
    ("Salaries Payable", A.LIABILITY), ("Salaries Expense", A.EXPENSE),
    ("Service Revenue", A.REVENUE), ("Sales Returns and Allowances", A.CONTRA_REVENUE),
    ("Cost of Goods Sold", A.COGS), ("J. Smith, Drawing", A.DRAWING), ("Dividends", A.DRAWING),
    ("Dividends Payable", A.LIABILITY), ("Common Stock", A.EQUITY),
    ("Paid-in Capital in Excess of Par", A.EQUITY), ("Retained Earnings", A.EQUITY),
    ("Interest Receivable", A.ASSET), ("Interest Revenue", A.REVENUE),
    ("Income Tax Expense", A.EXPENSE), ("Income Summary", A.CLOSING),
    ("Treasury Stock", A.CONTRA_EQUITY), ("Purchase Returns and Allowances", A.CONTRA_COGS),
    ("Gain on Sale of Equipment", A.REVENUE), ("Loss on Sale of Equipment", A.EXPENSE),
    ("Allowance for Doubtful Accounts", A.CONTRA_ASSET), ("Supplies", A.ASSET),
    ("Supplies Expense", A.EXPENSE), ("Notes Payable", A.LIABILITY),
])
def test_classify(name, expected):
    assert A.classify(name) == expected


@pytest.mark.parametrize("label,key", [
    ("Net Income (Loss)", "net income"), ("Add: Net income", "net income"),
    ("Total Stockholders' Equity", "total equity"),
    ("Total liabilities and owner's equity", "total liabilities and equity"),
    ("Retained earnings, ending", "ending retained earnings"),
    ("Gross profit", "gross profit"), ("Gross profit rate", "gross profit rate"),
])
def test_canonical(label, key):
    assert canonical(label) == key


def test_journal_with_credit_column(solved):
    result, wb = solved("journal_two_columns")
    tb = wb["Trial Balance"]
    assert find(tb, "Cash", 2) == 10300
    assert find(tb, "Totals", 2) == find(tb, "Totals", 3) == 11500
    assert wb["Trial Balance"]["A1"].value != "Description"
    assert find(wb["Balance Sheet"], "Check:", 3) == 0


def test_trial_balance_without_account_header(solved):
    result, wb = solved("tb_no_header")
    assert find(wb["Income Statement"], "Net income", 3) == 1000
    assert find(wb["Owner's Equity"], "ending", 3) == 5800
    assert find(wb["Balance Sheet"], "Check:", 3) == 0


def test_comparative_columns(solved):
    result, wb = solved("comparative")
    ws = wb.active
    assert find(ws, "Total assets", 2) == 18000
    assert find(ws, "Total assets", 3) == 16500
    assert find(ws, "Total liabilities and", 2) == 18000
    assert find(ws, "Total liabilities and", 3) == 16500


def test_plain_blank_merchandising_statement(solved):
    result, wb = solved("merchandising")
    ws = wb.active
    assert find(ws, "Net sales", 2) == 95000
    assert find(ws, "Gross profit", 2) == 35000
    assert find(ws, "Total operating expenses", 2) == 21000
    assert find(ws, "Income from operations", 2) == 14000
    assert find(ws, "Income before income taxes", 2) == 13000
    assert find(ws, "Net income", 2) == 10000
    assert find(ws, "Gross profit rate", 2) == pytest.approx(35000 / 95000)
    assert find(ws, "Profit margin", 2) == pytest.approx(10000 / 95000)
