"""The 10-column worksheet: Trial Balance | Adjustments | Adjusted TB |
Income Statement | Balance Sheet."""

import re

from openpyxl.utils import get_column_letter

from . import accounts
from .workbook import is_label, norm

ORDER = ["tb", "adj", "atb", "is", "bs"]


def group_roles(pairs):
    roles = {}
    for p in pairs:
        g = p.group
        if "unadjusted" in g or ("trial balance" in g and "adjusted" not in g):
            roles.setdefault("tb", p)
        elif "adjustment" in g:
            roles.setdefault("adj", p)
        elif "adjusted" in g:
            roles.setdefault("atb", p)
        elif "income" in g:
            roles.setdefault("is", p)
        elif "balance sheet" in g or "position" in g:
            roles.setdefault("bs", p)
    if len(roles) < 2:          # no group labels: use the standard left-to-right order
        order = ORDER if len(pairs) >= 5 else (["tb", "adj", "atb", "is"] if len(pairs) == 4
                                               else ["tb", "is", "bs"])
        roles = dict(zip(order, pairs))
    return roles


def _ref(col, row):
    return f"{get_column_letter(col)}{row}"


def fill_worksheet(ctx, t):
    ws = t.ws
    roles = group_roles(t.pairs)
    t.roles = roles
    tb, adj, atb = roles.get("tb"), roles.get("adj"), roles.get("atb")
    is_p, bs_p = roles.get("is"), roles.get("bs")
    for ln in t.account_lines:
        r = ln.row
        if atb and (tb or adj) and ctx.is_empty_cell(ws, r, atb.debit) and \
                ctx.is_empty_cell(ws, r, atb.credit):
            parts = []
            for p in (tb, adj):
                if p:
                    parts.append(f"({_ref(p.debit, r)}-{_ref(p.credit, r)})")
            net = "+".join(parts)
            ctx.fill(ws, r, atb.debit, f"=MAX(0,{net})", f"Adjusted debit balance of {ln.name}")
            ctx.fill(ws, r, atb.credit, f"=MAX(0,-({net}))",
                     f"Adjusted credit balance of {ln.name}")
        src = atb or tb
        if not src:
            continue
        if ln.acct_type == accounts.UNKNOWN:
            ctx.warn(f"{ws.title} row {r}: could not tell whether '{ln.name}' belongs on the "
                     f"income statement or balance sheet - add it to a 'Chart of Accounts' sheet.")
            continue
        dest = is_p if ln.acct_type in accounts.INCOME_STATEMENT_TYPES else bs_p
        if dest and ctx.is_empty_cell(ws, r, dest.debit) and ctx.is_empty_cell(ws, r, dest.credit):
            where = "income statement" if dest is is_p else "balance sheet"
            ctx.fill(ws, r, dest.debit, f"={_ref(src.debit, r)}",
                     f"{ln.name} extended to {where} columns")
            ctx.fill(ws, r, dest.credit, f"={_ref(src.credit, r)}",
                     f"{ln.name} extended to {where} columns")
    _net_income_rows(ctx, t, roles)


def _net_income_rows(ctx, t, roles):
    """Totals -> Net income (loss) -> final totals, below the accounts."""
    ws = t.ws
    totals = [ln.row for ln in t.lines if ln.total]
    if not totals:
        return
    first_total = totals[0]
    is_p, bs_p = roles.get("is"), roles.get("bs")
    ni_row = None
    for r in range(first_total + 1, first_total + 4):
        for c in t.acct_cols:
            v = ctx.value(ws, r, c)
            if is_label(v) and re.search(r"net (income|loss|profit|earnings)", norm(v)):
                ni_row = r
        if ni_row:
            break
    if not ni_row:
        return
    ctx.claim(ws, range(first_total, ni_row + 3))
    if is_p and bs_p:
        isd, isc = _ref(is_p.debit, first_total), _ref(is_p.credit, first_total)
        bsd, bsc = _ref(bs_p.debit, first_total), _ref(bs_p.credit, first_total)
        cells = {
            is_p.debit: (f"=MAX(0,{isc}-{isd})", "Net income: income statement credits minus debits"),
            is_p.credit: (f"=MAX(0,{isd}-{isc})", "Net loss: income statement debits minus credits"),
            bs_p.debit: (f"=MAX(0,{bsc}-{bsd})", "Net loss carried to the balance sheet"),
            bs_p.credit: (f"=MAX(0,{bsd}-{bsc})", "Net income carried to the balance sheet"),
        }
        for col, (formula, why) in cells.items():
            if ctx.is_empty_cell(ws, ni_row, col):
                ctx.fill(ws, ni_row, col, formula, why)
    # final totals row right under net income
    for r in range(ni_row + 1, ni_row + 3):
        labels = [ctx.value(ws, r, c) for c in t.acct_cols]
        if any(is_label(v) and norm(v).startswith("total") for v in labels):
            for p in (is_p, bs_p):
                if not p:
                    continue
                for col in (p.debit, p.credit):
                    if ctx.is_empty_cell(ws, r, col):
                        ctx.fill(ws, r, col, f"={_ref(col, first_total)}+{_ref(col, ni_row)}",
                                 "Totals including net income (should balance)")
            break
