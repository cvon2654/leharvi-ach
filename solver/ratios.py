"""Financial ratios: fill a blank next to "Current ratio", "Debt to equity", ..."""

import re

from .workbook import PERCENT_FMT, RATIO_FMT, norm

# (pattern, numerator concepts, denominator concepts, format, description)
# Each concept is a list of alternatives tried in order.
RATIOS = [
    (r"current ratio|working capital ratio", [["total current assets"]],
     [["total current liabilities"]], RATIO_FMT, "Current assets / current liabilities"),
    (r"quick ratio|acid[- ]test", [["cash"], ["short term investment", "marketable securities"],
                                   ["accounts receivable", "net accounts receivable"]],
     [["total current liabilities"]], RATIO_FMT,
     "(Cash + short-term investments + receivables) / current liabilities"),
    (r"debt[- ]to[- ](total )?assets|debt ratio", [["total liabilities"]], [["total assets"]],
     PERCENT_FMT, "Total liabilities / total assets"),
    (r"debt[- ]to[- ]equity", [["total liabilities"]], [["total equity"]], RATIO_FMT,
     "Total liabilities / total equity"),
    (r"gross (profit|margin) (rate|ratio|percentage|margin|%)|gross margin ratio|"
     r"gross profit margin", [["gross profit"]], [["net sales", "total revenues"]], PERCENT_FMT,
     "Gross profit / net sales"),
    (r"(net )?profit margin|return on sales|net margin", [["net income"]],
     [["net sales", "total revenues"]], PERCENT_FMT, "Net income / net sales"),
    (r"return on (total )?assets|\broa\b", [["net income"]], [["total assets"]], PERCENT_FMT,
     "Net income / total assets"),
    (r"return on (stockholders'?|owners'?|shareholders'?|common)? ?equity|\broe\b",
     [["net income"]], [["total equity"]], PERCENT_FMT, "Net income / total equity"),
    (r"asset turnover", [["net sales", "total revenues"]], [["total assets"]], RATIO_FMT,
     "Net sales / total assets"),
    (r"inventory turnover", [["cost of goods sold"]], [["inventory", "merchandise inventory"]],
     RATIO_FMT, "Cost of goods sold / inventory"),
    (r"(receivables?|accounts receivable) turnover",
     [["net sales", "total revenues"]], [["accounts receivable", "net accounts receivable"]],
     RATIO_FMT, "Net sales / accounts receivable"),
    (r"equity multiplier", [["total assets"]], [["total equity"]], RATIO_FMT,
     "Total assets / total equity"),
    (r"times interest earned", [["income before taxes"], ["interest expense"]],
     [["interest expense"]], RATIO_FMT, "(Income before taxes + interest) / interest expense"),
]
DAYS = [
    (r"days in inventory|days'? sales in inventory|days inventory", "inventory turnover"),
    (r"(average )?collection period|days'? sales outstanding|days in receivables|"
     r"days'? sales uncollected", "receivables turnover"),
]
WORKING_CAPITAL = r"^working capital$"


def is_ratio(label):
    t = norm(label)
    if len(t) > 60:
        return False
    return any(re.search(p, t) for p, *_ in RATIOS) or any(re.search(p, t) for p, _ in DAYS) \
        or bool(re.search(WORKING_CAPITAL, t))


def _book_total(ctx, name):
    """Totals a template may not show, rebuilt from the account balances."""
    from . import accounts as A
    book = ctx.book
    if book is None:
        return None
    assets = book.of_type(A.ASSET, A.CONTRA_ASSET)
    liabs = book.of_type(A.LIABILITY)
    picks = {
        "total current assets": [a for a in assets if A.is_current(a.name, a.acct_type)],
        "total assets": assets,
        "total current liabilities": [a for a in liabs if A.is_current(a.name, a.acct_type)],
        "total liabilities": liabs,
    }.get(name)
    if not picks:
        return None
    sign = {A.CONTRA_ASSET: "-"}
    parts = [("-" if a.acct_type in sign else "+") + a.normal_balance() for a in picks]
    return "(" + "".join(parts).lstrip("+") + ")"


def _concept(ctx, alternatives):
    for name in alternatives:
        got = ctx.lookup(name) or _book_total(ctx, name)
        if got:
            return got
        if ctx.book is not None:
            acct = ctx.book.get(name)
            if acct is None:
                # "inventory" matches "Merchandise inventory", "cash" matches "Cash"
                hits = [a for a in ctx.book.accounts.values() if name in a.name.lower()]
                acct = hits[0] if len(hits) == 1 else None
            if acct is not None:
                return acct.normal_balance()
    return None


def _ratio(ctx, pattern_label):
    t = norm(pattern_label)
    for pat, num, den, fmt, why in RATIOS:
        if re.search(pat, t):
            nums = [_concept(ctx, alt) for alt in num]
            nums = [n for n in nums if n]
            d = _concept(ctx, den[0])
            if not nums or not d:
                return None
            if pat.startswith("times interest") and len(nums) < 2:
                return None
            return f"({'+'.join(nums)})/{d}", fmt, why
    return None


def ratio_expr(ctx, label):
    t = norm(label)
    if re.search(WORKING_CAPITAL, t):
        a, l = ctx.lookup("total current assets"), ctx.lookup("total current liabilities")
        if a and l:
            return f"{a}-{l}", None, "Working capital = current assets - current liabilities"
        return None
    for pat, base in DAYS:
        if re.search(pat, t):
            got = _ratio(ctx, base)
            if got:
                return f"365/({got[0]})", RATIO_FMT, f"365 / {base}"
            return None
    return _ratio(ctx, label)
