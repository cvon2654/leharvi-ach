"""Depreciation, loan amortization and bond amortization.

Inputs are read from "label | value" pairs anywhere in the workbook
("Cost  50,000", "Useful life  5", "Market rate  8%").  Schedule tables
already in the workbook are filled; otherwise a new sheet is created."""

import re
from dataclasses import dataclass

from openpyxl.utils import get_column_letter

from .workbook import (BOLD, PERCENT_FMT, cell_ref, is_blank, is_label, is_number, norm,
                       set_widths, style_header, write_title)


@dataclass
class Input:
    label: str
    ws: object
    row: int
    col: int
    value: object

    def ref(self):
        return cell_ref(self.ws, self.row, self.col, absolute=True)

    def rate_ref(self):
        """Rates typed as 8 mean 8%."""
        if is_number(self.value) and self.value > 1:
            return f"({self.ref()}/100)"
        return self.ref()


class Inputs:
    def __init__(self, ctx):
        self.ctx = ctx
        self.items = []
        self.text = ""
        self._scan()

    def _scan(self):
        ctx = self.ctx
        words = []
        for ws in ctx.wb.worksheets:
            if ws.title in ctx.created:
                continue
            for r in range(1, min(ws.max_row, 400) + 1):
                if ctx.is_claimed(ws, r):
                    continue
                for c in range(1, min(ws.max_column, 15) + 1):
                    v = ctx.value(ws, r, c)
                    if not is_label(v):
                        continue
                    words.append(norm(v))
                    for cc in range(c + 1, min(c + 5, ws.max_column + 1)):
                        n = ctx.number(ws, r, cc)
                        raw = ctx.value(ws, r, cc)
                        if n is not None:
                            self.items.append(Input(norm(v), ws, r, cc, n))
                            break
                        if is_label(raw):
                            self.items.append(Input(norm(v), ws, r, cc, norm(raw)))
                            break
                        if not is_blank(raw):
                            break
        self.text = " | ".join(words)

    def find(self, pattern, exclude=None, numeric=True):
        for it in self.items:
            if numeric and not is_number(it.value):
                continue
            if re.search(pattern, it.label) and not (exclude and re.search(exclude, it.label)):
                return it
        return None

    def mentions(self, pattern):
        return bool(re.search(pattern, self.text))


# --------------------------------------------------------------------------
# Depreciation
# --------------------------------------------------------------------------

METHODS = [
    ("ddb", r"double[- ]declining|declining[- ]balance|\bddb\b|200%"),
    ("syd", r"sum[- ]of[- ](the[- ])?years|\bsyd\b"),
    ("units", r"units[- ]of[- ](production|activity|output)|activity method|"
              r"units[- ]of[- ]production|production method"),
    ("sl", r"straight[- ]line|\bsl\b"),
]
METHOD_NAMES = {"sl": "Straight-line", "ddb": "Double-declining-balance",
                "syd": "Sum-of-the-years'-digits", "units": "Units-of-activity"}


def _method_in(text):
    for key, pat in METHODS:
        if re.search(pat, text):
            return key
    return None


def depreciation_inputs(inp):
    cost = inp.find(r"^(the )?(asset |equipment |machine |machinery |truck |vehicle |building |"
                    r"original |historical |acquisition |initial |purchase )?(cost|price)"
                    r"( of (the )?\w+)?$|purchase price|acquisition cost|^cost\b|\bcost$",
                    exclude=r"goods|per unit|unit cost|sales|interest|depreciable|inventory|total|"
                            r"purchases|loan")
    life = inp.find(r"useful life|^life|estimated life|life in years|^years$|service life",
                    exclude=r"units|hours|miles|loan|bond")
    if not cost or not life:
        return None
    salvage = inp.find(r"salvage|residual|scrap|disposal value")
    units = inp.find(r"(total|estimated|lifetime).*(units|hours|miles|activity|output)|"
                     r"useful life.*(units|hours|miles)|life.*(units|hours|miles)")
    return {"cost": cost, "life": life, "salvage": salvage, "units": units}


def _dep_formulas(m, k, n, v, beg_bv):
    """Depreciation for year k of n under method m, as a formula expression."""
    C, L = v["cost"].ref(), v["life"].ref()
    S = v["salvage"].ref() if v["salvage"] else "0"
    if m == "sl":
        return f"({C}-{S})/{L}", f"1/{L}"
    if m == "syd":
        return f"({C}-{S})*({L}-{k}+1)/({L}*({L}+1)/2)", f"({L}-{k}+1)/({L}*({L}+1)/2)"
    if m == "ddb":
        if k == n:
            return f"MAX(0,{beg_bv}-{S})", f"2/{L}"
        return f"MAX(0,MIN({beg_bv}*2/{L},{beg_bv}-{S}))", f"2/{L}"
    return None, None


def _dep_columns(ctx, ws, r):
    roles = {}
    for c in range(1, min(ws.max_column, 20) + 1):
        v = ctx.value(ws, r, c)
        if not isinstance(v, str):
            continue
        t = norm(v)
        if re.search(r"accumulated", t):
            roles.setdefault("accum", c)
        elif re.search(r"(beginning|beg\.?|start).*(book|carrying|value)|(book|carrying) value.*"
                       r"(beginning|beg)", t):
            roles.setdefault("beg", c)
        elif re.search(r"book value|carrying|net value|\bnbv\b", t):
            roles.setdefault("end", c)
        elif re.search(r"rate", t):
            roles.setdefault("rate", c)
        elif re.search(r"depreciation", t):
            roles.setdefault("dep", c)
        elif re.search(r"units|hours|miles|activity|output|production", t):
            roles.setdefault("units", c)
        elif re.search(r"^(year|yr|period|end of year|date)", t):
            roles.setdefault("year", c)
        elif re.search(r"depreciable (cost|base|amount)", t):
            roles.setdefault("base", c)
    return roles


def fill_depreciation(ctx, inp):
    v = depreciation_inputs(inp)
    if not v:
        return
    life_n = int(round(v["life"].value)) if v["life"].value and v["life"].value < 100 else None
    global_m = _method_in(inp.text)
    tables = 0
    for ws in ctx.wb.worksheets:
        if ws.title in ctx.created:
            continue
        for r in range(1, min(ws.max_row, 400) + 1):
            if ctx.is_claimed(ws, r):
                continue
            roles = _dep_columns(ctx, ws, r)
            if "dep" not in roles or not ({"accum", "end"} & set(roles)):
                continue
            above = " ".join(norm(ctx.value(ws, rr, cc)) for rr in range(max(1, r - 3), r)
                             for cc in range(1, 12) if is_label(ctx.value(ws, rr, cc)))
            m = _method_in(above) or global_m or "sl"
            if m == "units" and not ("units" in roles and v["units"]):
                ctx.warn(f"{ws.title} row {r}: units-of-activity table needs total units and a "
                         f"units column - not filled.")
                continue
            _fill_dep_table(ctx, ws, r, roles, m, v, life_n)
            tables += 1
    targets = _dep_targets(ctx, inp, v, global_m or "sl")
    if tables or targets or not life_n:
        return
    if not (inp.mentions(r"depreciat|salvage|residual|useful life")):
        return
    ws = ctx.new_sheet("Depreciation")
    methods = [global_m] if global_m and global_m != "units" else ["sl", "ddb", "syd"]
    row = write_title(ws, ["Depreciation Schedules",
                           f"Cost {v['cost'].ref()}, salvage "
                           f"{v['salvage'].ref() if v['salvage'] else '0'}, life {v['life'].ref()}"])
    for m in methods:
        ws.cell(row=row, column=1, value=METHOD_NAMES[m]).font = BOLD
        row += 1
        hdr = row
        for c, text in enumerate(["Year", "Beginning book value", "Rate", "Depreciation expense",
                                  "Accumulated depreciation", "Ending book value"], 1):
            ws.cell(row=hdr, column=c, value=text)
        style_header(ws, hdr, range(1, 7))
        roles = {"year": 1, "beg": 2, "rate": 3, "dep": 4, "accum": 5, "end": 6}
        _fill_dep_table(ctx, ws, hdr, roles, m, v, life_n, created=True)
        row = hdr + life_n + 3
    set_widths(ws, [10, 20, 12, 20, 22, 18])
    ctx.info(f"Created sheet '{ws.title}' ({', '.join(METHOD_NAMES[m] for m in methods)}).")


def _fill_dep_table(ctx, ws, hdr, roles, m, v, life_n, created=False):
    put = (lambda r, c, f, why, fmt=None: ws.cell(row=r, column=c, value=f).__setattr__(
        "number_format", fmt or '#,##0.00_);(#,##0.00)')) if created else \
        (lambda r, c, f, why, fmt=None: ctx.fill(ws, r, c, f, why, fmt))
    C = v["cost"].ref()
    yc = roles.get("year")
    rows = []
    r = hdr + 1
    initial = None
    while r <= hdr + 60:
        yv = ctx.value(ws, r, yc) if yc else None
        if not created and yc and is_blank(yv):
            if rows or initial:
                break
            if life_n and all(is_blank(ctx.value(ws, rr, yc)) for rr in range(r, r + 3)):
                rows = list(range(r, r + life_n))
                break
            r += 1
            continue
        if created:
            rows = list(range(r, r + life_n))
            break
        if (is_number(yv) and yv == 0) or (isinstance(yv, str) and re.search(
                r"acqui|purchase|start|begin|issue|^0$", norm(yv))):
            initial = r
        else:
            t = norm(ctx.value(ws, r, yc)) if isinstance(ctx.value(ws, r, yc), str) else ""
            if t.startswith("total"):
                break
            rows.append(r)
        r += 1
    if not yc and life_n:
        rows = list(range(hdr + 1, hdr + 1 + life_n))
    if not rows:
        return
    ctx.claim(ws, range(hdr, rows[-1] + 2))
    n = len(rows)
    if initial and "end" in roles and ctx.is_empty_cell(ws, initial, roles["end"]):
        put(initial, roles["end"], f"={C}", "Book value at acquisition = cost")
    prev_end = f"{get_column_letter(roles['end'])}{initial}" if initial and "end" in roles else None
    prev_acc = None
    name = METHOD_NAMES[m]
    for k, r in enumerate(rows, 1):
        def empty(role):
            return role in roles and (created or ctx.is_empty_cell(ws, r, roles[role]))
        if yc and empty("year"):
            put(r, yc, k, "Year number", "0")
        beg = prev_end or C
        if empty("beg"):
            put(r, roles["beg"], f"={beg}", f"{name}: beginning book value")
        if "beg" in roles:
            beg = f"{get_column_letter(roles['beg'])}{r}"
        if m == "units":
            S = v["salvage"].ref() if v["salvage"] else "0"
            rate = f"({C}-{S})/{v['units'].ref()}"
            dep = f"{get_column_letter(roles['units'])}{r}*{rate}"
        else:
            dep, rate = _dep_formulas(m, k, n, v, beg)
        if empty("rate"):
            put(r, roles["rate"], f"={rate}", f"{name} rate", "0.0000" if m == "units" else
                PERCENT_FMT)
        if empty("dep"):
            put(r, roles["dep"], f"={dep}", f"{name} depreciation, year {k}")
        dep_ref = f"{get_column_letter(roles['dep'])}{r}"
        acc = f"{prev_acc}+{dep_ref}" if prev_acc else dep_ref
        if empty("accum"):
            put(r, roles["accum"], f"={acc}", "Accumulated depreciation = prior + this year")
        if "accum" in roles:
            acc = f"{get_column_letter(roles['accum'])}{r}"
            prev_acc = acc
        else:
            prev_acc = f"({acc})"
        if empty("end"):
            put(r, roles["end"], f"={C}-{acc}", "Book value = cost - accumulated depreciation")
        if "end" in roles:
            prev_end = f"{get_column_letter(roles['end'])}{r}"
        else:
            prev_end = f"({C}-{acc})"
    # a Total row right under the table
    tr = rows[-1] + 1
    if not created and yc and isinstance(ctx.value(ws, tr, yc), str) and \
            norm(ctx.value(ws, tr, yc)).startswith("total") and empty_cell(ctx, ws, tr,
                                                                           roles["dep"]):
        L = get_column_letter(roles["dep"])
        ctx.fill(ws, tr, roles["dep"], f"=SUM({L}{rows[0]}:{L}{rows[-1]})", "Total depreciation")


def empty_cell(ctx, ws, r, c):
    return ctx.is_empty_cell(ws, r, c)


def _blank_targets(ctx, inp, pattern, exclude=None):
    """Label cells matching pattern whose value cell to the right is blank."""
    out = []
    for ws in ctx.wb.worksheets:
        if ws.title in ctx.created:
            continue
        for r in range(1, min(ws.max_row, 400) + 1):
            if ctx.is_claimed(ws, r):
                continue
            for c in range(1, min(ws.max_column, 12) + 1):
                v = ctx.value(ws, r, c)
                if not is_label(v):
                    continue
                t = norm(v)
                if len(t) > 60 or not re.search(pattern, t) or (exclude and re.search(exclude,
                                                                                         t)):
                    break
                for cc in range(c + 1, min(c + 4, max(ws.max_column, c + 1) + 1)):
                    raw = ws.cell(row=r, column=cc).value
                    if is_blank(raw):
                        out.append((ws, r, cc, t))
                        break
                    if not isinstance(raw, str) or is_label(raw):
                        break
                break
    return out


def _dep_targets(ctx, inp, v, m):
    C = v["cost"].ref()
    S = v["salvage"].ref() if v["salvage"] else "0"
    L = v["life"].ref()
    n = 0
    for ws, r, c, t in _blank_targets(ctx, inp, r"depreciable (cost|base|amount)"):
        ctx.fill(ws, r, c, f"={C}-{S}", "Depreciable cost = cost - salvage")
        ctx.claim(ws, [r]); n += 1
    if v["units"]:
        for ws, r, c, t in _blank_targets(ctx, inp, r"(rate|depreciation) per (unit|hour|mile)"):
            ctx.fill(ws, r, c, f"=({C}-{S})/{v['units'].ref()}", "(Cost - salvage) / total units",
                     "0.0000")
            ctx.claim(ws, [r]); n += 1
    for ws, r, c, t in _blank_targets(ctx, inp, r"(annual|yearly|first year|year 1|1st year)"
                                                r".*depreciation|depreciation (expense|per year|"
                                                r"for (the )?(first )?year)",
                                      exclude=r"accumulated|per (unit|hour|mile)|rate"):
        mm = _method_in(t) or m
        dep, _ = _dep_formulas(mm if mm != "units" else "sl", 1, 2, v, C)
        ctx.fill(ws, r, c, f"={dep}", f"{METHOD_NAMES.get(mm, 'Straight-line')} depreciation, "
                                      f"first year")
        ctx.claim(ws, [r]); n += 1
    for ws, r, c, t in _blank_targets(ctx, inp, r"(straight[- ]line|depreciation) rate|"
                                                r"ddb rate|declining[- ]balance rate"):
        f = f"=2/{L}" if re.search(r"double|ddb|declin", t) else f"=1/{L}"
        ctx.fill(ws, r, c, f, "Depreciation rate", PERCENT_FMT)
        ctx.claim(ws, [r]); n += 1
    return n


# --------------------------------------------------------------------------
# Loans
# --------------------------------------------------------------------------

def _periods_per_year(inp, default):
    ppy = inp.find(r"(payments|periods|compound\w*) per year|times per year")
    if ppy:
        return ppy.ref(), int(ppy.value)
    if inp.mentions(r"monthly|per month|each month|months"):
        return "12", 12
    if inp.mentions(r"quarterly"):
        return "4", 4
    if inp.mentions(r"semi[- ]?annual|twice a year|every six months|semiannually"):
        return "2", 2
    if inp.mentions(r"\bannual(ly)?\b|per year|each year"):
        return "1", 1
    return str(default), default


def fill_loan(ctx, inp):
    if inp.mentions(r"\bbonds?\b") and inp.find(r"market|effective|yield"):
        return
    P = inp.find(r"loan amount|principal|amount borrowed|^loan$|amount of (the )?(loan|note)|"
                 r"mortgage( amount)?$|note amount|face (value|amount) of (the )?note|"
                 r"^borrowed$|^amount financed", exclude=r"payment|paid|reduction")
    rate = inp.find(r"interest rate|^rate$|\bapr\b|annual rate", exclude=r"market|stated|coupon")
    if not P or not rate:
        return
    term = inp.find(r"term|years|life of (the )?loan|number of (years|months|payments|periods)|"
                    r"months|periods|^n$", exclude=r"per year|rate")
    ppy_ref, ppy = _periods_per_year(inp, 1)
    if term:
        t = term.label
        if re.search(r"month|payments|periods|^n$", t):
            N, n_val = term.ref(), int(term.value)
        else:
            N, n_val = f"{term.ref()}*{ppy_ref}", int(round(term.value * ppy))
    else:
        N, n_val = None, None
    i = f"{rate.rate_ref()}/{ppy_ref}" if ppy != 1 else rate.rate_ref()
    pay_in = inp.find(r"(monthly |annual |quarterly |installment |loan )?payment( amount)?$")
    payment = pay_in.ref() if pay_in else (f"PMT({i},{N},-{P.ref()})" if N else None)
    if not payment:
        return
    for ws, r, c, t in _blank_targets(ctx, inp, r"^(monthly |annual |quarterly |periodic |"
                                                r"installment |loan )?(payment|installment)"
                                                r"( amount)?$"):
        ctx.fill(ws, r, c, f"=PMT({i},{N},-{P.ref()})", "Payment = PMT(rate per period, number "
                                                        "of payments, -principal)")
        ctx.claim(ws, [r])
        payment = cell_ref(ws, r, c, absolute=True)
    for ws, r, c, t in _blank_targets(ctx, inp, r"total interest"):
        if N:
            ctx.fill(ws, r, c, f"={payment}*{N}-{P.ref()}", "Total interest = payments - "
                                                             "principal")
            ctx.claim(ws, [r])
    found = False
    for ws in ctx.wb.worksheets:
        if ws.title in ctx.created:
            continue
        for r in range(1, min(ws.max_row, 400) + 1):
            if ctx.is_claimed(ws, r):
                continue
            roles = _loan_columns(ctx, ws, r)
            if {"int", "pay"} <= set(roles) and ({"end", "prin"} & set(roles)):
                _fill_loan_table(ctx, ws, r, roles, P.ref(), i, payment, n_val)
                found = True
    if not found and n_val and n_val <= 600:
        ws = ctx.new_sheet("Loan Amortization")
        row = write_title(ws, ["Loan Amortization Schedule"])
        for c, text in enumerate(["Period", "Beginning balance", "Payment", "Interest",
                                  "Principal", "Ending balance"], 1):
            ws.cell(row=row, column=c, value=text)
        style_header(ws, row, range(1, 7))
        roles = {"period": 1, "beg": 2, "pay": 3, "int": 4, "prin": 5, "end": 6}
        _fill_loan_table(ctx, ws, row, roles, P.ref(), i, payment, n_val, created=True)
        set_widths(ws, [10, 20, 16, 16, 16, 18])
        ctx.info(f"Created sheet '{ws.title}' ({n_val} payments).")


def _loan_columns(ctx, ws, r):
    roles = {}
    for c in range(1, min(ws.max_column, 20) + 1):
        v = ctx.value(ws, r, c)
        if not isinstance(v, str):
            continue
        t = norm(v)
        if re.search(r"carrying|premium|discount|cash paid", t):
            return {}
        if re.search(r"(beg|beginning|opening).*balance|balance.*(beg|beginning)", t):
            roles.setdefault("beg", c)
        elif re.search(r"interest", t):
            roles.setdefault("int", c)
        elif re.search(r"principal|reduction", t):
            roles.setdefault("prin", c)
        elif re.search(r"(ending|end|remaining|closing|new|loan|unpaid).*balance|^balance", t):
            roles.setdefault("end", c)
        elif re.search(r"^(period|month|year|payment (no|number|#)|no|#|date|pmt no)", t):
            roles.setdefault("period", c)
        elif re.search(r"payment|installment", t):
            roles.setdefault("pay", c)
    return roles


def _table_rows(ctx, ws, hdr, key_col, n_val, created):
    if created:
        return None, list(range(hdr + 1, hdr + 1 + n_val))
    initial, rows = None, []
    r = hdr + 1
    while r <= hdr + 800:
        v = ctx.value(ws, r, key_col) if key_col else None
        if key_col and is_blank(v):
            if rows:
                break
            if n_val and all(is_blank(ctx.value(ws, rr, key_col)) for rr in range(r, r + 3)):
                return initial, list(range(r, r + n_val))
            if r > hdr + 3:
                break
            r += 1
            continue
        if not key_col:
            return None, list(range(hdr + 1, hdr + 1 + (n_val or 0)))
        t = norm(v) if isinstance(v, str) else ""
        if t.startswith("total"):
            break
        if (is_number(v) and v == 0) or re.search(r"issue|start|begin|^0$|acqui|loan date", t):
            initial = r
        else:
            rows.append(r)
        r += 1
    return initial, rows


def _fill_loan_table(ctx, ws, hdr, roles, P, i, payment, n_val, created=False):
    def put(r, c, f, why):
        if created:
            ws.cell(row=r, column=c, value=f).number_format = '#,##0.00_);(#,##0.00)'
        elif ctx.is_empty_cell(ws, r, c):
            ctx.fill(ws, r, c, f, why)
    initial, rows = _table_rows(ctx, ws, hdr, roles.get("period"), n_val, created)
    if not rows:
        return
    ctx.claim(ws, range(hdr, rows[-1] + 2))
    L = {k: get_column_letter(c) for k, c in roles.items()}
    if initial and "end" in roles:
        put(initial, roles["end"], f"={P}", "Loan principal")
    prev = f"{L['end']}{initial}" if initial and "end" in roles else P
    for k, r in enumerate(rows, 1):
        if "period" in roles and (created or ctx.is_empty_cell(ws, r, roles["period"])):
            if created:
                ws.cell(row=r, column=roles["period"], value=k)
            else:
                ctx.fill(ws, r, roles["period"], k, "Period number", "0")
        beg = prev
        if "beg" in roles:
            put(r, roles["beg"], f"={prev}", "Beginning balance = prior ending balance")
            beg = f"{L['beg']}{r}"
        put(r, roles["pay"], f"={payment}", "Payment")
        put(r, roles["int"], f"={beg}*{i}", "Interest = beginning balance x rate per period")
        prin = f"{L['pay']}{r}-{L['int']}{r}"
        if "prin" in roles:
            put(r, roles["prin"], f"={prin}", "Principal = payment - interest")
            prin = f"{L['prin']}{r}"
        if "end" in roles:
            put(r, roles["end"], f"={beg}-{prin}", "Ending balance = beginning - principal")
            prev = f"{L['end']}{r}"
        else:
            prev = f"({beg}-{prin})"


# --------------------------------------------------------------------------
# Bonds
# --------------------------------------------------------------------------

def fill_bonds(ctx, inp):
    face = inp.find(r"face (value|amount)|par value|maturity value|bonds? (face|principal)|"
                    r"principal amount|^face$|^par$|bonds? payable|bonds? issued|"
                    r"amount of bonds")
    stated = inp.find(r"stated|coupon|contract(ual)?( interest)? rate|nominal|face rate|"
                      r"bond (interest )?rate")
    market = inp.find(r"market|effective|yield")
    if not (face and stated and market):
        return
    term = inp.find(r"term|years|maturity|life|periods|^n$", exclude=r"rate|per year")
    ppy_ref, ppy = _periods_per_year(inp, 2)
    if not inp.find(r"(payments|periods|compound\w*) per year|times per year"):
        if inp.mentions(r"semi[- ]?annual|twice a year|every six months"):
            ppy_ref, ppy = "2", 2
        elif inp.mentions(r"\bannual(ly)?\b"):
            ppy_ref, ppy = "1", 1
        else:
            ppy_ref, ppy = "2", 2
    F = face.ref()
    cash = f"{F}*{stated.rate_ref()}/{ppy_ref}"
    mi = f"{market.rate_ref()}/{ppy_ref}"
    if term:
        if re.search(r"periods|^n$", term.label):
            N, n_val = term.ref(), int(term.value)
        else:
            N, n_val = f"{term.ref()}*{ppy_ref}", int(round(term.value * ppy))
    else:
        return
    price = f"PV({mi},{N},-{cash},-{F})"
    targets = [
        (r"issue price|selling price|bond price|price of (the )?bonds|proceeds|"
         r"issuance price|present value of (the )?bonds|^price$|cash received|sales price",
         f"={price}", "Issue price = PV of interest payments + PV of face value (market rate)"),
        (r"present value of (the )?(principal|face|maturity)", f"={F}/(1+{mi})^{N}",
         "PV of face value at the market rate"),
        (r"present value of (the )?interest", f"=PV({mi},{N},-{cash})",
         "PV of the interest annuity at the market rate"),
        (r"(cash )?interest (payment|paid)|periodic interest payment|semiannual interest",
         f"={cash}", "Cash interest = face x stated rate / periods per year"),
        (r"(bond )?discount( on bonds)?$|amount of (the )?discount",
         f"=MAX(0,{F}-{price})", "Discount = face - issue price"),
        (r"(bond )?premium( on bonds)?$|amount of (the )?premium",
         f"=MAX(0,{price}-{F})", "Premium = issue price - face"),
    ]
    for pat, formula, why in targets:
        for ws, r, c, t in _blank_targets(ctx, inp, pat):
            ctx.fill(ws, r, c, formula, why)
            ctx.claim(ws, [r])
    found = False
    for ws in ctx.wb.worksheets:
        if ws.title in ctx.created:
            continue
        for r in range(1, min(ws.max_row, 400) + 1):
            if ctx.is_claimed(ws, r):
                continue
            roles = _bond_columns(ctx, ws, r)
            if "intexp" in roles and ({"cv", "amort"} & set(roles)):
                _fill_bond_table(ctx, ws, r, roles, F, cash, mi, price, n_val)
                found = True
    if not found and n_val <= 200:
        ws = ctx.new_sheet("Bond Amortization")
        row = write_title(ws, ["Bond Amortization Schedule (effective-interest method)"])
        for c, text in enumerate(["Period", "Cash paid", "Interest expense", "Amortization",
                                  "Unamortized discount/premium", "Carrying value"], 1):
            ws.cell(row=row, column=c, value=text)
        style_header(ws, row, range(1, 7))
        ws.cell(row=row + 1, column=1, value="Issue")
        roles = {"period": 1, "cash": 2, "intexp": 3, "amort": 4, "unamort": 5, "cv": 6}
        _fill_bond_table(ctx, ws, row, roles, F, cash, mi, price, n_val, created=True)
        set_widths(ws, [10, 16, 18, 16, 26, 18])
        ctx.info(f"Created sheet '{ws.title}' ({n_val} periods).")


def _bond_columns(ctx, ws, r):
    roles = {}
    for c in range(1, min(ws.max_column, 20) + 1):
        v = ctx.value(ws, r, c)
        if not isinstance(v, str):
            continue
        t = norm(v)
        if re.search(r"interest expense|effective interest|interest at market", t):
            roles.setdefault("intexp", c)
        elif re.search(r"unamortized|(discount|premium) balance|balance of", t):
            roles.setdefault("unamort", c)
        elif re.search(r"amortiz", t):
            roles.setdefault("amort", c)
        elif re.search(r"carrying|book value|bond value|net (bond|liability)", t):
            roles.setdefault("cv", c)
        elif re.search(r"cash|interest (paid|payment)|stated interest|interest at stated", t):
            roles.setdefault("cash", c)
        elif re.search(r"^(period|date|year|semiannual|payment|interest period|no|#)", t):
            roles.setdefault("period", c)
    return roles


def _fill_bond_table(ctx, ws, hdr, roles, F, cash, mi, price, n_val, created=False):
    def put(r, c, f, why):
        if created:
            ws.cell(row=r, column=c, value=f).number_format = '#,##0.00_);(#,##0.00)'
        elif ctx.is_empty_cell(ws, r, c):
            ctx.fill(ws, r, c, f, why)
    if created:
        initial, rows = hdr + 1, list(range(hdr + 2, hdr + 2 + n_val))
    else:
        initial, rows = _table_rows(ctx, ws, hdr, roles.get("period"), n_val, False)
    if not rows:
        return
    ctx.claim(ws, range(hdr, rows[-1] + 2))
    L = {k: get_column_letter(c) for k, c in roles.items()}
    cv_prev = None
    if initial:
        if "cv" in roles:
            put(initial, roles["cv"], f"={price}", "Carrying value at issue = issue price")
            cv_prev = f"{L['cv']}{initial}"
        if "unamort" in roles:
            put(initial, roles["unamort"], f"=ABS({F}-{price})",
                "Discount/premium at issue = |face - price|")
    cv_prev = cv_prev or price
    for k, r in enumerate(rows, 1):
        if created:
            ws.cell(row=r, column=roles["period"], value=k)
        elif "period" in roles and ctx.is_empty_cell(ws, r, roles["period"]):
            ctx.fill(ws, r, roles["period"], k, "Period number", "0")
        c_ref = f"({cash})"
        if "cash" in roles:
            put(r, roles["cash"], f"={cash}", "Cash paid = face x stated rate per period")
            c_ref = f"{L['cash']}{r}"
        last = k == len(rows)
        if last:   # final period: amortize whatever is left so carrying value = face
            ie = f"{c_ref}+({F}-{cv_prev})"
            why = "Interest expense (final period: plugged so carrying value ends at face)"
        else:
            ie = f"{cv_prev}*{mi}"
            why = "Interest expense = carrying value x market rate per period"
        put(r, roles["intexp"], f"={ie}", why)
        ie_ref = f"{L['intexp']}{r}"
        if "amort" in roles:
            put(r, roles["amort"], f"=ABS({ie_ref}-{c_ref})",
                "Amortization = interest expense - cash paid")
        cv = f"{cv_prev}+{ie_ref}-{c_ref}"
        if "cv" in roles:
            put(r, roles["cv"], f"={cv}", "Carrying value = prior + amortization")
            cv_prev = f"{L['cv']}{r}"
        else:
            cv_prev = f"({cv})"
        if "unamort" in roles:
            put(r, roles["unamort"], f"=ABS({F}-{cv_prev})", "Unamortized = |face - carrying "
                                                             "value|")
