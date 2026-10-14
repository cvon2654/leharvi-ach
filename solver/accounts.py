"""Chart-of-accounts knowledge: classify an account title into its type.

The classifier works from the words in the account name, the same way a
student would ("Accumulated depreciation" is a contra asset, anything
"payable" is a liability, ...).  A workbook can override the guesses with a
sheet named "Chart of Accounts" (see engine.load_chart_override).
"""

import re

ASSET = "asset"
CONTRA_ASSET = "contra asset"
LIABILITY = "liability"
EQUITY = "equity"
CONTRA_EQUITY = "contra equity"
DRAWING = "drawing"
REVENUE = "revenue"
CONTRA_REVENUE = "contra revenue"
COGS = "cost of goods sold"
EXPENSE = "expense"
CLOSING = "closing"
# Purchases returns/discounts reduce cost of goods sold (credit balance).
CONTRA_COGS = "contra cost of goods sold"
UNKNOWN = "unknown"

DEBIT_NORMAL = {ASSET, CONTRA_EQUITY, DRAWING, CONTRA_REVENUE, COGS, EXPENSE}
CREDIT_NORMAL = {CONTRA_ASSET, LIABILITY, EQUITY, REVENUE, CONTRA_COGS}
INCOME_STATEMENT_TYPES = {REVENUE, CONTRA_REVENUE, COGS, CONTRA_COGS, EXPENSE, CLOSING}

# Words that tell us a type when the user writes it in a chart-of-accounts sheet.
TYPE_WORDS = [
    ("contra asset", CONTRA_ASSET), ("contra-asset", CONTRA_ASSET),
    ("contra revenue", CONTRA_REVENUE), ("contra-revenue", CONTRA_REVENUE),
    ("contra equity", CONTRA_EQUITY), ("contra-equity", CONTRA_EQUITY),
    ("cost of goods", COGS), ("cogs", COGS),
    ("drawing", DRAWING), ("dividend", DRAWING), ("withdraw", DRAWING),
    ("asset", ASSET), ("liabilit", LIABILITY), ("equity", EQUITY), ("capital", EQUITY),
    ("revenue", REVENUE), ("income", REVENUE), ("expense", EXPENSE),
]


def _has(name, *words):
    return any(w in name for w in words)


def _clean(name):
    name = str(name).lower().replace("’", "'").replace("-", " ")
    return re.sub(r"\s+", " ", name).strip()


def classify(name):
    """Return the account type for an account title."""
    n = _clean(name)
    if not n:
        return UNKNOWN
    if _has(n, "income summary", "profit and loss summary", "revenue and expense summary"):
        return CLOSING
    if _has(n, "accumulated depreciation", "accumulated amortization", "accumulated depletion",
            "allowance for", "accum. dep", "accum dep", "acc. dep"):
        return CONTRA_ASSET
    if _has(n, "receivable"):
        return ASSET
    if _has(n, "purchase return", "purchase discount", "purchase allowance",
            "purchases return", "purchases discount", "purchases allowance"):
        return CONTRA_COGS
    if _has(n, "purchases", "freight in", "freight-in", "transportation in"):
        return COGS
    if _has(n, "sales return", "sales discount", "sales allowance", "returns and allowances",
            "sales returns"):
        return CONTRA_REVENUE
    if _has(n, "cost of goods sold", "cost of sales", "cost of merchandise sold", "cogs"):
        return COGS
    if _has(n, "payable", "unearned", "deferred revenue", "accrued liabilit", "mortgage",
            "customer deposit", "advances from customer", "warranty liabilit", "lease liabilit",
            "liability", "liabilities") or re.search(r"\bloans?\b(?! receivable)", n):
        return LIABILITY
    if _has(n, "treasury stock", "treasury shares"):
        return CONTRA_EQUITY
    if _has(n, "drawing", "withdrawal", "distributions") or \
            (_has(n, "dividend") and not _has(n, "revenue", "income", "payable")):
        return DRAWING
    if _has(n, "expense", "loss", "expenses"):
        return EXPENSE
    if _has(n, "revenue", "sales", "fees earned", "fee earned", "earned", "gain",
            "interest income", "rent income", "dividend income", "service income",
            "commission income", "income from", "other income", "investment income", "fees"):
        return REVENUE
    if re.search(r"(?<!pre)paid in", n) or _has(n, "capital", "common stock", "preferred stock",
                                                "retained earnings", "owner's equity", "owners equity", "share capital",
            "ordinary shares", "share premium", "stockholders' equity", "net assets"):
        return EQUITY
    if _has(n, "cash", "inventory", "merchandise", "supplies", "prepaid", "equipment", "land",
            "building", "furniture", "vehicle", "truck", "automobile", "machinery", "investment",
            "patent", "goodwill", "copyright", "trademark", "franchise", "petty cash", "computer",
            "fixtures", "intangible", "property", "plant", "notes receivable", "deposit",
            "bank", "office", "tools", "leasehold", "software", "short term investment",
            "marketable securities", "materials", "work in process", "finished goods"):
        return ASSET
    if _has(n, "salaries", "salary", "wages", "rent", "utilities", "depreciation", "insurance",
            "advertising", "repairs", "maintenance", "bad debt", "uncollectible", "telephone",
            "postage", "fuel", "gas", "travel", "freight out", "delivery", "commissions",
            "amortization", "tax", "taxes", "miscellaneous", "interest", "supplies used",
            "professional fees", "dues", "bank charge", "service charge", "cost"):
        return EXPENSE
    return UNKNOWN


def is_debit_normal(acct_type):
    return acct_type in DEBIT_NORMAL


def is_current(name, acct_type):
    """Rough current vs. long-term split used for classified balance sheets."""
    n = _clean(name)
    if acct_type in (ASSET, CONTRA_ASSET):
        if acct_type == CONTRA_ASSET and _has(n, "allowance"):
            return True
        return not _has(n, "equipment", "land", "building", "furniture", "vehicle", "truck",
                        "automobile", "machinery", "patent", "goodwill", "copyright",
                        "trademark", "franchise", "long term", "long-term", "property", "plant",
                        "fixtures", "computer", "intangible", "accumulated", "leasehold",
                        "software", "tools", "investment") or _has(n, "short term investment",
                                                                   "short-term investment")
    if acct_type == LIABILITY:
        return not _has(n, "bonds", "mortgage", "long term", "long-term", "lease liabilit",
                        "noncurrent", "non current", "non-current")
    return True


def type_from_words(text):
    """Map a user-written type ("Current Asset", "Expense", ...) to a type constant."""
    t = _clean(text)
    for word, acct_type in TYPE_WORDS:
        if word in t:
            return acct_type
    return None


_STOP = {"the", "a", "an", "of", "and", "&", "for", "less", "add", "plus", "deduct", "net"}


def account_key(name):
    """Loose matching key so "Salaries Expense" == "salary expense" == "Salaries exp."."""
    n = _clean(name)
    n = re.sub(r"^(less|add|plus|deduct)\s*:?\s*", "", n)
    n = n.replace("'s", "s").replace("'", "")
    n = re.sub(r"\bexp\b\.?", "expense", n)
    n = re.sub(r"\ba/r\b", "accounts receivable", n)
    n = re.sub(r"\ba/p\b", "accounts payable", n)
    n = re.sub(r"\baccum\b\.?", "accumulated", n)
    n = re.sub(r"\bdep\b\.?", "depreciation", n)
    n = re.sub(r"[^a-z0-9 ]", " ", n)
    words = []
    for w in n.split():
        if w in _STOP:
            continue
        if len(w) > 4 and w.endswith("ies"):
            w = w[:-3] + "y"
        elif len(w) > 3 and w.endswith("s") and not w.endswith("ss"):
            w = w[:-1]
        words.append(w)
    return " ".join(words)
