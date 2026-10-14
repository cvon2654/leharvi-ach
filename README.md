# Solver: fills in Excel accounting problems

Solver opens an Excel accounting problem, works out the empty answer boxes, and fills
them in **one box at a time**, explaining each one. It works **offline**. Nothing is sent
anywhere.

- Every answer is a real Excel **formula**, so you can click a box and see the work.
- Filled boxes turn **green**. A **Solver Notes** sheet lists every answer and any warnings.
- Your original file is never changed. Solver saves a copy named `… - Solved.xlsx`.
- Answer boxes do **not** need to be bold, bordered or marked. Solver reads the labels
  ("Total assets", "Net income", "Cash", …) and fills the empty cells next to them.
  Empty cells work, and so do cells that contain `?` or `____`.

## What it can solve

| Problem | What Solver does |
|---|---|
| **General journal** | Finds a missing debit or credit in an entry. Builds the **Trial Balance**, **General Ledger** (running balances), **Income Statement**, **Statement of Retained Earnings / Owner's Equity** and **Balance Sheet**. |
| **Trial balance** | Finds a missing account balance, fills the totals, and builds the statements. A blank trial balance next to a journal is filled from the journal. |
| **10-column worksheet** | Adjusted trial balance, then income statement and balance sheet columns, net income (loss), and final totals. |
| **Statements you were given to complete** | Fills account amounts, subtotals, totals, net sales, gross profit, income from operations, income before taxes, net income, beginning/ending retained earnings, and total liabilities and equity. Handles comparative (two-year) columns. |
| **Depreciation** | Straight-line, double-declining-balance, sum-of-the-years'-digits and units-of-activity schedules. Also depreciable cost, rates and annual depreciation. |
| **Loans / notes** | Payment (PMT) and amortization schedule (interest, principal, balance). |
| **Bonds** | Issue price, discount or premium, and the effective-interest amortization table. |
| **Inventory** | FIFO, LIFO and weighted-average cost of goods sold and ending inventory (periodic). |
| **Ratios** | Current, quick, debt-to-assets, debt-to-equity, gross profit rate, profit margin, ROA, ROE, asset / inventory / receivables turnover, days in inventory, collection period, working capital. |

Solver can't read a word problem and decide which journal entries to make. The numbers
must already be in the workbook, each with a label next to it.

## How to use it

### 1. Get it (once, while online)

**Easiest (Windows):** download `Solver-windows.zip` from this repo's **Actions** tab:
open the latest "Test and build Solver.exe" run and look under *Artifacts*. Unzip it
somewhere, for example `C:\Solver`. It contains `Solver.exe`, so you don't need to install
anything.

**Or with Python:** download this repo (green *Code* button → *Download ZIP*) and unzip it
to `C:\Solver`. Install Python from python.org and tick "Add python.exe to PATH". The
Excel library is already included in the `vendor` folder, so there's no `pip install`
and it runs offline.

### 2. Solve a file

- **Drag** your Excel file onto `Solver.bat` (or `Solver.exe`), or
- **Double-click** `Solver.bat`, then click *Browse…* and *Solve*.

The Solver window shows the answers **one box at a time**: click **Next box ›**, or
**Play** to watch them go. Then click *Open solved file in Excel*.

On a Mac, double-click `Solver.command` (it needs Python 3 from python.org).

### 3. Use it inside Excel (button in the ribbon, Windows)

This fills the boxes directly in the workbook you have open, one at a time. Each box is
selected and explained as it's filled.

1. In Excel press **Alt + F11** to open the VBA editor.
2. In the left panel, select **VBAProject (PERSONAL.XLSB)**. If you don't have it,
   record any macro once with *Store macro in: Personal Macro Workbook*, then stop.
3. **File → Import File…** and choose `excel\AccountingSolver.bas`.
4. At the top of the module, set `SOLVER_FOLDER` to where you unzipped Solver
   (default `C:\Solver`).
5. Close the editor. Go to **File → Options → Customize Ribbon** (or *Quick Access
   Toolbar*), choose *Macros*, and add **SolveStepByStep** (and **SolveAutoPlay** /
   **SolveAllAtOnce** if you like).

Now open a problem, **save it as .xlsx or .xlsm**, and click your button:

- **SolveStepByStep:** fills one box, shows what it is and its formula. Click **OK** for
  the next box, or **Cancel** to fill the rest at once.
- **SolveAutoPlay:** fills one box about every 0.7 seconds while you watch.
- **SolveAllAtOnce:** fills everything immediately.

> Excel also has a built-in "Solver" add-in for optimization (Data tab). That's a
> different tool. These macros are named `Solve…` so they don't clash with it.

### Command line

```
python run_solver.py problem.xlsx --cli            # solve, save "problem - Solved.xlsx"
python run_solver.py problem.xlsx --cli --steps    # walk through each box in the terminal
python run_solver.py problem.xlsx --cli --out answers.xlsx --no-highlight
```

## Tips for best results

- Put each label and its number on the **same row** (`Cost | 50,000`). Use column
  headers such as `Account | Debit | Credit`, or `Year | Depreciation expense |
  Accumulated depreciation | Book value`.
- Answer boxes can be left empty or contain `?`. No special formatting is needed.
- If Solver puts an account in the wrong place (for example an unusual account name),
  add a sheet named **Chart of Accounts** with two columns, `Account | Type` (Asset,
  Contra asset, Liability, Equity, Drawing/Dividends, Revenue, Contra revenue, Cost of
  goods sold, Expense). Solver will use your types.
- Always check the green boxes before you hand anything in. The **Check** line on a
  generated balance sheet should be 0.
- Old `.xls` files: in Excel use *File → Save As → Excel Workbook (.xlsx)* first.

## For developers

```
pip install vendor/*.whl pytest
python -m pytest tests          # needs LibreOffice to calculate the formulas
```

`build_exe.bat` builds `Solver.exe` locally. The GitHub workflow builds it on every push.

Code layout: `solver/journal.py` (journals, trial balances), `worksheet10.py`,
`statements.py` (new statement sheets), `totals.py` (filling existing templates),
`schedules.py` (depreciation, loans, bonds), `inventory.py`, `ratios.py`, `engine.py`
(runs everything, records the steps), `gui.py`, `cli.py`; `excel/AccountingSolver.bas`
is the Excel macro.
