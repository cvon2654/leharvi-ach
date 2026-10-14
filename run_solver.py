"""Start Solver.  Works offline: the libraries it needs are in the vendor folder.

    python run_solver.py                      -> opens the Solver window
    python run_solver.py problem.xlsx         -> opens the window and solves the file
    python run_solver.py problem.xlsx --cli   -> solves in the terminal
    python run_solver.py problem.xlsx --cli --steps   -> walk through one box at a time
"""

import glob
import os
import sys

HERE = os.path.dirname(os.path.abspath(sys.argv[0] if getattr(sys, "frozen", False)
                                       else __file__))
sys.path.insert(0, HERE)

try:
    import openpyxl  # noqa: F401
except ImportError:      # use the bundled copies (pure Python wheels import directly)
    for wheel in sorted(glob.glob(os.path.join(HERE, "vendor", "*.whl"))):
        sys.path.insert(0, wheel)

if sys.stdout is None:   # windowed .exe has no console
    sys.stdout = sys.stderr = open(os.devnull, "w")

from solver.cli import main  # noqa: E402

if __name__ == "__main__":
    args = sys.argv[1:]
    cli_mode = any(a in ("--cli", "--steps", "--steps-file", "--out", "-h", "--help")
                   or a.startswith("--steps-file") or a.startswith("--out") for a in args)
    if cli_mode:
        raise SystemExit(main([a for a in args if a != "--cli"]))
    raise SystemExit(main(["--gui"] + args[:1]))
