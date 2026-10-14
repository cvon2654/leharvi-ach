"""Command line:  python -m solver problem.xlsx [--steps] [--out solved.xlsx]

With no file, the Solver window opens."""

import argparse
import sys


def write_steps_file(result, path):
    """Tab-separated steps for the Excel macro: kind, sheet, cell, note."""
    with open(path, "w", encoding="utf-8", newline="\r\n") as f:
        for s in result.steps:
            note = str(s.note).replace("\t", " ").replace("\n", " ")
            f.write(f"{s.kind}\t{s.sheet}\t{s.cell}\t{note}\n")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="solver", description="Fill in an Excel accounting "
                                                            "problem.")
    ap.add_argument("file", nargs="?", help="the .xlsx / .xlsm workbook to solve")
    ap.add_argument("--out", help="where to save the solved workbook")
    ap.add_argument("--no-highlight", action="store_true", help="don't color filled cells")
    ap.add_argument("--steps", action="store_true",
                    help="walk through the answers one cell at a time in the terminal")
    ap.add_argument("--steps-file", help="also write the step list here (used by the Excel "
                                         "macro)")
    ap.add_argument("--gui", action="store_true", help="open the Solver window")
    args = ap.parse_args(argv)

    if args.gui or not args.file:
        from .gui import run_gui
        run_gui(args.file)
        return 0

    from .engine import solve
    try:
        result = solve(args.file, args.out, highlight=not args.no_highlight)
    except Exception as e:      # show a readable message, not a traceback
        print(f"Solver could not read that workbook: {e}", file=sys.stderr)
        if args.steps_file:
            with open(args.steps_file, "w", encoding="utf-8", newline="\r\n") as f:
                f.write(f"error\t\t\t{e}\n")
        return 1
    if args.steps_file:
        write_steps_file(result, args.steps_file)
    if args.steps:
        total = len(result.steps)
        for i, s in enumerate(result.steps, 1):
            where = f"{s.sheet}!{s.cell}" if s.cell else s.sheet
            print(f"[{i}/{total}] {where}: {s.note}")
            if s.kind == "fill":
                print(f"         {s.content}")
            try:
                input("         press Enter for the next box...")
            except EOFError:
                print()
    print(result.summary())
    return 0
