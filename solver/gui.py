"""The Solver window: pick a workbook, solve it, then step through each box."""

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__

PLAY_DELAY_MS = 900


def open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606 - opens the workbook in Excel
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class SolverApp:
    def __init__(self, root, path=None):
        self.root = root
        self.result = None
        self.index = -1
        self.playing = False
        root.title(f"Solver {__version__} - Excel accounting problem solver")
        root.geometry("900x620")
        root.minsize(640, 460)

        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Excel file:").pack(side="left")
        self.path_var = tk.StringVar(value=path or "")
        ttk.Entry(top, textvariable=self.path_var).pack(side="left", fill="x", expand=True,
                                                        padx=6)
        ttk.Button(top, text="Browse...", command=self.browse).pack(side="left")

        opts = ttk.Frame(root, padding=(10, 0))
        opts.pack(fill="x")
        self.highlight = tk.BooleanVar(value=True)
        self.open_after = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="Color the boxes Solver fills (green)",
                        variable=self.highlight).pack(side="left")
        ttk.Checkbutton(opts, text="Open the solved file in Excel when done",
                        variable=self.open_after).pack(side="left", padx=12)
        self.solve_btn = ttk.Button(opts, text="Solve", command=self.solve)
        self.solve_btn.pack(side="right")

        # step-by-step panel
        step = ttk.LabelFrame(root, text="Answers, one box at a time", padding=10)
        step.pack(fill="both", expand=True, padx=10, pady=10)
        self.counter = ttk.Label(step, text="Solve a workbook to see each answer here.",
                                 font=("Segoe UI", 11, "bold"))
        self.counter.pack(anchor="w")
        self.where = ttk.Label(step, text="", font=("Segoe UI", 14, "bold"), foreground="#1F6E43")
        self.where.pack(anchor="w", pady=(6, 0))
        self.why = ttk.Label(step, text="", wraplength=820, justify="left", font=("Segoe UI", 11))
        self.why.pack(anchor="w", pady=4)
        self.formula = tk.Text(step, height=3, wrap="word", font=("Consolas", 11),
                               relief="flat", background="#F3F6F4")
        self.formula.pack(fill="x")
        nav = ttk.Frame(step)
        nav.pack(fill="x", pady=8)
        ttk.Button(nav, text="< Back", command=lambda: self.go(self.index - 1)).pack(side="left")
        ttk.Button(nav, text="Next box >", command=lambda: self.go(self.index + 1)).pack(
            side="left", padx=6)
        self.play_btn = ttk.Button(nav, text="Play", command=self.toggle_play)
        self.play_btn.pack(side="left")
        ttk.Button(nav, text="Show all", command=lambda: self.go(len(self.steps()) - 1)).pack(
            side="left", padx=6)
        self.open_btn = ttk.Button(nav, text="Open solved file in Excel", command=self.open_out,
                                   state="disabled")
        self.open_btn.pack(side="right")

        cols = ("n", "where", "what")
        self.tree = ttk.Treeview(step, columns=cols, show="headings", height=8)
        for c, text, w in (("n", "#", 50), ("where", "Box", 200), ("what", "What it is", 560)):
            self.tree.heading(c, text=text)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        self.status = ttk.Label(root, text="Works offline. Your file is never changed - a "
                                           "'- Solved' copy is saved next to it.",
                                padding=(10, 0, 10, 8))
        self.status.pack(fill="x")
        root.bind("<Right>", lambda e: self.go(self.index + 1))
        root.bind("<Left>", lambda e: self.go(self.index - 1))
        if path:
            root.after(200, self.solve)

    # ------------------------------------------------------------------
    def browse(self):
        p = filedialog.askopenfilename(title="Choose the Excel problem",
                                       filetypes=[("Excel workbooks", "*.xlsx *.xlsm"),
                                                  ("All files", "*.*")])
        if p:
            self.path_var.set(p)

    def steps(self):
        return self.result.steps if self.result else []

    def solve(self):
        path = self.path_var.get().strip().strip('"')
        if not path:
            self.browse()
            path = self.path_var.get().strip()
            if not path:
                return
        if path.lower().endswith(".xls"):
            messagebox.showerror("Solver", "This is an old .xls file. In Excel use File > Save As "
                                           "> Excel Workbook (.xlsx), then try again.")
            return
        self.solve_btn.config(state="disabled")
        self.status.config(text="Solving...")

        def work():
            from .engine import solve
            try:
                res = solve(path, highlight=self.highlight.get())
                self.root.after(0, lambda: self.done(res))
            except Exception as e:     # show the problem instead of crashing
                msg = str(e)
                self.root.after(0, lambda: self.failed(msg))
        threading.Thread(target=work, daemon=True).start()

    def failed(self, msg):
        self.solve_btn.config(state="normal")
        self.status.config(text="")
        messagebox.showerror("Solver", f"Could not solve that file:\n\n{msg}\n\nIs it open in "
                                       f"Excel with unsaved changes, or not an .xlsx file?")

    def done(self, res):
        self.result = res
        self.solve_btn.config(state="normal")
        self.open_btn.config(state="normal")
        self.tree.delete(*self.tree.get_children())
        for i, s in enumerate(res.steps, 1):
            where = f"{s.sheet}!{s.cell}" if s.cell else f"{s.sheet} (new sheet)"
            self.tree.insert("", "end", iid=str(i - 1), values=(i, where, s.note))
        warn = [m for lvl, m in res.notes if lvl == "warning"]
        self.status.config(text=f"Saved: {res.output}" + (f"   ({len(warn)} note(s) - see the "
                                                           f"'Solver Notes' sheet)" if warn
                                                           else ""))
        self.index = -1
        if res.steps:
            self.go(0)
        else:
            self.counter.config(text="Nothing to fill was found.")
            self.why.config(text="\n".join(m for _, m in res.notes))
        if self.open_after.get():
            open_file(res.output)

    def go(self, i):
        steps = self.steps()
        if not steps:
            return
        i = max(0, min(i, len(steps) - 1))
        self.index = i
        s = steps[i]
        self.counter.config(text=f"Box {i + 1} of {len(steps)}")
        self.where.config(text=f"{s.sheet}  !  {s.cell}" if s.cell else f"New sheet: {s.sheet}")
        self.why.config(text=s.note)
        self.formula.delete("1.0", "end")
        if s.kind == "fill":
            self.formula.insert("1.0", str(s.content))
        self.tree.selection_set(str(i))
        self.tree.see(str(i))

    def on_select(self, _):
        sel = self.tree.selection()
        if sel and int(sel[0]) != self.index:
            self.go(int(sel[0]))

    def toggle_play(self):
        self.playing = not self.playing
        self.play_btn.config(text="Pause" if self.playing else "Play")
        if self.playing:
            self._tick()

    def _tick(self):
        if not self.playing:
            return
        if self.index >= len(self.steps()) - 1:
            self.toggle_play()
            return
        self.go(self.index + 1)
        self.root.after(PLAY_DELAY_MS, self._tick)

    def open_out(self):
        if self.result:
            open_file(self.result.output)


def run_gui(path=None):
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if sys.platform.startswith("win") else "clam")
    except tk.TclError:
        pass
    SolverApp(root, path)
    root.mainloop()
