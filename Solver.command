#!/bin/bash
# Solver for Mac: double-click to open the Solver window.
cd "$(dirname "$0")"
exec python3 run_solver.py "$@"
