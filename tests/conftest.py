import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from solver.engine import solve  # noqa: E402

from . import samples  # noqa: E402

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def recalc(path, outdir):
    """Let LibreOffice calculate the formulas, then read the values back."""
    outdir = Path(outdir) / "calc"
    outdir.mkdir(exist_ok=True)
    subprocess.run([SOFFICE, "--headless", "--calc", "--convert-to", "xlsx", "--outdir",
                    str(outdir), str(path)], check=True, capture_output=True, timeout=180)
    return load_workbook(outdir / Path(path).name, data_only=True)


@pytest.fixture
def solved(tmp_path):
    if not SOFFICE:
        pytest.skip("LibreOffice is needed to calculate the formulas")

    def run(name):
        src = tmp_path / f"{name}.xlsx"
        samples.ALL[name](src)
        result = solve(src)
        return result, recalc(result.output, tmp_path)
    return run
