"""Recalculates the workbook in LibreOffice and checks it matches the Python mirror.

Skipped when LibreOffice is not installed.
"""
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import pfmodel as pf  # noqa: E402

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
pytestmark = pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")

ROWS = {  # Model sheet label -> Python row key
    "Revenue": "revenue", "EBITDA": "ebitda", "Cash tax": "tax", "CFADS": "cfads",
    "Debt service": "ds", "Closing balance": "debt_close", "CCA claimed": "cca",
    "Equity cash flow (incl. ITC)": "equity_cf",
}


@pytest.fixture(scope="module")
def wb(tmp_path_factory):
    out = tmp_path_factory.mktemp("recalc")
    subprocess.run([SOFFICE, "--headless", "--convert-to", "xlsx", "--outdir", str(out),
                    str(ROOT / "model" / "BC_Wind_Project_Finance.xlsx")], check=True, capture_output=True, timeout=180)
    return load_workbook(out / "BC_Wind_Project_Finance.xlsx", data_only=True)


def test_rows_match_python(wb):
    ws = wb["Model"]
    res = pf.solve(pf.Inputs())["rows"]
    labels = {ws.cell(r, 2).value: r for r in range(1, ws.max_row + 1)}
    for label, key in ROWS.items():
        xl = [ws.cell(labels[label], 5 + i).value or 0 for i in range(pf.N_PERIODS)]
        gap = max(abs(a - b) for a, b in zip(xl, res[key]))
        assert gap < 1e-6, f"{label} off by {gap}"


def test_headline_outputs_match(wb):
    ws = wb["Outputs"]
    vals = {ws.cell(r, 2).value: ws.cell(r, 3).value for r in range(1, 40)}
    o = pf.solve(pf.Inputs())["out"]
    assert abs(vals["Equity IRR (nominal, after tax)"] - o["equity_irr"]) < 1e-8
    assert abs(vals["Senior debt ($M)"] - o["debt"]) < 1e-6
    assert abs(vals["Equity NPV at hurdle rate ($M)"] - o["equity_npv"]) < 1e-6


def test_all_checks_pass(wb):
    ws = wb["Outputs"]
    assert "OK" in [ws.cell(r, 6).value for r in range(1, 40)]
