import math
import sys
from dataclasses import replace
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import pfmodel as pf  # noqa: E402
import sensitivity as sens  # noqa: E402


@pytest.fixture(scope="module", params=list(pf.SCENARIOS))
def res(request):
    return pf.solve(pf.scenario(request.param))


def test_debt_fully_repaid(res):
    assert abs(res["out"]["final_debt"]) < 1e-6


def test_sources_equal_uses(res):
    r = res["rows"]
    assert math.isclose(sum(r["debt_draw"]) + sum(r["equity_in"]), res["out"]["capex_total"])


def test_tax_circularity_converged(res):
    assert max(abs(a - b) for a, b in zip(res["rows"]["tax"], res["tax_paste"])) < 1e-6


def test_dscr_never_below_target_when_dscr_binds(res):
    inp_target = pf.Inputs().dscr_target
    assert res["out"]["min_dscr"] >= inp_target - 1e-6


def test_gearing_cap_respected(res):
    assert res["out"]["gearing"] <= pf.Inputs().max_gearing + 1e-9


def test_no_debt_service_after_tenor(res):
    r = res["rows"]
    assert all(ds == 0 for ds, k in zip(r["ds"], r["ops_year"]) if k > pf.Inputs().tenor)


def test_loss_carryforward_never_negative(res):
    assert min(res["rows"]["loss_close"]) >= 0


def test_ucc_never_negative(res):
    assert min(res["rows"]["ucc_close"]) >= -1e-9


def test_irr_sign_and_npv_consistent(res):
    o = res["out"]
    assert (o["equity_irr"] > pf.Inputs().equity_rate) == (o["equity_npv"] > 0)


def test_higher_price_raises_irr():
    base = pf.solve(pf.Inputs())["out"]["equity_irr"]
    up = pf.solve(replace(pf.Inputs(), epa_price_2024=80))["out"]["equity_irr"]
    assert up > base


def test_gearing_cap_binds_when_set_low():
    o = pf.solve(replace(pf.Inputs(), max_gearing=0.40))["out"]
    assert math.isclose(o["gearing"], 0.40)
    assert o["min_dscr"] > pf.Inputs().dscr_target


def test_breakeven_price_gives_zero_npv():
    inp = pf.Inputs()
    p = sens.breakeven_price(inp)
    assert abs(pf.solve(replace(inp, epa_price_2024=p))["out"]["equity_npv"]) < 1e-3


def test_irr_matches_npv_definition():
    flows = [-100, 10, 10, 110]
    assert abs(pf.irr(flows) - 0.10) < 1e-9
