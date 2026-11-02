"""Python mirror of the Excel project finance model.

Every line here maps to a row on the Model sheet of model/BC_Wind_Project_Finance.xlsx,
so the two can be reconciled to the cent (see tests/test_reconcile.py).
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict

N_PERIODS = 32          # 2 construction years + 30 operating years
CON_YEARS = 2


@dataclass
class Inputs:
    first_year: int = 2027            # financial close / start of construction
    capacity_mw: float = 200.0
    capacity_factor: float = 0.36
    availability: float = 0.97
    degradation: float = 0.003
    epa_price_2024: float = 74.0      # $/MWh, 2024 dollars
    cpi: float = 0.02
    indexation: float = 1.0           # share of CPI passed through in the EPA price
    capex_per_kw: float = 3000.0
    capex_phase_y1: float = 0.40
    fom_per_kw: float = 45.0          # fixed O&M, $/kW-yr at COD
    land_pct: float = 0.03            # land, community and First Nations benefit payments, % revenue
    ptax_ins: float = 3.0             # property tax and insurance, $M in first operating year
    opex_esc: float = 0.02
    itc_rate: float = 0.30            # federal Clean Technology ITC
    itc_eligible: float = 0.90        # share of capex eligible for the ITC
    tax_rate: float = 0.27
    cca_rate: float = 0.50            # CCA class 43.2, declining balance
    interest: float = 0.0525
    tenor: int = 18
    dscr_target: float = 1.35
    max_gearing: float = 0.75
    equity_rate: float = 0.09


SCENARIOS = {
    "Base":     {},
    "Upside":   dict(capacity_factor=0.385, capex_per_kw=2800.0, fom_per_kw=42.0, interest=0.0475),
    "Downside": dict(capacity_factor=0.33, capex_per_kw=3400.0, fom_per_kw=52.0, interest=0.06),
}


def scenario(name: str, **overrides) -> Inputs:
    d = asdict(Inputs())
    d.update(SCENARIOS[name])
    d.update(overrides)
    return Inputs(**d)


def npv(rate, flows):
    """Excel NPV: first flow discounted one full period."""
    return sum(cf / (1 + rate) ** (i + 1) for i, cf in enumerate(flows))


def irr(flows, lo=-0.99, hi=1.0, tol=1e-12):
    f = lambda r: sum(cf / (1 + r) ** i for i, cf in enumerate(flows))
    flo, fhi = f(lo), f(hi)
    if flo * fhi > 0:
        return float("nan")
    for _ in range(300):
        mid = (lo + hi) / 2
        fm = f(mid)
        if flo * fm <= 0:
            hi = mid
        else:
            lo, flo = mid, fm
        if hi - lo < tol:
            break
    return (lo + hi) / 2


def run(inp: Inputs, tax_paste=None) -> dict:
    """One pass of the model given a pasted tax row (the circularity breaker)."""
    n = N_PERIODS
    tax_paste = tax_paste or [0.0] * n
    r = {k: [0.0] * n for k in (
        "con_flag", "ops_flag", "ops_year", "debt_flag", "gen", "price", "revenue", "fom",
        "land", "ptax", "opex", "ebitda", "capex", "itc", "sizing_cfads", "sculpted",
        "ds", "debt_open", "interest", "principal", "debt_close", "ucc_open", "ucc_add",
        "cca", "ucc_close", "ti", "loss_open", "loss_close", "taxable", "tax", "cfads",
        "dscr", "debt_draw", "equity_in", "equity_cf", "project_cf")}
    r["year"] = [inp.first_year + t for t in range(n)]

    capex_total = inp.capacity_mw * 1000 * inp.capex_per_kw / 1e6
    itc_total = capex_total * inp.itc_eligible * inp.itc_rate
    price_cod = inp.epa_price_2024 * (1 + inp.cpi * inp.indexation) ** (inp.first_year + CON_YEARS - 2024)
    for t in range(n):
        p = t + 1
        con = 1 if p <= CON_YEARS else 0
        ops = 1 - con
        k = (p - CON_YEARS) * ops
        r["con_flag"][t], r["ops_flag"][t], r["ops_year"][t] = con, ops, k
        r["debt_flag"][t] = 1 if (ops and k <= inp.tenor) else 0
        gen = inp.capacity_mw * 8.76 * inp.capacity_factor * inp.availability * (1 - inp.degradation) ** (k - 1) * ops
        price = price_cod * (1 + inp.cpi * inp.indexation) ** (k - 1) * ops
        rev = gen * price / 1000
        fom = inp.capacity_mw * 1000 * inp.fom_per_kw * (1 + inp.opex_esc) ** (k - 1) / 1e6 * ops
        land = rev * inp.land_pct
        ptax = inp.ptax_ins * (1 + inp.opex_esc) ** (k - 1) * ops
        opex = fom + land + ptax
        r["gen"][t], r["price"][t], r["revenue"][t] = gen, price, rev
        r["fom"][t], r["land"][t], r["ptax"][t], r["opex"][t] = fom, land, ptax, opex
        r["ebitda"][t] = rev - opex
        r["capex"][t] = capex_total * (inp.capex_phase_y1 if p == 1 else (1 - inp.capex_phase_y1) if p == 2 else 0)
        r["itc"][t] = itc_total if k == 1 else 0.0
        r["sizing_cfads"][t] = (r["ebitda"][t] - tax_paste[t]) * ops
        r["sculpted"][t] = r["sizing_cfads"][t] / inp.dscr_target * r["debt_flag"][t]

    pv_sculpted = npv(inp.interest, r["sculpted"][CON_YEARS:])
    debt = min(pv_sculpted, inp.max_gearing * capex_total)
    scale = debt / pv_sculpted if pv_sculpted else 0.0
    gearing = debt / capex_total

    loss = 0.0
    bal = 0.0
    ucc = 0.0
    for t in range(n):
        ops = r["ops_flag"][t]
        r["debt_draw"][t] = r["capex"][t] * gearing
        r["equity_in"][t] = r["capex"][t] - r["debt_draw"][t]
        r["ds"][t] = r["sculpted"][t] * scale
        r["debt_open"][t] = bal
        r["interest"][t] = bal * inp.interest * ops
        r["principal"][t] = (r["ds"][t] - r["interest"][t]) * ops
        bal = bal + r["debt_draw"][t] - r["principal"][t]
        r["debt_close"][t] = bal
        # tax: ITC reduces the UCC pool in the year it is received
        r["ucc_open"][t] = ucc
        r["ucc_add"][t] = r["capex"][t] - r["itc"][t]
        r["cca"][t] = inp.cca_rate * (ucc - r["itc"][t]) * ops
        ucc = ucc + r["ucc_add"][t] - r["cca"][t]
        r["ucc_close"][t] = ucc
        ti = r["ebitda"][t] - r["cca"][t] - r["interest"][t]
        r["ti"][t] = ti
        r["loss_open"][t] = loss
        r["taxable"][t] = max(0.0, ti - loss)
        loss = max(0.0, loss - ti)
        r["loss_close"][t] = loss
        r["tax"][t] = r["taxable"][t] * inp.tax_rate
        r["cfads"][t] = (r["ebitda"][t] - r["tax"][t]) * ops
        r["dscr"][t] = r["cfads"][t] / r["ds"][t] if r["ds"][t] > 0 else None
        r["equity_cf"][t] = -r["equity_in"][t] + r["cfads"][t] - r["ds"][t] + r["itc"][t]
        r["project_cf"][t] = -r["capex"][t] + r["ebitda"][t] + r["itc"][t]

    dscrs = [d for d in r["dscr"] if d is not None]
    cfads_loan = [r["cfads"][t] * r["debt_flag"][t] for t in range(CON_YEARS, n)]
    out = dict(
        capex_total=capex_total, itc_total=itc_total, price_cod=price_cod,
        debt=debt, pv_sculpted=pv_sculpted, gearing=gearing,
        equity=capex_total - debt,
        equity_irr=irr(r["equity_cf"]), project_irr=irr(r["project_cf"]),
        equity_npv=npv(inp.equity_rate, r["equity_cf"]),
        min_dscr=min(dscrs), avg_dscr=sum(dscrs) / len(dscrs),
        llcr=npv(inp.interest, cfads_loan) / debt if debt else None,
        final_debt=r["debt_close"][-1],
        yr1_revenue=r["revenue"][CON_YEARS], yr1_ebitda=r["ebitda"][CON_YEARS],
    )
    return {"rows": r, "out": out, "tax_paste": list(tax_paste)}


def solve(inp: Inputs, tol=1e-9, max_iter=100) -> dict:
    """Copy-paste iteration on the tax row, exactly what the VBA SculptDebt macro does."""
    paste = [0.0] * N_PERIODS
    for i in range(max_iter):
        res = run(inp, paste)
        new = res["rows"]["tax"]
        diff = max(abs(a - b) for a, b in zip(new, paste))
        if diff < tol:
            res["iterations"] = i + 1
            return res
        paste = list(new)
    raise RuntimeError("tax circularity did not converge")


if __name__ == "__main__":
    for name in SCENARIOS:
        o = solve(scenario(name))["out"]
        print(f"{name:9s} debt {o['debt']:7.1f}  gearing {o['gearing']:.1%}  "
              f"eqIRR {o['equity_irr']:.2%}  projIRR {o['project_irr']:.2%}  "
              f"minDSCR {o['min_dscr']:.2f}x  avgDSCR {o['avg_dscr']:.2f}x  LLCR {o['llcr']:.2f}x  NPV {o['equity_npv']:.1f}")
