"""Sensitivities, two way grid and breakeven price, mirroring the VBA macros."""
from __future__ import annotations
from dataclasses import asdict, replace

import pfmodel as pf

# (Excel flex name, Python field, label)
DRIVERS = [
    ("CF", "capacity_factor", "Net capacity factor"),
    ("EPAPrice", "epa_price_2024", "EPA price"),
    ("CapexKW", "capex_per_kw", "Capex per kW"),
    ("FOM", "fom_per_kw", "Fixed O&M"),
    ("InterestRate", "interest", "Interest rate"),
    ("DSCRTarget", "dscr_target", "DSCR target"),
]
FLEX = 0.10
PRICES = [66.0, 70.0, 74.0, 78.0, 82.0]
CFS = [0.32, 0.34, 0.36, 0.38, 0.40]


def flexed(inp, field, f):
    return replace(inp, **{field: getattr(inp, field) * (1 + f)})


def tornado(inp: pf.Inputs):
    base = pf.solve(inp)["out"]
    rows = []
    for flex, field, label in DRIVERS:
        lo = pf.solve(flexed(inp, field, -FLEX))["out"]
        hi = pf.solve(flexed(inp, field, FLEX))["out"]
        rows.append(dict(flex=flex, field=field, label=label,
                         irr_lo=lo["equity_irr"], irr_base=base["equity_irr"], irr_hi=hi["equity_irr"],
                         dscr_lo=lo["min_dscr"], dscr_hi=hi["min_dscr"]))
    return rows


def two_way(inp: pf.Inputs):
    grid = [[pf.solve(replace(inp, epa_price_2024=p, capacity_factor=c))["out"]["equity_irr"] for c in CFS]
            for p in PRICES]
    return dict(prices=PRICES, cfs=CFS, irr=grid)


def breakeven_price(inp: pf.Inputs, lo=40.0, hi=140.0, tol=1e-6):
    f = lambda p: pf.solve(replace(inp, epa_price_2024=p))["out"]["equity_npv"]
    for _ in range(100):
        mid = (lo + hi) / 2
        if f(mid) > 0:
            hi = mid
        else:
            lo = mid
        if hi - lo < tol:
            break
    return (lo + hi) / 2


def tornado_chart(rows, path):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    rows = sorted(rows, key=lambda r: abs(r["irr_hi"] - r["irr_lo"]))
    base = rows[0]["irr_base"] * 100
    fig, ax = plt.subplots(figsize=(8, 4.2))
    for i, r in enumerate(rows):
        lo, hi = r["irr_lo"] * 100, r["irr_hi"] * 100
        ax.barh(i, lo - base, left=base, color="#C0504D")
        ax.barh(i, hi - base, left=base, color="#1F3864")
    ax.set_yticks(range(len(rows)), [r["label"] for r in rows])
    ax.axvline(base, color="black", lw=0.8)
    ax.set_xlabel("Equity IRR (%)")
    ax.set_title(f"Equity IRR sensitivity, inputs flexed ±10% (base {base:.1f}%)")
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color="#C0504D", label="-10%"), Patch(color="#1F3864", label="+10%")], loc="lower right", frameon=False)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=160)


def cash_chart(res, path):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    r = res["rows"]
    s = pf.CON_YEARS
    yrs = r["year"][s:]
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.bar(yrs, r["cfads"][s:], color="#9DB3D9", label="CFADS")
    ax.plot(yrs, r["ds"][s:], color="#1F3864", lw=2, label="Debt service (sculpted)")
    ax.set_ylabel("C$ millions")
    ax.set_title("CFADS and sculpted debt service, base case")
    ax.legend(frameon=False)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=160)


if __name__ == "__main__":
    from pathlib import Path
    out = Path(__file__).resolve().parents[1] / "outputs"
    out.mkdir(exist_ok=True)
    inp = pf.Inputs()
    rows = tornado(inp)
    for r in rows:
        print(f"{r['label']:22s} {r['irr_lo']:.2%}  {r['irr_base']:.2%}  {r['irr_hi']:.2%}   DSCR {r['dscr_lo']:.2f} {r['dscr_hi']:.2f}")
    g = two_way(inp)
    for p, row in zip(g["prices"], g["irr"]):
        print(p, " ".join(f"{v:.2%}" for v in row))
    print("breakeven", round(breakeven_price(inp), 2))
    tornado_chart(rows, out / "tornado.png")
    cash_chart(pf.solve(inp), out / "cfads_vs_debt_service.png")
