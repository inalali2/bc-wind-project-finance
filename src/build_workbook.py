"""Builds model/BC_Wind_Project_Finance.xlsx with live formulas.

The workbook is the deliverable; this script just makes it reproducible.
The tax 'paste' row and the sensitivity tables are macro outputs, seeded here from the
Python mirror so the file opens fully converged.
"""
from __future__ import annotations
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

sys.path.insert(0, str(Path(__file__).parent))
import pfmodel as pf
import sensitivity as sens

OUT = Path(__file__).resolve().parents[1] / "model" / "BC_Wind_Project_Finance.xlsx"

FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF", size=10)
BLACK = Font(name=FONT, size=10)
GREEN = Font(name=FONT, color="008000", size=10)
BOLD = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SEC_FILL = PatternFill("solid", fgColor="D9E1F2")
KEY_FILL = PatternFill("solid", fgColor="FFFF00")
MACRO_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="808080")
TOP = Border(top=THIN)

M = '#,##0.0;(#,##0.0);"-"'
M2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
X = '0.00"x"'
INT = '0'

wb = Workbook()


def name(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def setw(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# --------------------------------------------------------------------------- Cover
cv = wb.active
cv.title = "Cover"
setw(cv, {"A": 2, "B": 30, "C": 90})
cv["B2"] = "Coastal Ridge Wind: Project Finance and Valuation Model"
cv["B2"].font = TITLE
rows = [
    ("Asset", "Hypothetical 200 MW onshore wind farm in northeast BC with a 30 year BC Hydro Electricity Purchase Agreement (EPA)"),
    ("Purpose", "Sell side style valuation and debt sizing for an infrastructure advisory mandate. Portfolio project, not investment advice."),
    ("Author", "Inal Ali | github.com/inalali2"),
    ("Units", "C$ millions unless stated. Annual periods, end of year cash flows. Valuation date is financial close (start of 2027)."),
    ("", ""),
    ("Tabs", ""),
    ("Inputs", "All assumptions. Pick the scenario in the yellow cell. Flex column is used by the sensitivity macro."),
    ("Model", "Annual timeline: revenue, opex, capex, sculpted debt, CCA tax, CFADS, equity cash flows."),
    ("Outputs", "Returns, coverage ratios, sources and uses, integrity checks and charts."),
    ("Sensitivity", "Tornado and two way grid written by the VBA macros, plus breakeven EPA price."),
    ("", ""),
    ("How to run", "1) Save as .xlsm. 2) Alt+F11, File > Import File, choose vba/ModelMacros.bas. "
                   "3) After changing any input, run SculptDebt (Alt+F8) so the tax circularity converges. "
                   "4) Run RunSensitivities, RunTwoWay and BreakevenPrice to refresh the Sensitivity tab."),
    ("Why a macro", "Debt is sculpted to after tax CFADS, but tax depends on interest, which depends on debt size. "
                    "Rather than turn on iterative calculation, the model pastes tax as values and loops until it converges, "
                    "the usual approach in lender models. Outputs!Checks flags when the paste is stale."),
    ("", ""),
    ("Colour code", "Blue text = hard coded input. Black = formula. Green = link from another sheet. "
                    "Yellow fill = key selector. Peach fill = values written by a macro."),
    ("Simplifications", "No interest during construction, DSRA, fees or decommissioning; ITC received in the first operating year; "
                        "100% project level equity (no partner waterfall). See README for extensions."),
]
for i, (a, b) in enumerate(rows, start=4):
    cv.cell(i, 2, a).font = BOLD
    c = cv.cell(i, 3, b)
    c.font = BLACK
    c.alignment = Alignment(wrap_text=True, vertical="top")

# --------------------------------------------------------------------------- Inputs
ws = wb.create_sheet("Inputs")
setw(ws, {"A": 2, "B": 42, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 10, "I": 90})
ws["B2"] = "Inputs"
ws["B2"].font = TITLE
ws["B4"] = "Scenario (1 Base, 2 Upside, 3 Downside)"
ws["B4"].font = BOLD
ws["C4"] = 1
ws["C4"].font = BLUE
ws["C4"].fill = KEY_FILL
ws["D4"] = '=CHOOSE(Scenario,"Base","Upside","Downside")'
ws["D4"].font = BOLD
name("Scenario", "Inputs!$C$4")

hdr = ["Input", "Live", "Unit", "Base", "Upside", "Downside", "Flex", "Source / rationale"]
for j, h in enumerate(hdr, start=2):
    c = ws.cell(6, j, h)
    c.font, c.fill = HDR, HDR_FILL

base, up, dn = pf.Inputs(), pf.SCENARIOS["Upside"], pf.SCENARIOS["Downside"]
# (excel name, python field, label, unit, fmt, flex?, note)
SPECS = [
    ("sec", "Timing"),
    ("FirstYear", "first_year", "Financial close year", "year", INT, False, "Assumed close in 2027 with a 2 year build, so COD is 2029."),
    ("ConYears", None, "Construction period", "years", INT, False, "Fixed at 2 in this version (capex phasing below assumes 2 years)."),
    ("sec", "Generation"),
    ("Capacity", "capacity_mw", "Installed capacity", "MW", "#,##0", False, "Close to the average project size in BC Hydro's 2024 Call for Power (nine wind projects, about 1,530 MW in total)."),
    ("CF", "capacity_factor", "Net capacity factor (P50)", "%", PCT, True, "Call for Power awards total about 4,830 GWh/yr from roughly 1.5 GW, which implies about 36%. Scenarios cover a P90 style downside."),
    ("Availability", "availability", "Availability", "%", PCT, True, "Typical turbine availability guarantee is 97%."),
    ("Degradation", "degradation", "Annual output degradation", "%", PCT2, True, "Assumption: 0.3% per year."),
    ("sec", "Revenue"),
    ("EPAPrice", "epa_price_2024", "EPA price (2024 dollars)", "$/MWh", "#,##0.00", True, "BC Hydro: weighted average levelized price of $74/MWh (2024$) for the 2024 Call for Power awards."),
    ("CPI", "cpi", "CPI assumption", "%", PCT, True, "Bank of Canada 2% inflation target."),
    ("Indexation", "indexation", "Share of CPI passed through to the EPA price", "%", PCT, True, "Assumption: fully indexed. Check against the EPA terms; the downside is partial indexation."),
    ("sec", "Costs"),
    ("CapexKW", "capex_per_kw", "Capex", "$/kW", "#,##0", True, "Assumption. BC reports $5 to 6B private capital for the awards (about $3,300 to 3,900/kW including interconnection, development and financing). The base case uses $3,000/kW of hard cost."),
    ("CapexPhaseY1", "capex_phase_y1", "Capex spent in construction year 1", "%", PCT, False, "Remainder spent in year 2."),
    ("FOM", "fom_per_kw", "Fixed O&M at COD", "$/kW-yr", "#,##0.0", True, "Assumption covering full service agreement, BOP and asset management."),
    ("LandPct", "land_pct", "Land, community and benefit payments", "% rev", PCT, True, "Assumption."),
    ("PtaxIns", "ptax_ins", "Property tax and insurance, first operating year", "$M", M2, True, "Assumption."),
    ("OpexEsc", "opex_esc", "Opex escalation", "%", PCT, True, "Matches CPI."),
    ("sec", "Tax"),
    ("ITCRate", "itc_rate", "Clean Technology ITC rate", "%", PCT, True, "Federal Clean Technology Investment Tax Credit, 30% of eligible cost."),
    ("ITCEligible", "itc_eligible", "Share of capex eligible for the ITC", "%", PCT, True, "Assumption: excludes ineligible items such as some interconnection and soft costs."),
    ("TaxRate", "tax_rate", "Combined corporate tax rate (BC)", "%", PCT, True, "15% federal plus 12% BC general rate."),
    ("CCARate", "cca_rate", "CCA rate (Class 43.2, declining balance)", "%", PCT, True, "50% declining balance. Accelerated investment incentive ignored for simplicity."),
    ("sec", "Financing"),
    ("InterestRate", "interest", "All in senior debt rate", "%", PCT2, True, "Assumption: swap rate plus a 150 to 175 bp margin for a contracted wind project."),
    ("Tenor", "tenor", "Debt tenor after COD", "years", INT, False, "Assumption: 18 years, leaving a 12 year tail on the 30 year EPA."),
    ("DSCRTarget", "dscr_target", "Sculpting DSCR target", "x", X, True, "Typical 1.30 to 1.40x for P50 contracted wind."),
    ("MaxGearing", "max_gearing", "Maximum gearing", "%", PCT, True, "Lender cap on debt as a share of capex."),
    ("EquityRate", "equity_rate", "Equity discount rate", "%", PCT, True, "Assumed hurdle for contracted renewables equity."),
]
r = 7
for spec in SPECS:
    if spec[0] == "sec":
        c = ws.cell(r, 2, spec[1])
        c.font = BOLD
        for j in range(2, 10):
            ws.cell(r, j).fill = SEC_FILL
        r += 1
        continue
    nm, fld, label, unit, fmt, flex, note = spec
    ws.cell(r, 2, label).font = BLACK
    ws.cell(r, 4, unit).font = BLACK
    bval = getattr(base, fld) if fld else pf.CON_YEARS
    e = ws.cell(r, 5, bval)
    e.font = BLUE
    if fld in up or fld in dn:
        ws.cell(r, 6, up.get(fld, bval)).font = BLUE
        ws.cell(r, 7, dn.get(fld, bval)).font = BLUE
        for col in (5, 6, 7):
            ws.cell(r, col).fill = PatternFill("solid", fgColor="EDEDED")
    else:
        ws.cell(r, 6, f"=E{r}").font = BLACK
        ws.cell(r, 7, f"=E{r}").font = BLACK
    if flex:
        h = ws.cell(r, 8, 0)
        h.font = BLUE
        h.number_format = PCT
        ws.cell(r, 3, f"=CHOOSE(Scenario,E{r},F{r},G{r})*(1+H{r})")
        name(f"Flex_{nm}", f"Inputs!$H${r}")
    else:
        ws.cell(r, 3, f"=CHOOSE(Scenario,E{r},F{r},G{r})")
    ws.cell(r, 3).font = BOLD
    for col in (3, 5, 6, 7):
        ws.cell(r, col).number_format = fmt
    n = ws.cell(r, 9, note)
    n.font = Font(name=FONT, size=9, italic=True, color="595959")
    name(nm, f"Inputs!$C${r}")
    r += 1

r += 1
ws.cell(r, 2, "Derived").font = BOLD
for j in range(2, 10):
    ws.cell(r, j).fill = SEC_FILL
r += 1
DERIVED = [
    ("CapexTotal", "Total capex", "$M", "=Capacity*1000*CapexKW/10^6", M),
    ("ITCTotal", "Clean Technology ITC", "$M", "=CapexTotal*ITCEligible*ITCRate", M),
    ("PriceCOD", "EPA price in first operating year", "$/MWh", "=EPAPrice*(1+CPI*Indexation)^(FirstYear+ConYears-2024)", "#,##0.00"),
]
for nm, label, unit, f, fmt in DERIVED:
    ws.cell(r, 2, label).font = BLACK
    c = ws.cell(r, 3, f)
    c.font, c.number_format = BOLD, fmt
    ws.cell(r, 4, unit).font = BLACK
    name(nm, f"Inputs!$C${r}")
    r += 1
ws.freeze_panes = "C7"

# --------------------------------------------------------------------------- Model
md = wb.create_sheet("Model")
N = pf.N_PERIODS
C0 = 5  # column E = period 1
COLS = [L(C0 + i) for i in range(N)]
FIRST, LAST = COLS[0], COLS[-1]
OPS1 = COLS[pf.CON_YEARS]  # first operating column (G)
setw(md, {"A": 2, "B": 40, "C": 9, "D": 11})
for cl in COLS:
    md.column_dimensions[cl].width = 9.5
md["B2"] = "Model"
md["B2"].font = TITLE
md["B3"] = "C$ millions unless stated. Column D is the total, or a single value for sizing rows."
md["B3"].font = Font(name=FONT, size=9, italic=True)

R = {}
row = 5


def section(title):
    global row
    row += 1
    for j in range(2, C0 + N):
        md.cell(row, j).fill = SEC_FILL
    md.cell(row, 2, title).font = BOLD
    row += 1


def line(key, label, unit, tmpl, fmt=M, total=False, first=None, font=BLACK, fill=None, values=None):
    """tmpl uses {c} for this column, {p} for the previous column and {key} for row numbers."""
    global row
    R[key] = row
    md.cell(row, 2, label).font = font if font is not BLACK else BLACK
    md.cell(row, 3, unit).font = Font(name=FONT, size=9, color="595959")
    for i, cl in enumerate(COLS):
        pc = COLS[i - 1] if i else None
        if values is not None:
            v = values[i]
        elif i == 0 and first is not None:
            v = first
        else:
            v = "=" + tmpl.format(c=cl, p=pc, **R)
        cell = md.cell(row, C0 + i, v)
        cell.number_format = fmt
        cell.font = font
        if fill:
            cell.fill = fill
    if total:
        t = md.cell(row, 4, f"=SUM({FIRST}{row}:{LAST}{row})")
        t.number_format, t.font = fmt, BOLD
    row += 1


def single(key, label, unit, formula, fmt, nm=None):
    global row
    R[key] = row
    md.cell(row, 2, label).font = BOLD
    md.cell(row, 3, unit).font = Font(name=FONT, size=9, color="595959")
    c = md.cell(row, 4, formula)
    c.number_format, c.font = fmt, BOLD
    if nm:
        name(nm, f"Model!$D${row}")
    row += 1


# header rows
R["period"] = row
md.cell(row, 2, "Period").font = BOLD
for i, cl in enumerate(COLS):
    md.cell(row, C0 + i, 1 if i == 0 else f"={COLS[i-1]}{row}+1").font = BOLD
row += 1
line("year", "Year", "", "FirstYear+{c}{period}-1", fmt=INT, font=BOLD)
for cl in COLS:
    md[f"{cl}{R['year']}"].fill = HDR_FILL
    md[f"{cl}{R['year']}"].font = HDR
line("con", "Construction flag", "flag", "IF({c}{period}<=ConYears,1,0)", fmt=INT)
line("ops", "Operations flag", "flag", "1-{c}{con}", fmt=INT)
line("opsyr", "Operating year", "#", "({c}{period}-ConYears)*{c}{ops}", fmt=INT)
line("debtflag", "Debt repayment flag", "flag", "IF(AND({c}{ops}=1,{c}{opsyr}<=Tenor),1,0)", fmt=INT)

section("Revenue")
line("gen", "Net generation", "GWh", "Capacity*8.76*CF*Availability*(1-Degradation)^({c}{opsyr}-1)*{c}{ops}", fmt="#,##0;(#,##0);\"-\"", total=True)
line("price", "EPA price", "$/MWh", "PriceCOD*(1+CPI*Indexation)^({c}{opsyr}-1)*{c}{ops}", fmt=M2)
line("rev", "Revenue", "$M", "{c}{gen}*{c}{price}/1000", total=True, font=BOLD)

section("Operating costs")
line("fom", "Fixed O&M", "$M", "Capacity*1000*FOM*(1+OpexEsc)^({c}{opsyr}-1)/10^6*{c}{ops}", total=True)
line("land", "Land, community and benefit payments", "$M", "{c}{rev}*LandPct", total=True)
line("ptax", "Property tax and insurance", "$M", "PtaxIns*(1+OpexEsc)^({c}{opsyr}-1)*{c}{ops}", total=True)
line("opex", "Total operating costs", "$M", "{c}{fom}+{c}{land}+{c}{ptax}", total=True)
line("ebitda", "EBITDA", "$M", "{c}{rev}-{c}{opex}", total=True, font=BOLD)
line("margin", "EBITDA margin", "%", "IF({c}{rev}>0,{c}{ebitda}/{c}{rev},0)", fmt=PCT)

section("Capex and ITC")
line("capex", "Construction capex", "$M", "CapexTotal*IF({c}{period}=1,CapexPhaseY1,IF({c}{period}=2,1-CapexPhaseY1,0))", total=True)
line("itc", "Clean Technology ITC received", "$M", "IF({c}{opsyr}=1,ITCTotal,0)", total=True)

section("Debt sizing (sculpted to target DSCR)")
base_res = pf.solve(pf.Inputs())
line("taxpaste", "Tax, pasted values (written by SculptDebt)", "$M", None, fill=MACRO_FILL, font=BLUE,
     values=[round(v, 10) for v in base_res["rows"]["tax"]], total=True)
md.cell(R["taxpaste"], 2).comment = Comment("Hard coded on purpose. SculptDebt copies the Tax row here and loops "
                                            "until the two match. See the circularity check on Outputs.", "Model")
line("sizecfads", "CFADS for sizing", "$M", "({c}{ebitda}-{c}{taxpaste})*{c}{ops}", total=True)
line("sculpted", "Sculpted debt service, unscaled", "$M", "{c}{sizecfads}/DSCRTarget*{c}{debtflag}", total=True)
single("pv", "PV of sculpted debt service at COD", "$M", f"=NPV(InterestRate,{OPS1}{R['sculpted']}:{LAST}{R['sculpted']})", M, "DebtCapacity")
single("debtamt", "Senior debt amount", "$M", "=MIN(DebtCapacity,MaxGearing*CapexTotal)", M, "DebtAmount")
single("gearing", "Gearing", "%", "=IF(CapexTotal>0,DebtAmount/CapexTotal,0)", PCT, "Gearing")
single("scale", "Debt service scaling factor", "x", "=IF(DebtCapacity>0,DebtAmount/DebtCapacity,0)", "0.0000", "DSScale")

section("Sources of funds")
line("draw", "Debt drawdown", "$M", "{c}{capex}*Gearing", total=True)
line("eqin", "Equity contribution", "$M", "{c}{capex}-{c}{draw}", total=True)

section("Senior debt")
line("ds", "Debt service", "$M", "{c}{sculpted}*DSScale", total=True)
R["dclose"] = row + 3   # forward reference, resolved when the row is written
line("dopen", "Opening balance", "$M", "{p}{dclose}", first=0)
line("int", "Interest", "$M", "{c}{dopen}*InterestRate*{c}{ops}", total=True)
line("prin", "Principal", "$M", "({c}{ds}-{c}{int})*{c}{ops}", total=True)
line("dclose", "Closing balance", "$M", "{c}{dopen}+{c}{draw}-{c}{prin}", font=BOLD)

section("Tax")
R["uclose"] = row + 3
line("uopen", "UCC opening", "$M", "{p}{uclose}", first=0)
line("uadd", "Additions net of ITC", "$M", "{c}{capex}-{c}{itc}", total=True)
line("cca", "CCA claimed", "$M", "CCARate*({c}{uopen}-{c}{itc})*{c}{ops}", total=True)
line("uclose", "UCC closing", "$M", "{c}{uopen}+{c}{uadd}-{c}{cca}")
line("ti", "Taxable income before losses", "$M", "{c}{ebitda}-{c}{cca}-{c}{int}", total=True)
R["lclose"] = row + 2
line("lopen", "Loss carryforward opening", "$M", "{p}{lclose}", first=0)
line("taxable", "Taxable income after losses", "$M", "MAX(0,{c}{ti}-{c}{lopen})", total=True)
line("lclose", "Loss carryforward closing", "$M", "MAX(0,{c}{lopen}-{c}{ti})")
line("tax", "Cash tax", "$M", "{c}{taxable}*TaxRate", total=True, font=BOLD)
line("taxdiff", "Circularity gap |tax - pasted tax|", "$M", "ABS({c}{tax}-{c}{taxpaste})", fmt="0.000000")

section("Cash flow waterfall")
line("cfads", "CFADS", "$M", "({c}{ebitda}-{c}{tax})*{c}{ops}", total=True, font=BOLD)
line("dscr", "DSCR", "x", 'IF({c}{ds}>0,{c}{cfads}/{c}{ds},"")', fmt=X)
line("cfadsloan", "CFADS during loan life", "$M", "{c}{cfads}*{c}{debtflag}", total=True)
line("eqdist", "Cash to equity after debt service", "$M", "({c}{cfads}-{c}{ds})*{c}{ops}", total=True)
line("eqcf", "Equity cash flow (incl. ITC)", "$M", "-{c}{eqin}+{c}{eqdist}+{c}{itc}", total=True, font=BOLD)
line("projcf", "Project cash flow, pre tax and unlevered (incl. ITC)", "$M", "-{c}{capex}+{c}{ebitda}+{c}{itc}", total=True)
line("cumeq", "Cumulative equity cash flow", "$M", "{p}{cumeq}+{c}{eqcf}", first=f"={FIRST}{R['eqcf']}")
md.freeze_panes = "E7"
md.column_dimensions.group(start="A", end="A")


def rng(key, start=FIRST):
    return f"Model!${start}${R[key]}:${LAST}${R[key]}"


for key, nm in [("taxpaste", "TaxPaste"), ("tax", "TaxCalc"), ("eqcf", "EquityCF"), ("projcf", "ProjectCF"), ("dscr", "DSCRRow"), ("taxdiff", "TaxGap")]:
    name(nm, rng(key))
name("CFADSLoan", rng("cfadsloan", OPS1))

# --------------------------------------------------------------------------- Outputs
op = wb.create_sheet("Outputs")
setw(op, {"A": 2, "B": 44, "C": 14, "D": 4, "E": 44, "F": 14})
op["B2"] = "Outputs"
op["B2"].font = TITLE
op["B3"] = '="Scenario: "&CHOOSE(Scenario,"Base","Upside","Downside")'
op["B3"].font = BOLD


def block(r0, col, title, items):
    c = op.cell(r0, col, title)
    c.font, c.fill = HDR, HDR_FILL
    op.cell(r0, col + 1).fill = HDR_FILL
    for i, (lab, f, fmt) in enumerate(items, start=1):
        op.cell(r0 + i, col, lab).font = BLACK
        v = op.cell(r0 + i, col + 1, f)
        v.number_format = fmt
        v.font = GREEN if f.startswith("=Model!") else BLACK
    return r0 + len(items) + 2


g = OPS1
nxt = block(5, 2, "Returns", [
    ("Equity IRR (nominal, after tax)", "=IRR(EquityCF)", PCT2),
    ("Project IRR (pre tax, unlevered)", "=IRR(ProjectCF)", PCT2),
    ("Equity NPV at hurdle rate ($M)", "=NPV(EquityRate,EquityCF)", M),
    ("Equity hurdle rate", "=EquityRate", PCT),
    ("Equity payback year", f'=IFERROR(INDEX(Model!${FIRST}${R["year"]}:${LAST}${R["year"]},MATCH(TRUE,INDEX(Model!${FIRST}${R["cumeq"]}:${LAST}${R["cumeq"]}>=0,0),0)),"n/a")', INT),
])
nxt = block(nxt, 2, "Debt", [
    ("Senior debt ($M)", "=DebtAmount", M),
    ("Gearing", "=Gearing", PCT),
    ("Binding constraint", '=IF(DebtCapacity<=MaxGearing*CapexTotal,"DSCR","Gearing")', "@"),
    ("Minimum DSCR", "=MIN(DSCRRow)", X),
    ("Average DSCR", "=AVERAGE(DSCRRow)", X),
    ("LLCR", "=IF(DebtAmount>0,NPV(InterestRate,CFADSLoan)/DebtAmount,0)", X),
    ("Debt fully repaid by", f"=FirstYear+ConYears+Tenor-1", INT),
])
block(5, 5, "Operations (first operating year)", [
    ("Net generation (GWh)", f"=Model!{g}{R['gen']}", "#,##0"),
    ("EPA price ($/MWh)", "=PriceCOD", M2),
    ("Revenue ($M)", f"=Model!{g}{R['rev']}", M),
    ("EBITDA ($M)", f"=Model!{g}{R['ebitda']}", M),
    ("EBITDA margin", f"=Model!{g}{R['margin']}", PCT),
])
nx2 = block(12, 5, "Sources and uses ($M)", [
    ("Uses: construction capex", "=CapexTotal", M),
    ("Sources: senior debt", f"=Model!D{R['draw']}", M),
    ("Sources: sponsor equity", f"=Model!D{R['eqin']}", M),
    ("Total sources", f"=Model!D{R['draw']}+Model!D{R['eqin']}", M),
    ("Clean Technology ITC (received end of operating year 1)", "=ITCTotal", M),
    ("Implied EV/EBITDA at close (capex / yr 1 EBITDA)", f"=CapexTotal/Model!{g}{R['ebitda']}", "0.0x"),
])
chk = block(nx2, 5, "Integrity checks (1 = pass)", [
    ("Sources equal uses", f"=IF(ABS(Model!D{R['draw']}+Model!D{R['eqin']}-CapexTotal)<0.0001,1,0)", INT),
    ("Debt fully repaid", f"=IF(ABS(Model!{LAST}{R['dclose']})<0.0001,1,0)", INT),
    ("Tax circularity converged (run SculptDebt if 0)", "=IF(MAX(TaxGap)<0.001,1,0)", INT),
    ("Min DSCR at or above target", "=IF(MIN(DSCRRow)>=DSCRTarget-0.0001,1,0)", INT),
])
op.cell(chk - 1, 5, "All checks").font = BOLD
c = op.cell(chk - 1, 6, f"=IF(SUM(F{chk-5}:F{chk-2})=4,\"OK\",\"CHECK\")")
c.font = BOLD
c.fill = KEY_FILL

# charts
yrs = Reference(md, min_col=C0 + pf.CON_YEARS, max_col=C0 + N - 1, min_row=R["year"])
bc = BarChart()
bc.title = "CFADS vs debt service ($M)"
bc.y_axis.title = "$M"
for key in ("cfads", "ds"):
    ref = Reference(md, min_col=C0 + pf.CON_YEARS, max_col=C0 + N - 1, min_row=R[key])
    bc.add_data(ref, from_rows=True, titles_from_data=False)
bc.series[0].tx = None
from openpyxl.chart.series import SeriesLabel
bc.series[0].tx = SeriesLabel(v="CFADS")
bc.series[1].tx = SeriesLabel(v="Debt service")
bc.set_categories(yrs)
bc.height, bc.width = 8, 18
op.add_chart(bc, "B28")
lc = LineChart()
lc.title = "Senior debt balance ($M)"
lc.add_data(Reference(md, min_col=C0, max_col=C0 + N - 1, min_row=R["dclose"]), from_rows=True, titles_from_data=False)
lc.series[0].tx = SeriesLabel(v="Closing balance")
lc.set_categories(Reference(md, min_col=C0, max_col=C0 + N - 1, min_row=R["year"]))
lc.height, lc.width = 8, 18
op.add_chart(lc, "B45")

# --------------------------------------------------------------------------- Sensitivity
sv = wb.create_sheet("Sensitivity")
setw(sv, {"A": 2, "B": 52, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 13, "I": 13})
sv["B2"] = "Sensitivity"
sv["B2"].font = TITLE
sv["B3"] = "Peach cells are written by the VBA macros (seeded from the Python mirror, base case). Re-run after changing inputs."
sv["B3"].font = Font(name=FONT, size=9, italic=True)

tor = sens.tornado(pf.Inputs())
hdrs = ["Driver (flexed ±10%)", "Flex name", "Equity IRR -10%", "Equity IRR base", "Equity IRR +10%",
        "Min DSCR -10%", "Min DSCR +10%", "IRR swing (pp)"]
for j, h in enumerate(hdrs, start=2):
    c = sv.cell(5, j, h)
    c.font, c.fill = HDR, HDR_FILL
    c.alignment = Alignment(wrap_text=True)
for i, t in enumerate(tor, start=6):
    sv.cell(i, 2, t["label"]).font = BLACK
    sv.cell(i, 3, t["flex"]).font = BLACK
    for j, k in enumerate(["irr_lo", "irr_base", "irr_hi", "dscr_lo", "dscr_hi"], start=4):
        c = sv.cell(i, j, t[k])
        c.fill, c.font = MACRO_FILL, BLUE
        c.number_format = PCT2 if "irr" in k else X
    c = sv.cell(i, 9, f"=ABS(F{i}-D{i})*100")
    c.number_format = "0.00"
tor_end = 5 + len(tor)

r2 = tor_end + 3
sv.cell(r2 - 1, 2, "Equity IRR: EPA price (2024$, rows) vs net capacity factor (columns)").font = BOLD
grid = sens.two_way(pf.Inputs())
sv.cell(r2, 2, "EPA price $/MWh  \\  CF").font = HDR
sv.cell(r2, 2).fill = HDR_FILL
for j, cf in enumerate(grid["cfs"], start=3):
    c = sv.cell(r2, j, cf)
    c.font, c.fill, c.number_format = HDR, HDR_FILL, PCT
for i, p in enumerate(grid["prices"], start=r2 + 1):
    c = sv.cell(i, 2, p)
    c.font, c.number_format = BOLD, "#,##0.00"
    for j, _ in enumerate(grid["cfs"], start=3):
        v = sv.cell(i, j, grid["irr"][i - r2 - 1][j - 3])
        v.fill, v.font, v.number_format = MACRO_FILL, BLUE, PCT2
name("TwoWayCorner", f"Sensitivity!$B${r2}")

r3 = r2 + len(grid["prices"]) + 3
sv.cell(r3, 2, "Breakeven EPA price (2024$) for equity NPV = 0 at hurdle").font = BOLD
c = sv.cell(r3, 3, sens.breakeven_price(pf.Inputs()))
c.fill, c.font, c.number_format = MACRO_FILL, BLUE, "#,##0.00"
name("BreakevenOut", f"Sensitivity!$C${r3}")
sv.cell(r3 + 1, 2, "Headroom vs awarded $74 average").font = BLACK
c = sv.cell(r3 + 1, 3, f"=EPAPrice/(1+Flex_EPAPrice)-C{r3}")
c.number_format = "#,##0.00"
name("TornadoTop", "Sensitivity!$B$6")

for wsx in wb.worksheets:
    wsx.sheet_view.showGridLines = False
wb.calculation.fullCalcOnLoad = True
OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print("saved", OUT, "rows:", {k: R[k] for k in ("taxpaste", "tax", "cfads", "eqcf", "dclose")})
