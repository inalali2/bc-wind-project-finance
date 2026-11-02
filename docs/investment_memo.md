# Coastal Ridge Wind: Financing and Valuation Memo

**To:** Sponsor investment committee | **From:** Inal Ali | **Date:** November 2026
**Asset:** Hypothetical 200 MW onshore wind farm, northeast BC, 30 year BC Hydro EPA
**Model:** `model/BC_Wind_Project_Finance.xlsx` (base case unless stated, C$ millions)

*Portfolio exercise. The asset is hypothetical and the assumptions are illustrative, anchored to public data on BC Hydro's 2024 Call for Power.*

## 1. Recommendation

Proceed to lender marketing on the base case structure. The project clears a 9% equity hurdle with a 10.3% equity IRR and $20.8M of NPV. The margin is thin, so the price is the main thing to defend. Breakeven is $70.93/MWh (2024$) against the $74 average awarded, which is about $3 or 4% of headroom.

## 2. Transaction summary

| Item | Base case |
|---|---|
| Capex ($3,000/kW) | $600.0M |
| Senior debt, sculpted at 1.35x DSCR, 18 year tenor | $338.0M (56% gearing) |
| Sponsor equity | $262.0M |
| Clean Technology ITC (30% of eligible capex) | $162.0M, received in operating year 1 |
| First year revenue / EBITDA | $50.0M / $36.5M (73% margin) |
| EPA price in 2029 (indexed from $74 in 2024$) | $81.70/MWh |
| Equity IRR / project IRR (pre tax, unlevered) | 10.3% / 8.2% |
| Min DSCR / LLCR | 1.35x / 1.35x |

Debt is limited by the DSCR test, not by the 75% gearing cap. Sculpting sets debt service equal to CFADS divided by 1.35 in every year, so DSCR is flat and the LLCR equals the target.

## 3. What drives value

**Price and resource dominate.** A 10% move in capacity factor or EPA price moves equity IRR from 7.4% to 14.4%. Those two inputs have the same effect because revenue is generation times price. Capex comes next (13.5% to 8.3% for minus or plus 10%). Opex, interest rate and DSCR target each move IRR by less than 2 points.

**The ITC carries the equity case.** The $162M credit returns about 62% of sponsor equity in the first operating year. Without it, equity IRR falls to 5.3% and NPV to negative $91.5M. Two consequences:
- Eligibility should be confirmed early, for example through a tax opinion.
- An ITC bridge loan during construction would cut peak equity and lift returns further. This is the first structuring extension to model.

**Indexation matters more than it looks.** If only half of CPI is passed through to the EPA price, IRR drops to 6.4%. The EPA indexation terms need checking before the price assumption is final.

## 4. Scenarios

| Scenario | Capacity factor | Capex $/kW | Fixed O&M $/kW-yr | Rate | Equity IRR | Gearing |
|---|---|---|---|---|---|---|
| Upside | 38.5% | 2,800 | 42 | 4.75% | 18.6% | 68% |
| Base | 36.0% | 3,000 | 45 | 5.25% | 10.3% | 56% |
| Downside | 33.0% | 3,400 | 52 | 6.00% | 4.9% | 41% |

Downside is a P90 style resource year combined with a 13% capex overrun and higher rates. Under DSCR based sculpting, lenders are protected in the downside because debt shrinks with cash flow, but equity takes the hit.

## 5. Financing levers worth testing

| Lever | Change | Equity IRR |
|---|---|---|
| Lower DSCR target | 1.35x to 1.30x | 10.6% |
| Longer tenor | 18 to 20 years | 11.4% |

A 20 year tenor still leaves a 10 year tail on the 30 year EPA, so lenders may accept it. Federal Indigenous loan guarantee programs may support the First Nations equity share (the Call for Power required at least 25% First Nations ownership, and most awarded projects have 51%). Lower cost partner equity would improve the blended cost of capital.

## 6. Key risks

1. **Wind resource.** The biggest driver. Mitigate with an independent energy yield assessment and P50 / P90 sizing.
2. **Construction cost.** Mitigate with a fixed price EPC or turbine supply agreement and contingency.
3. **ITC eligibility and timing.** Confirm early and structure a bridge.
4. **Curtailment and interconnection.** Not modelled. Review EPA curtailment compensation terms.
5. **Tail risk after year 30.** No merchant or repowering value is assumed, which is conservative.

## 7. Model notes

- Annual periods; financial close at the start of 2027 and COD in 2029.
- Tax: 27% combined BC rate; Class 43.2 CCA at 50% declining balance; the ITC reduces the UCC pool; losses carried forward. With CCA and losses, no cash tax is paid until 2044.
- Circularity: tax depends on interest, interest on debt size, and debt size on after tax CFADS. The `SculptDebt` macro breaks this by pasting tax as values and iterating to a $0.000001M tolerance. A check on the Outputs tab flags a stale paste.
- The Excel model reconciles to an independent Python version to within $0.000001M on every line (34 automated tests).

**Sources:** BC Government, "BC Hydro selects nine wind projects" (Dec 2024), news.gov.bc.ca/releases/2024ECS0048-001643 · BC Hydro 2024 Call for Power results ($74/MWh weighted average levelized price, 2024$, about 4,830 GWh/yr), reported by myeastkootenaynow.com/?p=43673 · BC Hydro, BCUC acceptance of 2024 Call for Power EPAs (2025).
