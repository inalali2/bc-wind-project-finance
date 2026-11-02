# Resume entry and interview prep

## Resume entry (replaces Alethia under PROJECT)

**Coastal Ridge Wind: Project Finance Model** | Excel, VBA, Python | Nov 2026
github.com/inalali2/bc-wind-project-finance
- Built a project finance model in Excel for a 200 MW BC wind farm with a 30 year BC Hydro EPA, sculpting $338M of senior debt to a 1.35x DSCR and valuing sponsor equity at a 10.3% IRR.
- Modelled Class 43.2 CCA, loss carryforwards and the 30% Clean Technology ITC, and wrote VBA macros to resolve the tax and debt circularity, run tornado and two way sensitivities, and solve a $70.93/MWh breakeven price.
- Reconciled the workbook line by line to an independent Python model with 34 automated tests, and wrote an investment memo identifying ITC timing and EPA indexation as the key structuring risks.

Keep the crude differential project above it; together they show energy markets plus infrastructure finance. Move Alethia off the page to make room. If space is tight, keep only the first two bullets.

**Skills line addition:** `Financial modelling: project finance, debt sculpting, DCF, VBA`

## Questions to be ready for

1. **Walk me through the model.** Inputs → generation and revenue → opex → EBITDA → tax → CFADS → debt sizing → debt schedule → equity cash flows → IRR and NPV.
2. **Why sculpt?** Sculpting matches debt service to cash flow so DSCR is flat. That maximises debt for a given DSCR covenant and raises equity IRR.
3. **Why is there a circularity, and how did you break it?** Debt size depends on after tax CFADS, tax depends on interest, and interest depends on debt size. A copy and paste macro is auditable and stable, while iterative calculation can hide errors.
4. **What binds, DSCR or gearing?** DSCR. Capacity at 1.35x is $338M, below the 75% cap of $450M.
5. **What would you change first?** An ITC bridge loan. The credit returns about 62% of equity in year 1, so bridging it during construction cuts peak equity.
6. **LLCR vs DSCR?** DSCR is one period. LLCR is the PV of CFADS over the loan life divided by debt. With full sculpting they are equal.
7. **Equity IRR vs project IRR?** Leverage at a 5.25% cost of debt below the 8.2% project return lifts the equity return.
8. **How would you value this asset in a sale?** DCF of equity cash flows at a market discount rate, cross checked against precedent $/MW and EV/EBITDA for contracted wind.

Rebuild the debt sizing and tax blocks by hand in a blank workbook at least once. You should be able to explain every formula without notes.
