# Solstice: a utility-scale solar-plus-storage investment model

What does an investment committee need before it commits equity to a renewable asset? The returns,
unlevered and levered. Whether the debt service holds through the amortization. What actually moves the
equity IRR. And, for a transition mandate, how much carbon the asset abates. Solstice is a working
project-finance model that answers all four, built in **Excel (formula-driven) with a Python engine** and
a one-page IC memo.

It models a **200 MW solar + 50 MW / 200 MWh storage** project on a 20-year contracted PPA with a merchant
tail, and underwrites it the way a deal team would.

## What it does
- **Builds a 30-year project cash flow** from first principles: generation (with annual degradation), a
  contracted PPA price stepping to a merchant tail, a storage capacity payment, fixed O&M, and capex.
- **Sculpts the capital structure:** debt sized to gearing, amortized over the tenor, with an annual
  **DSCR** so you can see the debt service holds.
- **Computes the returns the IC asks for:** unlevered project IRR and NPV, levered equity IRR, MOIC, and
  minimum DSCR.
- **Runs a sensitivity** of levered equity IRR against the two swing factors, **PPA price and capex**.
- **Scores the decarbonization impact:** tonnes of CO2 abated per year and over life, and an implied
  abatement cost, mapping return to a transition-investment framework.
- **Writes two artifacts:** `solstice_model.xlsx` (a real, auditable, formula-driven workbook with
  Assumptions, Model, and Sensitivity tabs) and `out/ic_memo.md` (the one-page recommendation).

## Headline result (base case)
On the assumptions in the model: **9.5% unlevered IRR**, **+$60M unlevered NPV** at a 7% discount rate,
**14.9% levered equity IRR**, **5.6x MOIC**, and a **1.9x minimum DSCR** on an $80.5M equity cheque.
Lifetime abatement is about **4.9 Mt CO2** (~175 kt/yr). All figures recompute live from the Assumptions
tab.

## The data
Real deal terms are private, so the project runs on **transparent, defensible assumptions** for a North
American utility-scale solar + battery asset (capacity factor, capex per kW, PPA price, gearing, tenor; all
on the Assumptions tab). Change an input and the whole model and the returns recompute. The figures are
illustrative; the **method** is the point: the cash-flow build, the sculpted debt, the returns, the
sensitivity, and the abatement read.

## Run it
Requires Python and `openpyxl`.
```bash
pip install -r requirements.txt
python solstice.py
```
You will see the returns in the console, and the script regenerates `solstice_model.xlsx` and
`out/ic_memo.md`. Open the workbook in Excel to audit every formula from the Assumptions tab through to the
returns.

## Concepts and references
- **Project finance / DSCR:** debt sized and amortized against contracted cash flow, with a debt-service
  coverage ratio as the lender's covenant; standard infrastructure underwriting.
- **Levered vs unlevered returns:** unlevered IRR/NPV measure the asset; levered equity IRR and MOIC
  measure the equity after the capital structure.
- **Contracted-to-merchant:** a PPA underwrites the early years; the post-PPA merchant tail is modelled
  conservatively as upside, not as the basis for the decision.
- **Transition framework:** scoring both financial return and tonnes abated mirrors how a transition fund
  evaluates an asset.

## Limitations
The cash flows are synthetic and the assumptions are illustrative, so the figures are not a real deal.
Depreciation and tax are simplified (straight-line, single rate; no ITC/PTC or MACRS/CCA timing), debt is
a level annuity rather than fully sculpted to a target DSCR, and the merchant tail is a flat assumption.
The point is to show I can build and audit a project-finance model end to end, pull the returns an IC
needs, and frame the decision, in Excel and in code.
