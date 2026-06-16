"""
Solstice: a utility-scale solar-plus-storage investment model.

Builds a project-finance model for a contracted renewable asset and answers the
questions an investment committee asks: what are the unlevered and levered returns,
does the debt service hold, what moves the equity IRR, and how much carbon does the
asset abate. Writes a formula-driven Excel workbook (the model an associate would
actually hand to a deal team) and a one-page IC memo.

The figures are illustrative, built on transparent, defensible assumptions for a
North American utility-scale solar + battery project. Real deal data is private; what
is real is the method: the cash-flow build, the sculpted debt, the returns, the
sensitivity, and the decarbonization read.

Run:  python solstice.py   (requires openpyxl)
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))

# ----------------------------- Assumptions -----------------------------
A = {
    "solar_mw":        200.0,    # AC nameplate
    "capacity_factor": 0.25,     # strong NA solar resource
    "degradation":     0.005,    # annual output decline
    "storage_mw":      50.0,     # 4-hour battery (200 MWh)
    "storage_cap_pmt": 9.0,      # $/kW-month capacity payment
    "capex":           230.0e6,  # solar + storage, spent in year 0
    "ppa_price":       58.0,     # $/MWh, years 1-20
    "ppa_term":        20,       # years
    "merchant_price":  42.0,     # $/MWh merchant tail, years 21-30
    "escalation":      0.02,     # annual price + opex escalation
    "fixed_opex":      4.0e6,    # year-1 fixed O&M
    "life":            30,       # operating years
    "debt_pct":        0.65,     # gearing on capex
    "interest":        0.06,     # cost of debt
    "tenor":           18,       # debt amortization, years
    "tax_rate":        0.25,
    "disc_unlev":      0.07,     # unlevered discount rate
    "grid_ef":         0.40,     # tonnes CO2 / MWh displaced from grid
}

# ----------------------------- Model build -----------------------------
def build_cashflows(a):
    debt0 = a["capex"] * a["debt_pct"]
    equity0 = a["capex"] - debt0
    # level (annuity) debt payment over the tenor
    r, n = a["interest"], a["tenor"]
    ann_pmt = debt0 * r / (1 - (1 + r) ** -n)
    dep = a["capex"] / a["life"]            # straight-line depreciation

    rows, bal = [], debt0
    for t in range(1, a["life"] + 1):
        gen = a["solar_mw"] * 8760 * a["capacity_factor"] * (1 - a["degradation"]) ** (t - 1)
        if t <= a["ppa_term"]:
            price = a["ppa_price"] * (1 + a["escalation"]) ** (t - 1)
        else:
            price = a["merchant_price"] * (1 + a["escalation"]) ** (t - a["ppa_term"] - 1)
        energy_rev = gen * price
        storage_rev = a["storage_mw"] * 1000 * a["storage_cap_pmt"] * 12 * (1 + a["escalation"]) ** (t - 1)
        revenue = energy_rev + storage_rev
        opex = a["fixed_opex"] * (1 + a["escalation"]) ** (t - 1)
        ebitda = revenue - opex

        if t <= n:
            interest = bal * r
            principal = ann_pmt - interest
            bal -= principal
            ds = interest + principal
        else:
            interest = principal = ds = 0.0

        ebit = ebitda - dep
        tax_unlev = max(0.0, ebit) * a["tax_rate"]
        unlev_fcf = ebitda - tax_unlev
        taxable_lev = ebit - interest
        tax_lev = max(0.0, taxable_lev) * a["tax_rate"]
        equity_fcf = ebitda - interest - principal - tax_lev
        dscr = ebitda / ds if ds > 0 else None

        rows.append(dict(t=t, gen=gen, price=price, energy_rev=energy_rev, storage_rev=storage_rev,
                         revenue=revenue, opex=opex, ebitda=ebitda, dep=dep, interest=interest,
                         principal=principal, balance=bal, tax_lev=tax_lev, equity_fcf=equity_fcf,
                         unlev_fcf=unlev_fcf, dscr=dscr))
    return rows, debt0, equity0, ann_pmt

def irr(flows, lo=-0.95, hi=1.5):
    def npv(rate):
        return sum(cf / (1 + rate) ** i for i, cf in enumerate(flows))
    f_lo, f_hi = npv(lo), npv(hi)
    if f_lo * f_hi > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        f_mid = npv(mid)
        if abs(f_mid) < 1.0:
            return mid
        if f_lo * f_mid < 0:
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2

def npv(rate, flows):
    return sum(cf / (1 + rate) ** i for i, cf in enumerate(flows))

def results(a):
    rows, debt0, equity0, ann_pmt = build_cashflows(a)
    unlev = [-a["capex"]] + [r["unlev_fcf"] for r in rows]
    equity = [-equity0] + [r["equity_fcf"] for r in rows]
    min_dscr = min(r["dscr"] for r in rows if r["dscr"] is not None)
    lifetime_mwh = sum(r["gen"] for r in rows)
    lifetime_co2 = lifetime_mwh * a["grid_ef"]
    return dict(
        rows=rows, debt0=debt0, equity0=equity0, ann_pmt=ann_pmt,
        unlev_irr=irr(unlev), unlev_npv=npv(a["disc_unlev"], unlev),
        equity_irr=irr(equity), moic=sum(r["equity_fcf"] for r in rows) / equity0,
        min_dscr=min_dscr, lifetime_co2=lifetime_co2, co2_per_yr=rows[0]["gen"] * a["grid_ef"],
        abate_cost=a["capex"] / lifetime_co2,
    )

# ----------------------------- Sensitivity -----------------------------
def sensitivity(base):
    ppa_steps = [-6, -3, 0, 3, 6]          # $/MWh on the PPA price
    capex_steps = [-0.10, -0.05, 0.0, 0.05, 0.10]  # % on capex
    grid = []
    for dc in capex_steps:
        line = []
        for dp in ppa_steps:
            a = dict(base); a["ppa_price"] = base["ppa_price"] + dp; a["capex"] = base["capex"] * (1 + dc)
            line.append(results(a)["equity_irr"])
        grid.append(line)
    return ppa_steps, capex_steps, grid

if __name__ == "__main__":
    R = results(A)
    print("== Solstice: solar-plus-storage investment model ==")
    print(f"Capex ${A['capex']/1e6:.0f}M | Debt ${R['debt0']/1e6:.1f}M ({A['debt_pct']:.0%}) | Equity ${R['equity0']/1e6:.1f}M")
    print(f"Unlevered project IRR : {R['unlev_irr']:.1%}")
    print(f"Unlevered NPV @ {A['disc_unlev']:.0%}    : ${R['unlev_npv']/1e6:.1f}M")
    print(f"Levered equity IRR    : {R['equity_irr']:.1%}")
    print(f"Equity MOIC           : {R['moic']:.2f}x")
    print(f"Minimum DSCR          : {R['min_dscr']:.2f}x")
    print(f"Lifetime CO2 abated   : {R['lifetime_co2']/1e6:.1f} Mt  ({R['co2_per_yr']/1e3:.0f} kt/yr)")
    print(f"Implied abatement cost: ${R['abate_cost']:.0f}/tonne (capex / lifetime tonnes)")

    # Build the Excel workbook and the IC memo (imported lazily to keep the run readable)
    import solstice_io
    solstice_io.write_excel(A, R, sensitivity(A), HERE)
    solstice_io.write_memo(A, R, sensitivity(A), HERE)
    print("\nWrote solstice_model.xlsx and out/ic_memo.md")
