"""Excel workbook (formula-driven) and IC-memo writers for Solstice."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

NAVY = "1F3864"
LIGHT = "D9E1F2"
bold_white = Font(bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor=NAVY)
sub_fill = PatternFill("solid", fgColor=LIGHT)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Assumptions sheet row map (label, key, number_format) — derived rows added after.
ROWS = [
    ("Solar nameplate (MW)",        "solar_mw",        "#,##0"),
    ("Capacity factor",             "capacity_factor", "0.0%"),
    ("Annual degradation",          "degradation",     "0.00%"),
    ("Storage power (MW)",          "storage_mw",      "#,##0"),
    ("Storage capacity pmt ($/kW-mo)", "storage_cap_pmt", "#,##0.00"),
    ("Capex ($)",                   "capex",           "#,##0"),
    ("PPA price ($/MWh)",           "ppa_price",       "#,##0.00"),
    ("PPA term (yrs)",              "ppa_term",        "#,##0"),
    ("Merchant price ($/MWh)",      "merchant_price",  "#,##0.00"),
    ("Escalation",                  "escalation",      "0.0%"),
    ("Fixed O&M, yr1 ($)",          "fixed_opex",      "#,##0"),
    ("Operating life (yrs)",        "life",            "#,##0"),
    ("Gearing (debt % capex)",      "debt_pct",        "0%"),
    ("Cost of debt",                "interest",        "0.0%"),
    ("Debt tenor (yrs)",            "tenor",           "#,##0"),
    ("Tax rate",                    "tax_rate",        "0%"),
    ("Unlevered discount rate",     "disc_unlev",      "0.0%"),
    ("Grid emissions factor (t/MWh)", "grid_ef",       "0.00"),
]

def write_excel(a, R, sens, here):
    wb = Workbook()
    # ---------------- Assumptions ----------------
    ws = wb.active; ws.title = "Assumptions"
    ws["A1"] = "Solstice | Assumptions"; ws["A1"].font = Font(bold=True, size=13, color=NAVY)
    ws.column_dimensions["A"].width = 34; ws.column_dimensions["B"].width = 18
    for i, (label, key, fmt) in enumerate(ROWS, start=2):
        ws.cell(i, 1, label)
        c = ws.cell(i, 2, a[key]); c.number_format = fmt; c.fill = sub_fill; c.border = border
    # derived (cells B21-B24, referenced by the model)
    base = len(ROWS) + 2  # = 20, so derived start row 21
    ws.cell(base + 1, 1, "Debt drawn ($)");   ws.cell(base + 1, 2).value = "=B7*B14";              ws.cell(base + 1, 2).number_format = "#,##0"
    ws.cell(base + 2, 1, "Equity ($)");        ws.cell(base + 2, 2).value = "=B7-B21";              ws.cell(base + 2, 2).number_format = "#,##0"
    ws.cell(base + 3, 1, "Level debt pmt ($)"); ws.cell(base + 3, 2).value = "=B21*B15/(1-(1+B15)^-B16)"; ws.cell(base + 3, 2).number_format = "#,##0"
    ws.cell(base + 4, 1, "Depreciation/yr ($)"); ws.cell(base + 4, 2).value = "=B7/B13";            ws.cell(base + 4, 2).number_format = "#,##0"

    # ---------------- Model ----------------
    m = wb.create_sheet("Model")
    headers = ["Year", "Generation (MWh)", "Energy price ($/MWh)", "Energy revenue ($)",
               "Storage revenue ($)", "Revenue ($)", "Opex ($)", "EBITDA ($)", "Depreciation ($)",
               "Interest ($)", "Principal ($)", "Debt balance ($)", "Tax-levered ($)",
               "Equity cash flow ($)", "Unlevered FCF ($)", "DSCR"]
    for j, h in enumerate(headers, start=1):
        c = m.cell(1, j, h); c.font = bold_white; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = border
        m.column_dimensions[chr(64 + j)].width = 15 if j > 1 else 6
    AS = "Assumptions!"
    # Year 0 (row 2)
    m.cell(2, 1, 0)
    m.cell(2, 12, f"={AS}$B$21")    # debt balance
    m.cell(2, 14, f"=-{AS}$B$22")   # equity outflow
    m.cell(2, 15, f"=-{AS}$B$7")    # capex outflow
    for j in (4, 5, 6, 7, 8, 9, 10, 11, 13):
        m.cell(2, j, 0)
    # Years 1..life
    for t in range(1, a["life"] + 1):
        r = t + 2; p = r - 1
        m.cell(r, 1, t)
        m.cell(r, 2, f"={AS}$B$2*8760*{AS}$B$3*(1-{AS}$B$4)^(A{r}-1)")
        m.cell(r, 3, f"=IF(A{r}<={AS}$B$9,{AS}$B$8*(1+{AS}$B$11)^(A{r}-1),{AS}$B$10*(1+{AS}$B$11)^(A{r}-{AS}$B$9-1))")
        m.cell(r, 4, f"=B{r}*C{r}")
        m.cell(r, 5, f"={AS}$B$5*1000*{AS}$B$6*12*(1+{AS}$B$11)^(A{r}-1)")
        m.cell(r, 6, f"=D{r}+E{r}")
        m.cell(r, 7, f"={AS}$B$12*(1+{AS}$B$11)^(A{r}-1)")
        m.cell(r, 8, f"=F{r}-G{r}")
        m.cell(r, 9, f"={AS}$B$24")
        m.cell(r, 10, f"=IF(A{r}<={AS}$B$16,L{p}*{AS}$B$15,0)")
        m.cell(r, 11, f"=IF(A{r}<={AS}$B$16,{AS}$B$23-J{r},0)")
        m.cell(r, 12, f"=IF(A{r}<={AS}$B$16,L{p}-K{r},0)")
        m.cell(r, 13, f"=MAX(0,H{r}-I{r}-J{r})*{AS}$B$17")
        m.cell(r, 14, f"=H{r}-J{r}-K{r}-M{r}")
        m.cell(r, 15, f"=H{r}-MAX(0,H{r}-I{r})*{AS}$B$17")
        m.cell(r, 16, f'=IF(A{r}<={AS}$B$16,H{r}/(J{r}+K{r}),"")')
        for j in range(2, 17):
            cell = m.cell(r, j)
            cell.number_format = "0.00" if j == 16 else ("#,##0.00" if j == 3 else "#,##0")
    last = a["life"] + 2
    # ---------------- Returns block ----------------
    rb = last + 2
    block = [
        ("Unlevered project IRR", f"=IRR(O2:O{last})", "0.0%"),
        ("Unlevered NPV @ disc rate ($)", f"=O2+NPV({AS}B18,O3:O{last})", "#,##0"),
        ("Levered equity IRR", f"=IRR(N2:N{last})", "0.0%"),
        ("Equity MOIC", f"=SUM(N3:N{last})/-N2", '0.00"x"'),
        ("Minimum DSCR", f"=MIN(P3:P{last})", '0.00"x"'),
    ]
    m.cell(rb, 1, "RETURNS").font = Font(bold=True, color=NAVY)
    for i, (lab, f, fmt) in enumerate(block, start=rb + 1):
        m.cell(i, 1, lab); c = m.cell(i, 2, f); c.number_format = fmt; c.font = Font(bold=True)

    # ---------------- Sensitivity ----------------
    s = wb.create_sheet("Sensitivity")
    ppa_steps, capex_steps, grid = sens
    s["A1"] = "Levered equity IRR sensitivity"; s["A1"].font = Font(bold=True, size=12, color=NAVY)
    s["A2"] = "Rows: capex change   |   Columns: PPA price change ($/MWh)"
    s.cell(4, 1, "capex \\ PPA").font = Font(bold=True)
    for j, dp in enumerate(ppa_steps, start=2):
        c = s.cell(4, j, f"{dp:+d}"); c.font = bold_white; c.fill = hdr_fill; c.alignment = Alignment(horizontal="center")
    for i, dc in enumerate(capex_steps):
        rr = 5 + i
        lab = s.cell(rr, 1, f"{dc:+.0%}"); lab.font = bold_white; lab.fill = hdr_fill
        for j, val in enumerate(grid[i], start=2):
            c = s.cell(rr, j, val); c.number_format = "0.0%"; c.border = border
            if ppa_steps[j - 2] == 0 and dc == 0:
                c.fill = sub_fill; c.font = Font(bold=True)
    s.column_dimensions["A"].width = 14

    wb.save(os.path.join(here, "solstice_model.xlsx"))

def write_memo(a, R, sens, here):
    os.makedirs(os.path.join(here, "out"), exist_ok=True)
    ppa_steps, capex_steps, grid = sens
    g = lambda i, j: f"{grid[i][j]:.1%}"
    md = f"""# Investment Committee memo: Project Solstice

**Asset:** {a['solar_mw']:.0f} MW utility-scale solar + {a['storage_mw']:.0f} MW / {a['storage_mw']*4:.0f} MWh storage, North America
**Ask:** approve ${R['equity0']/1e6:.1f}M equity commitment | **Total capex:** ${a['capex']/1e6:.0f}M ({a['debt_pct']:.0%} geared)

## Recommendation
Proceed. The project clears an unlevered IRR of **{R['unlev_irr']:.1%}** and a levered equity IRR of
**{R['equity_irr']:.1%}** ({R['moic']:.2f}x MOIC) on a {a['ppa_term']}-year contracted PPA, with a minimum
DSCR of **{R['min_dscr']:.2f}x** that holds the debt comfortably through the amortization period. Returns
sit in the band expected for contracted renewables, and the merchant tail past year {a['ppa_term']} is
upside, not underwriting.

## Returns summary
| Metric | Value |
|---|---|
| Unlevered project IRR | {R['unlev_irr']:.1%} |
| Unlevered NPV @ {a['disc_unlev']:.0%} | ${R['unlev_npv']/1e6:.1f}M |
| Levered equity IRR | {R['equity_irr']:.1%} |
| Equity MOIC | {R['moic']:.2f}x |
| Minimum DSCR | {R['min_dscr']:.2f}x |
| Equity cheque | ${R['equity0']/1e6:.1f}M |
| Debt drawn | ${R['debt0']/1e6:.1f}M |

## Key value drivers and risks
- **PPA price and term** are the dominant value driver: the contract underwrites ~{a['ppa_term']} years of
  cash flow, after which the asset runs merchant. The sensitivity below isolates PPA price and capex.
- **Capex** is the other swing factor; procurement timing on modules and batteries moves the equity IRR.
- **Resource and degradation:** a {a['capacity_factor']:.0%} capacity factor declining {a['degradation']:.1%}/yr.
- **Refinancing / merchant tail:** years {a['ppa_term']+1}-{a['life']} are modelled at a conservative
  ${a['merchant_price']:.0f}/MWh; a contracted extension is upside.

## Levered equity IRR sensitivity (PPA price vs capex)
| capex \\ PPA | {ppa_steps[0]:+d} | {ppa_steps[1]:+d} | {ppa_steps[2]:+d} | {ppa_steps[3]:+d} | {ppa_steps[4]:+d} |
|---|---|---|---|---|---|
| {capex_steps[0]:+.0%} | {g(0,0)} | {g(0,1)} | {g(0,2)} | {g(0,3)} | {g(0,4)} |
| {capex_steps[1]:+.0%} | {g(1,0)} | {g(1,1)} | {g(1,2)} | {g(1,3)} | {g(1,4)} |
| {capex_steps[2]:+.0%} | {g(2,0)} | {g(2,1)} | **{g(2,2)}** | {g(2,3)} | {g(2,4)} |
| {capex_steps[3]:+.0%} | {g(3,0)} | {g(3,1)} | {g(3,2)} | {g(3,3)} | {g(3,4)} |
| {capex_steps[4]:+.0%} | {g(4,0)} | {g(4,1)} | {g(4,2)} | {g(4,3)} | {g(4,4)} |

## Decarbonization impact
Displacing grid power at {a['grid_ef']:.2f} t CO2/MWh, the asset abates about **{R['co2_per_yr']/1e3:.0f} kt CO2/yr**
and **{R['lifetime_co2']/1e6:.1f} Mt over its life**, an implied **${R['abate_cost']:.0f}/tonne** on capex alone.
This maps to a transition-investment framework that scores both return and abatement.

*Figures are illustrative, built on transparent assumptions (see the Assumptions tab of the model). The
method, not the specific numbers, is the point.*
"""
    with open(os.path.join(here, "out", "ic_memo.md"), "w", encoding="utf-8") as f:
        f.write(md)
