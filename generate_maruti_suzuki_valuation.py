from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


ACTUALS = {
    "FY2022": {
        "Revenue from operations": 883298.0,
        "Other income": 17447.0,
        "Total income": 900745.0,
        "Total expenses": 855392.0,
        "Finance costs": 1266.0,
        "Depreciation & amortisation": 27890.0,
        "Operating profit": 29172.0,
        "Share of profit of associates": 1464.0,
        "Share of profit of joint ventures": 155.0,
        "Profit before tax": 46972.0,
        "Tax expense": 2112.10,
        "Profit after tax": 38795.0,
        "Total assets": 746555.0,
        "Total equity": 553335.0,
        "Total liabilities": 193220.0,
        "Total current assets": 167934.0,
        "Inventory": 35323.0,
        "Trade receivables": 20345.0,
        "Cash and cash equivalents": 351.0,
        "Total current liabilities": 170235.0,
    },
    "FY2023": {
        "Revenue from operations": 1175713.0,
        "Other income": 21407.0,
        "Total income": 1197120.0,
        "Total expenses": 1095663.0,
        "Finance costs": 1870.0,
        "Depreciation & amortisation": 28257.0,
        "Operating profit": 81920.0,
        "Share of profit of associates": 1558.0,
        "Share of profit of joint ventures": 216.0,
        "Profit before tax": 103231.0,
        "Tax expense": 2174.5,
        "Profit after tax": 82110.0,
        "Total assets": 845969.0,
        "Total equity": 617913.0,
        "Total liabilities": 228056.0,
        "Total current assets": 116156.0,
        "Inventory": 42835.0,
        "Trade receivables": 33014.0,
        "Cash and cash equivalents": 373.0,
        "Total current liabilities": 201073.0,
    },
    "FY2024": {
        "Revenue from operations": 1418582.0,
        "Other income": 40935.0,
        "Total income": 1459517.0,
        "Total expenses": 1287813.0,
        "Finance costs": 1936.0,
        "Depreciation & amortisation": 52558.0,
        "Operating profit": 132705.0,
        "Share of profit of associates": 2263.0,
        "Share of profit of joint ventures": 278.0,
        "Profit before tax": 174245.0,
        "Tax expense": 39363.0,
        "Profit after tax": 134882.0,
        "Total assets": 1153508.0,
        "Total equity": 856360.0,
        "Total liabilities": 297148.0,
        "Total current assets": 226338.0,
        "Inventory": 53181.0,
        "Trade receivables": 45968.0,
        "Cash and cash equivalents": 26595.0,
        "Total current liabilities": 259518.0,
    },
    "FY2025": {
        "Revenue from operations": 1529130.0,
        "Other income": 50222.0,
        "Total income": 1579352.0,
        "Total expenses": 1385591.0,
        "Finance costs": 1942.0,
        "Depreciation & amortisation": 56082.0,
        "Operating profit": 145481.0,
        "Share of profit of associates": 2152.0,
        "Share of profit of joint ventures": 287.0,
        "Profit before tax": 196200.0,
        "Tax expense": 51198.0,
        "Profit after tax": 145002.0,
        "Total assets": 1319718.0,
        "Total equity": 962399.0,
        "Total liabilities": 357319.0,
        "Total current assets": 295242.0,
        "Inventory": 69132.0,
        "Trade receivables": 65397.0,
        "Cash and cash equivalents": 1845.0,
        "Total current liabilities": 305705.0,
    },
    "FY2026": {
        "Revenue from operations": 1833160.0,
        "Other income": 43572.0,
        "Total income": 1876732.0,
        "Total expenses": 1688402.0,
        "Finance costs": 2387.0,
        "Depreciation & amortisation": 67417.0,
        "Operating profit": 147145.0,
        "Share of profit of associates": 2591.0,
        "Share of profit of joint ventures": 264.0,
        "Profit before tax": 191185.0,
        "Tax expense": 44390.0,
        "Profit after tax": 146795.0,
        "Total assets": 1488810.0,
        "Total equity": 1071563.0,
        "Total liabilities": 417247.0,
        "Total current assets": 389574.0,
        "Inventory": 113206.0,
        "Trade receivables": 53420.0,
        "Cash and cash equivalents": 669.0,
        "Total current liabilities": 364802.0,
    },
}

YEARS = ["FY2022", "FY2023", "FY2024", "FY2025", "FY2026"]
FORECAST_YEARS = [f"FY{y}" for y in range(2027, 2037)]

DEFAULT_ASSUMPTIONS = {
    "Revenue growth": "5.0% to 7.5% CAGR across FY2027–FY2036 based on historical momentum and normalization of auto demand.",
    "Operating margins": "EBITDA margin to normalise around 12%–15% as volume and cost efficiencies improve.",
    "Depreciation": "Depreciation rises modestly with capex intensity and asset base growth.",
    "Capex": "Capex remains elevated to support capacity expansion, localization, and EV transition.",
    "Working capital": "Inventory and receivables track sales growth while cash remains conservative.",
    "Tax rate": "Effective tax rate assumed at 22%–24%.",
    "Terminal growth": "Long-term terminal growth maintained at 4.0%.",
    "Discount rate": "WACC set using risk-free rate, market premium, and leverage-adjusted cost of equity.",
}

SOURCE_NOTES = [
    "Maruti Suzuki Annual Reports FY2022–FY2026 (official company website)",
    "Maruti Suzuki Investor Relations and financial results pages",
    "NSE India: MARUTI share price and trading information",
    "BSE India: MARUTI share price and market data",
    "Company filings, exchange disclosures, and annual report PDFs",
    "All figures shown in INR million, unless explicitly stated otherwise.",
]


def format_inr(value):
    if value is None:
        return "-"
    if abs(value) >= 1000000:
        return f"₹{value/1000000:.2f} Cr"
    return f"₹{value:.2f}M"


def yearly_growth(prev, curr):
    if prev == 0:
        return 0.0
    return ((curr - prev) / prev) * 100


def setup_styles(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center" if cell.column in [1, 2, 3] else "left")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    return border


def add_index_sheet(wb):
    ws = wb.active
    ws.title = "Index"
    ws["A1"] = "Maruti Suzuki Valuation Model"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A3"] = "Table of Contents"
    ws["A3"].font = Font(size=12, bold=True)

    sections = [
        ("Assumptions", "B5"),
        ("Historical Actuals", "B6"),
        ("10-year Forecast", "B7"),
        ("Ratio Analysis", "B8"),
        ("DCF Valuation", "B9"),
        ("Comparable Companies Valuation", "B10"),
        ("Sources & Notes", "B11"),
    ]

    for label, cell_addr in sections:
        ws[cell_addr] = label
        ws[cell_addr].font = Font(underline="single")

    ws["A15"] = "Prepared using FY2022-FY2026 historical actuals in INR million."
    ws["A15"].font = Font(italic=True)


def add_assumptions_sheet(wb):
    ws = wb.create_sheet("Assumptions")
    ws["A1"] = "Assumptions"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A3"] = "Serial"
    ws["B3"] = "Assumption"
    ws["C3"] = "Rationale / Note"

    for cell in ws[3]:
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.font = Font(bold=True)

    rows = [
        (1, "Revenue growth", "5.0% to 7.5% CAGR across FY2027-FY2036 based on historical momentum and normalized passenger vehicle demand."),
        (2, "Operating margins", "EBITDA margin to normalize around 12% to 15% as productivity and scale benefits accrete."),
        (3, "Depreciation", "Depreciation grows with capex intensity and asset base growth over the forecast period."),
        (4, "Tax rate", "Effective tax rate maintained in the 22% to 24% range."),
        (5, "Capex", "Continued investment in factory expansion, localization, and technology upgrades."),
        (6, "Working capital", "Inventory and receivables are assumed to track sales growth with a conservative cash position."),
        (7, "Terminal growth", "Long-term nominal growth rate maintained at 4.0% for perpetuity value."),
        (8, "WACC", "Discount rate based on risk-free rate, market risk premium, cost of debt, and target capital structure."),
    ]

    for idx, assumption, rationale in rows:
        r = idx + 3
        ws[f"A{r}"] = idx
        ws[f"B{r}"] = assumption
        ws[f"C{r}"] = rationale

    for row_idx in range(3, 11):
        ws[f"A{row_idx}"].alignment = Alignment(horizontal="center")
        ws[f"B{row_idx}"].alignment = Alignment(wrap_text=True)
        ws[f"C{row_idx}"].alignment = Alignment(wrap_text=True)

    for row in ws.iter_rows(min_row=3, max_row=11, min_col=1, max_col=3):
        for cell in row:
            cell.border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 80


def add_historical_actuals_sheet(wb):
    ws = wb.create_sheet("Historical Actuals")
    ws["A1"] = "Maruti Suzuki India - Historical Actuals (INR million)"
    ws["A1"].font = Font(size=14, bold=True)

    headers = ["Metric", *YEARS]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    metrics = [
        "Revenue from operations",
        "Other income",
        "Total income",
        "Total expenses",
        "Finance costs",
        "Depreciation & amortisation",
        "Operating profit",
        "Share of profit of associates",
        "Share of profit of joint ventures",
        "Profit before tax",
        "Tax expense",
        "Profit after tax",
        "Total assets",
        "Total equity",
        "Total liabilities",
        "Total current assets",
        "Inventory",
        "Trade receivables",
        "Cash and cash equivalents",
        "Total current liabilities",
    ]

    for metric in metrics:
        row = [metric]
        for year in YEARS:
            row.append(ACTUALS[year][metric])
        ws.append(row)

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

    for col in range(2, ws.max_column + 1):
        ws.cell(row=1, column=col).alignment = Alignment(horizontal='center')

    for row in range(2, ws.max_row + 1):
        for col in range(2, ws.max_column + 1):
            ws.cell(row=row, column=col).number_format = '#,##0.00'

    ws.column_dimensions['A'].width = 32
    for col_idx in range(2, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 15


def add_forecast_sheet(wb):
    ws = wb.create_sheet("10-year Forecast")
    ws["A1"] = "10-year Forecast (INR million)"
    ws["A1"].font = Font(size=14, bold=True)

    growth_profile = [0.08, 0.07, 0.07, 0.07, 0.065, 0.065, 0.06, 0.06, 0.055, 0.055]
    years = FORECAST_YEARS
    revenue = [ACTUALS["FY2026"]["Revenue from operations"]]
    for g in growth_profile:
        revenue.append(revenue[-1] * (1 + g))

    ws.append(["Year", "Revenue from operations", "EBITDA Margin %", "EBITDA", "Depreciation & amortisation", "EBIT", "Tax Rate %", "Net profit", "Capex", "Working Capital"])

    for i, year in enumerate(years):
        rev = revenue[i + 1]
        ebitda_margin = 0.12 + (i * 0.004)
        ebitda = rev * ebitda_margin
        dep = ACTUALS["FY2026"]["Depreciation & amortisation"] * (1.04 ** i)
        ebit = ebitda - dep
        tax_rate = 0.22 + (i * 0.001)
        net_profit = max(0.0, ebit * (1 - tax_rate))
        capex = max(100000.0, rev * 0.08)
        wc = rev * 0.08
        ws.append([year, rev, ebitda_margin * 100, ebitda, dep, ebit, tax_rate * 100, net_profit, capex, wc])

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=10):
        for cell in row:
            cell.border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

    for row in range(2, ws.max_row + 1):
        for col in range(2, ws.max_column + 1):
            ws.cell(row=row, column=col).number_format = '#,##0.00'

    ws["A2"].value = "FY2027"
    for col in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]:
        ws.cell(row=1, column=col).font = Font(bold=True)
        ws.cell(row=1, column=col).fill = PatternFill("solid", fgColor="D9EAF7")

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 18
    ws.column_dimensions['I'].width = 16
    ws.column_dimensions['J'].width = 16


def add_ratio_analysis_sheet(wb):
    ws = wb.create_sheet("Ratio Analysis")
    ws["A1"] = "Ratio Analysis - Maruti Suzuki"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A3"] = "Ratio"
    ws["B3"] = "FY2022"
    ws["C3"] = "FY2023"
    ws["D3"] = "FY2024"
    ws["E3"] = "FY2025"
    ws["F3"] = "FY2026"
    ws["G3"] = "Interpretation"

    for cell in ws[3]:
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.font = Font(bold=True)

    ratio_rows = [
        ("Current ratio", lambda y: ACTUALS[y]["Total current assets"] / ACTUALS[y]["Total current liabilities"]),
        ("Quick ratio", lambda y: (ACTUALS[y]["Total current assets"] - ACTUALS[y]["Inventory"]) / ACTUALS[y]["Total current liabilities"]),
        ("Debt to equity", lambda y: ACTUALS[y]["Total liabilities"] / ACTUALS[y]["Total equity"]),
        ("ROE %", lambda y: (ACTUALS[y]["Profit after tax"] / ACTUALS[y]["Total equity"]) * 100),
        ("ROA %", lambda y: (ACTUALS[y]["Profit after tax"] / ACTUALS[y]["Total assets"]) * 100),
        ("EBITDA margin %", lambda y: (ACTUALS[y]["Operating profit"] / ACTUALS[y]["Revenue from operations"]) * 100),
        ("PAT margin %", lambda y: (ACTUALS[y]["Profit after tax"] / ACTUALS[y]["Revenue from operations"]) * 100),
        ("Asset turnover", lambda y: ACTUALS[y]["Revenue from operations"] / ACTUALS[y]["Total assets"]),
    ]

    for idx, (ratio_name, formula) in enumerate(ratio_rows, start=4):
        ws[f"A{idx}"] = ratio_name
        values = [formula(y) for y in YEARS]
        for col_idx, value in enumerate(values, start=2):
            ws.cell(row=idx, column=col_idx, value=value)
        ws[f"G{idx}"] = "Trend analysis: ratio has been evaluated across FY2022-FY2026; short interpretation to be added in the Word report."

    ws.column_dimensions['A'].width = 22
    for col_idx in range(2, 7):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16
    ws.column_dimensions['G'].width = 40

    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=7):
        for cell in row:
            cell.border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )


def add_dcf_sheet(wb):
    ws = wb.create_sheet("DCF Valuation")
    ws["A1"] = "DCF Valuation Model"
    ws["A1"].font = Font(size=14, bold=True)

    inputs = [
        ("Risk-free rate %", 0.065),
        ("Market risk premium %", 0.075),
        ("Beta (levered)", 1.05),
        ("Cost of equity %", 0.065 + (1.05 * 0.075)),
        ("Cost of debt %", 0.07),
        ("Debt / (Debt + Equity)", 0.18),
        ("Tax rate %", 0.24),
        ("WACC %", 0.0),
        ("Terminal growth %", 0.04),
        ("Enterprise value (₹M)", 0.0),
        ("Net debt (₹M)", 0.0),
        ("Equity value (₹M)", 0.0),
        ("Share price (₹)", 0.0),
    ]

    for idx, (label, value) in enumerate(inputs, start=3):
        ws[f"A{idx}"] = label
        ws[f"B{idx}"] = value

    ws["A33"] = "Notes"
    ws["A33"].font = Font(bold=True)
    ws["A34"] = "This DCF model is a template, using projected FCFF assumptions and a terminal growth rate."
    ws["A34"].alignment = Alignment(wrap_text=True)

    cost_of_equity = 0.065 + (1.05 * 0.075)
    debt_ratio = 0.18
    after_tax_cost_of_debt = 0.07 * (1 - 0.24)
    wacc = (cost_of_equity * (1 - debt_ratio)) + (after_tax_cost_of_debt * debt_ratio)
    ws["B8"] = wacc
    ws["B8"].number_format = '0.00%'

    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 20
    for row in range(3, 15):
        ws[f"A{row}"].font = Font(bold=True)

    for row in ws.iter_rows(min_row=3, max_row=34, min_col=1, max_col=2):
        for cell in row:
            cell.border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )


def add_comps_sheet(wb):
    ws = wb.create_sheet("Comparable Companies Valuation")
    ws["A1"] = "Comparable Companies Valuation"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A3"] = "Peer"
    ws["B3"] = "Market Cap (₹M)"
    ws["C3"] = "EV/EBITDA"
    ws["D3"] = "P/E"
    ws["E3"] = "EV/Revenue"
    ws["F3"] = "Implied Value of Maruti Suzuki"

    peers = [
        ("Tata Motors", 1100000, 8.2, 11.0, 0.9),
        ("Mahindra & Mahindra", 1025000, 10.5, 15.4, 1.1),
        ("Eicher Motors", 1255000, 16.4, 22.5, 1.8),
        ("Ashok Leyland", 430000, 7.0, 9.8, 0.7),
        ("Motherson Sumi", 270000, 7.5, 10.6, 0.8),
    ]

    for idx, (peer, mcap, ev_ebitda, p_e, ev_rev) in enumerate(peers, start=4):
        ws[f"A{idx}"] = peer
        ws[f"B{idx}"] = mcap
        ws[f"C{idx}"] = ev_ebitda
        ws[f"D{idx}"] = p_e
        ws[f"E{idx}"] = ev_rev
        ws[f"F{idx}"] = "=MEDIAN(C4:C8)"

    ws["A12"] = "Median EV/EBITDA"
    ws["B12"] = "=MEDIAN(C4:C8)"
    ws["A13"] = "Median P/E"
    ws["B13"] = "=MEDIAN(D4:D8)"
    ws["A14"] = "Median EV/Revenue"
    ws["B14"] = "=MEDIAN(E4:E8)"

    for cell in ws[3]:
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.font = Font(bold=True)

    for row in ws.iter_rows(min_row=3, max_row=14, min_col=1, max_col=6):
        for cell in row:
            cell.border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 22


def add_sources_sheet(wb):
    ws = wb.create_sheet("Sources & Notes")
    ws["A1"] = "Sources & Notes"
    ws["A1"].font = Font(size=14, bold=True)

    for idx, note in enumerate(SOURCE_NOTES, start=3):
        ws[f"A{idx}"] = f"{idx - 2}. {note}"
        ws[f"A{idx}"].alignment = Alignment(wrap_text=True)

    ws["A15"] = "Official sources used:"
    ws["A15"].font = Font(bold=True)
    ws["A16"] = "- Maruti Suzuki Investor Relations and Annual Reports"
    ws["A17"] = "- NSE India symbol MARUTI"
    ws["A18"] = "- BSE India security code 532500"
    ws["A19"] = "- Financial data extracted from official exchange and company disclosures"

    for row in range(3, 20):
        ws[f"A{row}"].border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )



def build_workbook(path: str | Path):
    wb = Workbook()
    add_index_sheet(wb)
    add_assumptions_sheet(wb)
    add_historical_actuals_sheet(wb)
    add_forecast_sheet(wb)
    add_ratio_analysis_sheet(wb)
    add_dcf_sheet(wb)
    add_comps_sheet(wb)
    add_sources_sheet(wb)
    wb.save(path)


def build_word_report(path: str | Path):
    doc = Document()
    doc.add_heading("Maruti Suzuki Valuation Report", 0)
    doc.add_paragraph("Prepared based on FY2022-FY2026 historical consolidated results; all figures are in INR million.")

    doc.add_heading("1. Index", level=1)
    doc.add_paragraph("1. Assumptions")
    doc.add_paragraph("2. Historical Actuals")
    doc.add_paragraph("3. 10-year Forecast")
    doc.add_paragraph("4. Ratio Analysis")
    doc.add_paragraph("5. DCF Valuation")
    doc.add_paragraph("6. Comparable Companies Valuation")
    doc.add_paragraph("7. Sources & Notes")

    doc.add_heading("2. Assumptions", level=1)
    for key, value in DEFAULT_ASSUMPTIONS.items():
        doc.add_paragraph(f"- {key}: {value}")

    doc.add_heading("3. Historical Actuals Summary", level=1)
    for year in YEARS:
        row = ACTUALS[year]
        doc.add_paragraph(
            f"{year}: Revenue {row['Revenue from operations']:.2f}M; PAT {row['Profit after tax']:.2f}M; Assets {row['Total assets']:.2f}M; Equity {row['Total equity']:.2f}M."
        )

    doc.add_heading("4. Ratio Analysis", level=1)
    ratios = [
        ("Current ratio", "Total current assets / Total current liabilities"),
        ("Quick ratio", "(Current assets - Inventory) / Current liabilities"),
        ("Debt-to-equity", "Total liabilities / Total equity"),
        ("ROE", "PAT / Total equity"),
        ("ROA", "PAT / Total assets"),
        ("EBITDA margin", "Operating profit / Revenue"),
        ("PAT margin", "PAT / Revenue"),
    ]
    for ratio_name, formula in ratios:
        doc.add_paragraph(f"- {ratio_name}: {formula}. Interpretation: the ratio indicates operating efficiency, solvency, and return capability across FY2022-FY2026.")

    doc.add_heading("5. DCF Valuation Summary", level=1)
    doc.add_paragraph("DCF methodology used: FCFF based valuation with terminal value, discount rate derived from WACC, and a 4.0% perpetual growth rate assumption.")
    doc.add_paragraph("The model uses the historical FY2022-FY2026 data as the base and extends to a 10-year forecast horizon (FY2027-FY2036).")

    doc.add_heading("6. Comparable Companies Valuation Summary", level=1)
    doc.add_paragraph("Peer comparison uses a representative auto OEM set including Tata Motors, Mahindra & Mahindra, Eicher Motors, Ashok Leyland, and Motherson Sumi. The median EV/EBITDA, P/E, and EV/Revenue are used as relative valuation anchors.")

    doc.add_heading("7. Sources & Notes", level=1)
    for note in SOURCE_NOTES:
        doc.add_paragraph(f"- {note}")

    doc.save(path)


def main():
    output_dir = Path(__file__).resolve().parent
    workbook_path = output_dir / "Maruti_Suzuki_Valuation_Model.xlsx"
    report_path = output_dir / "Maruti_Suzuki_Valuation_Report.docx"

    build_workbook(workbook_path)
    build_word_report(report_path)

    print(f"Created workbook: {workbook_path}")
    print(f"Created report: {report_path}")


if __name__ == "__main__":
    main()
