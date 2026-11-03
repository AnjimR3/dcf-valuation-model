"""Write the valuation as a formula-driven Excel model."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Arial"
NAVY = "1F3864"
LIGHT = "DCE6F1"
INPUT = "0000FF"
LINK = "008000"
MONEY = "$#,##0.0;($#,##0.0);-"
SHARE = "$#,##0.00;($#,##0.00);-"
PERCENT = "0.0%;(0.0%);-"
FORECAST_YEARS = 5


def font(**kwargs) -> Font:
    return Font(name=FONT, size=kwargs.pop("size", 10), **kwargs)


def put(ws, ref, value, number_format=None, color=None, bold=False, fill=None, align=None):
    cell = ws[ref]
    cell.value = value
    cell.font = font(color=color, bold=bold)
    if number_format:
        cell.number_format = number_format
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    if align:
        cell.alignment = Alignment(horizontal=align)
    return cell


def title(ws, heading, subtitle):
    ws["A1"] = heading
    ws["A1"].font = font(size=14, bold=True, color=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = font(italic=True, color="595959")
    ws.sheet_view.showGridLines = False


def header(ws, row, labels, start=1):
    for offset, label in enumerate(labels):
        cell = ws.cell(row, start + offset, label)
        cell.font = font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="left" if offset == 0 and start == 1 else "right", wrap_text=True)


def emphasise(ws, row, columns):
    for index in range(1, columns + 1):
        cell = ws.cell(row, index)
        color = cell.font.color.rgb if cell.font.color else None
        cell.font = font(bold=True, color=color[-6:] if isinstance(color, str) else "000000")
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        cell.border = Border(top=Side(style="thin", color=NAVY))


def landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def write_assumptions(ws, data) -> dict:
    title(ws, f"{data['company']}: Assumptions", "Fictional company, USD millions. Blue cells are inputs. Change the yellow cell to switch scenario.")
    names = list(data["scenarios"])
    put(ws, "A4", "Active scenario (1, 2 or 3)", bold=True)
    put(ws, "B4", 1, "0", INPUT, True, "FFFF00", "right")
    validation = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    ws.add_data_validation(validation)
    validation.add("B4")
    put(ws, "C4", f'=CHOOSE($B$4,"{names[0]}","{names[1]}","{names[2]}")', bold=True, align="right")

    header(ws, 6, ["General inputs", "Value"])
    general = [("Tax rate", data["tax_rate"], PERCENT), ("D&A, % of revenue", data["da_pct"], PERCENT),
               ("Capex, % of revenue", data["capex_pct"], PERCENT), ("Net working capital, % of revenue", data["nwc_pct"], PERCENT),
               ("Net debt", data["net_debt"], MONEY), ("Shares outstanding (millions)", data["shares"], "#,##0.0")]
    cells = {}
    for offset, (label, number, fmt) in enumerate(general):
        row = 7 + offset
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", number, fmt, INPUT)
        cells[label] = f"Assumptions!$B${row}"

    header(ws, 14, ["Cost of capital", "Value"])
    inputs = data["wacc"]
    lines = [("Risk-free rate", inputs["risk_free"], PERCENT, True), ("Beta", inputs["beta"], "0.00", True),
             ("Equity risk premium", inputs["equity_premium"], PERCENT, True), ("Pre-tax cost of debt", inputs["cost_of_debt"], PERCENT, True),
             ("Debt weight", inputs["debt_weight"], PERCENT, True), ("Cost of equity", "=B15+B16*B17", PERCENT, False),
             ("After-tax cost of debt", "=B18*(1-B7)", PERCENT, False), ("WACC", "=B20*(1-B19)+B21*B19", PERCENT, False)]
    for offset, (label, number, fmt, is_input) in enumerate(lines):
        row = 15 + offset
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", number, fmt, INPUT if is_input else None)
    emphasise(ws, 22, 2)

    years = [data["history"]["years"][-1] + year for year in range(1, FORECAST_YEARS + 1)]
    row = 24
    positions = {}
    for number, (name, plan) in enumerate(data["scenarios"].items(), start=1):
        header(ws, row, [f"{number}. {name}", *[f"{year}E" for year in years]])
        put(ws, f"A{row + 1}", "Revenue growth")
        put(ws, f"A{row + 2}", "EBIT margin")
        for year in range(FORECAST_YEARS):
            letter = col(year + 2)
            put(ws, f"{letter}{row + 1}", plan["growth"][year], PERCENT, INPUT)
            put(ws, f"{letter}{row + 2}", plan["margin"][year], PERCENT, INPUT)
        put(ws, f"A{row + 3}", "Terminal growth")
        put(ws, f"B{row + 3}", plan["terminal_growth"], PERCENT, INPUT)
        positions[name] = row
        row += 5

    ws.column_dimensions["A"].width = 36
    for index in range(2, 8):
        ws.column_dimensions[col(index)].width = 12
    landscape(ws)
    return {"general": cells, "scenarios": positions, "names": names}


def cash_flow_block(ws, row, data, refs, growth, margin, prefix=""):
    """Write revenue-to-cash-flow rows. growth and margin return a formula for a forecast column."""
    last_history = 3 + 1  # column D holds the last historical year
    general = refs["general"]
    labels = ["Revenue", "EBIT", "Taxes on EBIT", "NOPAT", "D&A", "Capex", "Change in NWC", "Free cash flow"]
    rows = {label: row + offset for offset, label in enumerate(labels)}
    for label, position in rows.items():
        put(ws, f"A{position}", prefix + label)
    for index in range(FORECAST_YEARS):
        column = last_history + 1 + index
        c, p = col(column), col(column - 1)
        r = rows
        formulas = {
            "Revenue": f"={p}{r['Revenue']}*(1+{growth(index)})",
            "EBIT": f"={c}{r['Revenue']}*{margin(index)}",
            "Taxes on EBIT": f"={c}{r['EBIT']}*{general['Tax rate']}",
            "NOPAT": f"={c}{r['EBIT']}-{c}{r['Taxes on EBIT']}",
            "D&A": f"={c}{r['Revenue']}*{general['D&A, % of revenue']}",
            "Capex": f"={c}{r['Revenue']}*{general['Capex, % of revenue']}",
            "Change in NWC": f"=({c}{r['Revenue']}-{p}{r['Revenue']})*{general['Net working capital, % of revenue']}",
            "Free cash flow": f"={c}{r['NOPAT']}+{c}{r['D&A']}-{c}{r['Capex']}-{c}{r['Change in NWC']}",
        }
        for label, formula in formulas.items():
            put(ws, f"{c}{r[label]}", formula, MONEY)
    return rows


def write_forecast(ws, data, refs) -> dict:
    title(ws, f"{data['company']}: Forecast", "USD millions. Historical years are inputs, forecast years follow the active scenario.")
    put(ws, "A3", '="Active scenario: "&Assumptions!$C$4', bold=True)
    history = data["history"]
    years = [f"{year}A" for year in history["years"]]
    years += [f"{history['years'][-1] + year}E" for year in range(1, FORECAST_YEARS + 1)]
    header(ws, 4, ["USD millions", *years])

    starts = refs["scenarios"]
    names = refs["names"]

    def choice(offset, index):
        letter = col(index + 2)
        options = ",".join(f"Assumptions!{letter}${starts[name] + offset}" for name in names)
        return f"CHOOSE(Assumptions!$B$4,{options})"

    rows = cash_flow_block(ws, 5, data, refs, lambda i: choice(1, i), lambda i: choice(2, i))
    for index in range(3):
        c = col(index + 2)
        put(ws, f"{c}{rows['Revenue']}", history["revenue"][index], MONEY, INPUT)
        put(ws, f"{c}{rows['EBIT']}", history["ebit"][index], MONEY, INPUT)
    emphasise(ws, rows["Free cash flow"], 9)

    row = rows["Free cash flow"] + 2
    put(ws, f"A{row}", "Revenue growth")
    put(ws, f"A{row + 1}", "EBIT margin")
    put(ws, f"A{row + 2}", "Free cash flow margin")
    put(ws, f"A{row + 3}", "Year number")
    put(ws, f"A{row + 4}", "Discount factor")
    put(ws, f"A{row + 5}", "Present value of free cash flow")
    for column in range(2, 10):
        c, p = col(column), col(column - 1)
        if column > 2:
            put(ws, f"{c}{row}", f"={c}{rows['Revenue']}/{p}{rows['Revenue']}-1", PERCENT)
        put(ws, f"{c}{row + 1}", f"=IF({c}{rows['Revenue']}=0,0,{c}{rows['EBIT']}/{c}{rows['Revenue']})", PERCENT)
        if column > 4:
            put(ws, f"{c}{row + 2}", f"={c}{rows['Free cash flow']}/{c}{rows['Revenue']}", PERCENT)
            put(ws, f"{c}{row + 3}", column - 4, "0")
            put(ws, f"{c}{row + 4}", f"=1/(1+Assumptions!$B$22)^{c}{row + 3}", "0.000")
            put(ws, f"{c}{row + 5}", f"={c}{rows['Free cash flow']}*{c}{row + 4}", MONEY)
    ws.column_dimensions["A"].width = 32
    for index in range(2, 10):
        ws.column_dimensions[col(index)].width = 11
    ws.freeze_panes = "B5"
    landscape(ws)
    return {**rows, "year": row + 3, "pv": row + 5}


def write_scenarios(ws, data, refs, forecast) -> dict:
    title(ws, f"{data['company']}: Scenario Comparison", "All three scenarios are calculated at once, whichever one is active.")
    header(ws, 4, ["Value per share", "Enterprise value", "Equity value", "Terminal share of EV", "Value per share"], start=1)
    general = refs["general"]
    row = 11
    outputs = {}
    for name in refs["names"]:
        start = refs["scenarios"][name]
        put(ws, f"A{row}", f"{name} scenario", bold=True, color=NAVY)
        last = data["history"]["years"][-1]
        header(ws, row + 1, ["USD millions", "", "", f"{last}A", *[f"{last + y}E" for y in range(1, 6)]])
        put(ws, f"D{row + 2}", "=Forecast!D5", MONEY, LINK)
        rows = cash_flow_block(
            ws, row + 2, data, refs,
            lambda i, s=start: f"Assumptions!{col(i + 2)}${s + 1}",
            lambda i, s=start: f"Assumptions!{col(i + 2)}${s + 2}",
        )
        fcf = rows["Free cash flow"]
        emphasise(ws, fcf, 9)
        wacc = "Assumptions!$B$22"
        growth = f"Assumptions!$B${start + 3}"
        flows = f"$E${fcf}:$I${fcf}"
        pv = f"SUMPRODUCT({flows}/(1+{wacc})^Forecast!$E${forecast['year']}:$I${forecast['year']})"
        terminal = f"$I${fcf}*(1+{growth})/({wacc}-{growth})/(1+{wacc})^5"
        outputs[name] = {"ev": f"={pv}+{terminal}", "terminal": terminal}
        row = fcf + 2
    for offset, name in enumerate(refs["names"]):
        target = 5 + offset
        put(ws, f"A{target}", name)
        put(ws, f"B{target}", outputs[name]["ev"], MONEY)
        put(ws, f"C{target}", f"=B{target}-{general['Net debt']}", MONEY)
        put(ws, f"D{target}", f"=({outputs[name]['terminal']})/B{target}", PERCENT)
        put(ws, f"E{target}", f"=C{target}/{general['Shares outstanding (millions)']}", SHARE, bold=True)
    ws.column_dimensions["A"].width = 26
    for index in range(2, 10):
        ws.column_dimensions[col(index)].width = 13
    ws["A4"].value = "Scenario"

    chart = BarChart()
    chart.type = "bar"
    chart.title = "Value per share by scenario (USD)"
    chart.add_data(Reference(ws, min_col=5, min_row=5, max_row=7), titles_from_data=False)
    chart.set_categories(Reference(ws, min_col=1, min_row=5, max_row=7))
    chart.legend = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.scaling.orientation = "maxMin"
    chart.y_axis.numFmt = "$#,##0"
    chart.y_axis.majorGridlines = None
    chart.series[0].graphicalProperties.solidFill = NAVY
    chart.height, chart.width = 5.5, 15
    ws.add_chart(chart, "K4")
    landscape(ws)
    return {"first": 5}


def write_valuation(ws, data, refs, forecast) -> None:
    title(ws, f"{data['company']}: DCF Valuation", "Fictional company, USD millions except per share. Not investment advice.")
    put(ws, "A3", '="Active scenario: "&Assumptions!$C$4', bold=True)
    names = refs["names"]
    starts = refs["scenarios"]
    growth_options = ",".join(f"Assumptions!$B${starts[name] + 3}" for name in names)
    fcf = forecast["Free cash flow"]
    lines = [
        ("WACC", "=Assumptions!$B$22", PERCENT, LINK),
        ("Terminal growth", f"=CHOOSE(Assumptions!$B$4,{growth_options})", PERCENT, LINK),
        ("Sum of present values, 2026E to 2030E", f"=SUM(Forecast!$E${forecast['pv']}:$I${forecast['pv']})", MONEY, LINK),
        ("Terminal value at 2030E", f"=Forecast!$I${fcf}*(1+B6)/(B5-B6)", MONEY, None),
        ("Present value of terminal value", "=B8/(1+B5)^5", MONEY, None),
        ("Enterprise value", "=B7+B9", MONEY, None),
        ("Less: net debt", f"={refs['general']['Net debt']}", MONEY, LINK),
        ("Equity value", "=B10-B11", MONEY, None),
        ("Shares outstanding (millions)", f"={refs['general']['Shares outstanding (millions)']}", "#,##0.0", LINK),
        ("Value per share", "=B12/B13", SHARE, None),
    ]
    header(ws, 4, ["Valuation", "Value"])
    for offset, (label, formula, fmt, color) in enumerate(lines):
        row = 5 + offset
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", formula, fmt, color)
    emphasise(ws, 10, 2)
    emphasise(ws, 12, 2)
    emphasise(ws, 14, 2)

    header(ws, 16, ["Reasonableness checks", "Value"])
    ebitda = f"(Forecast!$I${forecast['EBIT']}+Forecast!$I${forecast['D&A']})"
    checks = [
        ("Terminal value as % of enterprise value", "=B9/B10", PERCENT),
        ("Implied EV / 2026E EBITDA", f"=B10/(Forecast!$E${forecast['EBIT']}+Forecast!$E${forecast['D&A']})", "0.0x"),
        ("Implied terminal EV / EBITDA", f"=B8/{ebitda}", "0.0x"),
        ("WACC above terminal growth", '=IF(B5>B6,"OK","CHECK")', None),
        ("Active scenario matches comparison", '=IF(ABS(B14-INDEX(Scenarios!$E$5:$E$7,Assumptions!$B$4))<0.0001,"OK","CHECK")', None),
    ]
    for offset, (label, formula, fmt) in enumerate(checks):
        row = 17 + offset
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", formula, fmt, align="right")
    ws.conditional_formatting.add("B20:B21", CellIsRule(operator="equal", formula=['"OK"'], font=Font(color="007A33", bold=True)))
    ws.conditional_formatting.add("B20:B21", CellIsRule(operator="equal", formula=['"CHECK"'], font=Font(color="C00000", bold=True)))
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 14
    landscape(ws)


def write_sensitivity(ws, data, refs, forecast) -> None:
    title(ws, f"{data['company']}: Sensitivity", "Value per share in USD for the active scenario. Change the blue steps to widen the grid.")
    put(ws, "A3", '="Active scenario: "&Assumptions!$C$4', bold=True)
    put(ws, "A4", "WACC step")
    put(ws, "B4", 0.005, PERCENT, INPUT)
    put(ws, "A5", "Terminal growth step")
    put(ws, "B5", 0.005, PERCENT, INPUT)
    put(ws, "A7", "WACC (down), terminal growth (across)", bold=True)
    for index, step in enumerate(range(-2, 3)):
        cell = put(ws, f"{col(index + 2)}7", f"=Valuation!$B$6+({step})*$B$5", PERCENT, bold=True, fill=LIGHT)
        cell.alignment = Alignment(horizontal="right")
    fcf = forecast["Free cash flow"]
    flows = f"Forecast!$E${fcf}:$I${fcf}"
    years = f"Forecast!$E${forecast['year']}:$I${forecast['year']}"
    net_debt = refs["general"]["Net debt"]
    shares = refs["general"]["Shares outstanding (millions)"]
    for row_index, step in enumerate(range(-2, 3)):
        row = 8 + row_index
        put(ws, f"A{row}", f"=Valuation!$B$5+({step})*$B$4", "0.00%", bold=True, fill=LIGHT)
        for column_index in range(5):
            c = col(column_index + 2)
            wacc, growth = f"$A{row}", f"{c}$7"
            formula = (
                f"=(SUMPRODUCT({flows}/(1+{wacc})^{years})"
                f"+Forecast!$I${fcf}*(1+{growth})/({wacc}-{growth})/(1+{wacc})^5-{net_debt})/{shares}"
            )
            put(ws, f"{c}{row}", formula, SHARE)
    ws.conditional_formatting.add("B8:F12", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="9DC3E6"))
    ws["D10"].font = font(bold=True)
    ws.column_dimensions["A"].width = 38
    for index in range(2, 7):
        ws.column_dimensions[col(index)].width = 11
    landscape(ws)


def build_workbook(data: dict, output: str | Path, scenario: int = 1) -> Path:
    workbook = Workbook()
    valuation = workbook.active
    valuation.title = "Valuation"
    assumptions = workbook.create_sheet("Assumptions")
    forecast_sheet = workbook.create_sheet("Forecast")
    scenarios = workbook.create_sheet("Scenarios")
    sensitivity = workbook.create_sheet("Sensitivity")

    refs = write_assumptions(assumptions, data)
    assumptions["B4"] = scenario
    forecast = write_forecast(forecast_sheet, data, refs)
    write_scenarios(scenarios, data, refs, forecast)
    write_valuation(valuation, data, refs, forecast)
    write_sensitivity(sensitivity, data, refs, forecast)

    workbook.calculation.fullCalcOnLoad = True
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return output
