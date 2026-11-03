"""Independent Python version of the valuation, used to check the workbook."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


def load_assumptions(path: str | Path) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def wacc(assumptions: dict) -> float:
    """Weighted average cost of capital, using the CAPM for cost of equity."""
    inputs = assumptions["wacc"]
    equity_cost = inputs["risk_free"] + inputs["beta"] * inputs["equity_premium"]
    debt_cost = inputs["cost_of_debt"] * (1 - assumptions["tax_rate"])
    return equity_cost * (1 - inputs["debt_weight"]) + debt_cost * inputs["debt_weight"]


def free_cash_flows(assumptions: dict, scenario: str) -> list[float]:
    """Unlevered free cash flow for each forecast year."""
    plan = assumptions["scenarios"][scenario]
    prior_revenue = assumptions["history"]["revenue"][-1]
    flows = []
    for growth, margin in zip(plan["growth"], plan["margin"]):
        revenue = prior_revenue * (1 + growth)
        after_tax_profit = revenue * margin * (1 - assumptions["tax_rate"])
        depreciation = revenue * assumptions["da_pct"]
        capex = revenue * assumptions["capex_pct"]
        working_capital = (revenue - prior_revenue) * assumptions["nwc_pct"]
        flows.append(after_tax_profit + depreciation - capex - working_capital)
        prior_revenue = revenue
    return flows


@dataclass
class Valuation:
    enterprise_value: float
    equity_value: float
    value_per_share: float
    terminal_share: float


def value(flows: list[float], discount: float, growth: float, net_debt: float, shares: float) -> Valuation:
    """Discount year-end cash flows and a Gordon growth terminal value."""
    if discount <= growth:
        raise ValueError("The discount rate must be above terminal growth")
    present = sum(flow / (1 + discount) ** year for year, flow in enumerate(flows, start=1))
    terminal = flows[-1] * (1 + growth) / (discount - growth) / (1 + discount) ** len(flows)
    enterprise = present + terminal
    equity = enterprise - net_debt
    return Valuation(enterprise, equity, equity / shares, terminal / enterprise)


def value_scenario(assumptions: dict, scenario: str) -> Valuation:
    growth = assumptions["scenarios"][scenario]["terminal_growth"]
    return value(
        free_cash_flows(assumptions, scenario), wacc(assumptions), growth,
        assumptions["net_debt"], assumptions["shares"],
    )
