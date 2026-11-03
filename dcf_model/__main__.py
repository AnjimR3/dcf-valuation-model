"""Command line entry point: python -m dcf_model"""

from __future__ import annotations

import argparse

from .model import load_assumptions, value_scenario, wacc
from .workbook import build_workbook


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the DCF valuation workbook.")
    parser.add_argument("--assumptions", default="data/assumptions.json")
    parser.add_argument("--output", default="output/dcf_model.xlsx")
    args = parser.parse_args()

    data = load_assumptions(args.assumptions)
    path = build_workbook(data, args.output)
    print(f"WACC: {wacc(data):.2%}")
    for name in data["scenarios"]:
        result = value_scenario(data, name)
        print(f"{name}: EV ${result.enterprise_value:,.1f}m, ${result.value_per_share:,.2f} per share, terminal share {result.terminal_share:.1%}")
    print(f"Workbook written to {path}")


if __name__ == "__main__":
    main()
