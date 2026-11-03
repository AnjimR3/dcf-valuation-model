import pytest

from dcf_model.model import free_cash_flows, load_assumptions, value, value_scenario, wacc

DATA = load_assumptions("data/assumptions.json")


def test_wacc_formula():
    assert wacc(DATA) == pytest.approx(0.10 * 0.8 + 0.04875 * 0.2)


def test_first_year_cash_flow_by_hand():
    revenue = 980 * 1.09
    expected = revenue * 0.16 * 0.75 + revenue * 0.04 - revenue * 0.05 - (revenue - 980) * 0.12
    assert free_cash_flows(DATA, "Base")[0] == pytest.approx(expected)


def test_single_period_value_by_hand():
    result = value([100.0], 0.10, 0.0, 50.0, 10.0)
    assert result.enterprise_value == pytest.approx(100 / 1.1 + 1000 / 1.1)
    assert result.value_per_share == pytest.approx((1000 - 50) / 10)


def test_discount_rate_must_exceed_growth():
    with pytest.raises(ValueError):
        value([100.0], 0.03, 0.03, 0.0, 1.0)


def test_upside_is_worth_more_than_base_and_base_more_than_downside():
    results = [value_scenario(DATA, name).value_per_share for name in ("Upside", "Base", "Downside")]
    assert results == sorted(results, reverse=True)
