"""Recalculate the workbook and compare it with the independent Python model."""

import shutil
import subprocess

import pytest
from openpyxl import load_workbook

from dcf_model.model import free_cash_flows, load_assumptions, value, value_scenario, wacc
from dcf_model.workbook import build_workbook

DATA = load_assumptions("data/assumptions.json")
NAMES = list(DATA["scenarios"])
OFFICE = shutil.which("soffice") or shutil.which("libreoffice")
pytestmark = pytest.mark.skipif(OFFICE is None, reason="LibreOffice is needed to recalculate formulas")


@pytest.fixture(scope="module", params=[1, 2, 3])
def solved(request, tmp_path_factory):
    folder = tmp_path_factory.mktemp(f"scenario{request.param}")
    built = build_workbook(DATA, folder / "model.xlsx", request.param)
    subprocess.run(
        [OFFICE, "--headless", f"-env:UserInstallation=file://{folder}/profile",
         "--convert-to", "xlsx", "--outdir", str(folder / "out"), str(built)],
        check=True, capture_output=True, timeout=180,
    )
    return request.param, load_workbook(folder / "out/model.xlsx", data_only=True)


def test_no_formula_errors(solved):
    _, workbook = solved
    for sheet in workbook:
        for row in sheet.iter_rows():
            for cell in row:
                assert not (isinstance(cell.value, str) and cell.value.startswith("#")), f"{sheet.title}!{cell.coordinate}"


def test_active_scenario_matches_python(solved):
    number, workbook = solved
    expected = value_scenario(DATA, NAMES[number - 1])
    sheet = workbook["Valuation"]
    assert sheet["B5"].value == pytest.approx(wacc(DATA))
    assert sheet["B10"].value == pytest.approx(expected.enterprise_value)
    assert sheet["B14"].value == pytest.approx(expected.value_per_share)
    assert sheet["B17"].value == pytest.approx(expected.terminal_share)
    assert sheet["B20"].value == "OK" and sheet["B21"].value == "OK"


def test_scenario_table_matches_python(solved):
    _, workbook = solved
    sheet = workbook["Scenarios"]
    for row, name in enumerate(NAMES, start=5):
        assert sheet[f"A{row}"].value == name
        assert sheet[f"E{row}"].value == pytest.approx(value_scenario(DATA, name).value_per_share)


def test_sensitivity_grid_matches_python(solved):
    number, workbook = solved
    name = NAMES[number - 1]
    sheet = workbook["Sensitivity"]
    flows = free_cash_flows(DATA, name)
    centre = value_scenario(DATA, name).value_per_share
    assert sheet["D10"].value == pytest.approx(centre)
    for row in range(8, 13):
        for column in "BCDEF":
            expected = value(flows, sheet[f"A{row}"].value, sheet[f"{column}7"].value, DATA["net_debt"], DATA["shares"])
            assert sheet[f"{column}{row}"].value == pytest.approx(expected.value_per_share)
