import pytest
import pandas as pd
from uhxl import UhExcelFile

DATA = "tests/data/test.xlsx"

HIDDEN_SHEETS = ["hidden1", "hidden2"]
VISIBLE_SHEETS = ["sheet1", "sheet2"]
ALL_SHEETS = HIDDEN_SHEETS + VISIBLE_SHEETS

FILE = UhExcelFile(DATA, hide_sheets=False)
UHFILE = UhExcelFile(DATA, hide_sheets=True)


def test_read_all():
    df = pd.read_excel(FILE, sheet_name=None)
    assert isinstance(df, dict)
    assert all(s in df for s in ALL_SHEETS)


def test_read_all_visible():
    df = pd.read_excel(UHFILE, sheet_name=None)
    assert isinstance(df, dict)
    assert all(s in df for s in VISIBLE_SHEETS)
    assert not any(s in df for s in HIDDEN_SHEETS)


def test_read_sheet():
    df = pd.read_excel(FILE, sheet_name=HIDDEN_SHEETS[0])
    assert isinstance(df, pd.DataFrame)


def test_read_skipped_sheet():
    with pytest.raises(KeyError):
        pd.read_excel(UHFILE, sheet_name=HIDDEN_SHEETS[0])


def test_read_visible_sheet():
    df = pd.read_excel(UHFILE, sheet_name=VISIBLE_SHEETS[0])
    assert isinstance(df, pd.DataFrame)


def test_read_skipped_and_visible():
    with pytest.raises(KeyError):
        pd.read_excel(UHFILE, sheet_name=[0, HIDDEN_SHEETS[0]])
