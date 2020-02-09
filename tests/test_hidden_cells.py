import pandas as pd
from uhxl import UhExcelFile

DATA = "tests/data/test.xlsx"
FILE = UhExcelFile(DATA, hide_rows=True, hide_columns=True)


def test_read_excel_sheet1():
    df = pd.read_excel(FILE, sheet_name="sheet1")
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"A": ["a1", "a2"], "B": ["b1", "b2"]}))


def test_read_excel_sheet2():
    df = pd.read_excel(FILE, sheet_name="sheet2")
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"A": [13], "C": [15]}))


def test_header():
    df = pd.read_excel(FILE, sheet_name="sheet1", header=1)
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"a1": ["a2"], "b1": ["b2"]}))


def test_usecols_str():
    df = pd.read_excel(FILE, sheet_name="sheet1", usecols="A:B")
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"A": ["a1", "a2"]}))


def test_usecols_int():
    df = pd.read_excel(FILE, sheet_name="sheet1", usecols=[0, 1])
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"A": ["a1", "a2"]}))


def test_nrows():
    df = pd.read_excel(FILE, sheet_name="sheet1", nrows=1)
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"A": ["a1"], "B": ["b1"]}))


def test_skiprows_one():
    df = pd.read_excel(FILE, sheet_name="sheet1", skiprows=(1,))
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"a1": ["a2"], "b1": ["b2"]}))


def test_skiprows_many():
    df = pd.read_excel(FILE, sheet_name="sheet1", skiprows=(0, 2, 3, 4))
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"A": ["a2"], "B": ["b2"]}))
