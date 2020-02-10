import pandas as pd
from uhxl import UhExcelFile

DATA = "tests/data/test_merged_cell.xlsx"
FILE = UhExcelFile(DATA)


def test_merged_cell():
    df = pd.read_excel(FILE)
    assert isinstance(df, pd.DataFrame)
    assert df.equals(pd.DataFrame({"merged": ["col1", "a"], None: ["col2", "b"]}))


def test_merged_cell_multi_columns():
    df = pd.read_excel(FILE, header=(0, 1))
    assert isinstance(df, pd.DataFrame)
    columns = pd.MultiIndex.from_tuples([("merged", "col1"), ("merged", "col2")])
    assert df.equals(pd.DataFrame([["a", "b"]], columns=columns))
