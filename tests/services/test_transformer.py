import pandas as pd

from pathlib import Path
from pytest import mark
from pandas.testing import assert_frame_equal

from src.services.transformer import Transformer
from src.domains.excel import ExcelData, RectangularRange

TEST_DATA_DIR = Path("./tests/data")
TEST_EXCEL_1 = ExcelData(TEST_DATA_DIR / "test.xlsx", "Sheet1", RectangularRange("B2", "D4"))
TEST_EXCEL_2 = ExcelData(TEST_DATA_DIR / "test.xlsx", "Sheet1", RectangularRange("B5", "C7"))


class TestTransformer:

    def test_to_dataframe_1(self):
        expected = [["a", "b", "c"],
                    ["f", "g", "h"],
                    ["k", "l", "m"]]
        expected = pd.DataFrame(expected)
        tf = Transformer(TEST_DATA_DIR / "settings.yml")
        actual = tf.to_dataframe(TEST_EXCEL_1)
        assert_frame_equal(actual, expected)

    def test_to_dataframe_2(self):
        expected = [[1, 2],
                    [6, 7],
                    ["=B5+B6", "=C5+C6"]]
        expected = pd.DataFrame(expected)
        tf = Transformer(TEST_DATA_DIR / "settings.yml")
        actual = tf.to_dataframe(TEST_EXCEL_2)
        assert_frame_equal(actual, expected)

    @mark.parametrize("src, dest, expected", (
            [ExcelData(TEST_DATA_DIR / "test.xlsx", "Sheet1", RectangularRange("B2", "F4")),
             ExcelData(TEST_DATA_DIR / "test.xlsx", "Sheet2", RectangularRange("A1", "E3")),
             [["a", "b", "c", "d", "e"], ["f", "g", "h", "i", "j"], ["k", "l", "m", None, "o"]]],
    ))
    def test_overwrite_excel(self, src, dest, expected):
        expected = pd.DataFrame(expected)
        tf = Transformer(TEST_DATA_DIR / "settings.yml")
        tf.overwrite_excel(src, dest)
        actual = tf.to_dataframe(dest)
        assert_frame_equal(actual, expected)
