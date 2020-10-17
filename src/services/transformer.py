from pathlib import Path

import openpyxl
import pandas as pd
import yaml

from src.domains.excel import ExcelData
from src.services.utils import fetch_merged_cell_value


class Transformer:
    """Transform data"""

    def __init__(self, settings: Path) -> None:
        """
        Args:
            settings(Path): setting file's path
        """
        with settings.open(mode="r", encoding="utf-8") as yml:
            self.settings = yaml.safe_load(yml)

    @staticmethod
    def to_dataframe(excel: ExcelData) -> pd.DataFrame:
        """transform Excel to pandas DataFrame

        Args:
            excel(ExcelData): Target Excel range

        Returns:
            (pd.DataFrame) The DataFrame of select range data in Excel
        """
        wb = openpyxl.load_workbook(excel.path, data_only=False)
        sheet = wb[excel.sheet]
        rectangular_range = sheet[excel.target_range.start:excel.target_range.end]

        cells = []
        for row in rectangular_range:
            rows = []
            for cell in row:
                rows.append(fetch_merged_cell_value(sheet, cell))
            cells.append(rows)

        df = pd.DataFrame(cells)
        return df

    def overwrite_excel(self, src: ExcelData, dest: ExcelData) -> None:
        """overwrite target Excel range

        Args:
            src(ExcelData): Posting source Excel range
            dest(ExcelData): Posting destination Excel range

        Notes:
            Images and charts are deleted when save
        """
        print(f"\t{src.path}[{src.sheet}] -> {dest.path}[{dest.sheet}]")
        src = self.to_dataframe(src)

        dest_wb = openpyxl.load_workbook(dest.path, keep_vba=True)
        dest_ws = dest_wb[dest.sheet]

        for rows, values in zip(dest_ws[dest.target_range.start:dest.target_range.end],
                                [x for x in src.values.tolist()]):
            for dest_cell, src_cell in zip(rows, values):
                # Post only those that have no formula in the destination cell and have a value in the source cell
                if (dest_cell is not None) and (type(dest_cell) is not str):
                    dest_cell.value = src_cell
        dest_wb.save(dest.path)
