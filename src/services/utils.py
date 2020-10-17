from typing import Union

from openpyxl.worksheet.worksheet import Worksheet, Cell
from openpyxl.utils.cell import rows_from_range


def fetch_merged_cell_value(sheet: Worksheet, cell: Cell) -> Union[int, float, str]:
    """fetch a value from a merged cell

    Args:
        sheet(worksheet): The target Excel's sheet name
        cell(cell): The target cell

    Returns:
        (Union[int, float, str]) the merged cell value
    """
    cell_index = cell.coordinate
    for range_ in sheet.merged_cells.ranges:
        merged_cells = list(rows_from_range(str(range_)))
        for row in merged_cells:
            if cell_index in row:
                return sheet[merged_cells[0][0]].value
    return sheet[cell_index].value
