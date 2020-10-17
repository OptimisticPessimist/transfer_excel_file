from dataclasses import dataclass
from pathlib import Path


@dataclass
class RectangularRange:
    """The Range of Excel's start and end
    Attributes:
        start(str): The top left cell
        end(str): The bottom right cell

    Examples:
        RectangularRange("A1", "B16")
            It's means select a rectangular range from the cell A1 to the cell B16.
    """
    start: str
    end: str


@dataclass
class ExcelData:
    """Excel metadata class
    Attributes:
        path(Path): Target Excel's file path
        sheet(str): Target Excel's sheet name
        target_range(RectangularRange): Target Excel's range
    """
    path: Path
    sheet: str
    target_range: RectangularRange
