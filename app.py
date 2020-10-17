from pathlib import Path

from src.domains.excel import ExcelData, RectangularRange
from src.services.transformer import Transformer


DIR = Path("./data")


if __name__ == "__main__":
    settings_path = DIR / "settings.yml"

    tf = Transformer(settings_path)

    settings = tf.settings["dest"]

    for setting in settings:
        src_book = DIR / tf.settings["src"]
        src_sheet = settings[setting]["src"]["sheet"]
        src_start = settings[setting]["src"]["target"]["start"]
        src_end = settings[setting]["src"]["target"]["end"]
        src_range = RectangularRange(src_start, src_end)
        src = ExcelData(src_book, src_sheet, src_range)

        dest_book = DIR / settings[setting]["dest"]["book"]
        dest_sheet = settings[setting]["dest"]["sheet"]
        dest_start = settings[setting]["dest"]["target"]["start"]
        dest_end = settings[setting]["dest"]["target"]["end"]
        dest_range = RectangularRange(dest_start, dest_end)
        dest = ExcelData(dest_book, dest_sheet, dest_range)

        tf.overwrite_excel(src, dest)
