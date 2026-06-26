"""Быстрый осмотр структуры файла: листы, видимость, картинки, диапазоны.

Запуск:
    python scripts/inspect_xlsx.py "data/Новая таблица.xlsx"
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook


def main(path: str):
    wb = load_workbook(path, data_only=True)
    print(f"=== {Path(path).name} ===")
    for s in wb.worksheets:
        print(f"  {s.title!r:30} state={s.sheet_state:8} dims={s.dimensions}")

    with zipfile.ZipFile(path) as z:
        media = [n for n in z.namelist() if "media/" in n]
        print("media:", media)
        wbxml = z.read("xl/workbook.xml").decode()
        names = re.findall(r'<definedName name="([^"]+)"[^>]*>([^<]*)', wbxml)
        if names:
            print("defined names:")
            for n, ref in names:
                print(f"  {n} = {ref}")


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else str(
        Path(__file__).resolve().parents[1] / "data" / "Новая таблица.xlsx"
    )
    main(path)
