"""Загрузить генплан с топографией на лист «Ген. план»."""
from __future__ import annotations

import shutil
from pathlib import Path

import scripts.gsheet as g
from scripts.overlay_genplan_topo import OUT, build

SHEET = "Ген. план"
SHEET_ID = 10627592
IMAGE_CELL = "C10"

GITHUB_RAW = (
    "https://raw.githubusercontent.com/Belikof/chateau-shahmatka/main/"
    "taplink/assets/genplan_topo.png"
)
PAGES_URL = "https://belikof.github.io/chateau-shahmatka/assets/genplan_topo.png"


def apply_image_formula(ss, url: str) -> None:
    ws = ss.worksheet(SHEET)
    ws.update(
        IMAGE_CELL,
        [[f'=IMAGE("{url}"; 1)']],
        value_input_option="USER_ENTERED",
    )
    reqs = []
    for col in range(10):
        reqs.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": SHEET_ID,
                        "dimension": "COLUMNS",
                        "startIndex": col,
                        "endIndex": col + 1,
                    },
                    "properties": {"pixelSize": 42},
                    "fields": "pixelSize",
                }
            }
        )
    for row in range(9, 48):
        reqs.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": SHEET_ID,
                        "dimension": "ROWS",
                        "startIndex": row,
                        "endIndex": row + 1,
                    },
                    "properties": {"pixelSize": 22},
                    "fields": "pixelSize",
                }
            }
        )
    ss.batch_update({"requests": reqs})
    print(f"IMAGE в {IMAGE_CELL}")


def add_topo_legend(ss) -> None:
    ws = ss.worksheet(SHEET)
    ws.update(
        "K6:R6",
        [["РЕЛЬЕФ: 80 м (запад) → 130 м Поливадина (восток) · градиент на карте"]],
        value_input_option="USER_ENTERED",
    )
    ss.batch_update(
        {
            "requests": [
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": SHEET_ID,
                            "startRowIndex": 5,
                            "endRowIndex": 6,
                            "startColumnIndex": 10,
                            "endColumnIndex": 18,
                        },
                        "mergeType": "MERGE_ALL",
                    }
                },
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": SHEET_ID,
                            "startRowIndex": 5,
                            "endRowIndex": 6,
                            "startColumnIndex": 10,
                            "endColumnIndex": 18,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 0.93,
                                    "green": 0.96,
                                    "blue": 1,
                                },
                                "textFormat": {"fontSize": 10, "bold": True},
                            }
                        },
                        "fields": "userEnteredFormat",
                    }
                },
            ]
        }
    )


def main() -> None:
    build()
    assets = Path(__file__).resolve().parent.parent / "taplink" / "assets"
    assets.mkdir(parents=True, exist_ok=True)
    shutil.copy2(OUT, assets / "genplan_topo.png")

    ss = g.open_spreadsheet()
    applied = False
    try:
        import requests

        for url in (GITHUB_RAW, PAGES_URL):
            if requests.head(url, timeout=10).status_code == 200:
                apply_image_formula(ss, url)
                applied = True
                break
    except Exception:
        pass

    if not applied:
        print(
            "Сначала запушь taplink/assets/genplan_topo.png в main, "
            "потом снова: python scripts/apply_genplan_topo.py"
        )
    add_topo_legend(ss)
    print("Готово.")


if __name__ == "__main__":
    main()
