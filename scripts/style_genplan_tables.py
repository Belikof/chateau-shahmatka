"""Справочник участков на «Ген. план»: № + Зона + читаемая легенда."""
from __future__ import annotations

import scripts.gsheet as g
from scripts.sync_genplan_plots import END_ROW, START_ROW, plots_by_block

SHEET = "Ген. план"
SHEET_ID = 10627592

LEGEND_TITLE_ROW = 7
LEGEND_ZONE_ROW = 8      # М + В
LEGEND_ZONE2_ROW = 9     # МВ
LEGEND_HOUSES_ROW = 10   # дома
BLOCK_TITLE_ROW = 11
HEADER_ROW = 12
# START_ROW = 13

BLOCKS = (("K", "L"), ("M", "N"), ("O", "P"), ("Q", "R"))
BLOCK_TITLES = ("Блок 1", "Блок 2", "Блок 3", "Блок 4")

BLOCK_HDR = ("#1C4587", "#2E75B6", "#38761D", "#674EA7")
WHITE = {"red": 1, "green": 1, "blue": 1}
TITLE_BG = {"red": 0.25, "green": 0.38, "blue": 0.52}
LEGEND_BG = {"red": 0.98, "green": 0.99, "blue": 1.0}
ZONE_CLR = {
    "М": {"red": 0.82, "green": 0.88, "blue": 0.95},
    "В": {"red": 0.85, "green": 0.92, "blue": 0.82},
    "МВ": {"red": 0.99, "green": 0.9, "blue": 0.8},
}

# зоны: (код, текст, col_code, col_text_start, col_text_end)
ZONE_CELLS = (
    ("М", "средняя · дома 1–2 эт.", "K", "L", "N"),
    ("В", "видовая · виноградники, лиман", "O", "P", "R"),
)
ZONE_ROW2 = ("МВ", "панорама · море, виноград, лиман, вулканы", "K", "L", "R")

HOUSES_LINE = (
    "Дома на каждом участке:  "
    "1 эт. — Костеро, Линия 40, Линия 72, Логос, Логос мини, Арбор, Баня  ·  "
    "2 эт. — Палаццо, Эдем, Вилла Маре"
)


def _hex_rgb(hex6: str) -> dict:
    h = hex6.lstrip("#")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}


def zone_formula(plot_col: str, row: int) -> str:
    """Код зоны из колонки «Вид» (D), без технических «1 ф ВЛВ»."""
    p = f"{plot_col}{row}"
    v = f"VLOOKUP(VALUE({p});'Данные участки'!$A:$D;4;FALSE)"
    return (
        f'=IF({p}="";"";'
        f'IF({v}="Панорама";"МВ";'
        f'IF({v}="Вид на море";"М";'
        f'IF({v}="Виноградники / лиман / вулканы";"В";"—")))'
    )


def _grid(row: int, col: str) -> tuple[int, int]:
    from openpyxl.utils import column_index_from_string as ci
    return row - 1, ci(col) - 1


def _rng(r0: int, c0: int, r1: int, c1: int) -> dict:
    return {
        "sheetId": SHEET_ID,
        "startRowIndex": r0,
        "endRowIndex": r1,
        "startColumnIndex": c0,
        "endColumnIndex": c1,
    }


def unmerge_area(ss) -> list[dict]:
    meta = ss.fetch_sheet_metadata({"fields": "sheets(properties(sheetId,title),merges)"})
    req: list[dict] = []
    for s in meta["sheets"]:
        if s["properties"]["title"] != SHEET:
            continue
        for m in s.get("merges", []):
            c0, c1 = m["startColumnIndex"], m["endColumnIndex"]
            if c1 <= 10 or c0 > 22:
                continue
            if m["endRowIndex"] <= 6:
                continue
            req.append({"unmergeCells": {"range": m}})
    return req


def unmerge_trailing(ss) -> list[dict]:
    from openpyxl.utils import column_index_from_string as ci
    c_s = ci("S") - 1
    req: list[dict] = []
    for s in ss.fetch_sheet_metadata({"fields": "sheets(properties(sheetId,title),merges)"})["sheets"]:
        if s["properties"]["title"] != SHEET:
            continue
        for m in s.get("merges", []):
            if m["startColumnIndex"] >= c_s:
                req.append({"unmergeCells": {"range": m}})
    return req


def clear_trailing_cols() -> list[dict]:
    from openpyxl.utils import get_column_letter as L, column_index_from_string as ci
    blank = [[""] for _ in range(END_ROW - LEGEND_TITLE_ROW + 1)]
    return [
        {"range": f"{L(c)}{LEGEND_TITLE_ROW}:{L(c)}{END_ROW}", "values": blank}
        for c in range(ci("S"), ci("BB") + 1)
    ]


def legend_value_updates() -> list[dict]:
    u: list[dict] = [{"range": f"K{LEGEND_TITLE_ROW}", "values": [["КАК ЧИТАТЬ ТАБЛИЦУ"]]}]
    for code, text, cc, cs, ce in ZONE_CELLS:
        u.append({"range": f"{cc}{LEGEND_ZONE_ROW}", "values": [[code]]})
        u.append({"range": f"{cs}{LEGEND_ZONE_ROW}", "values": [[text]]})
    code, text, cc, cs, ce = ZONE_ROW2
    u.append({"range": f"{cc}{LEGEND_ZONE2_ROW}", "values": [[code]]})
    u.append({"range": f"{cs}{LEGEND_ZONE2_ROW}", "values": [[text]]})
    u.append({"range": f"K{LEGEND_HOUSES_ROW}", "values": [[HOUSES_LINE]]})
    return u


def build_requests() -> list[dict]:
    req: list[dict] = []
    border = {"style": "SOLID", "width": 1, "color": {"red": 0.78, "green": 0.78, "blue": 0.8}}
    thick = {"style": "SOLID", "width": 2, "color": {"red": 0.5, "green": 0.5, "blue": 0.52}}

    for col, px in (("K", 58), ("L", 40), ("M", 58), ("N", 40),
                    ("O", 58), ("P", 40), ("Q", 58), ("R", 40)):
        _, c0 = _grid(1, col)
        req.append({
            "updateDimensionProperties": {
                "range": {"sheetId": SHEET_ID, "dimension": "COLUMNS",
                          "startIndex": c0, "endIndex": c0 + 1},
                "properties": {"pixelSize": px},
                "fields": "pixelSize",
            }
        })

    # --- легенда ---
    # заголовок K7:R7
    r0, c0 = _grid(LEGEND_TITLE_ROW, "K")
    _, c1 = _grid(LEGEND_TITLE_ROW, "R")
    req.append({"mergeCells": {"range": _rng(r0, c0, r0 + 1, c1 + 1), "mergeType": "MERGE_ALL"}})
    req.append({
        "repeatCell": {
            "range": _rng(r0, c0, r0 + 1, c1 + 1),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": TITLE_BG,
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": WHITE},
                }
            },
            "fields": "userEnteredFormat",
        }
    })
    req.append({
        "updateDimensionProperties": {
            "range": {"sheetId": SHEET_ID, "dimension": "ROWS", "startIndex": r0, "endIndex": r0 + 1},
            "properties": {"pixelSize": 26},
            "fields": "pixelSize",
        }
    })

    # зоны строка 8: М + текст | В + текст
    zr, _ = _grid(LEGEND_ZONE_ROW, "K")
    req.append({
        "updateDimensionProperties": {
            "range": {"sheetId": SHEET_ID, "dimension": "ROWS", "startIndex": zr, "endIndex": zr + 1},
            "properties": {"pixelSize": 28},
            "fields": "pixelSize",
        }
    })
    for code, text, cc, cs, ce in ZONE_CELLS:
        _, cc_i = _grid(LEGEND_ZONE_ROW, cc)
        _, cs_i = _grid(LEGEND_ZONE_ROW, cs)
        _, ce_i = _grid(LEGEND_ZONE_ROW, ce)
        req.append({"mergeCells": {"range": _rng(zr, cs_i, zr + 1, ce_i + 1), "mergeType": "MERGE_ALL"}})
        req.append({
            "repeatCell": {
                "range": _rng(zr, cc_i, zr + 1, cc_i + 1),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": ZONE_CLR[code],
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True, "fontSize": 12},
                    }
                },
                "fields": "userEnteredFormat",
            }
        })
        req.append({
            "repeatCell": {
                "range": _rng(zr, cs_i, zr + 1, ce_i + 1),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": ZONE_CLR[code],
                        "horizontalAlignment": "LEFT",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"fontSize": 10},
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": "userEnteredFormat",
            }
        })

    # зона МВ строка 9
    zr2, _ = _grid(LEGEND_ZONE2_ROW, "K")
    _, cc_i = _grid(LEGEND_ZONE2_ROW, "K")
    _, cs_i = _grid(LEGEND_ZONE2_ROW, "L")
    _, ce_i = _grid(LEGEND_ZONE2_ROW, "R")
    req.append({"mergeCells": {"range": _rng(zr2, cs_i, zr2 + 1, ce_i + 1), "mergeType": "MERGE_ALL"}})
    req.append({
        "repeatCell": {
            "range": _rng(zr2, cc_i, zr2 + 1, cc_i + 1),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": ZONE_CLR["МВ"],
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "textFormat": {"bold": True, "fontSize": 12},
                }
            },
            "fields": "userEnteredFormat",
        }
    })
    req.append({
        "repeatCell": {
            "range": _rng(zr2, cs_i, zr2 + 1, ce_i + 1),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": ZONE_CLR["МВ"],
                    "horizontalAlignment": "LEFT",
                    "verticalAlignment": "MIDDLE",
                    "textFormat": {"fontSize": 10},
                }
            },
            "fields": "userEnteredFormat",
        }
    })
    req.append({
        "updateDimensionProperties": {
            "range": {"sheetId": SHEET_ID, "dimension": "ROWS", "startIndex": zr2, "endIndex": zr2 + 1},
            "properties": {"pixelSize": 28},
            "fields": "pixelSize",
        }
    })

    # дома строка 10
    hr, hc0 = _grid(LEGEND_HOUSES_ROW, "K")
    _, hc1 = _grid(LEGEND_HOUSES_ROW, "R")
    req.append({"mergeCells": {"range": _rng(hr, hc0, hr + 1, hc1 + 1), "mergeType": "MERGE_ALL"}})
    req.append({
        "repeatCell": {
            "range": _rng(hr, hc0, hr + 1, hc1 + 1),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": LEGEND_BG,
                    "horizontalAlignment": "LEFT",
                    "verticalAlignment": "MIDDLE",
                    "wrapStrategy": "WRAP",
                    "textFormat": {"fontSize": 10, "foregroundColor": {"red": 0.15, "green": 0.18, "blue": 0.25}},
                    "borders": {
                        "top": border, "bottom": border, "left": border, "right": border,
                    },
                }
            },
            "fields": "userEnteredFormat",
        }
    })
    req.append({
        "updateDimensionProperties": {
            "range": {"sheetId": SHEET_ID, "dimension": "ROWS", "startIndex": hr, "endIndex": hr + 1},
            "properties": {"pixelSize": 44},
            "fields": "pixelSize",
        }
    })

    # рамка легенды K7:R10
    req.append({
        "updateBorders": {
            "range": _rng(LEGEND_TITLE_ROW - 1, c0, LEGEND_HOUSES_ROW, c1 + 1),
            "top": thick, "bottom": thick, "left": thick, "right": thick,
            "innerHorizontal": border,
        }
    })

    # --- блоки таблицы ---
    for i, (plot_c, zone_c) in enumerate(BLOCKS):
        bg = _hex_rgb(BLOCK_HDR[i])
        br, bc0 = _grid(BLOCK_TITLE_ROW, plot_c)
        _, bc1 = _grid(BLOCK_TITLE_ROW, zone_c)
        req.append({"mergeCells": {"range": _rng(br, bc0, br + 1, bc1 + 1), "mergeType": "MERGE_ALL"}})
        req.append({
            "repeatCell": {
                "range": _rng(br, bc0, br + 1, bc1 + 1),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": bg,
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": WHITE},
                    }
                },
                "fields": "userEnteredFormat",
            }
        })
        hr, hc0 = _grid(HEADER_ROW, plot_c)
        _, hc1 = _grid(HEADER_ROW, zone_c)
        hdr_bg = {"red": bg["red"] * 0.75 + 0.25, "green": bg["green"] * 0.75 + 0.25, "blue": bg["blue"] * 0.75 + 0.25}
        req.append({
            "repeatCell": {
                "range": _rng(hr, hc0, hr + 1, hc1 + 1),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": hdr_bg,
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"bold": True, "fontSize": 9, "foregroundColor": WHITE},
                        "borders": {"bottom": thick},
                    }
                },
                "fields": "userEnteredFormat",
            }
        })
        dr, dc0 = _grid(START_ROW, plot_c)
        _, dc1 = _grid(END_ROW, zone_c)
        req.append({
            "repeatCell": {
                "range": _rng(dr, dc0, END_ROW, dc1 + 1),
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"fontSize": 10},
                        "borders": {
                            "top": border, "bottom": border, "left": border, "right": border,
                        },
                    }
                },
                "fields": "userEnteredFormat",
            }
        })
        req.append({
            "repeatCell": {
                "range": _rng(dr, dc0, END_ROW, dc0 + 1),
                "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                "fields": "userEnteredFormat.textFormat",
            }
        })
        req.append({
            "updateBorders": {
                "range": _rng(br, bc0, END_ROW, bc1 + 1),
                "top": thick, "bottom": thick, "left": thick, "right": thick,
            }
        })

    for _, zone_c in BLOCKS:
        _, c0 = _grid(START_ROW, zone_c)
        for val, clr in ZONE_CLR.items():
            req.append({
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [_rng(START_ROW - 1, c0, END_ROW, c0 + 1)],
                        "booleanRule": {
                            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": val}]},
                            "format": {"backgroundColor": clr, "textFormat": {"bold": True}},
                        },
                    },
                    "index": 0,
                }
            })

    # очистка S:BB
    from openpyxl.utils import column_index_from_string as ci
    white = {"red": 1, "green": 1, "blue": 1}
    req.append({
        "repeatCell": {
            "range": _rng(LEGEND_TITLE_ROW - 1, ci("S") - 1, END_ROW, ci("BB")),
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": white,
                    "borders": {
                        "top": {"style": "NONE"}, "bottom": {"style": "NONE"},
                        "left": {"style": "NONE"}, "right": {"style": "NONE"},
                    },
                }
            },
            "fields": "userEnteredFormat(backgroundColor,borders)",
        }
    })

    return req


def main() -> None:
    blocks = plots_by_block()
    ws = g.worksheet(SHEET)
    ss = g.open_spreadsheet()

    ss.batch_update({"requests": unmerge_area(ss)})

    n = END_ROW - START_ROW + 1
    updates: list[dict] = legend_value_updates()
    updates.extend(clear_trailing_cols())
    # очистить старые строки легенды
    for r in (LEGEND_ZONE_ROW, LEGEND_ZONE2_ROW, LEGEND_HOUSES_ROW):
        updates.append({"range": f"K{r}:R{r}", "values": [[""] * 8]})

    for i, (plot_c, zone_c) in enumerate(BLOCKS):
        updates.append({"range": f"{plot_c}{BLOCK_TITLE_ROW}", "values": [[BLOCK_TITLES[i]]]})
        updates.append({"range": f"{plot_c}{HEADER_ROW}:{zone_c}{HEADER_ROW}", "values": [["№", "Зона"]]})
        nums = blocks[i]
        updates.append({
            "range": f"{plot_c}{START_ROW}:{plot_c}{END_ROW}",
            "values": [[nums[j] if j < len(nums) else ""] for j in range(n)],
        })
        updates.append({
            "range": f"{zone_c}{START_ROW}:{zone_c}{END_ROW}",
            "values": [[zone_formula(plot_c, r)] for r in range(START_ROW, END_ROW + 1)],
        })

    ws.batch_update(updates, value_input_option="USER_ENTERED")
    ss.batch_update({"requests": unmerge_trailing(ss) + build_requests()})

    print(f"Готово: легенда {LEGEND_TITLE_ROW}–{LEGEND_HOUSES_ROW}, таблица с {START_ROW}")


if __name__ == "__main__":
    main()
