"""Структура листа «Данные участки» для брокеров + инструкция."""
from __future__ import annotations

import scripts.gsheet as g

DATA = "Данные участки"
ARCHIVE = "Архив участков"
GUIDE = "Инструкция для брокеров"
ACTIVE = ("Айвазовского", "Луначарского")
INSTR_LINES = 7  # строк инструкции сверху

INSTRUCTION = (
    "КАК ВНЕСТИ БРОНЬ ИЛИ ИЗМЕНИТЬ СТАТУС\n\n"
    "1. Ctrl+F (Cmd+F) → номер участка\n"
    "2. Столбец E «Статус» → выбрать из списка\n"
    "3. ПКМ по номеру (столбец A) → примечание: дата, брокер, срок брони\n"
    "4. Столбцы F–L не трогать — считаются автоматически\n\n"
    "Свободно — в продаже  |  Бронь — забронирован  |  "
    "Резерв — на согласовании  |  Продано — сделка закрыта\n"
    "Старые улицы (продано) → лист «Архив участков»"
)

GUIDE_LINES = [
    ["ИНСТРУКЦИЯ ДЛЯ БРОКЕРОВ — Шато Дель Маре"],
    [""],
    ["Где что лежит"],
    ["• Генплан и лоты — листы «Ген. план», «Квартал 1»"],
    ["• Бронь и статусы — лист «Данные участки» (жёлтый блок вверху)"],
    ["• Презентации и файлы — https://belikof.github.io/chateau-shahmatka/"],
    [""],
    ["Как поставить БРОНЬ"],
    ["1. Откройте «Данные участки»."],
    ["2. Ctrl+F / Cmd+F → номер участка → Enter."],
    ["3. Столбец E «Статус» → выберите «Бронь»."],
    ["4. ПКМ по номеру (столбец A) → «Вставить примечание», например:"],
    ["   12.07.2026, Иванова А., брокер Смирнов, бронь до 19.07"],
    ["5. Готово — статус виден на генплане."],
    [""],
    ["Другие статусы"],
    ["• Резерв — клиент думает, место придерживаем"],
    ["• Свободно — снова в продаже"],
    ["• Продано — сделка закрыта (только куратор)"],
    [""],
    ["Не редактировать: столбцы с ценами (J–L), лист «Каталог»."],
]


def rgb(h: str) -> dict:
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


def _sid(ss, title: str) -> int:
    for s in ss.fetch_sheet_metadata()["sheets"]:
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]
    raise KeyError(title)


def _row_street(row: list) -> str:
    return row[1].strip() if len(row) > 1 else ""


def _plot_num(row: list) -> int:
    try:
        return int(float(str(row[0]).replace(" ", "").replace("₽", "")))
    except ValueError:
        return 0


def restructure(ss) -> None:
    ws = ss.worksheet(DATA)
    sid = _sid(ss, DATA)
    rows = ws.get_all_values()
    if not rows:
        return

    header = rows[0]
    data_rows = [r for r in rows[1:] if r and r[0].strip()]
    active = [r for r in data_rows if _row_street(r) in ACTIVE]
    archive = [r for r in data_rows if _row_street(r) not in ACTIVE]

    # архив на отдельный лист
    try:
        aws = ss.worksheet(ARCHIVE)
    except Exception:
        aws = ss.add_worksheet(ARCHIVE, rows=max(len(archive) + 5, 50), cols=12)
    aws.clear()
    if archive:
        aws.update("A1", [header[:12]] + [r[:12] for r in archive])

    # удалить строки неактивных улиц снизу вверх (обычно блок 2–43)
    archive_indices = [
        i for i, r in enumerate(data_rows, start=2) if _row_street(r) not in ACTIVE
    ]
    if archive_indices:
        start = min(archive_indices) - 1  # 0-based
        end = max(archive_indices)  # exclusive 0-based
        ss.batch_update(
            {
                "requests": [
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": sid,
                                "dimension": "ROWS",
                                "startIndex": start,
                                "endIndex": end,
                            }
                        }
                    }
                ]
            }
        )

    ws = ss.worksheet(DATA)
    # инструкция сверху (если ещё нет)
    top = ws.acell("A1").value or ""
    if "КАК ВНЕСТИ БРОНЬ" not in str(top):
        ss.batch_update(
            {
                "requests": [
                    {
                        "insertDimension": {
                            "range": {
                                "sheetId": sid,
                                "dimension": "ROWS",
                                "startIndex": 0,
                                "endIndex": INSTR_LINES,
                            }
                        }
                    }
                ]
            }
        )
        ws = ss.worksheet(DATA)

    hdr = INSTR_LINES + 1
    data_start = hdr + 1

    ws.update("A1", [[INSTRUCTION]], value_input_option="USER_ENTERED")
    ws.update(f"E{hdr}", [["Статус ▾ (менять здесь)"]], value_input_option="USER_ENTERED")

    # сортировка активных участков
    n_active = len(active)
    if n_active > 1:
        ss.batch_update(
            {
                "requests": [
                    {
                        "sortRange": {
                            "range": {
                                "sheetId": sid,
                                "startRowIndex": hdr,
                                "endRowIndex": hdr + n_active,
                                "startColumnIndex": 0,
                                "endColumnIndex": 12,
                            },
                            "sortSpecs": [
                                {"dimensionIndex": 1, "sortOrder": "ASCENDING"},
                                {"dimensionIndex": 0, "sortOrder": "ASCENDING"},
                            ],
                        }
                    }
                ]
            }
        )

    # столбец P — калькулятор
    plots = sorted(_plot_num(r) for r in active if _plot_num(r))
    ws.update(f"P{hdr}", [["Участки (калькулятор)"]], value_input_option="USER_ENTERED")
    if plots:
        ws.update(
            f"P{hdr + 1}",
            [[str(p)] for p in plots],
            value_input_option="USER_ENTERED",
        )
    last_p = hdr + len(plots)

    def gr(r0, r1, c0, c1):
        return {
            "sheetId": sid,
            "startRowIndex": r0,
            "endRowIndex": r1,
            "startColumnIndex": c0,
            "endColumnIndex": c1,
        }

    reqs = [
        {"mergeCells": {"range": gr(0, INSTR_LINES, 0, 12), "mergeType": "MERGE_ALL"}},
        {
            "repeatCell": {
                "range": gr(0, INSTR_LINES, 0, 12),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": rgb("FFF8E7"),
                        "wrapStrategy": "WRAP",
                        "verticalAlignment": "TOP",
                        "textFormat": {"fontSize": 10},
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "repeatCell": {
                "range": gr(hdr - 1, hdr, 0, 12),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": rgb("1C4587"),
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": rgb("FFFFFF"),
                            "fontSize": 10,
                        },
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "repeatCell": {
                "range": gr(hdr - 1, hdr, 4, 5),
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": rgb("FFD966"),
                        "textFormat": {"bold": True, "fontSize": 10},
                    }
                },
                "fields": "userEnteredFormat",
            }
        },
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": hdr}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
        {
            "setBasicFilter": {
                "filter": {"range": gr(hdr - 1, hdr + n_active, 0, 12)}
            }
        },
        {
            "setDataValidation": {
                "range": gr(hdr, hdr + n_active, 4, 5),
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [
                            {"userEnteredValue": x}
                            for x in ("Свободно", "Бронь", "Резерв", "Продано")
                        ],
                    },
                    "showCustomUi": True,
                    "strict": True,
                },
            }
        },
    ]
    ss.batch_update({"requests": reqs})

    # именованный диапазон для калькулятора
    try:
        ss.batch_update(
            {
                "requests": [
                    {
                        "addNamedRange": {
                            "namedRange": {
                                "name": "СписокУчастков",
                                "range": gr(hdr, last_p, 15, 16),
                            }
                        }
                    }
                ]
            }
        )
    except Exception:
        pass


def write_guide(ss) -> None:
    try:
        ws = ss.worksheet(GUIDE)
    except Exception:
        ws = ss.add_worksheet(GUIDE, rows=40, cols=4)
    ws.clear()
    ws.update("A1", GUIDE_LINES, value_input_option="USER_ENTERED")
    gid = ws.id
    ss.batch_update(
        {
            "requests": [
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": gid,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 4,
                        },
                        "mergeType": "MERGE_ALL",
                    }
                },
                {
                    "updateSheetProperties": {
                        "properties": {"sheetId": gid, "index": 0},
                        "fields": "index",
                    }
                },
            ]
        }
    )
    try:
        old = ss.worksheet("Как пользоваться")
        ss.batch_update(
            {
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {"sheetId": old.id, "hidden": True},
                            "fields": "hidden",
                        }
                    }
                ]
            }
        )
    except Exception:
        pass


def main() -> None:
    ss = g.open_spreadsheet()
    print("Структура «Данные участки»…")
    restructure(ss)
    print("Инструкция для брокеров…")
    write_guide(ss)
    print("Готово.")


if __name__ == "__main__":
    main()
