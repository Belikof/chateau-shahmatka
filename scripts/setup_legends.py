"""Легенды, примечания, скрытие SPACE — одноразовая настройка живой таблицы."""
from __future__ import annotations

import scripts.gsheet as g

NOTE_PANORAMA = "Вид: море и лиман (панорама). На генплане: МВ."
NOTE_BRON = "Бронь. Менять статус в колонке E («Данные участки»)."
NOTE_AREA = (
    "Площадь по схеме посёлка: {area} м². "
    "На генплане участок может казаться крупнее соседей — ориентируемся на эту цифру."
)
CATALOG_NOTE = (
    "«Дом м²» — жилая площадь по проекту. Антресоль/мансарда указана на планировке; "
    "если не выделена отдельно — входит в площадь этажа."
)
HOWTO = [
    "Как работать с шахматкой",
    "",
    "1. Статусы — только лист «Данные участки», колонка E:",
    "   Свободно | Бронь | Резерв | Продано",
    "2. Бронь: статус «Бронь» + примечание (дата, брокер, срок).",
    "3. Двойной вид (море + лиман): в колонке D — «Панорама», примечание уточняет.",
    "4. Генплан: коды зон — см. блок «КАК ЧИТАТЬ ТАБЛИЦУ».",
    "5. Листы SPACE — другой продукт, скрыты.",
]


def rgb(h: str) -> dict:
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


def text_fmt(bold: bool = False, size: int = 9, color: str = "000000") -> dict:
    return {
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP",
        "textFormat": {
            "bold": bold,
            "fontSize": size,
            "foregroundColor": rgb(color),
        },
    }


def setup_genplan_legend(ss) -> None:
    ws = ss.worksheet("Ген. план")
    sid = ws.id
    meta = ss.fetch_sheet_metadata()
    merges = []
    for s in meta["sheets"]:
        if s["properties"]["sheetId"] == sid:
            merges = s.get("merges", [])
            break

    def gr(r0, r1, c0, c1):
        return {
            "sheetId": sid,
            "startRowIndex": r0,
            "endRowIndex": r1,
            "startColumnIndex": c0,
            "endColumnIndex": c1,
        }

    reqs = []
    for m in merges:
        if m["startRowIndex"] >= 7 and m["startRowIndex"] <= 9 and m["startColumnIndex"] >= 10:
            reqs.append({"unmergeCells": {"range": m}})

    labels = [
        (7, 8, 10, 14, "М — вид на море", "D1E0F2"),
        (7, 8, 14, 18, "В — виноградники, лиман", "D8EAD1"),
        (8, 9, 10, 18, "МВ — море и лиман (панорама)", "FCE5CC"),
        (9, 10, 10, 18, "— без видовых преимуществ", "F9FCFF"),
    ]
    for r0, r1, c0, c1, text, hx in labels:
        reqs.append({"mergeCells": {"range": gr(r0, r1, c0, c1), "mergeType": "MERGE_ALL"}})
        reqs.append(
            {
                "repeatCell": {
                    "range": gr(r0, r1, c0, c1),
                    "cell": {
                        "userEnteredValue": {"stringValue": text},
                        "userEnteredFormat": {
                            **text_fmt(size=9),
                            "backgroundColor": rgb(hx),
                        },
                    },
                    "fields": "userEnteredValue,userEnteredFormat",
                }
            }
        )
    ss.batch_update({"requests": reqs})


def setup_kvartal1_legend(ss) -> None:
    ws = ss.worksheet("Квартал 1")
    ws.batch_update(
        [
            {"range": "A11", "values": [["Легенда / статусы:"]]},
            {"range": "A12", "values": [["Продан — сделка закрыта"]]},
            {"range": "D12", "values": [["Бронь — клиент забронировал"]]},
            {"range": "A13", "values": [["Резерв — на согласовании"]]},
            {"range": "D13", "values": [["Свободно — в продаже"]]},
            {
                "range": "A14",
                "values": [[
                    "Бронь вносим в «Данные участки», колонка E + примечание к номеру."
                ]],
            },
        ],
        value_input_option="USER_ENTERED",
    )


def setup_catalog_note(ss) -> None:
    ws = ss.worksheet("Каталог")
    ws.update_note("C1", CATALOG_NOTE)
    ws.batch_update(
        [{"range": "C1", "values": [["Дом м²"]]}],
        value_input_option="USER_ENTERED",
    )


def setup_howto(ss) -> None:
    ws = ss.worksheet("Как пользоваться")
    rows = [[line] for line in HOWTO]
    ws.update(range_name="A1", values=rows, value_input_option="USER_ENTERED")


def hide_space_sheets(ss) -> None:
    reqs = []
    for title in ("Данные SPACE", "Шахматка SPACE"):
        try:
            ws = ss.worksheet(title)
        except Exception:
            continue
        reqs.append(
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": ws.id, "hidden": True},
                    "fields": "hidden",
                }
            }
        )
    if reqs:
        ss.batch_update({"requests": reqs})


def add_plot_notes(ss) -> None:
    ws = ss.worksheet("Данные участки")
    rows = ws.get_all_values()
    header = rows[0]
    col = {name: i for i, name in enumerate(header)}

    for r, row in enumerate(rows[1:], start=2):
        if not row or not row[col["№ участка"]]:
            continue
        try:
            num = int(float(row[col["№ участка"]]))
        except ValueError:
            continue
        view = row[col["Вид"]] if col.get("Вид", -1) < len(row) else ""
        status = row[col["Статус"]] if col.get("Статус", -1) < len(row) else ""
        area = row[col["Площадь м²"]] if col.get("Площадь м²", -1) < len(row) else ""

        cell = f"A{r}"
        notes = []
        if view == "Панорама":
            notes.append(NOTE_PANORAMA)
        if status == "Бронь":
            notes.append(NOTE_BRON)
        if num == 1705 and area:
            notes.append(NOTE_AREA.format(area=int(float(area))))
        if notes:
            ws.update_note(cell, "\n".join(notes))


def main() -> None:
    ss = g.open_spreadsheet()
    print("Генплан: легенда зон…")
    setup_genplan_legend(ss)
    print("Квартал 1: легенда статусов…")
    setup_kvartal1_legend(ss)
    print("Каталог: примечание по площади…")
    setup_catalog_note(ss)
    print("Как пользоваться…")
    setup_howto(ss)
    print("Скрыть SPACE…")
    hide_space_sheets(ss)
    print("Примечания к участкам…")
    add_plot_notes(ss)
    print("Готово.")


if __name__ == "__main__":
    main()
