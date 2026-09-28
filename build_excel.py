"""Genera el Excel público de tipo de cambio digital desde la fuente CSV."""
from __future__ import annotations

import csv
import io
import statistics
import sys
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

SOURCE_URL = "https://raw.githubusercontent.com/mauforonda/indicadores_dolar/main/dolar_sell.csv"
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template" / "tc hoy.xlsx"
OUTPUT_PATH = BASE_DIR / "public" / "tipo_cambio_digital.xlsx"
SHEET_NAME = "TC Digital actualizado"


def fetch_daily_averages() -> list[tuple[datetime, float, int, datetime]]:
    response = requests.get(SOURCE_URL, timeout=60)
    response.raise_for_status()
    grouped: dict[str, list[tuple[float, datetime]]] = defaultdict(list)

    for row in csv.DictReader(io.StringIO(response.text)):
        timestamp = datetime.fromisoformat(row["timestamp"].replace("Z", "+00:00"))
        value = float(row["vwap_sale"])
        if value > 0:
            # La fecha se mantiene en UTC para coincidir con la serie histórica.
            grouped[timestamp.date().isoformat()].append((value, timestamp))

    result = []
    for date_text, values in sorted(grouped.items()):
        result.append((
            datetime.fromisoformat(date_text),
            statistics.fmean(value for value, _ in values),
            len(values),
            max(timestamp for _, timestamp in values).astimezone(timezone.utc).replace(tzinfo=None),
        ))
    return result


def build_workbook(rows: list[tuple[datetime, float, int, datetime]]) -> None:
    workbook = load_workbook(TEMPLATE_PATH)
    if SHEET_NAME in workbook.sheetnames:
        del workbook[SHEET_NAME]
    sheet = workbook.create_sheet(SHEET_NAME, 0)
    sheet.append(["Fecha (UTC)", "TC digital promedio", "Observaciones", "Última actualización (UTC)"])
    for row in rows:
        sheet.append(row)

    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    sheet.column_dimensions["A"].width = 15
    sheet.column_dimensions["B"].width = 22
    sheet.column_dimensions["C"].width = 16
    sheet.column_dimensions["D"].width = 29
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row, 1).number_format = "yyyy-mm-dd"
        sheet.cell(row, 2).number_format = "0.0000"
        sheet.cell(row, 3).number_format = "0"
        sheet.cell(row, 4).number_format = "yyyy-mm-dd hh:mm"

    table = Table(displayName="tblTC_Digital", ref=f"A1:D{sheet.max_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    sheet.add_table(table)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT_PATH)


if __name__ == "__main__":
    try:
        build_workbook(fetch_daily_averages())
    except Exception as error:
        print(f"No se pudo generar el Excel: {error}", file=sys.stderr)
        raise
