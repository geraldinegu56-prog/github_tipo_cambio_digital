"""Genera el Excel público con tipo de cambio digital y oficial."""
from __future__ import annotations

import csv
import io
import math
import statistics
import unicodedata
from collections import defaultdict
from datetime import date, datetime, timezone
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

SOURCE_URL = "https://raw.githubusercontent.com/mauforonda/indicadores_dolar/main/dolar_sell.csv"
BCB_URL = "https://www.bcb.gob.bo/tiposDeCambioHistorico/xls.php?anio={year}"
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template" / "tc hoy.xlsx"
OUTPUT_PATH = BASE_DIR / "public" / "tipo_cambio_digital.xlsx"
SHEET_NAME = "TC Digital y oficial"
MONTHS = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
          "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}


def normalize(value: object) -> str:
    value = unicodedata.normalize("NFKD", str(value))
    return " ".join("".join(char if char.isalnum() else " " for char in value).casefold().split())


def official_rates(years: set[int]) -> dict[date, float]:
    rates: dict[date, tuple[int, float]] = {}
    for year in sorted(years):
        response = requests.get(BCB_URL.format(year=year), timeout=60)
        response.raise_for_status()
        table = pd.read_excel(BytesIO(response.content), header=None)
        headers = [(r, c) for r in range(table.shape[0]) for c in range(table.shape[1]) if normalize(table.iat[r, c]) == "dias"]
        if len(headers) != 1:
            raise ValueError(f"BCB {year}: no se encontró la cabecera DIAS")
        header_row, day_col = headers[0]
        columns = []
        for col in range(day_col + 1, table.shape[1]):
            month = MONTHS.get(normalize(table.iat[header_row, col]))
            quote = normalize(table.iat[header_row + 1, col])
            if month and quote in {"venta", "oficial"}:
                columns.append((col, month, 1 if quote == "oficial" else 0))
        for row in range(header_row + 2, table.shape[0]):
            day = pd.to_numeric(table.iat[row, day_col], errors="coerce")
            if pd.isna(day) or int(day) != day:
                continue
            for col, month, priority in columns:
                value = pd.to_numeric(table.iat[row, col], errors="coerce")
                if pd.isna(value) or not math.isfinite(float(value)) or float(value) <= 0:
                    continue
                try:
                    current_date = date(year, month, int(day))
                except ValueError:
                    continue
                current = rates.get(current_date)
                if current is None or priority > current[0]:
                    rates[current_date] = (priority, float(value))
    return {current_date: value for current_date, (_, value) in rates.items()}


def fetch_daily_averages() -> list[tuple[datetime, float, int, datetime]]:
    response = requests.get(SOURCE_URL, timeout=60)
    response.raise_for_status()
    grouped: dict[str, list[tuple[float, datetime]]] = defaultdict(list)
    for row in csv.DictReader(io.StringIO(response.text)):
        timestamp = datetime.fromisoformat(row["timestamp"].replace("Z", "+00:00"))
        value = float(row["vwap_sale"])
        if value > 0:
            # La fecha UTC coincide con la serie histórica del archivo base.
            grouped[timestamp.date().isoformat()].append((value, timestamp))
    return [(
        datetime.fromisoformat(day), statistics.fmean(value for value, _ in values), len(values),
        max(timestamp for _, timestamp in values).astimezone(timezone.utc).replace(tzinfo=None),
    ) for day, values in sorted(grouped.items())]


def build_workbook(rows: list[tuple[datetime, float, int, datetime]]) -> None:
    official = official_rates({day.year for day, _, _, _ in rows})
    workbook = load_workbook(TEMPLATE_PATH)
    if SHEET_NAME in workbook.sheetnames:
        del workbook[SHEET_NAME]
    sheet = workbook.create_sheet(SHEET_NAME, 0)
    sheet.append(["Fecha (UTC)", "TC oficial BCB (venta)", "TC digital promedio", "Observaciones", "Prima digital (%)", "Ultima actualizacion (UTC)"])
    for current_date, digital, observations, updated_at in rows:
        official_rate = official.get(current_date.date())
        premium = digital / official_rate - 1 if official_rate else None
        sheet.append([current_date, official_rate, digital, observations, premium, updated_at])
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for column, width in {"A": 15, "B": 23, "C": 22, "D": 16, "E": 18, "F": 29}.items():
        sheet.column_dimensions[column].width = width
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row, 1).number_format = "yyyy-mm-dd"
        sheet.cell(row, 2).number_format = "0.0000"
        sheet.cell(row, 3).number_format = "0.0000"
        sheet.cell(row, 4).number_format = "0"
        sheet.cell(row, 5).number_format = "0.00%"
        sheet.cell(row, 6).number_format = "yyyy-mm-dd hh:mm"
    table = Table(displayName="tblTC_DigitalOficial", ref=f"A1:F{sheet.max_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    sheet.add_table(table)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(OUTPUT_PATH)


if __name__ == "__main__":
    build_workbook(fetch_daily_averages())
