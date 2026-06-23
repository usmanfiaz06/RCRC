"""
Build a polished Excel workbook from the cached RCRC + Open-Meteo data.

Sheets:
  1. Cover            - title, range, sources, notes
  2. Stations         - all 17 RCRC stations with metadata + lat/lon
  3. Monthly Summary  - cross-station monthly averages (PM2.5, PM10, AQI, etc.)
  4. Daily Summary    - per-station daily averages
  5. Compliance       - WHO + NCEC guideline comparison
  6. <Station> Hourly - one sheet per station with full hourly time series

Output: rcrc_riyadh_2026_jan_to_may.xlsx
"""
from __future__ import annotations
import json
import os
import statistics
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

SCRATCH = Path(os.environ.get(
    "RCRC_SCRATCH",
    "/tmp/claude-0/-home-user-RCRC/3e7d6508-8244-5a9f-b105-f6f5e0138312/scratchpad/raw",
))
OUTPUT = Path("rcrc_riyadh_2026_jan_to_may.xlsx")

# ===== Brand palette =====
BRAND = "1DA1F2"        # primary
BRAND_DARK = "0D6FA8"
BRAND_LIGHT = "E1F1FB"
COVER_BG = "0B2A45"
WHITE = "FFFFFF"
GREY = "F4F6F8"
DARK_TEXT = "1B2733"

# AQI band colors (US AQI)
AQI_BANDS = [
    (0,  50,   "00E400", "Good"),
    (51, 100,  "FFFF00", "Moderate"),
    (101,150,  "FF7E00", "Unhealthy (SG)"),
    (151,200,  "FF0000", "Unhealthy"),
    (201,300,  "8F3F97", "Very Unhealthy"),
    (301,1000, "7E0023", "Hazardous"),
]

WHO_LIMITS = {       # 24-hour
    "PM2.5 (µg/m³)": 15,
    "PM10 (µg/m³)": 45,
    "NO2 (µg/m³)": 25,
    "SO2 (µg/m³)": 40,
    "O3 (µg/m³)": 100,
    "CO (mg/m³)": 4,
}
NCEC_LIMITS = {
    "PM2.5 (µg/m³)": 35,
    "PM10 (µg/m³)": 340,
    "NO2 (µg/m³)": 660,
    "SO2 (µg/m³)": 365,
    "O3 (µg/m³)": 120,
    "CO (mg/m³)": 10,
}


# ===== Style helpers =====
THIN = Side(style="thin", color="C6CFD8")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def header_fill():
    return PatternFill("solid", fgColor=BRAND)


def header_font():
    return Font(name="Calibri", size=11, bold=True, color=WHITE)


def body_font():
    return Font(name="Calibri", size=11, color=DARK_TEXT)


def title_font():
    return Font(name="Calibri", size=26, bold=True, color=WHITE)


def cover_subfont(size: int = 12, bold: bool = False, color: str = WHITE):
    return Font(name="Calibri", size=size, bold=bold, color=color)


def write_header(ws, row: int, headers: list[str]):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = header_fill()
        cell.font = header_font()
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 28


def autosize(ws, max_w: int = 28):
    for col_idx, col in enumerate(ws.columns, start=1):
        letter = get_column_letter(col_idx)
        best = 8
        for c in col:
            v = c.value
            if v is None:
                continue
            s = str(v)
            for line in s.splitlines():
                if len(line) > best:
                    best = len(line)
        ws.column_dimensions[letter].width = min(max(best + 2, 10), max_w)


def aqi_fill(aqi):
    if aqi is None:
        return None
    try:
        v = float(aqi)
    except (TypeError, ValueError):
        return None
    for lo, hi, color, _ in AQI_BANDS:
        if lo <= v <= hi:
            return PatternFill("solid", fgColor=color)
    return None


# ===== Data loaders =====
def load_stations() -> list[dict]:
    return json.loads((SCRATCH / "rcrc_stations.json").read_text())["results"]


def slugify(name: str) -> str:
    return (
        name.lower()
        .replace(" ", "_")
        .replace("/", "_")
        .replace("(", "")
        .replace(")", "")
        .replace(",", "")
        .replace("__", "_")
        .strip("_")
    )


def load_series(slug: str) -> tuple[dict, dict]:
    aq = json.loads((SCRATCH / f"aq_{slug}.json").read_text())
    wx = json.loads((SCRATCH / f"wx_{slug}.json").read_text())
    return aq, wx


# ===== Sheet builders =====
def build_cover(wb: Workbook, stations: list[dict], generated: str, row_count: int):
    ws = wb.create_sheet("Cover", 0)
    ws.sheet_view.showGridLines = False

    # Banner background
    for r in range(1, 13):
        for c in range(1, 12):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=COVER_BG)

    ws.merge_cells("B2:K2")
    t = ws.cell(row=2, column=2, value="RCRC · Riyadh Air Quality Dataset")
    t.font = title_font()
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 44

    ws.merge_cells("B3:K3")
    s = ws.cell(row=3, column=2,
                value="Royal Commission for Riyadh City — Air Quality Monitoring Stations")
    s.font = cover_subfont(14, color="9BD0F0")
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 22

    ws.merge_cells("B4:K4")
    ws.cell(row=4, column=2, value="Reporting period: 1 January 2026 — 31 May 2026 (151 days, hourly)") \
        .font = cover_subfont(12, color="C6E4F4")

    ws.merge_cells("B5:K5")
    ws.cell(row=5, column=2, value=f"Generated: {generated}") \
        .font = cover_subfont(11, color="C6E4F4")

    # Stat panels
    panels = [
        ("Stations", str(len(stations))),
        ("Pollutants tracked", "6 (PM2.5, PM10, NO₂, SO₂, O₃, CO)"),
        ("Weather variables", "9 (temp, humidity, wind, pressure...)"),
        ("Total hourly rows", f"{row_count:,}"),
    ]
    for i, (label, value) in enumerate(panels):
        c = 2 + i * 2
        ws.merge_cells(start_row=7, start_column=c, end_row=7, end_column=c + 1)
        ws.merge_cells(start_row=8, start_column=c, end_row=9, end_column=c + 1)
        lab = ws.cell(row=7, column=c, value=label)
        lab.font = cover_subfont(10, bold=True, color="9BD0F0")
        lab.alignment = Alignment(horizontal="center", vertical="center")
        val = ws.cell(row=8, column=c, value=value)
        val.font = cover_subfont(16, bold=True, color=WHITE)
        val.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 18
    ws.row_dimensions[8].height = 28
    ws.row_dimensions[9].height = 16

    # Sources / methodology
    ws.cell(row=15, column=2, value="Data sources").font = Font(name="Calibri", size=13, bold=True, color=BRAND_DARK)
    sources = [
        ("RCRC Open Data Portal",
         "Station inventory — https://opendata.rcrc.gov.sa/explore/dataset/air-quality-stations-in-riyadh-2025/"),
        ("Open-Meteo Air Quality (CAMS)",
         "Hourly PM2.5, PM10, NO₂, SO₂, O₃, CO, dust, UV, US-AQI sub-indices — https://open-meteo.com/en/docs/air-quality-api"),
        ("Open-Meteo Historical Weather (ERA5)",
         "Hourly temperature, humidity, wind, pressure, precipitation — https://open-meteo.com/en/docs/historical-weather-api"),
        ("WAQI / aqicn.org",
         "Real-time snapshot reference — https://aqicn.org/city/riyadh/"),
    ]
    r = 16
    for name, desc in sources:
        a = ws.cell(row=r, column=2, value=name)
        a.font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=11)
        b = ws.cell(row=r, column=3, value=desc)
        b.font = Font(name="Calibri", size=11, color=DARK_TEXT)
        b.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 20
        r += 1

    # Notes
    ws.cell(row=r + 1, column=2, value="Notes").font = Font(name="Calibri", size=13, bold=True, color=BRAND_DARK)
    notes = [
        "Pollutant concentrations are hourly modelled values from CAMS reanalysis sampled at each station's coordinates.",
        "Weather variables come from ERA5 reanalysis (Open-Meteo \"Historical Weather\" archive).",
        "Times are local (Asia/Riyadh, GMT+3).",
        "WHO 2021 limits and NCEC Saudi limits are 24-hour averages unless otherwise noted.",
        "AQI bands follow US-AQI: 0–50 Good, 51–100 Moderate, 101–150 USG, 151–200 Unhealthy, 201–300 Very Unhealthy, 301+ Hazardous.",
    ]
    for i, n in enumerate(notes):
        ws.merge_cells(start_row=r + 2 + i, start_column=2, end_row=r + 2 + i, end_column=11)
        cell = ws.cell(row=r + 2 + i, column=2, value=f"• {n}")
        cell.font = Font(name="Calibri", size=11, color=DARK_TEXT)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r + 2 + i].height = 18

    # AQI legend
    legend_r = r + 2 + len(notes) + 2
    ws.cell(row=legend_r, column=2, value="AQI Legend").font = Font(
        name="Calibri", size=13, bold=True, color=BRAND_DARK,
    )
    for i, (lo, hi, color, label) in enumerate(AQI_BANDS):
        cell = ws.cell(row=legend_r + 1, column=2 + i, value=f"{label}\n{lo}–{hi}")
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[legend_r + 1].height = 36

    # Column widths
    for col_letter in "ABCDEFGHIJKL":
        ws.column_dimensions[col_letter].width = 18
    ws.column_dimensions["A"].width = 3


def build_stations(wb: Workbook, stations: list[dict]):
    ws = wb.create_sheet("Stations")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:G1")
    t = ws.cell(row=1, column=1, value="RCRC Air Quality Monitoring Stations — Riyadh")
    t.font = Font(name="Calibri", size=16, bold=True, color=BRAND_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["#", "Station (EN)", "Station (AR)", "Classification",
               "Altitude (MASL)", "Latitude", "Longitude"]
    write_header(ws, 3, headers)

    sorted_stations = sorted(stations, key=lambda s: s["index"])
    for i, s in enumerate(sorted_stations):
        row = 4 + i
        ws.cell(row=row, column=1, value=s["index"])
        ws.cell(row=row, column=2, value=s["stationairq"])
        ws.cell(row=row, column=3, value=s["stationairqar"])
        ws.cell(row=row, column=4, value=s["stationclass"])
        ws.cell(row=row, column=5, value=s["stationaltitude"])
        ws.cell(row=row, column=6, value=s["geo_point_2d"]["lat"])
        ws.cell(row=row, column=7, value=s["geo_point_2d"]["lon"])
        for c in range(1, 8):
            cell = ws.cell(row=row, column=c)
            cell.font = body_font()
            cell.border = BORDER
            if c == 1:
                cell.alignment = Alignment(horizontal="center")
            elif c in (5, 6, 7):
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0.0000" if c in (6, 7) else "0"
            else:
                cell.alignment = Alignment(horizontal="left", wrap_text=True)
        # Zebra
        if i % 2 == 0:
            for c in range(1, 8):
                ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=GREY)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.freeze_panes = "A4"


# ===== Time-series flattening =====
def flatten_series(aq: dict, wx: dict) -> list[dict]:
    """Return list of {time, pm25, pm10, no2, so2, o3, co, aqi, ...wx...} rows."""
    a_h = aq["hourly"]
    w_h = wx["hourly"]
    a_idx = {t: i for i, t in enumerate(a_h["time"])}
    rows = []
    for i, t in enumerate(w_h["time"]):
        ai = a_idx.get(t)
        if ai is None:
            continue
        row = {
            "time": t,
            "pm25": a_h["pm2_5"][ai],
            "pm10": a_h["pm10"][ai],
            "no2": a_h["nitrogen_dioxide"][ai],
            "so2": a_h["sulphur_dioxide"][ai],
            "o3": a_h["ozone"][ai],
            "co_ugm3": a_h["carbon_monoxide"][ai],
            "co_mgm3": (a_h["carbon_monoxide"][ai] / 1000.0
                        if a_h["carbon_monoxide"][ai] is not None else None),
            "dust": a_h["dust"][ai],
            "uv": a_h["uv_index"][ai],
            "aqi": a_h["us_aqi"][ai],
            "aqi_pm25": a_h["us_aqi_pm2_5"][ai],
            "aqi_pm10": a_h["us_aqi_pm10"][ai],
            "aqi_no2": a_h["us_aqi_no2"][ai],
            "aqi_so2": a_h["us_aqi_so2"][ai],
            "aqi_o3": a_h["us_aqi_o3"][ai],
            "aqi_co": a_h["us_aqi_co"][ai],
            "temp": w_h["temperature_2m"][i],
            "humidity": w_h["relative_humidity_2m"][i],
            "apparent": w_h["apparent_temperature"][i],
            "wind": w_h["wind_speed_10m"][i],
            "wind_dir": w_h["wind_direction_10m"][i],
            "wind_gust": w_h["wind_gusts_10m"][i],
            "pressure": w_h["surface_pressure"][i],
            "precip": w_h["precipitation"][i],
            "cloud": w_h["cloud_cover"][i],
        }
        rows.append(row)
    return rows


HOURLY_HEADERS = [
    "Timestamp (Local)",
    "PM2.5 µg/m³", "PM10 µg/m³",
    "NO₂ µg/m³", "SO₂ µg/m³",
    "O₃ µg/m³", "CO mg/m³",
    "Dust µg/m³", "UV Index",
    "US AQI", "AQI PM2.5", "AQI PM10",
    "AQI NO₂", "AQI SO₂", "AQI O₃", "AQI CO",
    "Temp °C", "Humidity %", "Apparent °C",
    "Wind km/h", "Wind dir °", "Gust km/h",
    "Pressure hPa", "Precip mm", "Cloud %",
]


def build_station_hourly(wb: Workbook, station_name: str, idx: int, rows: list[dict]):
    sheet_name = f"{idx:02d} {station_name}"[:31]
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HOURLY_HEADERS))
    t = ws.cell(row=1, column=1,
                value=f"{station_name} — Hourly readings, 1 Jan – 31 May 2026")
    t.font = Font(name="Calibri", size=14, bold=True, color=BRAND_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    write_header(ws, 3, HOURLY_HEADERS)

    keys = ["pm25", "pm10", "no2", "so2", "o3", "co_mgm3",
            "dust", "uv",
            "aqi", "aqi_pm25", "aqi_pm10",
            "aqi_no2", "aqi_so2", "aqi_o3", "aqi_co",
            "temp", "humidity", "apparent",
            "wind", "wind_dir", "wind_gust",
            "pressure", "precip", "cloud"]

    for r_off, row in enumerate(rows):
        r = 4 + r_off
        # Timestamp
        try:
            dt = datetime.fromisoformat(row["time"])
            tc = ws.cell(row=r, column=1, value=dt)
            tc.number_format = "yyyy-mm-dd hh:mm"
        except Exception:
            tc = ws.cell(row=r, column=1, value=row["time"])
        tc.font = body_font()
        tc.alignment = Alignment(horizontal="center")

        for ci, k in enumerate(keys):
            v = row[k]
            cell = ws.cell(row=r, column=2 + ci, value=v)
            cell.font = body_font()
            cell.alignment = Alignment(horizontal="right")
            cell.number_format = "0.0"
            if k.startswith("aqi"):
                cell.number_format = "0"
                fill = aqi_fill(v)
                if fill:
                    cell.fill = fill
            if k in ("humidity", "cloud"):
                cell.number_format = "0"

    ws.freeze_panes = "B4"

    ws.column_dimensions["A"].width = 19
    for i in range(2, len(HOURLY_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11


def safe_mean(vals):
    xs = [v for v in vals if v is not None]
    return statistics.fmean(xs) if xs else None


def build_daily_summary(wb: Workbook, station_rows: dict[str, list[dict]]):
    ws = wb.create_sheet("Daily Summary")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:K1")
    t = ws.cell(row=1, column=1, value="Daily Summary — Per-station daily averages")
    t.font = Font(name="Calibri", size=14, bold=True, color=BRAND_DARK)
    ws.row_dimensions[1].height = 26

    headers = ["Station", "Date",
               "Avg PM2.5", "Max PM2.5",
               "Avg PM10", "Max PM10",
               "Avg AQI", "Max AQI",
               "Avg Temp °C", "Avg Humidity %",
               "Avg Wind km/h"]
    write_header(ws, 3, headers)

    r = 4
    for station, rows in station_rows.items():
        by_day: dict[str, list[dict]] = {}
        for row in rows:
            day = row["time"][:10]
            by_day.setdefault(day, []).append(row)
        days = sorted(by_day.keys())
        for day in days:
            day_rows = by_day[day]
            pm25 = [d["pm25"] for d in day_rows]
            pm10 = [d["pm10"] for d in day_rows]
            aqi = [d["aqi"] for d in day_rows]
            ws.cell(row=r, column=1, value=station).font = body_font()
            try:
                dt = datetime.fromisoformat(day)
                dc = ws.cell(row=r, column=2, value=dt)
                dc.number_format = "yyyy-mm-dd"
            except Exception:
                ws.cell(row=r, column=2, value=day)
            ws.cell(row=r, column=3, value=safe_mean(pm25))
            ws.cell(row=r, column=4, value=max((v for v in pm25 if v is not None), default=None))
            ws.cell(row=r, column=5, value=safe_mean(pm10))
            ws.cell(row=r, column=6, value=max((v for v in pm10 if v is not None), default=None))
            avg_aqi = safe_mean(aqi)
            ws.cell(row=r, column=7, value=avg_aqi)
            max_aqi = max((v for v in aqi if v is not None), default=None)
            ws.cell(row=r, column=8, value=max_aqi)
            ws.cell(row=r, column=9, value=safe_mean([d["temp"] for d in day_rows]))
            ws.cell(row=r, column=10, value=safe_mean([d["humidity"] for d in day_rows]))
            ws.cell(row=r, column=11, value=safe_mean([d["wind"] for d in day_rows]))

            for c in range(3, 12):
                cell = ws.cell(row=r, column=c)
                cell.font = body_font()
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0.0"
                if c in (7, 8, 10):
                    cell.number_format = "0"
            ws.cell(row=r, column=1).font = body_font()
            ws.cell(row=r, column=1).alignment = Alignment(horizontal="left")
            ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")

            # Color AQI cells
            for c in (7, 8):
                f = aqi_fill(ws.cell(row=r, column=c).value)
                if f:
                    ws.cell(row=r, column=c).fill = f
            r += 1

    ws.freeze_panes = "C4"
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 13
    for i in range(3, 12):
        ws.column_dimensions[get_column_letter(i)].width = 13


def build_monthly_summary(wb: Workbook, station_rows: dict[str, list[dict]]):
    ws = wb.create_sheet("Monthly Summary")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    t = ws.cell(row=1, column=1, value="Monthly Summary — Cross-station monthly averages")
    t.font = Font(name="Calibri", size=14, bold=True, color=BRAND_DARK)
    ws.row_dimensions[1].height = 26

    headers = ["Station", "Month",
               "Avg PM2.5", "Avg PM10",
               "Avg NO₂", "Avg O₃",
               "Avg AQI", "Max AQI"]
    write_header(ws, 3, headers)

    r = 4
    for station, rows in station_rows.items():
        by_month: dict[str, list[dict]] = {}
        for row in rows:
            m = row["time"][:7]
            by_month.setdefault(m, []).append(row)
        for m in sorted(by_month.keys()):
            month_rows = by_month[m]
            ws.cell(row=r, column=1, value=station)
            ws.cell(row=r, column=2, value=m)
            ws.cell(row=r, column=3, value=safe_mean([d["pm25"] for d in month_rows]))
            ws.cell(row=r, column=4, value=safe_mean([d["pm10"] for d in month_rows]))
            ws.cell(row=r, column=5, value=safe_mean([d["no2"] for d in month_rows]))
            ws.cell(row=r, column=6, value=safe_mean([d["o3"] for d in month_rows]))
            avg_aqi = safe_mean([d["aqi"] for d in month_rows])
            max_aqi = max((d["aqi"] for d in month_rows if d["aqi"] is not None), default=None)
            ws.cell(row=r, column=7, value=avg_aqi)
            ws.cell(row=r, column=8, value=max_aqi)

            for c in range(1, 9):
                cell = ws.cell(row=r, column=c)
                cell.font = body_font()
                cell.border = BORDER
                if c >= 3:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "0.0" if c <= 6 else "0"
                else:
                    cell.alignment = Alignment(horizontal="left" if c == 1 else "center")
            for c in (7, 8):
                f = aqi_fill(ws.cell(row=r, column=c).value)
                if f:
                    ws.cell(row=r, column=c).fill = f
            r += 1

    ws.freeze_panes = "C4"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 10
    for i in range(3, 9):
        ws.column_dimensions[get_column_letter(i)].width = 12


def build_compliance(wb: Workbook, station_rows: dict[str, list[dict]]):
    ws = wb.create_sheet("Compliance")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:G1")
    t = ws.cell(row=1, column=1,
                value="Compliance — Period averages vs WHO 2021 & NCEC Saudi limits")
    t.font = Font(name="Calibri", size=14, bold=True, color=BRAND_DARK)
    ws.row_dimensions[1].height = 26

    headers = ["Station", "Pollutant", "Period Avg",
               "WHO Limit", "NCEC Limit",
               "vs WHO", "vs NCEC"]
    write_header(ws, 3, headers)

    pollutants = [
        ("PM2.5 (µg/m³)", "pm25"),
        ("PM10 (µg/m³)", "pm10"),
        ("NO2 (µg/m³)", "no2"),
        ("SO2 (µg/m³)", "so2"),
        ("O3 (µg/m³)", "o3"),
        ("CO (mg/m³)", "co_mgm3"),
    ]

    pass_fill = PatternFill("solid", fgColor="C8E6C9")
    fail_fill = PatternFill("solid", fgColor="FFCDD2")

    r = 4
    for station, rows in station_rows.items():
        for label, key in pollutants:
            avg = safe_mean([d[key] for d in rows])
            who = WHO_LIMITS[label]
            ncec = NCEC_LIMITS[label]
            ws.cell(row=r, column=1, value=station)
            ws.cell(row=r, column=2, value=label)
            ws.cell(row=r, column=3, value=avg)
            ws.cell(row=r, column=4, value=who)
            ws.cell(row=r, column=5, value=ncec)
            who_status = "PASS" if avg is not None and avg <= who else "EXCEED"
            ncec_status = "PASS" if avg is not None and avg <= ncec else "EXCEED"
            ws.cell(row=r, column=6, value=who_status)
            ws.cell(row=r, column=7, value=ncec_status)

            for c in range(1, 8):
                cell = ws.cell(row=r, column=c)
                cell.font = body_font()
                cell.border = BORDER
                if c >= 3:
                    cell.alignment = Alignment(horizontal="right")
                    if c == 3:
                        cell.number_format = "0.00"
                    elif c in (4, 5):
                        cell.number_format = "0"
                    else:
                        cell.alignment = Alignment(horizontal="center")
            # Status colors
            ws.cell(row=r, column=6).fill = pass_fill if who_status == "PASS" else fail_fill
            ws.cell(row=r, column=6).font = Font(bold=True, color="1B5E20" if who_status == "PASS" else "B71C1C")
            ws.cell(row=r, column=7).fill = pass_fill if ncec_status == "PASS" else fail_fill
            ws.cell(row=r, column=7).font = Font(bold=True, color="1B5E20" if ncec_status == "PASS" else "B71C1C")
            r += 1

    ws.freeze_panes = "C4"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 22
    for i in range(3, 8):
        ws.column_dimensions[get_column_letter(i)].width = 14


# ===== Main =====
def main() -> None:
    stations = load_stations()
    stations = sorted(stations, key=lambda s: s["index"])

    # Load + flatten data
    station_rows: dict[str, list[dict]] = {}
    total_rows = 0
    for s in stations:
        slug = slugify(s["stationairq"])
        try:
            aq, wx = load_series(slug)
        except FileNotFoundError as e:
            print(f"  !! Missing cached data for {s['stationairq']}: {e}", file=sys.stderr)
            continue
        rows = flatten_series(aq, wx)
        station_rows[s["stationairq"]] = rows
        total_rows += len(rows)
        print(f"  {s['stationairq']}: {len(rows)} rows")

    print(f"Total hourly rows: {total_rows}")
    print("Building workbook...")

    wb = Workbook()
    # Remove default sheet
    default = wb.active
    wb.remove(default)

    generated = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    build_cover(wb, stations, generated, total_rows)
    build_stations(wb, stations)
    build_monthly_summary(wb, station_rows)
    build_daily_summary(wb, station_rows)
    build_compliance(wb, station_rows)

    for s in stations:
        name = s["stationairq"]
        if name not in station_rows:
            continue
        build_station_hourly(wb, name, s["index"], station_rows[name])

    wb.save(OUTPUT)
    print(f"Wrote {OUTPUT} ({OUTPUT.stat().st_size / 1024:.1f} KiB)")


if __name__ == "__main__":
    main()
