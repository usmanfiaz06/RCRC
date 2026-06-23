"""
Build a focused Wadi Hanifa workbook covering 2026-01-01 to 2026-05-31.

Sheets:
  1. Cover           - station identity card + headline statistics
  2. Hourly          - 3,624 hourly readings (25 columns, 2-decimal precision)
  3. Daily Averages  - 151 daily aggregates (avg/min/max/std for every variable)
  4. Monthly Summary - 5 monthly aggregates with WHO/NCEC compliance
  5. Compliance      - period averages vs WHO 2021 + NCEC limits, plus exceedance hours
  6. Pollutant Stats - distribution stats per pollutant (mean, median, p95, p99, max)

Output: wadi_hanifa_2026_jan_to_may.xlsx
"""
from __future__ import annotations
import json
import os
import statistics
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SCRATCH = Path(os.environ.get(
    "RCRC_SCRATCH",
    "/tmp/claude-0/-home-user-RCRC/3e7d6508-8244-5a9f-b105-f6f5e0138312/scratchpad/raw",
))
OUTPUT = Path("wadi_hanifa_2026_jan_to_may.xlsx")

STATION_SLUG = "wadi_hanifa"
STATION_NAME = "Wadi Hanifa"
STATION_AR = "وادي حنيفة"
STATION_CLASS = "Background Station"
STATION_ALT = 672
STATION_LAT = 24.77438
STATION_LON = 46.529476

BRAND = "1DA1F2"
BRAND_DARK = "0D6FA8"
BRAND_LIGHT = "E1F1FB"
COVER_BG = "0B2A45"
WHITE = "FFFFFF"
GREY = "F4F6F8"
DARK_TEXT = "1B2733"

AQI_BANDS = [
    (0,   50,  "00E400", "Good"),
    (51,  100, "FFFF00", "Moderate"),
    (101, 150, "FF7E00", "Unhealthy (SG)"),
    (151, 200, "FF0000", "Unhealthy"),
    (201, 300, "8F3F97", "Very Unhealthy"),
    (301, 9999, "7E0023", "Hazardous"),
]

WHO_LIMITS = {
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

THIN = Side(style="thin", color="C6CFD8")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def header_fill():
    return PatternFill("solid", fgColor=BRAND)


def write_header(ws, row, headers):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = header_fill()
        cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def body_font():
    return Font(name="Calibri", size=11, color=DARK_TEXT)


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


def zebra(ws, row, end_col):
    for c in range(1, end_col + 1):
        ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=GREY)


def fmt_cell(cell, fmt="0.00", align="right", color_aqi=False):
    cell.font = body_font()
    cell.alignment = Alignment(horizontal=align)
    cell.number_format = fmt
    cell.border = BORDER
    if color_aqi:
        f = aqi_fill(cell.value)
        if f:
            cell.fill = f


# ===== Load data =====
def load_series():
    aq = json.loads((SCRATCH / f"aq_{STATION_SLUG}.json").read_text())
    wx = json.loads((SCRATCH / f"wx_{STATION_SLUG}.json").read_text())
    return aq, wx


def flatten(aq, wx):
    a_h = aq["hourly"]
    w_h = wx["hourly"]
    a_idx = {t: i for i, t in enumerate(a_h["time"])}
    rows = []
    for i, t in enumerate(w_h["time"]):
        ai = a_idx.get(t)
        if ai is None:
            continue
        co_ug = a_h["carbon_monoxide"][ai]
        rows.append({
            "time": t,
            "pm25": a_h["pm2_5"][ai],
            "pm10": a_h["pm10"][ai],
            "no2": a_h["nitrogen_dioxide"][ai],
            "so2": a_h["sulphur_dioxide"][ai],
            "o3": a_h["ozone"][ai],
            "co_mg": (co_ug / 1000.0) if co_ug is not None else None,
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
        })
    return rows


# ===== Stats helpers =====
def _clean(vals):
    return [v for v in vals if v is not None]


def mean(vals):
    xs = _clean(vals)
    return statistics.fmean(xs) if xs else None


def stddev(vals):
    xs = _clean(vals)
    return statistics.pstdev(xs) if len(xs) >= 2 else None


def pct(vals, q):
    xs = sorted(_clean(vals))
    if not xs:
        return None
    k = max(0, min(len(xs) - 1, int(round((q / 100) * (len(xs) - 1)))))
    return xs[k]


def vmin(vals):
    xs = _clean(vals)
    return min(xs) if xs else None


def vmax(vals):
    xs = _clean(vals)
    return max(xs) if xs else None


def count_above(vals, threshold):
    return sum(1 for v in vals if v is not None and v > threshold)


# ===== Sheets =====
def build_cover(wb, rows, generated):
    ws = wb.create_sheet("Cover", 0)
    ws.sheet_view.showGridLines = False

    for r in range(1, 13):
        for c in range(1, 12):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=COVER_BG)

    ws.merge_cells("B2:K2")
    t = ws.cell(row=2, column=2, value="Wadi Hanifa — Air Quality Dataset")
    t.font = Font(name="Calibri", size=28, bold=True, color=WHITE)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 48

    ws.merge_cells("B3:K3")
    s = ws.cell(row=3, column=2,
                value=f"RCRC Station 7 · {STATION_AR} · Background Station · 672 MASL")
    s.font = Font(name="Calibri", size=14, color="9BD0F0")
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 24

    ws.merge_cells("B4:K4")
    ws.cell(row=4, column=2,
            value="Reporting period: 1 January 2026 — 31 May 2026 (151 days · hourly)") \
        .font = Font(name="Calibri", size=12, color="C6E4F4")

    ws.merge_cells("B5:K5")
    ws.cell(row=5, column=2,
            value=f"Coordinates: {STATION_LAT}°N, {STATION_LON}°E   ·   Generated: {generated}") \
        .font = Font(name="Calibri", size=11, color="C6E4F4")

    # Headline KPIs
    pm25_avg = mean([r["pm25"] for r in rows])
    pm10_avg = mean([r["pm10"] for r in rows])
    aqi_avg = mean([r["aqi"] for r in rows])
    aqi_max = vmax([r["aqi"] for r in rows])
    panels = [
        ("Total hourly rows", f"{len(rows):,}"),
        ("Avg PM2.5", f"{pm25_avg:.2f} µg/m³"),
        ("Avg PM10", f"{pm10_avg:.2f} µg/m³"),
        ("Avg / Peak AQI", f"{aqi_avg:.0f} / {aqi_max:.0f}"),
    ]
    for i, (label, value) in enumerate(panels):
        c = 2 + i * 2
        ws.merge_cells(start_row=7, start_column=c, end_row=7, end_column=c + 1)
        ws.merge_cells(start_row=8, start_column=c, end_row=9, end_column=c + 1)
        lab = ws.cell(row=7, column=c, value=label)
        lab.font = Font(name="Calibri", size=10, bold=True, color="9BD0F0")
        lab.alignment = Alignment(horizontal="center", vertical="center")
        val = ws.cell(row=8, column=c, value=value)
        val.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
        val.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 18
    ws.row_dimensions[8].height = 30
    ws.row_dimensions[9].height = 18

    # Sources
    ws.cell(row=15, column=2, value="Data sources").font = Font(
        name="Calibri", size=14, bold=True, color=BRAND_DARK)
    sources = [
        ("RCRC Open Data Portal",
         "Station identity — opendata.rcrc.gov.sa/explore/dataset/air-quality-stations-in-riyadh-2025/"),
        ("Open-Meteo Air Quality (CAMS)",
         "Hourly PM2.5, PM10, NO₂, SO₂, O₃, CO, dust, UV, US AQI — air-quality-api.open-meteo.com"),
        ("Open-Meteo Historical Weather (ERA5)",
         "Hourly temp, humidity, wind, pressure, precip, cloud — archive-api.open-meteo.com"),
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
    ws.cell(row=r + 1, column=2, value="Notes").font = Font(
        name="Calibri", size=14, bold=True, color=BRAND_DARK)
    notes = [
        "All pollutant concentrations are hourly modelled values from CAMS reanalysis sampled at the station coordinates.",
        "Carbon monoxide is reported in mg/m³ for regulatory comparison (CAMS raw values in µg/m³ ÷ 1000).",
        "All numeric columns carry two-decimal precision in the Hourly sheet; aggregates use 3-decimal precision where meaningful.",
        "Timestamps are local Riyadh time (Asia/Riyadh, GMT+3).",
        "WHO 2021 and NCEC Saudi limits are 24-hour averages unless otherwise noted.",
        "US-AQI bands: 0–50 Good · 51–100 Moderate · 101–150 USG · 151–200 Unhealthy · 201–300 Very Unhealthy · 301+ Hazardous.",
    ]
    for i, n in enumerate(notes):
        ws.merge_cells(start_row=r + 2 + i, start_column=2, end_row=r + 2 + i, end_column=11)
        cell = ws.cell(row=r + 2 + i, column=2, value=f"• {n}")
        cell.font = Font(name="Calibri", size=11, color=DARK_TEXT)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r + 2 + i].height = 20

    # AQI legend
    legend_r = r + 2 + len(notes) + 2
    ws.cell(row=legend_r, column=2, value="AQI Legend").font = Font(
        name="Calibri", size=14, bold=True, color=BRAND_DARK)
    for i, (lo, hi, color, label) in enumerate(AQI_BANDS):
        cell = ws.cell(row=legend_r + 1, column=2 + i, value=f"{label}\n{lo}–{hi}")
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[legend_r + 1].height = 36

    for letter in "ABCDEFGHIJKL":
        ws.column_dimensions[letter].width = 18
    ws.column_dimensions["A"].width = 3


HOURLY_HEADERS = [
    "Timestamp (Local)",
    "PM2.5 µg/m³", "PM10 µg/m³",
    "NO₂ µg/m³", "SO₂ µg/m³",
    "O₃ µg/m³", "CO mg/m³",
    "Dust µg/m³", "UV Index",
    "US AQI", "AQI PM2.5", "AQI PM10",
    "AQI NO₂", "AQI SO₂", "AQI O₃", "AQI CO",
    "Temp °C", "Humidity %", "Apparent °C",
    "Wind km/h", "Wind Dir °", "Gust km/h",
    "Pressure hPa", "Precip mm", "Cloud %",
]
HOURLY_KEYS = ["pm25", "pm10", "no2", "so2", "o3", "co_mg",
               "dust", "uv",
               "aqi", "aqi_pm25", "aqi_pm10",
               "aqi_no2", "aqi_so2", "aqi_o3", "aqi_co",
               "temp", "humidity", "apparent",
               "wind", "wind_dir", "wind_gust",
               "pressure", "precip", "cloud"]


def build_hourly(wb, rows):
    ws = wb.create_sheet("Hourly")
    ws.sheet_view.showGridLines = False

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HOURLY_HEADERS))
    t = ws.cell(row=1, column=1, value="Wadi Hanifa — Hourly readings · 1 Jan – 31 May 2026")
    t.font = Font(name="Calibri", size=15, bold=True, color=BRAND_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    write_header(ws, 3, HOURLY_HEADERS)

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
        tc.border = BORDER

        for ci, k in enumerate(HOURLY_KEYS):
            v = row[k]
            cell = ws.cell(row=r, column=2 + ci, value=v)
            if k.startswith("aqi"):
                fmt_cell(cell, "0", color_aqi=True)
            elif k in ("humidity", "cloud", "wind_dir"):
                fmt_cell(cell, "0")
            elif k == "uv":
                fmt_cell(cell, "0.00")
            else:
                fmt_cell(cell, "0.00")

    ws.freeze_panes = "B4"
    ws.column_dimensions["A"].width = 19
    for i in range(2, len(HOURLY_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 12


DAILY_HEADERS = [
    "Date", "Hours", "DOW",
    "PM2.5 Avg", "PM2.5 Min", "PM2.5 Max", "PM2.5 Std",
    "PM10 Avg", "PM10 Min", "PM10 Max", "PM10 Std",
    "NO₂ Avg", "NO₂ Max", "SO₂ Avg", "SO₂ Max",
    "O₃ Avg", "O₃ Max", "CO Avg", "CO Max",
    "AQI Avg", "AQI Max", "AQI Min",
    "Temp Avg °C", "Temp Min °C", "Temp Max °C",
    "Humidity Avg %", "Wind Avg km/h", "Wind Max km/h",
    "Pressure Avg hPa", "Precip Total mm",
]


def build_daily(wb, rows):
    ws = wb.create_sheet("Daily Averages")
    ws.sheet_view.showGridLines = False

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(DAILY_HEADERS))
    t = ws.cell(row=1, column=1,
                value="Wadi Hanifa — Daily averages with min/max/std · 1 Jan – 31 May 2026")
    t.font = Font(name="Calibri", size=15, bold=True, color=BRAND_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    write_header(ws, 3, DAILY_HEADERS)

    by_day = {}
    for row in rows:
        d = row["time"][:10]
        by_day.setdefault(d, []).append(row)
    days = sorted(by_day.keys())

    for i, d in enumerate(days):
        r = 4 + i
        day_rows = by_day[d]
        dt = datetime.fromisoformat(d)
        dow = dt.strftime("%a")

        pm25 = [x["pm25"] for x in day_rows]
        pm10 = [x["pm10"] for x in day_rows]
        no2 = [x["no2"] for x in day_rows]
        so2 = [x["so2"] for x in day_rows]
        o3 = [x["o3"] for x in day_rows]
        co = [x["co_mg"] for x in day_rows]
        aqi = [x["aqi"] for x in day_rows]
        temp = [x["temp"] for x in day_rows]
        hum = [x["humidity"] for x in day_rows]
        wind = [x["wind"] for x in day_rows]
        press = [x["pressure"] for x in day_rows]
        precip = [x["precip"] for x in day_rows]

        ws.cell(row=r, column=1, value=dt).number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=2, value=len(day_rows))
        ws.cell(row=r, column=3, value=dow)
        vals = [
            mean(pm25), vmin(pm25), vmax(pm25), stddev(pm25),
            mean(pm10), vmin(pm10), vmax(pm10), stddev(pm10),
            mean(no2), vmax(no2), mean(so2), vmax(so2),
            mean(o3), vmax(o3), mean(co), vmax(co),
            mean(aqi), vmax(aqi), vmin(aqi),
            mean(temp), vmin(temp), vmax(temp),
            mean(hum), mean(wind), vmax(wind),
            mean(press), sum(p for p in precip if p is not None),
        ]
        for ci, v in enumerate(vals):
            cell = ws.cell(row=r, column=4 + ci, value=v)
            fmt = "0.00"
            # AQI columns use integer format with color
            if 4 + ci in (23, 24, 25):
                fmt = "0"
                fmt_cell(cell, fmt, color_aqi=True)
            elif 4 + ci in (26,):  # humidity
                fmt_cell(cell, "0")
            else:
                fmt_cell(cell, fmt)

        ws.cell(row=r, column=1).font = body_font()
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=1).border = BORDER
        ws.cell(row=r, column=2).font = body_font()
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2).border = BORDER
        ws.cell(row=r, column=3).font = body_font()
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3).border = BORDER
        if i % 2 == 0:
            for c in (1, 2, 3):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=GREY)

    ws.freeze_panes = "D4"
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 7
    ws.column_dimensions["C"].width = 6
    for i in range(4, len(DAILY_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 12


def build_monthly(wb, rows):
    ws = wb.create_sheet("Monthly Summary")
    ws.sheet_view.showGridLines = False

    headers = ["Month", "Days", "Hours",
               "PM2.5 Avg", "PM2.5 Max",
               "PM10 Avg", "PM10 Max",
               "NO₂ Avg", "SO₂ Avg",
               "O₃ Avg", "CO Avg",
               "AQI Avg", "AQI Max",
               "Temp Avg", "Humidity Avg",
               "Wind Avg", "Precip Total"]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    t = ws.cell(row=1, column=1, value="Wadi Hanifa — Monthly summary")
    t.font = Font(name="Calibri", size=15, bold=True, color=BRAND_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    write_header(ws, 3, headers)

    by_month = {}
    for row in rows:
        m = row["time"][:7]
        by_month.setdefault(m, []).append(row)

    for i, m in enumerate(sorted(by_month.keys())):
        r = 4 + i
        month_rows = by_month[m]
        days = len({x["time"][:10] for x in month_rows})
        pm25 = [x["pm25"] for x in month_rows]
        pm10 = [x["pm10"] for x in month_rows]

        vals = [
            m, days, len(month_rows),
            mean(pm25), vmax(pm25),
            mean(pm10), vmax(pm10),
            mean([x["no2"] for x in month_rows]), mean([x["so2"] for x in month_rows]),
            mean([x["o3"] for x in month_rows]), mean([x["co_mg"] for x in month_rows]),
            mean([x["aqi"] for x in month_rows]), vmax([x["aqi"] for x in month_rows]),
            mean([x["temp"] for x in month_rows]), mean([x["humidity"] for x in month_rows]),
            mean([x["wind"] for x in month_rows]),
            sum(x["precip"] for x in month_rows if x["precip"] is not None),
        ]
        for ci, v in enumerate(vals):
            cell = ws.cell(row=r, column=1 + ci, value=v)
            cell.font = body_font()
            cell.border = BORDER
            if ci == 0:
                cell.alignment = Alignment(horizontal="center")
            elif ci in (1, 2):
                cell.alignment = Alignment(horizontal="center")
                cell.number_format = "0"
            elif ci in (11, 12):  # AQI Avg / Max
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0"
                f = aqi_fill(v)
                if f:
                    cell.fill = f
            elif ci == 14:  # humidity %
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0"
            else:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0.000"

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 7
    ws.column_dimensions["C"].width = 7
    for i in range(4, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13


def build_compliance(wb, rows):
    ws = wb.create_sheet("Compliance")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    t = ws.cell(row=1, column=1,
                value="Wadi Hanifa — Compliance vs WHO 2021 & NCEC Saudi limits")
    t.font = Font(name="Calibri", size=15, bold=True, color=BRAND_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["Pollutant", "Period Avg", "WHO Limit", "NCEC Limit",
               "vs WHO", "vs NCEC", "Hrs > WHO", "Hrs > NCEC"]
    write_header(ws, 3, headers)

    pollutants = [
        ("PM2.5 (µg/m³)", "pm25"),
        ("PM10 (µg/m³)", "pm10"),
        ("NO2 (µg/m³)", "no2"),
        ("SO2 (µg/m³)", "so2"),
        ("O3 (µg/m³)", "o3"),
        ("CO (mg/m³)", "co_mg"),
    ]

    pass_fill = PatternFill("solid", fgColor="C8E6C9")
    fail_fill = PatternFill("solid", fgColor="FFCDD2")

    for i, (label, key) in enumerate(pollutants):
        r = 4 + i
        avg = mean([x[key] for x in rows])
        who = WHO_LIMITS[label]
        ncec = NCEC_LIMITS[label]
        who_status = "PASS" if avg is not None and avg <= who else "EXCEED"
        ncec_status = "PASS" if avg is not None and avg <= ncec else "EXCEED"
        hrs_who = count_above([x[key] for x in rows], who)
        hrs_ncec = count_above([x[key] for x in rows], ncec)

        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=avg)
        ws.cell(row=r, column=3, value=who)
        ws.cell(row=r, column=4, value=ncec)
        ws.cell(row=r, column=5, value=who_status)
        ws.cell(row=r, column=6, value=ncec_status)
        ws.cell(row=r, column=7, value=hrs_who)
        ws.cell(row=r, column=8, value=hrs_ncec)

        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.font = body_font()
            cell.border = BORDER
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="left")
        ws.cell(row=r, column=2).number_format = "0.000"
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=3).number_format = "0"
        ws.cell(row=r, column=4).number_format = "0"
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=4).alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=7).number_format = "0"
        ws.cell(row=r, column=8).number_format = "0"
        ws.cell(row=r, column=7).alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=8).alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=6).alignment = Alignment(horizontal="center")

        ws.cell(row=r, column=5).fill = pass_fill if who_status == "PASS" else fail_fill
        ws.cell(row=r, column=5).font = Font(bold=True,
                                             color="1B5E20" if who_status == "PASS" else "B71C1C")
        ws.cell(row=r, column=6).fill = pass_fill if ncec_status == "PASS" else fail_fill
        ws.cell(row=r, column=6).font = Font(bold=True,
                                             color="1B5E20" if ncec_status == "PASS" else "B71C1C")

    ws.column_dimensions["A"].width = 20
    for i in range(2, 9):
        ws.column_dimensions[get_column_letter(i)].width = 14


def build_distribution(wb, rows):
    ws = wb.create_sheet("Pollutant Stats")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:I1")
    t = ws.cell(row=1, column=1,
                value="Wadi Hanifa — Pollutant distribution statistics (hourly)")
    t.font = Font(name="Calibri", size=15, bold=True, color=BRAND_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["Variable", "Unit", "Count", "Mean", "Std", "Min",
               "Median (p50)", "p95", "p99", "Max"]
    write_header(ws, 3, headers)

    items = [
        ("PM2.5", "µg/m³", "pm25"),
        ("PM10", "µg/m³", "pm10"),
        ("NO₂", "µg/m³", "no2"),
        ("SO₂", "µg/m³", "so2"),
        ("O₃", "µg/m³", "o3"),
        ("CO", "mg/m³", "co_mg"),
        ("Dust", "µg/m³", "dust"),
        ("UV Index", "—", "uv"),
        ("US AQI", "—", "aqi"),
        ("Temperature", "°C", "temp"),
        ("Humidity", "%", "humidity"),
        ("Wind Speed", "km/h", "wind"),
        ("Pressure", "hPa", "pressure"),
    ]
    for i, (label, unit, key) in enumerate(items):
        r = 4 + i
        vals = [x[key] for x in rows]
        cleaned = _clean(vals)
        ws.cell(row=r, column=1, value=label).font = Font(bold=True, color=DARK_TEXT)
        ws.cell(row=r, column=2, value=unit)
        ws.cell(row=r, column=3, value=len(cleaned))
        ws.cell(row=r, column=4, value=mean(vals))
        ws.cell(row=r, column=5, value=stddev(vals))
        ws.cell(row=r, column=6, value=vmin(vals))
        ws.cell(row=r, column=7, value=pct(vals, 50))
        ws.cell(row=r, column=8, value=pct(vals, 95))
        ws.cell(row=r, column=9, value=pct(vals, 99))
        ws.cell(row=r, column=10, value=vmax(vals))
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c)
            if c != 1:
                cell.font = body_font()
            cell.border = BORDER
            if c in (1, 2):
                cell.alignment = Alignment(horizontal="left")
            elif c == 3:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0"
            else:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = "0.000"
        if i % 2 == 0:
            for c in range(1, 11):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=GREY)

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 9
    for i in range(3, 11):
        ws.column_dimensions[get_column_letter(i)].width = 13


# ===== Main =====
def main():
    aq, wx = load_series()
    rows = flatten(aq, wx)
    print(f"Loaded {len(rows)} hourly rows for {STATION_NAME}")

    wb = Workbook()
    wb.remove(wb.active)

    generated = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    build_cover(wb, rows, generated)
    build_hourly(wb, rows)
    build_daily(wb, rows)
    build_monthly(wb, rows)
    build_compliance(wb, rows)
    build_distribution(wb, rows)

    wb.save(OUTPUT)
    print(f"Wrote {OUTPUT} ({OUTPUT.stat().st_size / 1024:.1f} KiB)")


if __name__ == "__main__":
    main()
