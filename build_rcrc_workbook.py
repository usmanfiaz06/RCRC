"""
Build an Excel workbook of 18 months of air-quality + weather data for the
Royal Commission for Riyadh City (RCRC) air-quality monitoring stations.

Station registry: RCRC Open Data Portal
  https://opendata.rcrc.gov.sa/api/explore/v2.1/catalog/datasets/
  air-quality-stations-in-riyadh-2025/records

Pollutant + AQI readings: Copernicus Atmosphere Monitoring Service (CAMS) global
reanalysis, served by Open-Meteo (https://open-meteo.com/en/docs/air-quality-api).

Weather: ECMWF ERA5 reanalysis, served by Open-Meteo
  (https://open-meteo.com/en/docs/historical-weather-api).

The CAMS reanalysis is model+satellite derived at ~10 km resolution; values
sampled at each RCRC station's coordinates approximate the air over that
location, not the station's own instrument readings. This is documented
prominently on the Cover sheet of the workbook.
"""

from __future__ import annotations

import json
import time
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import requests
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

START_DATE = date(2024, 11, 17)
END_DATE = date(2026, 5, 17)
OUTPUT_PATH = Path(__file__).parent / "rcrc_stations_18mo.xlsx"

RCRC_DATASET_URL = (
    "https://opendata.rcrc.gov.sa/api/explore/v2.1/catalog/datasets/"
    "air-quality-stations-in-riyadh-2025/records"
)
OM_AIR_URL = "https://air-quality-api.open-meteo.com/v1/air-quality"
OM_WX_URL = "https://archive-api.open-meteo.com/v1/archive"

AIR_HOURLY = [
    "pm10",
    "pm2_5",
    "carbon_monoxide",
    "nitrogen_dioxide",
    "sulphur_dioxide",
    "ozone",
    "us_aqi",
]
WX_HOURLY = [
    "temperature_2m",
    "relative_humidity_2m",
    "wind_speed_10m",
    "wind_direction_10m",
    "surface_pressure",
]

# Brand palette (Royal Commission green tones + AQI bands)
BRAND_PRIMARY = "0E5C3A"   # deep green header bar
BRAND_ACCENT = "1F9D55"    # accent green
BRAND_LIGHT = "E7F4ED"     # zebra row fill
HEADER_FONT_COLOR = "FFFFFF"

AQI_BANDS = [
    (0, 50, "00E400", "000000", "Good"),
    (51, 100, "FFFF00", "000000", "Moderate"),
    (101, 150, "FF7E00", "000000", "Unhealthy for Sensitive Groups"),
    (151, 200, "FF0000", "FFFFFF", "Unhealthy"),
    (201, 300, "8F3F97", "FFFFFF", "Very Unhealthy"),
    (301, 500, "7E0023", "FFFFFF", "Hazardous"),
]

# WHO 2021 AQGs / Saudi NCEC ambient air-quality limits, annual mean (µg/m³)
LIMITS = {
    "pm2_5": {"WHO": 5.0, "NCEC": 15.0},
    "pm10": {"WHO": 15.0, "NCEC": 35.0},
    "no2": {"WHO": 10.0, "NCEC": 40.0},
    "so2": {"WHO": 40.0, "NCEC": 50.0},  # WHO 24h, NCEC annual approx
    "o3": {"WHO": 60.0, "NCEC": 100.0},  # peak season
    "co": {"WHO": 4000.0, "NCEC": 10000.0},  # 24-h, µg/m³
}

THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------


def http_get(url: str, params: dict[str, Any], retries: int = 5) -> dict:
    backoff = 2
    for attempt in range(retries):
        try:
            r = requests.get(url, params=params, timeout=60)
            if r.status_code == 200:
                return r.json()
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(backoff)
                backoff *= 2
                continue
            r.raise_for_status()
        except requests.RequestException as exc:
            if attempt == retries - 1:
                raise
            print(f"  retry {attempt + 1}/{retries}: {exc}")
            time.sleep(backoff)
            backoff *= 2
    raise RuntimeError(f"GET {url} failed after {retries} retries")


# ---------------------------------------------------------------------------
# RCRC station registry
# ---------------------------------------------------------------------------


def fetch_stations() -> list[dict]:
    print("Fetching RCRC station registry...")
    out: list[dict] = []
    offset = 0
    while True:
        data = http_get(RCRC_DATASET_URL, {"limit": 100, "offset": offset})
        results = data.get("results", [])
        out.extend(results)
        if len(results) < 100:
            break
        offset += 100
    out.sort(key=lambda s: s.get("index") or 0)
    print(f"  {len(out)} stations retrieved")
    return out


# ---------------------------------------------------------------------------
# Open-Meteo CAMS + ERA5
# ---------------------------------------------------------------------------


def chunk_dates(start: date, end: date, months: int = 6) -> list[tuple[date, date]]:
    """Yield (chunk_start, chunk_end) windows so the API doesn't time out."""
    chunks = []
    cursor = start
    while cursor <= end:
        nxt = cursor + timedelta(days=months * 30)
        if nxt > end:
            nxt = end
        chunks.append((cursor, nxt))
        cursor = nxt + timedelta(days=1)
    return chunks


def fetch_hourly(url: str, lat: float, lon: float, vars_: list[str]) -> dict[str, list]:
    merged: dict[str, list] = {"time": []}
    for v in vars_:
        merged[v] = []
    for cs, ce in chunk_dates(START_DATE, END_DATE):
        params = {
            "latitude": lat,
            "longitude": lon,
            "hourly": ",".join(vars_),
            "start_date": cs.isoformat(),
            "end_date": ce.isoformat(),
            "timezone": "Asia/Riyadh",
        }
        data = http_get(url, params)
        hourly = data.get("hourly") or {}
        if not hourly.get("time"):
            continue
        for k in merged:
            merged[k].extend(hourly.get(k, []))
    return merged


def fetch_station_data(lat: float, lon: float) -> list[dict]:
    """Return one row per hour with pollutants + weather aligned by timestamp."""
    air = fetch_hourly(OM_AIR_URL, lat, lon, AIR_HOURLY)
    wx = fetch_hourly(OM_WX_URL, lat, lon, WX_HOURLY)
    wx_idx = {t: i for i, t in enumerate(wx.get("time", []))}
    rows: list[dict] = []
    for i, t in enumerate(air.get("time", [])):
        wi = wx_idx.get(t)
        rows.append(
            {
                "time": t,
                "pm2_5": air["pm2_5"][i],
                "pm10": air["pm10"][i],
                "o3": air["ozone"][i],
                "no2": air["nitrogen_dioxide"][i],
                "so2": air["sulphur_dioxide"][i],
                "co": air["carbon_monoxide"][i],
                "us_aqi": air["us_aqi"][i],
                "temperature": wx["temperature_2m"][wi] if wi is not None else None,
                "humidity": wx["relative_humidity_2m"][wi] if wi is not None else None,
                "wind_speed": wx["wind_speed_10m"][wi] if wi is not None else None,
                "wind_dir": wx["wind_direction_10m"][wi] if wi is not None else None,
                "pressure": wx["surface_pressure"][wi] if wi is not None else None,
            }
        )
    return rows


# ---------------------------------------------------------------------------
# Workbook styling helpers
# ---------------------------------------------------------------------------


def style_header_row(ws, row: int, ncols: int, fill: str = BRAND_PRIMARY) -> None:
    fill_obj = PatternFill("solid", fgColor=fill)
    font = Font(bold=True, color=HEADER_FONT_COLOR, size=11)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill_obj
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 28


def autosize(ws, widths: dict[int, int]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def add_aqi_conditional(ws, col_letter: str, first: int, last: int) -> None:
    rng = f"{col_letter}{first}:{col_letter}{last}"
    for lo, hi, fg, fc, _label in AQI_BANDS:
        ws.conditional_formatting.add(
            rng,
            CellIsRule(
                operator="between",
                formula=[str(lo), str(hi)],
                fill=PatternFill("solid", fgColor=fg),
                font=Font(color=fc, bold=False),
            ),
        )


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------


def build_cover(wb: Workbook, stations: list[dict], total_records: int) -> None:
    ws = wb.create_sheet("Cover", 0)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Royal Commission for Riyadh City"
    ws["A1"].font = Font(bold=True, size=22, color=BRAND_PRIMARY)
    ws["A2"] = "Air Quality Monitoring Stations — 18-month dataset"
    ws["A2"].font = Font(bold=True, size=14, color=BRAND_ACCENT)
    ws.merge_cells("A1:F1")
    ws.merge_cells("A2:F2")

    rows = [
        ("Date range", f"{START_DATE.isoformat()} → {END_DATE.isoformat()}"),
        ("Generation timestamp", datetime.now().isoformat(timespec="seconds")),
        ("Station count", len(stations)),
        ("Total hourly records", total_records),
        ("Station registry", "RCRC Open Data Portal — air-quality-stations-in-riyadh-2025"),
        ("Registry URL", "https://opendata.rcrc.gov.sa/explore/dataset/air-quality-stations-in-riyadh-2025/"),
        ("Readings source", "Copernicus Atmosphere Monitoring Service (CAMS) reanalysis via Open-Meteo"),
        ("Readings URL", "https://open-meteo.com/en/docs/air-quality-api"),
        ("Weather source", "ECMWF ERA5 reanalysis via Open-Meteo"),
        ("Weather URL", "https://open-meteo.com/en/docs/historical-weather-api"),
        ("Timezone", "Asia/Riyadh (GMT+3)"),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws.cell(row=i, column=2, value=v)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=6)

    note_row = 4 + len(rows) + 1
    ws.cell(
        row=note_row, column=1,
        value="Important note on data provenance"
    ).font = Font(bold=True, size=12, color="A04000")
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)

    note_body = (
        "RCRC publishes the station registry only — its open data portal does not expose "
        "historical pollutant readings. The hourly pollutant + weather values in this workbook "
        "are sampled from Copernicus CAMS / ECMWF ERA5 global reanalysis at each RCRC station's "
        "coordinates. CAMS data is model + satellite derived at ~10 km horizontal resolution, so "
        "stations close to each other in Riyadh will have similar values. Use these readings as a "
        "regional indicator, not as the stations' own instrument observations."
    )
    ws.cell(row=note_row + 1, column=1, value=note_body).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=note_row + 1, start_column=1, end_row=note_row + 5, end_column=6)

    autosize(ws, {1: 26, 2: 22, 3: 22, 4: 22, 5: 22, 6: 22})


def build_stations(wb: Workbook, stations: list[dict]) -> None:
    ws = wb.create_sheet("Stations")
    headers = [
        "Index", "Station (EN)", "Station (AR)", "Classification (EN)",
        "Classification (AR)", "Altitude (m)", "Latitude", "Longitude",
        "Location (EN)", "Location (AR)",
    ]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    for s in stations:
        geo = s.get("geo_point_2d") or {}
        ws.append([
            s.get("index"),
            s.get("stationairq"),
            s.get("stationairqar"),
            s.get("stationclass"),
            s.get("stationclassar"),
            s.get("stationaltitude"),
            geo.get("lat"),
            geo.get("lon"),
            s.get("stationairqlocation"),
            s.get("stationairqlocationar"),
        ])
    ws.freeze_panes = "A2"
    last = ws.max_row
    for r in range(2, last + 1):
        if r % 2 == 0:
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=BRAND_LIGHT)
        ws.cell(row=r, column=7).number_format = "0.0000"
        ws.cell(row=r, column=8).number_format = "0.0000"
    autosize(ws, {1: 7, 2: 28, 3: 28, 4: 18, 5: 22, 6: 12, 7: 11, 8: 11, 9: 60, 10: 60})

    table = Table(displayName="Stations", ref=f"A1:{get_column_letter(len(headers))}{last}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=False, showColumnStripes=False
    )
    ws.add_table(table)


def safe_sheet_name(name: str) -> str:
    bad = '[]:*?/\\'
    out = "".join("_" if ch in bad else ch for ch in name)
    return out[:31]


def build_station_sheet(wb: Workbook, station: dict, rows: list[dict]) -> None:
    name = safe_sheet_name(station["stationairq"])
    ws = wb.create_sheet(name)
    headers = [
        "Timestamp (Riyadh)", "PM2.5 (µg/m³)", "PM10 (µg/m³)", "O₃ (µg/m³)",
        "NO₂ (µg/m³)", "SO₂ (µg/m³)", "CO (µg/m³)", "US AQI",
        "Temp (°C)", "Humidity (%)", "Wind (km/h)", "Wind dir (°)", "Pressure (hPa)",
    ]
    title = f"{station['stationairq']}  ({station.get('stationclass')})"
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=14, color=BRAND_PRIMARY)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.row_dimensions[1].height = 22

    for c, h in enumerate(headers, start=1):
        ws.cell(row=2, column=c, value=h)
    style_header_row(ws, 2, len(headers))

    for i, r in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=r["time"])
        ws.cell(row=i, column=2, value=r["pm2_5"])
        ws.cell(row=i, column=3, value=r["pm10"])
        ws.cell(row=i, column=4, value=r["o3"])
        ws.cell(row=i, column=5, value=r["no2"])
        ws.cell(row=i, column=6, value=r["so2"])
        ws.cell(row=i, column=7, value=r["co"])
        ws.cell(row=i, column=8, value=r["us_aqi"])
        ws.cell(row=i, column=9, value=r["temperature"])
        ws.cell(row=i, column=10, value=r["humidity"])
        ws.cell(row=i, column=11, value=r["wind_speed"])
        ws.cell(row=i, column=12, value=r["wind_dir"])
        ws.cell(row=i, column=13, value=r["pressure"])

    last = ws.max_row
    for c in range(2, 8):
        col = get_column_letter(c)
        for r in range(3, last + 1):
            ws[f"{col}{r}"].number_format = "0.0"
    for c in (9, 10, 11, 12, 13):
        col = get_column_letter(c)
        for r in range(3, last + 1):
            ws[f"{col}{r}"].number_format = "0.0"
    for r in range(3, last + 1):
        ws.cell(row=r, column=1).number_format = "yyyy-mm-dd hh:mm"
    add_aqi_conditional(ws, "H", 3, last)

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{last}"
    autosize(ws, {
        1: 20, 2: 12, 3: 12, 4: 11, 5: 11, 6: 11, 7: 11, 8: 10,
        9: 10, 10: 11, 11: 11, 12: 11, 13: 12,
    })


def build_summary(wb: Workbook, per_station: dict[str, tuple[dict, list[dict]]]) -> None:
    ws = wb.create_sheet("Summary")
    headers = [
        "Station", "Year-Month", "Avg PM2.5", "Avg PM10", "Avg O₃", "Avg NO₂",
        "Avg SO₂", "Avg CO", "Avg US AQI", "WHO PM2.5 ✓", "NCEC PM2.5 ✓",
    ]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    for stn_name, (_meta, rows) in per_station.items():
        buckets: dict[str, dict[str, list[float]]] = defaultdict(lambda: defaultdict(list))
        for r in rows:
            ts = r["time"]
            if not ts:
                continue
            ym = ts[:7]
            for k in ("pm2_5", "pm10", "o3", "no2", "so2", "co", "us_aqi"):
                v = r.get(k)
                if v is not None:
                    buckets[ym][k].append(v)
        for ym in sorted(buckets.keys()):
            b = buckets[ym]
            def avg(k: str) -> float | None:
                return sum(b[k]) / len(b[k]) if b.get(k) else None
            avg_pm25 = avg("pm2_5")
            who_ok = "✓" if avg_pm25 is not None and avg_pm25 <= LIMITS["pm2_5"]["WHO"] else "✗"
            ncec_ok = "✓" if avg_pm25 is not None and avg_pm25 <= LIMITS["pm2_5"]["NCEC"] else "✗"
            ws.append([
                stn_name, ym, avg_pm25, avg("pm10"), avg("o3"), avg("no2"),
                avg("so2"), avg("co"), avg("us_aqi"), who_ok, ncec_ok,
            ])

    last = ws.max_row
    for c in range(3, 10):
        col = get_column_letter(c)
        for r in range(2, last + 1):
            ws[f"{col}{r}"].number_format = "0.0"
    add_aqi_conditional(ws, "I", 2, last)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last}"
    autosize(ws, {
        1: 28, 2: 11, 3: 12, 4: 12, 5: 11, 6: 11,
        7: 11, 8: 11, 9: 12, 10: 14, 11: 14,
    })


def build_legend(wb: Workbook) -> None:
    ws = wb.create_sheet("Legend")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Legend"
    ws["A1"].font = Font(bold=True, size=16, color=BRAND_PRIMARY)
    ws["A3"] = "US AQI bands"
    ws["A3"].font = Font(bold=True, size=12)
    headers = ["Range", "Category", "Health implications"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=4, column=c, value=h)
    style_header_row(ws, 4, len(headers))

    descriptions = {
        "Good": "Air quality is satisfactory; little to no risk.",
        "Moderate": "Acceptable; minor risk for unusually sensitive individuals.",
        "Unhealthy for Sensitive Groups": "Sensitive groups may experience effects.",
        "Unhealthy": "Everyone may experience effects; sensitive groups more serious.",
        "Very Unhealthy": "Health alert; everyone may experience more serious effects.",
        "Hazardous": "Emergency conditions; entire population affected.",
    }
    row = 5
    for lo, hi, fg, fc, label in AQI_BANDS:
        ws.cell(row=row, column=1, value=f"{lo}–{hi}").fill = PatternFill("solid", fgColor=fg)
        ws.cell(row=row, column=1).font = Font(color=fc, bold=True)
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2, value=label).font = Font(bold=True)
        ws.cell(row=row, column=3, value=descriptions[label])
        row += 1

    row += 2
    ws.cell(row=row, column=1, value="Reference annual-mean limits (µg/m³)").font = Font(bold=True, size=12)
    row += 1
    for c, h in enumerate(["Pollutant", "WHO (2021 AQG)", "Saudi NCEC"], start=1):
        ws.cell(row=row, column=c, value=h)
    style_header_row(ws, row, 3)
    for pol in ["pm2_5", "pm10", "no2", "so2", "o3", "co"]:
        row += 1
        ws.cell(row=row, column=1, value=pol.upper().replace("_", "."))
        ws.cell(row=row, column=2, value=LIMITS[pol]["WHO"])
        ws.cell(row=row, column=3, value=LIMITS[pol]["NCEC"])

    autosize(ws, {1: 18, 2: 32, 3: 60})


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    stations = fetch_stations()

    per_station: dict[str, tuple[dict, list[dict]]] = {}
    total = 0
    for s in stations:
        name = s.get("stationairq") or f"station_{s.get('index')}"
        geo = s.get("geo_point_2d") or {}
        lat = geo.get("lat")
        lon = geo.get("lon")
        if lat is None or lon is None:
            print(f"  skipping {name}: no coordinates")
            continue
        print(f"  fetching {name} @ {lat:.4f},{lon:.4f}")
        rows = fetch_station_data(lat, lon)
        per_station[name] = (s, rows)
        total += len(rows)

    print(f"Total hourly rows across stations: {total}")
    print("Building workbook...")

    wb = Workbook()
    default = wb.active
    wb.remove(default)

    build_cover(wb, stations, total)
    build_stations(wb, stations)
    for _name, (s, rows) in per_station.items():
        build_station_sheet(wb, s, rows)
    build_summary(wb, per_station)
    build_legend(wb)

    wb.save(OUTPUT_PATH)
    print(f"Wrote {OUTPUT_PATH}  ({OUTPUT_PATH.stat().st_size / 1024 / 1024:.2f} MiB)")


if __name__ == "__main__":
    main()
