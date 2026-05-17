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
    # core pollutants
    "pm10", "pm2_5",
    "carbon_monoxide", "carbon_dioxide",
    "nitrogen_dioxide", "sulphur_dioxide", "ozone",
    # atmospheric composition
    "dust", "aerosol_optical_depth",
    "uv_index", "uv_index_clear_sky",
    # AQI: overall + per-pollutant sub-indices
    "us_aqi",
    "us_aqi_pm2_5", "us_aqi_pm10",
    "us_aqi_nitrogen_dioxide", "us_aqi_ozone",
    "us_aqi_sulphur_dioxide", "us_aqi_carbon_monoxide",
]
WX_HOURLY = [
    # temperature + humidity
    "temperature_2m", "apparent_temperature",
    "relative_humidity_2m", "dew_point_2m",
    # pressure
    "surface_pressure", "pressure_msl",
    # precipitation + cloud
    "precipitation", "rain", "snowfall",
    "cloud_cover", "cloud_cover_low", "cloud_cover_mid", "cloud_cover_high",
    "weather_code",
    # wind
    "wind_speed_10m", "wind_speed_100m",
    "wind_direction_10m", "wind_direction_100m",
    "wind_gusts_10m",
    # radiation + sunshine
    "shortwave_radiation", "direct_radiation",
    "diffuse_radiation", "direct_normal_irradiance",
    "sunshine_duration", "is_day",
    # soil + agro
    "soil_temperature_0_to_7cm", "soil_moisture_0_to_7cm",
    "vapour_pressure_deficit", "et0_fao_evapotranspiration",
]

# Brand palette (Royal Commission green tones + AQI bands)
BRAND_PRIMARY = "0E5C3A"   # deep green header bar
BRAND_ACCENT = "1F9D55"    # accent green
BRAND_LIGHT = "E7F4ED"     # zebra row fill
BRAND_GOLD = "C9A227"      # accent gold
HEADER_FONT_COLOR = "FFFFFF"

# Category-band colors for grouped headers
CAT_COLORS = {
    "pollutants": "1F6FB2",  # blue
    "atmos": "B45F06",       # brown
    "aqi": "7E0023",         # maroon (matches hazardous AQI)
    "temp": "B83232",        # warm red
    "press": "5D3A8E",       # purple
    "precip": "0E7C7B",      # teal
    "cloud": "5B7C99",       # slate
    "wind": "2F8F4F",        # green
    "rad": "C9A227",         # gold
    "soil": "6B4226",        # earth
    "agro": "4C7A2F",        # leaf
}

# Column groups (label, category, list of (key, header, unit, fmt))
COLUMN_GROUPS: list[tuple[str, str, list[tuple[str, str, str, str]]]] = [
    ("Time", "pollutants", [
        ("time", "Timestamp (Riyadh)", "", "yyyy-mm-dd hh:mm"),
    ]),
    ("Core Pollutants", "pollutants", [
        ("pm2_5", "PM2.5", "µg/m³", "0.0"),
        ("pm10", "PM10", "µg/m³", "0.0"),
        ("carbon_monoxide", "CO", "µg/m³", "0.0"),
        ("carbon_dioxide", "CO₂", "ppm", "0.0"),
        ("nitrogen_dioxide", "NO₂", "µg/m³", "0.0"),
        ("sulphur_dioxide", "SO₂", "µg/m³", "0.0"),
        ("ozone", "O₃", "µg/m³", "0.0"),
    ]),
    ("Atmospheric", "atmos", [
        ("dust", "Dust", "µg/m³", "0.0"),
        ("aerosol_optical_depth", "AOD", "—", "0.000"),
        ("uv_index", "UV Index", "—", "0.0"),
        ("uv_index_clear_sky", "UV (clear)", "—", "0.0"),
    ]),
    ("AQI (US)", "aqi", [
        ("us_aqi", "US AQI", "—", "0"),
        ("us_aqi_pm2_5", "AQI PM2.5", "—", "0"),
        ("us_aqi_pm10", "AQI PM10", "—", "0"),
        ("us_aqi_nitrogen_dioxide", "AQI NO₂", "—", "0"),
        ("us_aqi_ozone", "AQI O₃", "—", "0"),
        ("us_aqi_sulphur_dioxide", "AQI SO₂", "—", "0"),
        ("us_aqi_carbon_monoxide", "AQI CO", "—", "0"),
    ]),
    ("Temperature & Humidity", "temp", [
        ("temperature_2m", "Temp", "°C", "0.0"),
        ("apparent_temperature", "Feels Like", "°C", "0.0"),
        ("relative_humidity_2m", "Humidity", "%", "0"),
        ("dew_point_2m", "Dew Point", "°C", "0.0"),
    ]),
    ("Pressure", "press", [
        ("surface_pressure", "Surface P", "hPa", "0.0"),
        ("pressure_msl", "MSL P", "hPa", "0.0"),
    ]),
    ("Precipitation", "precip", [
        ("precipitation", "Precip", "mm", "0.00"),
        ("rain", "Rain", "mm", "0.00"),
        ("snowfall", "Snow", "cm", "0.00"),
    ]),
    ("Cloud", "cloud", [
        ("cloud_cover", "Cloud", "%", "0"),
        ("cloud_cover_low", "Low", "%", "0"),
        ("cloud_cover_mid", "Mid", "%", "0"),
        ("cloud_cover_high", "High", "%", "0"),
        ("weather_code", "WMO Code", "—", "0"),
    ]),
    ("Wind", "wind", [
        ("wind_speed_10m", "Wind 10m", "km/h", "0.0"),
        ("wind_speed_100m", "Wind 100m", "km/h", "0.0"),
        ("wind_direction_10m", "Dir 10m", "°", "0"),
        ("wind_direction_100m", "Dir 100m", "°", "0"),
        ("wind_gusts_10m", "Gusts", "km/h", "0.0"),
    ]),
    ("Radiation & Sunshine", "rad", [
        ("shortwave_radiation", "Shortwave", "W/m²", "0"),
        ("direct_radiation", "Direct", "W/m²", "0"),
        ("diffuse_radiation", "Diffuse", "W/m²", "0"),
        ("direct_normal_irradiance", "DNI", "W/m²", "0"),
        ("sunshine_duration", "Sunshine", "s", "0"),
        ("is_day", "Day?", "0/1", "0"),
    ]),
    ("Soil", "soil", [
        ("soil_temperature_0_to_7cm", "Soil T", "°C", "0.0"),
        ("soil_moisture_0_to_7cm", "Soil M", "m³/m³", "0.000"),
    ]),
    ("Agro", "agro", [
        ("vapour_pressure_deficit", "VPD", "kPa", "0.00"),
        ("et0_fao_evapotranspiration", "ET₀", "mm", "0.00"),
    ]),
]

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


def _chunk_vars(vars_: list[str], size: int = 15) -> list[list[str]]:
    return [vars_[i:i + size] for i in range(0, len(vars_), size)]


def fetch_hourly(url: str, lat: float, lon: float, vars_: list[str]) -> dict[str, list]:
    merged: dict[str, list] = {"time": []}
    for v in vars_:
        merged[v] = []
    for cs, ce in chunk_dates(START_DATE, END_DATE):
        # Split variables into smaller groups so each URL stays under server limits
        per_chunk_time: list[str] | None = None
        for vgroup in _chunk_vars(vars_):
            params = {
                "latitude": lat,
                "longitude": lon,
                "hourly": ",".join(vgroup),
                "start_date": cs.isoformat(),
                "end_date": ce.isoformat(),
                "timezone": "Asia/Riyadh",
            }
            data = http_get(url, params)
            hourly = data.get("hourly") or {}
            if not hourly.get("time"):
                continue
            if per_chunk_time is None:
                per_chunk_time = hourly["time"]
                merged["time"].extend(per_chunk_time)
            for k in vgroup:
                merged[k].extend(hourly.get(k, [None] * len(per_chunk_time)))
    return merged


def fetch_station_data(lat: float, lon: float) -> list[dict]:
    """Return one row per hour with every CAMS + ERA5 variable aligned by timestamp."""
    air = fetch_hourly(OM_AIR_URL, lat, lon, AIR_HOURLY)
    wx = fetch_hourly(OM_WX_URL, lat, lon, WX_HOURLY)
    wx_idx = {t: i for i, t in enumerate(wx.get("time", []))}
    rows: list[dict] = []
    for i, t in enumerate(air.get("time", [])):
        wi = wx_idx.get(t)
        row: dict[str, Any] = {"time": t}
        for v in AIR_HOURLY:
            arr = air.get(v) or []
            row[v] = arr[i] if i < len(arr) else None
        for v in WX_HOURLY:
            arr = wx.get(v) or []
            row[v] = arr[wi] if wi is not None and wi < len(arr) else None
        rows.append(row)
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


def _kpi_tile(ws, top_row: int, left_col: int, label: str, value: str,
              fill: str, accent: str = HEADER_FONT_COLOR) -> None:
    label_cell = ws.cell(row=top_row, column=left_col, value=label)
    value_cell = ws.cell(row=top_row + 1, column=left_col, value=value)
    ws.merge_cells(start_row=top_row, start_column=left_col,
                   end_row=top_row, end_column=left_col + 2)
    ws.merge_cells(start_row=top_row + 1, start_column=left_col,
                   end_row=top_row + 2, end_column=left_col + 2)
    label_cell.fill = PatternFill("solid", fgColor=fill)
    label_cell.font = Font(bold=True, color=accent, size=10)
    label_cell.alignment = Alignment(horizontal="center", vertical="center")
    value_cell.fill = PatternFill("solid", fgColor="FFFFFF")
    value_cell.font = Font(bold=True, size=20, color=fill)
    value_cell.alignment = Alignment(horizontal="center", vertical="center")
    value_cell.border = Border(
        left=Side(style="medium", color=fill),
        right=Side(style="medium", color=fill),
        bottom=Side(style="medium", color=fill),
    )
    ws.row_dimensions[top_row].height = 18
    ws.row_dimensions[top_row + 1].height = 22
    ws.row_dimensions[top_row + 2].height = 22


def build_cover(wb: Workbook, stations: list[dict], total_records: int) -> None:
    ws = wb.create_sheet("Cover", 0)
    ws.sheet_view.showGridLines = False

    # Title bar
    ws.merge_cells("A1:L3")
    title = ws["A1"]
    title.value = "RCRC  ·  Air Quality Monitoring Stations"
    title.font = Font(bold=True, size=26, color="FFFFFF")
    title.alignment = Alignment(horizontal="center", vertical="center")
    title.fill = PatternFill("solid", fgColor=BRAND_PRIMARY)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 30

    ws.merge_cells("A4:L4")
    sub = ws["A4"]
    sub.value = "Royal Commission for Riyadh City · 18-month hourly dataset"
    sub.font = Font(italic=True, size=12, color="FFFFFF")
    sub.alignment = Alignment(horizontal="center", vertical="center")
    sub.fill = PatternFill("solid", fgColor=BRAND_ACCENT)
    ws.row_dimensions[4].height = 22

    # KPI tiles
    days = (END_DATE - START_DATE).days + 1
    total_air_vars = sum(len(g[2]) for g in COLUMN_GROUPS if g[1] in ("pollutants", "atmos", "aqi")) - 1
    total_wx_vars = sum(len(g[2]) for g in COLUMN_GROUPS if g[1] not in ("pollutants", "atmos", "aqi"))
    kpis = [
        ("STATIONS", str(len(stations)), BRAND_PRIMARY),
        ("HOURLY ROWS", f"{total_records:,}", BRAND_ACCENT),
        ("DAYS COVERED", str(days), CAT_COLORS["temp"]),
        ("PARAMETERS / ROW", str(total_air_vars + total_wx_vars), BRAND_GOLD),
    ]
    for i, (lbl, val, color) in enumerate(kpis):
        _kpi_tile(ws, top_row=6, left_col=1 + i * 3, label=lbl, value=val, fill=color)

    # Metadata table
    meta = [
        ("Date range", f"{START_DATE.isoformat()}  →  {END_DATE.isoformat()}"),
        ("Generated", datetime.now().isoformat(timespec="seconds")),
        ("Timezone", "Asia/Riyadh (GMT+3)"),
        ("Air-quality parameters", f"{total_air_vars} (CAMS reanalysis)"),
        ("Weather parameters", f"{total_wx_vars} (ERA5 reanalysis)"),
        ("Station registry", "RCRC Open Data Portal — air-quality-stations-in-riyadh-2025"),
        ("Registry URL", "https://opendata.rcrc.gov.sa/explore/dataset/air-quality-stations-in-riyadh-2025/"),
        ("Air-quality source", "Copernicus Atmosphere Monitoring Service (CAMS) via Open-Meteo"),
        ("Air-quality URL", "https://open-meteo.com/en/docs/air-quality-api"),
        ("Weather source", "ECMWF ERA5 reanalysis via Open-Meteo"),
        ("Weather URL", "https://open-meteo.com/en/docs/historical-weather-api"),
    ]
    start_row = 11
    ws.cell(row=start_row, column=1, value="DATASET METADATA").font = Font(
        bold=True, size=11, color=BRAND_PRIMARY)
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=12)
    for i, (k, v) in enumerate(meta, start=start_row + 1):
        key_cell = ws.cell(row=i, column=1, value=k)
        key_cell.font = Font(bold=True, color="333333")
        key_cell.fill = PatternFill("solid", fgColor=BRAND_LIGHT)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
        val_cell = ws.cell(row=i, column=4, value=v)
        val_cell.font = Font(color="222222")
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=12)
        for c in range(1, 13):
            ws.cell(row=i, column=c).border = BORDER

    note_row = start_row + len(meta) + 2
    note_header = ws.cell(row=note_row, column=1, value="⚠  DATA PROVENANCE")
    note_header.font = Font(bold=True, size=12, color="FFFFFF")
    note_header.fill = PatternFill("solid", fgColor="A04000")
    note_header.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=12)
    ws.row_dimensions[note_row].height = 22

    note_body = (
        "RCRC publishes the station registry only — its open-data portal does NOT expose "
        "historical pollutant readings. The hourly pollutant + weather values in this workbook "
        "are sampled from Copernicus CAMS / ECMWF ERA5 global reanalysis at each RCRC station's "
        "coordinates. CAMS air-quality data is model + satellite derived at ~10 km horizontal "
        "resolution, so stations within a few km of each other in Riyadh will show similar values. "
        "Use these readings as a regional indicator, not as the stations' own instrument observations."
    )
    body = ws.cell(row=note_row + 1, column=1, value=note_body)
    body.alignment = Alignment(wrap_text=True, vertical="top")
    body.font = Font(size=11, color="222222")
    body.fill = PatternFill("solid", fgColor="FFF6E5")
    ws.merge_cells(start_row=note_row + 1, start_column=1, end_row=note_row + 5, end_column=12)
    for r in range(note_row + 1, note_row + 6):
        for c in range(1, 13):
            ws.cell(row=r, column=c).border = BORDER

    autosize(ws, {c: 13 for c in range(1, 13)})


CLASS_COLORS = {
    "Traffic": ("B83232", "FFFFFF"),
    "Suburban": ("2F8F4F", "FFFFFF"),
    "Background stations": ("1F6FB2", "FFFFFF"),
    "Mobile": ("C9A227", "222222"),
}


def build_stations(wb: Workbook, stations: list[dict]) -> None:
    ws = wb.create_sheet("Stations")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:J2")
    title = ws["A1"]
    title.value = "RCRC Air Quality Monitoring Stations"
    title.font = Font(bold=True, size=18, color="FFFFFF")
    title.alignment = Alignment(horizontal="center", vertical="center")
    title.fill = PatternFill("solid", fgColor=BRAND_PRIMARY)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22

    headers = [
        "#", "Station (EN)", "Station (AR)", "Classification",
        "تصنيف", "Altitude (m)", "Latitude", "Longitude",
        "Location (EN)", "Location (AR)",
    ]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=3, column=c, value=h)
    style_header_row(ws, 3, len(headers))

    for i, s in enumerate(stations, start=4):
        geo = s.get("geo_point_2d") or {}
        ws.cell(row=i, column=1, value=s.get("index"))
        ws.cell(row=i, column=2, value=s.get("stationairq"))
        ws.cell(row=i, column=3, value=s.get("stationairqar"))
        cls = s.get("stationclass") or ""
        cls_cell = ws.cell(row=i, column=4, value=cls)
        fg, fc = CLASS_COLORS.get(cls, ("888888", "FFFFFF"))
        cls_cell.fill = PatternFill("solid", fgColor=fg)
        cls_cell.font = Font(bold=True, color=fc)
        cls_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=i, column=5, value=s.get("stationclassar")).alignment = Alignment(
            horizontal="right")
        ws.cell(row=i, column=6, value=s.get("stationaltitude")).alignment = Alignment(
            horizontal="center")
        ws.cell(row=i, column=7, value=geo.get("lat")).number_format = "0.0000"
        ws.cell(row=i, column=8, value=geo.get("lon")).number_format = "0.0000"
        ws.cell(row=i, column=9, value=s.get("stationairqlocation"))
        ws.cell(row=i, column=10, value=s.get("stationairqlocationar")).alignment = Alignment(
            horizontal="right")
        if i % 2 == 1:
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=i, column=c)
                if cell.fill.fgColor.rgb in (None, "00000000", "FFFFFFFF"):
                    cell.fill = PatternFill("solid", fgColor=BRAND_LIGHT)
        for c in range(1, len(headers) + 1):
            ws.cell(row=i, column=c).border = BORDER

    ws.freeze_panes = "A4"
    autosize(ws, {1: 5, 2: 28, 3: 32, 4: 22, 5: 22, 6: 12, 7: 11, 8: 11, 9: 65, 10: 65})


def safe_sheet_name(name: str) -> str:
    bad = '[]:*?/\\'
    out = "".join("_" if ch in bad else ch for ch in name)
    return out[:31]


def _flat_columns() -> list[tuple[str, str, str, str, str]]:
    """Flatten COLUMN_GROUPS into (key, header, unit, fmt, category) tuples."""
    out = []
    for _grp_label, cat, items in COLUMN_GROUPS:
        for key, header, unit, fmt in items:
            out.append((key, header, unit, fmt, cat))
    return out


def build_station_sheet(wb: Workbook, station: dict, rows: list[dict]) -> None:
    name = safe_sheet_name(station["stationairq"])
    ws = wb.create_sheet(name)
    cols = _flat_columns()
    ncols = len(cols)
    geo = station.get("geo_point_2d") or {}

    # Row 1: title block
    title_text = f"  {station['stationairq']}"
    sub_text = f"{station.get('stationclass') or '—'}  ·  {geo.get('lat'):.4f}, {geo.get('lon'):.4f}  ·  alt {station.get('stationaltitude') or '—'} m"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    title = ws.cell(row=1, column=1, value=title_text)
    title.font = Font(bold=True, size=18, color="FFFFFF")
    title.alignment = Alignment(horizontal="left", vertical="center")
    title.fill = PatternFill("solid", fgColor=BRAND_PRIMARY)
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    subc = ws.cell(row=2, column=1, value=sub_text)
    subc.font = Font(italic=True, size=11, color="FFFFFF")
    subc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    subc.fill = PatternFill("solid", fgColor=BRAND_ACCENT)
    ws.row_dimensions[2].height = 18

    # KPI tiles (compact, span cols)
    pm25_vals = [r["pm2_5"] for r in rows if r.get("pm2_5") is not None]
    pm10_vals = [r["pm10"] for r in rows if r.get("pm10") is not None]
    aqi_vals = [r["us_aqi"] for r in rows if r.get("us_aqi") is not None]
    dust_vals = [r["dust"] for r in rows if r.get("dust") is not None]
    avg = lambda xs: sum(xs) / len(xs) if xs else 0
    kpi_specs = [
        ("AVG PM2.5", f"{avg(pm25_vals):.1f} µg/m³", BRAND_PRIMARY),
        ("AVG PM10", f"{avg(pm10_vals):.0f} µg/m³", BRAND_ACCENT),
        ("AVG AQI", f"{avg(aqi_vals):.0f}", CAT_COLORS["aqi"]),
        ("PEAK AQI", f"{max(aqi_vals or [0]):.0f}", "B83232"),
        ("AVG DUST", f"{avg(dust_vals):.0f} µg/m³", CAT_COLORS["atmos"]),
        ("HOURLY ROWS", f"{len(rows):,}", BRAND_GOLD),
    ]
    width_per = max(3, ncols // len(kpi_specs))
    for i, (lbl, val, color) in enumerate(kpi_specs):
        left = 1 + i * width_per
        if left + width_per - 1 > ncols:
            break
        lbl_cell = ws.cell(row=4, column=left, value=lbl)
        val_cell = ws.cell(row=5, column=left, value=val)
        ws.merge_cells(start_row=4, start_column=left,
                       end_row=4, end_column=left + width_per - 1)
        ws.merge_cells(start_row=5, start_column=left,
                       end_row=5, end_column=left + width_per - 1)
        lbl_cell.fill = PatternFill("solid", fgColor=color)
        lbl_cell.font = Font(bold=True, color="FFFFFF", size=9)
        lbl_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.fill = PatternFill("solid", fgColor="FFFFFF")
        val_cell.font = Font(bold=True, size=14, color=color)
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 22

    # Row 6: grouped category headers
    col = 1
    for grp_label, cat, items in COLUMN_GROUPS:
        span = len(items)
        ws.merge_cells(start_row=6, start_column=col,
                       end_row=6, end_column=col + span - 1)
        gc = ws.cell(row=6, column=col, value=grp_label)
        gc.fill = PatternFill("solid", fgColor=CAT_COLORS[cat])
        gc.font = Font(bold=True, color="FFFFFF", size=10)
        gc.alignment = Alignment(horizontal="center", vertical="center")
        gc.border = BORDER
        col += span
    ws.row_dimensions[6].height = 20

    # Row 7: column headers with unit
    for c, (_key, header, unit, _fmt, cat) in enumerate(cols, start=1):
        label = f"{header}" if not unit or unit == "—" else f"{header}\n({unit})"
        cell = ws.cell(row=7, column=c, value=label)
        cell.fill = PatternFill("solid", fgColor=CAT_COLORS[cat])
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[7].height = 32

    # Data rows
    data_start = 8
    for i, r in enumerate(rows):
        rr = data_start + i
        for c, (key, _h, _u, fmt, _cat) in enumerate(cols, start=1):
            val = r.get(key)
            cell = ws.cell(row=rr, column=c, value=val)
            cell.number_format = fmt
            if (i % 2) == 1:
                cell.fill = PatternFill("solid", fgColor=BRAND_LIGHT)

    last = data_start + len(rows) - 1

    # Conditional formatting on every AQI column
    for c, (key, _h, _u, _f, _cat) in enumerate(cols, start=1):
        if key.startswith("us_aqi"):
            add_aqi_conditional(ws, get_column_letter(c), data_start, last)

    ws.freeze_panes = ws.cell(row=data_start, column=2).coordinate
    ws.auto_filter.ref = f"A7:{get_column_letter(ncols)}{last}"

    # Auto-size widths
    widths = {1: 18}
    for c, (_k, header, unit, _f, _cat) in enumerate(cols[1:], start=2):
        widths[c] = max(9, min(14, max(len(header), len(unit) + 2) + 2))
    autosize(ws, widths)


def build_summary(wb: Workbook, per_station: dict[str, tuple[dict, list[dict]]]) -> None:
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:N2")
    title = ws.cell(row=1, column=1, value="  Monthly Averages by Station")
    title.font = Font(bold=True, size=18, color="FFFFFF")
    title.alignment = Alignment(horizontal="left", vertical="center")
    title.fill = PatternFill("solid", fgColor=BRAND_PRIMARY)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22

    headers = [
        ("Station", "pollutants"),
        ("Year-Month", "pollutants"),
        ("Avg PM2.5", "pollutants"),
        ("Avg PM10", "pollutants"),
        ("Avg O₃", "pollutants"),
        ("Avg NO₂", "pollutants"),
        ("Avg SO₂", "pollutants"),
        ("Avg CO", "pollutants"),
        ("Avg Dust", "atmos"),
        ("Avg AOD", "atmos"),
        ("Avg AQI", "aqi"),
        ("Peak AQI", "aqi"),
        ("WHO PM2.5", "aqi"),
        ("NCEC PM2.5", "aqi"),
    ]
    for c, (h, cat) in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.fill = PatternFill("solid", fgColor=CAT_COLORS[cat])
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[3].height = 28

    row_idx = 4
    for stn_name, (_meta, rows) in per_station.items():
        buckets: dict[str, dict[str, list[float]]] = defaultdict(lambda: defaultdict(list))
        for r in rows:
            ts = r["time"]
            if not ts:
                continue
            ym = ts[:7]
            for k in ("pm2_5", "pm10", "ozone", "nitrogen_dioxide", "sulphur_dioxide",
                      "carbon_monoxide", "dust", "aerosol_optical_depth", "us_aqi"):
                v = r.get(k)
                if v is not None:
                    buckets[ym][k].append(v)
        for ym in sorted(buckets.keys()):
            b = buckets[ym]
            def avg(k: str) -> float | None:
                return sum(b[k]) / len(b[k]) if b.get(k) else None
            avg_pm25 = avg("pm2_5")
            peak_aqi = max(b["us_aqi"]) if b.get("us_aqi") else None
            who_ok = "✓" if avg_pm25 is not None and avg_pm25 <= LIMITS["pm2_5"]["WHO"] else "✗"
            ncec_ok = "✓" if avg_pm25 is not None and avg_pm25 <= LIMITS["pm2_5"]["NCEC"] else "✗"
            vals = [
                stn_name, ym, avg_pm25, avg("pm10"), avg("ozone"), avg("nitrogen_dioxide"),
                avg("sulphur_dioxide"), avg("carbon_monoxide"),
                avg("dust"), avg("aerosol_optical_depth"),
                avg("us_aqi"), peak_aqi, who_ok, ncec_ok,
            ]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=row_idx, column=c, value=v)
                cell.border = BORDER
                if row_idx % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=BRAND_LIGHT)
            # PM2.5 compliance coloring
            who_cell = ws.cell(row=row_idx, column=13)
            ncec_cell = ws.cell(row=row_idx, column=14)
            who_cell.fill = PatternFill("solid", fgColor="00E400" if who_ok == "✓" else "FF0000")
            who_cell.font = Font(bold=True, color="000000" if who_ok == "✓" else "FFFFFF")
            who_cell.alignment = Alignment(horizontal="center")
            ncec_cell.fill = PatternFill("solid", fgColor="00E400" if ncec_ok == "✓" else "FF0000")
            ncec_cell.font = Font(bold=True, color="000000" if ncec_ok == "✓" else "FFFFFF")
            ncec_cell.alignment = Alignment(horizontal="center")
            row_idx += 1

    last = row_idx - 1
    for c in range(3, 11):
        col = get_column_letter(c)
        for r in range(4, last + 1):
            fmt = "0.000" if c == 10 else "0.0"
            ws[f"{col}{r}"].number_format = fmt
    for c in (11, 12):
        col = get_column_letter(c)
        for r in range(4, last + 1):
            ws[f"{col}{r}"].number_format = "0"
    add_aqi_conditional(ws, "K", 4, last)
    add_aqi_conditional(ws, "L", 4, last)
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{last}"
    autosize(ws, {
        1: 28, 2: 11, 3: 11, 4: 11, 5: 10, 6: 10,
        7: 10, 8: 10, 9: 11, 10: 10, 11: 10, 12: 10, 13: 12, 14: 12,
    })


def build_legend(wb: Workbook) -> None:
    ws = wb.create_sheet("Legend")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:D2")
    title = ws.cell(row=1, column=1, value="  Legend  ·  Reference Tables")
    title.font = Font(bold=True, size=18, color="FFFFFF")
    title.alignment = Alignment(horizontal="left", vertical="center")
    title.fill = PatternFill("solid", fgColor=BRAND_PRIMARY)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22

    # --- AQI band table
    ws.cell(row=4, column=1, value="US AQI BANDS").font = Font(
        bold=True, size=12, color=BRAND_PRIMARY)
    headers = ["Swatch", "Range", "Category", "Health Implications"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.fill = PatternFill("solid", fgColor=BRAND_PRIMARY)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[5].height = 22

    descriptions = {
        "Good": "Air quality is satisfactory; little to no risk.",
        "Moderate": "Acceptable; minor risk for unusually sensitive individuals.",
        "Unhealthy for Sensitive Groups": "Sensitive groups may experience effects.",
        "Unhealthy": "Everyone may experience effects; sensitive groups more serious.",
        "Very Unhealthy": "Health alert; everyone may experience more serious effects.",
        "Hazardous": "Emergency conditions; entire population affected.",
    }
    row = 6
    for lo, hi, fg, fc, label in AQI_BANDS:
        sw = ws.cell(row=row, column=1, value="")
        sw.fill = PatternFill("solid", fgColor=fg)
        sw.border = BORDER
        rng = ws.cell(row=row, column=2, value=f"{lo}–{hi}")
        rng.font = Font(bold=True)
        rng.alignment = Alignment(horizontal="center")
        rng.border = BORDER
        cat = ws.cell(row=row, column=3, value=label)
        cat.font = Font(bold=True, color=fc)
        cat.fill = PatternFill("solid", fgColor=fg)
        cat.alignment = Alignment(horizontal="center")
        cat.border = BORDER
        desc = ws.cell(row=row, column=4, value=descriptions[label])
        desc.alignment = Alignment(wrap_text=True, vertical="center")
        desc.border = BORDER
        ws.row_dimensions[row].height = 24
        row += 1

    # --- Reference limits table
    row += 2
    ws.cell(row=row, column=1, value="REFERENCE ANNUAL-MEAN LIMITS").font = Font(
        bold=True, size=12, color=BRAND_PRIMARY)
    row += 1
    for c, h in enumerate(["Pollutant", "WHO (2021 AQG)", "Saudi NCEC", "Unit"], start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = PatternFill("solid", fgColor=BRAND_PRIMARY)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 20
    for i, pol in enumerate(["pm2_5", "pm10", "no2", "so2", "o3", "co"]):
        row += 1
        unit = "µg/m³"
        ws.cell(row=row, column=1, value=pol.upper().replace("_", "."))
        ws.cell(row=row, column=2, value=LIMITS[pol]["WHO"])
        ws.cell(row=row, column=3, value=LIMITS[pol]["NCEC"])
        ws.cell(row=row, column=4, value=unit)
        for c in range(1, 5):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).alignment = Alignment(horizontal="center")
            if i % 2 == 1:
                ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=BRAND_LIGHT)

    # --- Category swatches
    row += 2
    ws.cell(row=row, column=1, value="COLUMN CATEGORY COLOR KEY").font = Font(
        bold=True, size=12, color=BRAND_PRIMARY)
    row += 1
    for c, h in enumerate(["Swatch", "Category", "Description", ""], start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = PatternFill("solid", fgColor=BRAND_PRIMARY)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    cat_descs = {
        "pollutants": "Core gas + particulate pollutants",
        "atmos": "Atmospheric composition (dust, AOD, UV)",
        "aqi": "US Air Quality Index (overall + per-pollutant)",
        "temp": "Temperature and humidity",
        "press": "Surface and mean sea-level pressure",
        "precip": "Precipitation (rain, snow)",
        "cloud": "Cloud cover at altitude bands",
        "wind": "Wind speed, gusts and direction",
        "rad": "Solar radiation and sunshine",
        "soil": "Soil temperature and moisture",
        "agro": "Agro indices (VPD, ET₀)",
    }
    for cat, descr in cat_descs.items():
        row += 1
        sw = ws.cell(row=row, column=1, value="")
        sw.fill = PatternFill("solid", fgColor=CAT_COLORS[cat])
        sw.border = BORDER
        ws.cell(row=row, column=2, value=cat.upper()).font = Font(bold=True)
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2).border = BORDER
        ws.cell(row=row, column=3, value=descr).border = BORDER
        ws.cell(row=row, column=4, value="").border = BORDER

    autosize(ws, {1: 12, 2: 18, 3: 60, 4: 12})


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
