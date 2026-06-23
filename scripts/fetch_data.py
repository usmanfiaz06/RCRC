"""
Fetch RCRC station list + Open-Meteo historical air quality & weather data
for all RCRC monitoring stations in Riyadh, covering 2026-01-01 to 2026-05-31.
"""
from __future__ import annotations
import json
import os
import sys
import time
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

SCRATCH = Path(os.environ.get(
    "RCRC_SCRATCH",
    "/tmp/claude-0/-home-user-RCRC/3e7d6508-8244-5a9f-b105-f6f5e0138312/scratchpad/raw",
))
SCRATCH.mkdir(parents=True, exist_ok=True)

START_DATE = "2026-01-01"
END_DATE = "2026-05-31"
TIMEZONE = "Asia/Riyadh"

AQ_VARS = [
    "pm10", "pm2_5", "carbon_monoxide", "nitrogen_dioxide",
    "sulphur_dioxide", "ozone",
    "us_aqi", "us_aqi_pm2_5", "us_aqi_pm10",
    "us_aqi_no2", "us_aqi_o3", "us_aqi_so2", "us_aqi_co",
    "dust", "uv_index",
]
WX_VARS = [
    "temperature_2m", "relative_humidity_2m", "apparent_temperature",
    "wind_speed_10m", "wind_direction_10m", "wind_gusts_10m",
    "surface_pressure", "precipitation", "cloud_cover",
]


def http_get_json(url: str, retries: int = 6) -> dict:
    last_err: Exception | None = None
    for i in range(retries):
        try:
            req = Request(url, headers={"User-Agent": "RCRC-data-export/1.0"})
            with urlopen(req, timeout=60) as r:
                return json.load(r)
        except Exception as e:  # noqa: BLE001
            last_err = e
            wait = min(60, 3 * (2 ** i))
            print(f"  ! retry {i+1}/{retries} after {wait}s ({e})", file=sys.stderr)
            time.sleep(wait)
    raise RuntimeError(f"GET failed: {url}: {last_err}")


THROTTLE_S = 8.0


def fetch_rcrc_stations() -> list[dict]:
    cache = SCRATCH / "rcrc_stations.json"
    if cache.exists():
        return json.loads(cache.read_text())["results"]
    url = (
        "https://opendata.rcrc.gov.sa/api/explore/v2.1/catalog/datasets/"
        "air-quality-stations-in-riyadh-2025/records?limit=100"
    )
    data = http_get_json(url)
    cache.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    return data["results"]


def fetch_waqi_riyadh_snapshot() -> dict:
    cache = SCRATCH / "waqi_snapshot.json"
    if cache.exists():
        return json.loads(cache.read_text())
    out: dict[str, object] = {}
    token = "40ab14872a11434ab8fa725d0ce972b474f56ced"
    keywords = ["riyadh", "saudi"]
    for kw in keywords:
        try:
            s = http_get_json(
                f"https://api.waqi.info/search/?token={token}&keyword={kw}"
            )
            out[f"search_{kw}"] = s
        except Exception as e:
            out[f"search_{kw}_error"] = str(e)
    # Pull a few active station feeds
    feeds = {}
    for uid in [12928, 11464, 11650, 11617, 11620, 11619]:
        try:
            feeds[uid] = http_get_json(
                f"https://api.waqi.info/feed/@{uid}/?token={token}"
            )
        except Exception as e:
            feeds[uid] = {"error": str(e)}
    out["feeds"] = feeds
    cache.write_text(json.dumps(out, ensure_ascii=False, indent=2))
    return out


def fetch_open_meteo_aq(lat: float, lon: float, slug: str) -> dict:
    cache = SCRATCH / f"aq_{slug}.json"
    if cache.exists():
        return json.loads(cache.read_text())
    qs = urlencode({
        "latitude": lat,
        "longitude": lon,
        "hourly": ",".join(AQ_VARS),
        "start_date": START_DATE,
        "end_date": END_DATE,
        "timezone": TIMEZONE,
    })
    data = http_get_json(f"https://air-quality-api.open-meteo.com/v1/air-quality?{qs}")
    cache.write_text(json.dumps(data))
    return data


def fetch_open_meteo_wx(lat: float, lon: float, slug: str) -> dict:
    cache = SCRATCH / f"wx_{slug}.json"
    if cache.exists():
        return json.loads(cache.read_text())
    qs = urlencode({
        "latitude": lat,
        "longitude": lon,
        "hourly": ",".join(WX_VARS),
        "start_date": START_DATE,
        "end_date": END_DATE,
        "timezone": TIMEZONE,
    })
    data = http_get_json(f"https://archive-api.open-meteo.com/v1/archive?{qs}")
    cache.write_text(json.dumps(data))
    return data


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


def main() -> None:
    print("Fetching RCRC station list...")
    stations = fetch_rcrc_stations()
    print(f"  -> {len(stations)} stations")

    print("Fetching WAQI snapshot (context only)...")
    fetch_waqi_riyadh_snapshot()

    for s in stations:
        name = s["stationairq"]
        lat = s["geo_point_2d"]["lat"]
        lon = s["geo_point_2d"]["lon"]
        slug = slugify(name)
        print(f"Fetching {name} ({lat}, {lon})")
        aq_cache = SCRATCH / f"aq_{slug}.json"
        wx_cache = SCRATCH / f"wx_{slug}.json"
        if not aq_cache.exists():
            fetch_open_meteo_aq(lat, lon, slug)
            time.sleep(THROTTLE_S)
        if not wx_cache.exists():
            fetch_open_meteo_wx(lat, lon, slug)
            time.sleep(THROTTLE_S)

    print("Done.")


if __name__ == "__main__":
    main()
