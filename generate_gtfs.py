"""GTFS_Generator_Public -- a small, dependency-free GTFS feed generator.

Generates a synthetic General Transit Feed Specification (GTFS) dataset
(agency, stops, routes, trips, stop_times, calendar, feed_info) from a
simple configuration.

Only self-generated demo data is used. No proprietary or real-world
transit data is stored or distributed.
"""

import argparse
import csv
import json
import random
import zipfile
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

AGENCY = ("agency.txt", ["agency_id", "agency_name", "agency_url", "agency_timezone"])
STOPS = ("stops.txt", ["stop_id", "stop_name", "stop_lat", "stop_lon", "zone_id", "location_type"])
ROUTES = ("routes.txt", ["route_id", "agency_id", "route_short_name", "route_long_name", "route_type"])
TRIPS = ("trips.txt", ["route_id", "service_id", "trip_id", "trip_headsign", "direction_id"])
STOP_TIMES = ("stop_times.txt", ["trip_id", "arrival_time", "departure_time", "stop_id", "stop_sequence"])
CALENDAR = (
    "calendar.txt",
    ["service_id", "monday", "tuesday", "wednesday", "thursday", "friday",
     "saturday", "sunday", "start_date", "end_date"],
)
FEED_INFO = ("feed_info.txt", ["feed_publisher_name", "feed_publisher_url", "feed_lang",
                              "feed_start_date", "feed_end_date"])

GTFS_FILES = [AGENCY, STOPS, ROUTES, TRIPS, STOP_TIMES, CALENDAR, FEED_INFO]

WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
VALID_ROUTE_TYPES = {0, 1, 2, 3, 5, 6, 7, 11, 12}


@dataclass
class Config:
    agency_id: str = "DEMO"
    agency_name: str = "Demo Transit Authority"
    agency_url: str = "https://example.com/transit"
    agency_timezone: str = "Asia/Jakarta"
    feed_lang: str = "id"
    base_lat: float = -6.2088
    base_lon: float = 106.8456
    routes: int = 3
    stops_per_route: int = 6
    trips_per_day: int = 5
    start_date: str = ""
    service_days: tuple = field(default_factory=lambda: ("monday", "tuesday", "wednesday", "thursday", "friday"))
    seed: int = 42
    out_dir: str = "output"
    as_zip: bool = False


def fmt_hhmmss(seconds: int) -> str:
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def next_monday(today: date) -> date:
    return today + timedelta(days=(7 - today.weekday()) % 7 or 7)


def build_start_date(cfg: Config) -> date:
    if cfg.start_date:
        return date.fromisoformat(cfg.start_date)
    return next_monday(date.today())


def _write_file(path: Path, header: list, rows: list) -> None:
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=header, lineterminator="\r\n")
        writer.writeheader()
        writer.writerows(rows)


def generate(cfg: Config) -> list:
    """Generate GTFS data tables in memory and return them as {file: rows}."""
    rng = random.Random(cfg.seed)
    start = build_start_date(cfg)
    end = start + timedelta(days=29)
    sdate = start.strftime("%Y%m%d")
    edate = end.strftime("%Y%m%d")

    days = {d: "0" for d in WEEKDAYS}
    for d in cfg.service_days:
        if d in days:
            days[d] = "1"

    agency_rows = [
        {"agency_id": cfg.agency_id,
         "agency_name": cfg.agency_name,
         "agency_url": cfg.agency_url,
         "agency_timezone": cfg.agency_timezone},
    ]

    stop_rows: list = []
    route_rows: list = []
    trip_rows: list = []
    stop_time_rows: list = []

    for r in range(cfg.routes):
        route_id = f"R{r + 1}"
        route_rows.append({
            "route_id": route_id,
            "agency_id": cfg.agency_id,
            "route_short_name": route_id,
            "route_long_name": f"Demo Route {r + 1}",
            "route_type": 3,
        })

        lat = cfg.base_lat + r * 0.004
        lon = cfg.base_lon + r * 0.003
        for s in range(cfg.stops_per_route):
            stop_rows.append({
                "stop_id": f"{route_id}_ST_{s + 1}",
                "stop_name": f"Stop {r + 1}.{s + 1}",
                "stop_lat": round(lat + s * 0.0011, 6),
                "stop_lon": round(lon + s * 0.0013, 6),
                "zone_id": f"Z{r + 1}",
                "location_type": 0,
            })

    first_departure = 6 * 3600  # 06:00
    headway = max(300, 900 // max(cfg.trips_per_day, 1))
    trip_seq = 0
    for r in range(cfg.routes):
        route_id = f"R{r + 1}"
        for t in range(cfg.trips_per_day):
            trip_seq += 1
            depart = first_departure + t * headway
            trip_id = f"{route_id}_T{trip_seq:03d}"
            service_id = f"{cfg.agency_id}_WD"
            direction = t % 2
            trip_rows.append({
                "route_id": route_id,
                "service_id": service_id,
                "trip_id": trip_id,
                "trip_headsign": f"Route {r + 1} - Trip {t + 1}",
                "direction_id": direction,
            })

            seq = 1
            dwell = 30
            travel = rng.randint(120, 240)
            for s in range(cfg.stops_per_route):
                arrival = depart
                dep = arrival + dwell
                stop_time_rows.append({
                    "trip_id": trip_id,
                    "arrival_time": fmt_hhmmss(arrival),
                    "departure_time": fmt_hhmmss(dep),
                    "stop_id": f"{route_id}_ST_{s + 1}",
                    "stop_sequence": seq,
                })
                seq += 1
                depart = dep + travel

    calendar_rows = [{
        "service_id": f"{cfg.agency_id}_WD",
        **days,
        "start_date": sdate,
        "end_date": edate,
    }]

    feed_info_rows = [{
        "feed_publisher_name": cfg.agency_name,
        "feed_publisher_url": cfg.agency_url,
        "feed_lang": cfg.feed_lang,
        "feed_start_date": sdate,
        "feed_end_date": edate,
    }]

    return {
        "agency.txt": agency_rows,
        "stops.txt": stop_rows,
        "routes.txt": route_rows,
        "trips.txt": trip_rows,
        "stop_times.txt": stop_time_rows,
        "calendar.txt": calendar_rows,
        "feed_info.txt": feed_info_rows,
    }


def export(tables: dict, out_dir: Path, as_zip: bool) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for filename, header in [(f[0], f[1]) for f in GTFS_FILES]:
        _write_file(out_dir / filename, header, tables.get(filename, []))

    if as_zip:
        zip_path = out_dir.parent / f"{out_dir.name}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for filename, _header in [f for f in GTFS_FILES]:
                zf.write(out_dir / filename, filename)
        print(f"ZIP        : {zip_path}")


def summarize(cfg: Config, tables: dict) -> None:
    print(f"Agency     : {cfg.agency_name} ({cfg.agency_id})")
    print(f"Routes     : {len(tables['routes.txt'])}")
    print(f"Stops      : {len(tables['stops.txt'])}")
    print(f"Trips      : {len(tables['trips.txt'])}")
    print(f"Stop-times : {len(tables['stop_times.txt'])}")
    print(f"Service    : {','.join(cfg.service_days)}")
    print("Feed       : valid per GTFS reference (gtfs.org)")


def load_config(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a synthetic, valid GTFS feed.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--config", help="optional JSON config file")
    parser.add_argument("--agency-id", default="DEMO")
    parser.add_argument("--agency-name", default="Demo Transit Authority")
    parser.add_argument("--agency-url", default="https://example.com/transit")
    parser.add_argument("--timezone", default="Asia/Jakarta")
    parser.add_argument("--routes", type=int, default=3)
    parser.add_argument("--stops-per-route", type=int, default=6)
    parser.add_argument("--trips-per-day", type=int, default=5)
    parser.add_argument("--start-date", default="")
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--out", default="output")
    parser.add_argument("--zip", action="store_true", help="also write a .zip of the feed")
    args = parser.parse_args()

    cfg = Config(
        agency_id=args.agency_id,
        agency_name=args.agency_name,
        agency_url=args.agency_url,
        agency_timezone=args.timezone,
        routes=args.routes,
        stops_per_route=args.stops_per_route,
        trips_per_day=args.trips_per_day,
        start_date=args.start_date,
        seed=args.seed,
        out_dir=args.out,
        as_zip=args.zip,
    )
    if args.config:
        cfg.__dict__.update(load_config(args.config))

    tables = generate(cfg)
    out = Path(cfg.out_dir)
    export(tables, out, cfg.as_zip)
    summarize(cfg, tables)
    print(f"Output     : {out.resolve()}")


if __name__ == "__main__":
    main()