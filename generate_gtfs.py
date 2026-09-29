"""GTFS Generator KCI -- build a valid GTFS feed for KAI Commuter (Jabodetabek).

Two modes:

1. Demo mode (no input file) -- emits a small synthetic feed. Dependency-free,
   so the tool stays runnable on a bare Python install.
2. KCI mode -- parses a real KCI timetable, either the GAPEKA Excel workbook
   (``--excel``) or the public Commuter Line schedule PDF (``--pdf`` /
   ``--pdf-url``), and writes a full feed.

The KCI feed contains 7 files, mirroring KCI's reference feed exactly:
agency.txt, calendar.txt, levels.txt, routes.txt, stops.txt, trips.txt,
stop_times.txt.

Optional: reading a timetable needs one extra package, imported lazily and
only when that input mode is used.

    pip install openpyxl   # for --excel
    pip install pymupdf    # for --pdf / --pdf-url

Station and route data describe the public KCI network. It is provided so the
generated feeds are usable as-is; feed validation should follow the operator's
own source timetable.
"""

import argparse
import csv
import io
import json
import re
import sys
import urllib.request
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path

import gtfs_edit

try:  # optional: only needed for --excel
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None

try:  # optional: only needed for --pdf / --pdf-url
    import pymupdf
except ImportError:  # pragma: no cover
    pymupdf = None


# ===========================================================================
# GTFS file layout
# ===========================================================================
AGENCY = ("agency.txt", "agency_id,agency_name,agency_url,agency_timezone,agency_phone,agency_lang")
CALENDAR = ("calendar.txt", "end_date,friday,monday,saturday,service_id,start_date,sunday,thursday,tuesday,wednesday")
LEVELS = ("levels.txt", "level_id,level_index,level_name")
ROUTES = ("routes.txt", "agency_id,route_desc,route_id,route_long_name,route_short_name,route_type")
STOPS = ("stops.txt", "level_id,location_type,parent_station,stop_code,stop_desc,stop_id,stop_lat,stop_lon,stop_name,wheelchair_boarding,zone_id")
TRIPS = ("trips.txt", "route_id,service_id,trip_id,trip_short_name,trip_headsign,direction_id,shape_id")
STOP_TIMES = ("stop_times.txt", "trip_id,pickup_type,departure_time,stop_id,arrival_time,stop_sequence,drop_off_type,departure_time_fixed")

GTFS_FILES = [AGENCY, CALENDAR, LEVELS, ROUTES, STOPS, TRIPS, STOP_TIMES]

LEVELS_BODY = "level_id,level_index,level_name\nL0,0,Street\nL1,-1,Lower Ground\nL2,1,Upper Ground"

FEED_START_DATE = "20210816"
FEED_END_DATE = "20261231"

WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

PDF_DEFAULT_URL = (
    "https://kci.id/files/download/documents/"
    "Jadwal%20Commuter%20Line%20Jabodetabek%20-%20Update%2010%20September%202026.pdf"
)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0 Safari/537.36"
)


@dataclass
class Config:
    excel: str = ""
    pdf: str = ""
    pdf_url: str = ""
    out: str = "output"
    as_zip: bool = False
    filter_services: tuple = field(default_factory=tuple)
    demo_routes: int = 3
    demo_stops_per_route: int = 6
    demo_trips_per_day: int = 5
    seed: int = 42
    start_date: str = ""


# ===========================================================================
# KCI network data
# ===========================================================================
AGENCY_INFO = {
    "agency_id": "1",
    "agency_name": "PT. Kereta Commuter Indonesia",
    "agency_url": "https://jakartakci.co.id/",
    "agency_timezone": "Asia/Jakarta",
    "agency_phone": "",
    "agency_lang": "id",
}

ROUTE_LIST = [
    {"agency_id": "1", "route_id": "1001", "route_short_name": "BO", "route_long_name": "Bogor tujuan Jakarta Kota", "route_desc": "Bogor tujuan Jakarta Kota", "route_type": "0"},
    {"agency_id": "1", "route_id": "1002", "route_short_name": "BO", "route_long_name": "Jakarta Kota tujuan Nambo/Bogor", "route_desc": "Jakarta Kota tujuan Nambo/Bogor", "route_type": "0"},
    {"agency_id": "1", "route_id": "2001", "route_short_name": "BE", "route_long_name": "Bekasi tujuan Kampung Bandan-Bekasi/Cikarang", "route_desc": "Bekasi tujuan Kampung Bandan-Bekasi/Cikarang", "route_type": "0"},
    {"agency_id": "1", "route_id": "2002", "route_short_name": "BE", "route_long_name": "Kampung Bandan tujuan Bekasi/Cikarang", "route_desc": "Kampung Bandan tujuan Bekasi/Cikarang", "route_type": "0"},
    {"agency_id": "1", "route_id": "3001", "route_short_name": "RA", "route_long_name": "Rangkasbitung tujuan Tanah Abang", "route_desc": "Rangkasbitung tujuan Tanah Abang", "route_type": "0"},
    {"agency_id": "1", "route_id": "3002", "route_short_name": "RA", "route_long_name": "Tanah Abang tujuan Rangkasbitung", "route_desc": "Tanah Abang tujuan Rangkasbitung", "route_type": "0"},
    {"agency_id": "1", "route_id": "4001", "route_short_name": "TA", "route_long_name": "Tangerang tujuan Duri", "route_desc": "Tangerang tujuan Duri", "route_type": "0"},
    {"agency_id": "1", "route_id": "4002", "route_short_name": "TA", "route_long_name": "Duri tujuan Tangerang", "route_desc": "Duri tujuan Tangerang", "route_type": "0"},
    {"agency_id": "1", "route_id": "5001", "route_short_name": "TP", "route_long_name": "Tanjung Priok tujuan Jakarta Kota", "route_desc": "Tanjung Priok tujuan Jakarta Kota", "route_type": "0"},
    {"agency_id": "1", "route_id": "5002", "route_short_name": "TP", "route_long_name": "Jakarta Kota tujuan Tanjung Priok", "route_desc": "Jakarta Kota tujuan Tanjung Priok", "route_type": "0"},
]

SHEET_ROUTE_MAP = {
    "JAKK-BOO": {
        "route_id": "1002",
        "route_short_name": "BO",
        "route_long_name": "Jakarta Kota tujuan Nambo/Bogor",
        "route_desc": "Jakarta Kota tujuan Nambo/Bogor",
        "direction": "downline",
    },
    "BOO-JAKK": {
        "route_id": "1001",
        "route_short_name": "BO",
        "route_long_name": "Bogor tujuan Jakarta Kota",
        "route_desc": "Bogor tujuan Jakarta Kota",
        "direction": "upline",
    },
    "CKR-MRI-KPB": {
        "route_id": "2001",
        "route_short_name": "BE",
        "route_long_name": "Bekasi tujuan Kampung Bandan-Bekasi/Cikarang",
        "route_desc": "Bekasi tujuan Kampung Bandan-Bekasi/Cikarang",
        "direction": "upline",
    },
    "KPB-PSE-CKR": {
        "route_id": "2002",
        "route_short_name": "BE",
        "route_long_name": "Kampung Bandan tujuan Bekasi/Cikarang",
        "route_desc": "Kampung Bandan tujuan Bekasi/Cikarang",
        "direction": "downline",
    },
    "SRP-THB": {
        "route_id": "3001",
        "route_short_name": "RA",
        "route_long_name": "Rangkasbitung tujuan Tanah Abang",
        "route_desc": "Rangkasbitung tujuan Tanah Abang",
        "direction": "upline",
    },
    "THB-SRP": {
        "route_id": "3002",
        "route_short_name": "RA",
        "route_long_name": "Tanah Abang tujuan Rangkasbitung",
        "route_desc": "Tanah Abang tujuan Rangkasbitung",
        "direction": "downline",
    },
    "TNG-DU": {
        "route_id": "4001",
        "route_short_name": "TA",
        "route_long_name": "Tangerang tujuan Duri",
        "route_desc": "Tangerang tujuan Duri",
        "direction": "upline",
    },
    "DU-TNG": {
        "route_id": "4002",
        "route_short_name": "TA",
        "route_long_name": "Duri tujuan Tangerang",
        "route_desc": "Duri tujuan Tangerang",
        "direction": "downline",
    },
}

# The JAKK-TPK sheet carries both directions in a single sheet.
JAKK_TPK_MAP = {
    "left": {
        "route_id": "5002",
        "route_short_name": "TP",
        "route_long_name": "Jakarta Kota tujuan Tanjung Priok",
        "route_desc": "Jakarta Kota tujuan Tanjung Priok",
        "direction": "downline",
    },
    "right": {
        "route_id": "5001",
        "route_short_name": "TP",
        "route_long_name": "Tanjung Priok tujuan Jakarta Kota",
        "route_desc": "Tanjung Priok tujuan Jakarta Kota",
        "direction": "upline",
    },
}

# Stations marked as skipped by GAPEKA sheets (no scheduled stop of their own).
SKIP_STATIONS = {"GMR", "GPI"}

LS_VALUES = ("Ls", "Ls.", "LS", "ls", "lS", "LS.", "Ls..")

STATIONS = {
    'AC': {"stop_id": 'AC', "stop_code": '10001', "stop_name": 'STASIUN KCI ANCOL', "stop_lat": -6.127971259, "stop_lon": 106.8451331, "zone_id": 'AC', "level_id": 'L0', "is_active": 1},
    'AK': {"stop_id": 'AK', "stop_code": '10002', "stop_name": 'STASIUN KCI ANGKE', "stop_lat": -6.144529555, "stop_lon": 106.8008054, "zone_id": 'AK', "level_id": 'L0', "is_active": 1},
    'BJD': {"stop_id": 'BJD', "stop_code": '10003', "stop_name": 'STASIUN KCI BOJONGGEDE', "stop_lat": -6.492615357, "stop_lon": 106.7951425, "zone_id": 'BJD', "level_id": 'L0', "is_active": 1},
    'BKS': {"stop_id": 'BKS', "stop_code": '10004', "stop_name": 'STASIUN KCI BEKASI', "stop_lat": -6.236332831, "stop_lon": 106.9991907, "zone_id": 'BKS', "level_id": 'L0', "is_active": 1},
    'BKST': {"stop_id": 'BKST', "stop_code": '10005', "stop_name": 'STASIUN KCI BEKASI TIMUR', "stop_lat": -6.247755104, "stop_lon": 107.0168618, "zone_id": 'BKST', "level_id": 'L0', "is_active": 1},
    'BOI': {"stop_id": 'BOI', "stop_code": '10006', "stop_name": 'STASIUN KCI BOJONGINDAH', "stop_lat": -6.160257526, "stop_lon": 106.735313, "zone_id": 'BOI', "level_id": 'L0', "is_active": 1},
    'BOO': {"stop_id": 'BOO', "stop_code": '10007', "stop_name": 'STASIUN KCI BOGOR', "stop_lat": -6.59553519, "stop_lon": 106.7903721, "zone_id": 'BOO', "level_id": 'L0', "is_active": 1},
    'BPR': {"stop_id": 'BPR', "stop_code": '10008', "stop_name": 'STASIUN KCI BATUCEPER', "stop_lat": -6.172160295, "stop_lon": 106.6644325, "zone_id": 'BPR', "level_id": 'L0', "is_active": 1},
    'BUA': {"stop_id": 'BUA', "stop_code": '10009', "stop_name": 'STASIUN KCI BUARAN', "stop_lat": -6.215541198, "stop_lon": 106.9232196, "zone_id": 'BUA', "level_id": 'L0', "is_active": 1},
    'CBN': {"stop_id": 'CBN', "stop_code": '10010', "stop_name": 'STASIUN KCI CIBINONG', "stop_lat": -6.464093825, "stop_lon": 106.8524803, "zone_id": 'CBN', "level_id": 'L0', "is_active": 1},
    'CC': {"stop_id": 'CC', "stop_code": '10012', "stop_name": 'STASIUN KCI CICAYUR', "stop_lat": -6.329312109, "stop_lon": 106.61885, "zone_id": 'CC', "level_id": 'L0', "is_active": 1},
    'CIT': {"stop_id": 'CIT', "stop_code": '10011', "stop_name": 'STASIUN KCI CIBITUNG', "stop_lat": -6.261175162, "stop_lon": 107.083598, "zone_id": 'CIT', "level_id": 'L0', "is_active": 1},
    'CJT': {"stop_id": 'CJT', "stop_code": '10013', "stop_name": 'STASIUN KCI CILEJIT', "stop_lat": -6.354161863, "stop_lon": 106.5096197, "zone_id": 'CJT', "level_id": 'L0', "is_active": 1},
    'CKI': {"stop_id": 'CKI', "stop_code": '10014', "stop_name": 'STASIUN KCI CIKINI', "stop_lat": -6.198088457, "stop_lon": 106.8408286, "zone_id": 'CKI', "level_id": 'L0', "is_active": 1},
    'CKR': {"stop_id": 'CKR', "stop_code": '10015', "stop_name": 'STASIUN KCI CIKARANG', "stop_lat": -6.255277811, "stop_lon": 107.1449188, "zone_id": 'CKR', "level_id": 'L0', "is_active": 1},
    'CKY': {"stop_id": 'CKY', "stop_code": '10016', "stop_name": 'STASIUN KCI CIKOYA', "stop_lat": -6.335516842, "stop_lon": 106.4117844, "zone_id": 'CKY', "level_id": 'L0', "is_active": 1},
    'CLT': {"stop_id": 'CLT', "stop_code": '10017', "stop_name": 'STASIUN KCI CILEBUT', "stop_lat": -6.53033637, "stop_lon": 106.8005607, "zone_id": 'CLT', "level_id": 'L0', "is_active": 1},
    'CSK': {"stop_id": 'CSK', "stop_code": '10018', "stop_name": 'STASIUN KCI CISAUK', "stop_lat": -6.324631625, "stop_lon": 106.6409369, "zone_id": 'CSK', "level_id": 'L0', "is_active": 1},
    'CTA': {"stop_id": 'CTA', "stop_code": '10019', "stop_name": 'STASIUN KCI CITAYAM', "stop_lat": -6.448517911, "stop_lon": 106.8025176, "zone_id": 'CTA', "level_id": 'L0', "is_active": 1},
    'CTR': {"stop_id": 'CTR', "stop_code": '10020', "stop_name": 'STASIUN KCI CITERAS', "stop_lat": -6.334649739, "stop_lon": 106.3326455, "zone_id": 'CTR', "level_id": 'L0', "is_active": 1},
    'CUK': {"stop_id": 'CUK', "stop_code": '10021', "stop_name": 'STASIUN KCI CAKUNG', "stop_lat": -6.21896241, "stop_lon": 106.9523934, "zone_id": 'CUK', "level_id": 'L0', "is_active": 1},
    'CW': {"stop_id": 'CW', "stop_code": '10022', "stop_name": 'STASIUN KCI CAWANG', "stop_lat": -6.242415643, "stop_lon": 106.8587954, "zone_id": 'CW', "level_id": 'L0', "is_active": 1},
    'DAR': {"stop_id": 'DAR', "stop_code": '10023', "stop_name": 'STASIUN KCI DARU', "stop_lat": -6.337834057, "stop_lon": 106.4922818, "zone_id": 'DAR', "level_id": 'L0', "is_active": 1},
    'DP': {"stop_id": 'DP', "stop_code": '10024', "stop_name": 'STASIUN KCI DEPOK', "stop_lat": -6.404387311, "stop_lon": 106.8172077, "zone_id": 'DP', "level_id": 'L0', "is_active": 1},
    'DPB': {"stop_id": 'DPB', "stop_code": '10025', "stop_name": 'STASIUN KCI DEPOKBARU', "stop_lat": -6.390185488, "stop_lon": 106.8217138, "zone_id": 'DPB', "level_id": 'L0', "is_active": 1},
    'DRN': {"stop_id": 'DRN', "stop_code": '10026', "stop_name": 'STASIUN KCI DURENKALIBATA', "stop_lat": -6.255470408, "stop_lon": 106.8551521, "zone_id": 'DRN', "level_id": 'L0', "is_active": 1},
    'DU': {"stop_id": 'DU', "stop_code": '10027', "stop_name": 'STASIUN KCI DURI', "stop_lat": -6.155049989, "stop_lon": 106.8015022, "zone_id": 'DU', "level_id": 'L0', "is_active": 1},
    'GDD': {"stop_id": 'GDD', "stop_code": '10028', "stop_name": 'STASIUN KCI GONDANGDIA', "stop_lat": -6.185866268, "stop_lon": 106.832647, "zone_id": 'GDD', "level_id": 'L0', "is_active": 1},
    'GRG': {"stop_id": 'GRG', "stop_code": '10029', "stop_name": 'STASIUN KCI GROGOL', "stop_lat": -6.161717357, "stop_lon": 106.7894467, "zone_id": 'GRG', "level_id": 'L0', "is_active": 1},
    'GST': {"stop_id": 'GST', "stop_code": '10030', "stop_name": 'STASIUN KCI GANGSENTIONG', "stop_lat": -6.185913632, "stop_lon": 106.8506798, "zone_id": 'GST', "level_id": 'L0', "is_active": 1},
    'JAKK': {"stop_id": 'JAKK', "stop_code": '10031', "stop_name": 'STASIUN KCI JAKARTAKOTA', "stop_lat": -6.137360549, "stop_lon": 106.8145207, "zone_id": 'JAKK', "level_id": 'L0', "is_active": 1},
    'JAY': {"stop_id": 'JAY', "stop_code": '10032', "stop_name": 'STASIUN KCI JAYAKARTA', "stop_lat": -6.141178723, "stop_lon": 106.8231755, "zone_id": 'JAY', "level_id": 'L0', "is_active": 1},
    'JIS': {"stop_id": 'JIS', "stop_code": '10085', "stop_name": 'STASIUN KCI JAKARTA INTERNASIONAL STADIUM', "stop_lat": -6.122866, "stop_lon": 106.861705, "zone_id": 'JIS', "level_id": 'L0', "is_active": 1},
    'JMU': {"stop_id": 'JMU', "stop_code": '10034', "stop_name": 'STASIUN KCI JURANGMANGU', "stop_lat": -6.288617831, "stop_lon": 106.7291409, "zone_id": 'JMU', "level_id": 'L0', "is_active": 1},
    'JNG': {"stop_id": 'JNG', "stop_code": '10033', "stop_name": 'STASIUN KCI JATINEGARA', "stop_lat": -6.215049189, "stop_lon": 106.871031, "zone_id": 'JNG', "level_id": 'L0', "is_active": 1},
    'JTK': {"stop_id": 'JTK', "stop_code": '10084', "stop_name": 'STASIUN KCI JATAKE', "stop_lat": -6.334748, "stop_lon": 106.607319, "zone_id": 'JTK', "level_id": 'L0', "is_active": 1},
    'JUA': {"stop_id": 'JUA', "stop_code": '10035', "stop_name": 'STASIUN KCI JUANDA', "stop_lat": -6.166122924, "stop_lon": 106.8303577, "zone_id": 'JUA', "level_id": 'L0', "is_active": 1},
    'KAT': {"stop_id": 'KAT', "stop_code": '10038', "stop_name": 'STASIUN KCI KARET', "stop_lat": -6.200621674, "stop_lon": 106.8159111, "zone_id": 'KAT', "level_id": 'L0', "is_active": 1},
    'KBY': {"stop_id": 'KBY', "stop_code": '10036', "stop_name": 'STASIUN KCI KEBAYORAN', "stop_lat": -6.237002193, "stop_lon": 106.7825547, "zone_id": 'KBY', "level_id": 'L0', "is_active": 1},
    'KDS': {"stop_id": 'KDS', "stop_code": '10037', "stop_name": 'STASIUN KCI KALIDERES', "stop_lat": -6.165763759, "stop_lon": 106.7037899, "zone_id": 'KDS', "level_id": 'L0', "is_active": 1},
    'KLD': {"stop_id": 'KLD', "stop_code": '10039', "stop_name": 'STASIUN KCI KLENDER', "stop_lat": -6.211446441, "stop_lon": 106.8987328, "zone_id": 'KLD', "level_id": 'L0', "is_active": 1},
    'KLDB': {"stop_id": 'KLDB', "stop_code": '10040', "stop_name": 'STASIUN KCI KLENDERBARU', "stop_lat": -6.215627438, "stop_lon": 106.9398457, "zone_id": 'KLDB', "level_id": 'L0', "is_active": 1},
    'KMO': {"stop_id": 'KMO', "stop_code": '10041', "stop_name": 'STASIUN KCI KEMAYORAN', "stop_lat": -6.161677992, "stop_lon": 106.8415485, "zone_id": 'KMO', "level_id": 'L0', "is_active": 1},
    'KMT': {"stop_id": 'KMT', "stop_code": '10042', "stop_name": 'STASIUN KCI KRAMAT', "stop_lat": -6.193589471, "stop_lon": 106.8564668, "zone_id": 'KMT', "level_id": 'L0', "is_active": 1},
    'KPB': {"stop_id": 'KPB', "stop_code": '10043', "stop_name": 'STASIUN KCI KAMPUNGBANDAN', "stop_lat": -6.132995351, "stop_lon": 106.8285578, "zone_id": 'KPB', "level_id": 'L0', "is_active": 1},
    'KRI': {"stop_id": 'KRI', "stop_code": '10044', "stop_name": 'STASIUN KCI KRANJI', "stop_lat": -6.224277853, "stop_lon": 106.9797923, "zone_id": 'KRI', "level_id": 'L0', "is_active": 1},
    'LNA': {"stop_id": 'LNA', "stop_code": '10045', "stop_name": 'STASIUN KCI LENTENGAGUNG', "stop_lat": -6.330749058, "stop_lon": 106.834803, "zone_id": 'LNA', "level_id": 'L0', "is_active": 1},
    'MGB': {"stop_id": 'MGB', "stop_code": '10046', "stop_name": 'STASIUN KCI MANGGABESAR', "stop_lat": -6.149467954, "stop_lon": 106.8270907, "zone_id": 'MGB', "level_id": 'L0', "is_active": 1},
    'MJ': {"stop_id": 'MJ', "stop_code": '10047', "stop_name": 'STASIUN KCI MAJA', "stop_lat": -6.332109491, "stop_lon": 106.3965719, "zone_id": 'MJ', "level_id": 'L0', "is_active": 1},
    'MRI': {"stop_id": 'MRI', "stop_code": '10048', "stop_name": 'STASIUN KCI MANGGARAI', "stop_lat": -6.209717941, "stop_lon": 106.8502363, "zone_id": 'MRI', "level_id": 'L0', "is_active": 1},
    'MTR': {"stop_id": 'MTR', "stop_code": '10081', "stop_name": 'STASIUN KCI MATRAMAN', "stop_lat": -6.212370651, "stop_lon": 106.8600678, "zone_id": 'MTR', "level_id": 'L0', "is_active": 1},
    'NMO': {"stop_id": 'NMO', "stop_code": '10050', "stop_name": 'STASIUN KCI NAMBO', "stop_lat": -6.466037205, "stop_lon": 106.9061938, "zone_id": 'NMO', "level_id": 'L0', "is_active": 1},
    'PDJ': {"stop_id": 'PDJ', "stop_code": '10051', "stop_name": 'STASIUN KCI PONDOKRANJI', "stop_lat": -6.276196763, "stop_lon": 106.744919, "zone_id": 'PDJ', "level_id": 'L0', "is_active": 1},
    'PDRG': {"stop_id": 'PDRG', "stop_code": '10083', "stop_name": 'STASIUN KCI PONDOK RAJEG', "stop_lat": -6.4606892, "stop_lon": 106.8248719, "zone_id": 'PDRG', "level_id": 'L0', "is_active": 1},
    'PI': {"stop_id": 'PI', "stop_code": '10052', "stop_name": 'STASIUN KCI PORIS', "stop_lat": -6.169736028, "stop_lon": 106.6799735, "zone_id": 'PI', "level_id": 'L0', "is_active": 1},
    'PLM': {"stop_id": 'PLM', "stop_code": '10053', "stop_name": 'STASIUN KCI PALMERAH', "stop_lat": -6.207103809, "stop_lon": 106.7974836, "zone_id": 'PLM', "level_id": 'L0', "is_active": 1},
    'POC': {"stop_id": 'POC', "stop_code": '10054', "stop_name": 'STASIUN KCI PONDOKCINA', "stop_lat": -6.368769279, "stop_lon": 106.8322617, "zone_id": 'POC', "level_id": 'L0', "is_active": 1},
    'POK': {"stop_id": 'POK', "stop_code": '10055', "stop_name": 'STASIUN KCI PONDOKJATI', "stop_lat": -6.208969643, "stop_lon": 106.862284, "zone_id": 'POK', "level_id": 'L0', "is_active": 1},
    'PRP': {"stop_id": 'PRP', "stop_code": '10056', "stop_name": 'STASIUN KCI PARUNGPANJANG', "stop_lat": -6.344037522, "stop_lon": 106.5699005, "zone_id": 'PRP', "level_id": 'L0', "is_active": 1},
    'PSE': {"stop_id": 'PSE', "stop_code": '10057', "stop_name": 'STASIUN KCI PASARSENEN', "stop_lat": -6.174532997, "stop_lon": 106.844259, "zone_id": 'PSE', "level_id": 'L0', "is_active": 1},
    'PSG': {"stop_id": 'PSG', "stop_code": '10058', "stop_name": 'STASIUN KCI PESING', "stop_lat": -6.161020358, "stop_lon": 106.771574, "zone_id": 'PSG', "level_id": 'L0', "is_active": 1},
    'PSM': {"stop_id": 'PSM', "stop_code": '10059', "stop_name": 'STASIUN KCI PASARMINGGU', "stop_lat": -6.284032208, "stop_lon": 106.8445679, "zone_id": 'PSM', "level_id": 'L0', "is_active": 1},
    'PSMB': {"stop_id": 'PSMB', "stop_code": '10060', "stop_name": 'STASIUN KCI PASARMINGGUBARU', "stop_lat": -6.262679461, "stop_lon": 106.8517352, "zone_id": 'PSMB', "level_id": 'L0', "is_active": 1},
    'RJW': {"stop_id": 'RJW', "stop_code": '10061', "stop_name": 'STASIUN KCI RAJAWALI', "stop_lat": -6.144862986, "stop_lon": 106.8367707, "zone_id": 'RJW', "level_id": 'L0', "is_active": 1},
    'RK': {"stop_id": 'RK', "stop_code": '10062', "stop_name": 'STASIUN KCI RANGKAS BITUNG', "stop_lat": -6.352475988, "stop_lon": 106.2515019, "zone_id": 'RK', "level_id": 'L0', "is_active": 1},
    'RU': {"stop_id": 'RU', "stop_code": '10063', "stop_name": 'STASIUN KCI RAWABUNTU', "stop_lat": -6.314813086, "stop_lon": 106.6760449, "zone_id": 'RU', "level_id": 'L0', "is_active": 1},
    'RW': {"stop_id": 'RW', "stop_code": '10064', "stop_name": 'STASIUN KCI RAWABUAYA', "stop_lat": -6.162525592, "stop_lon": 106.723649, "zone_id": 'RW', "level_id": 'L0', "is_active": 1},
    'SDM': {"stop_id": 'SDM', "stop_code": '10065', "stop_name": 'STASIUN KCI SUDIMARA', "stop_lat": -6.296781713, "stop_lon": 106.7127431, "zone_id": 'SDM', "level_id": 'L0', "is_active": 1},
    'SRP': {"stop_id": 'SRP', "stop_code": '10066', "stop_name": 'STASIUN KCI SERPONG', "stop_lat": -6.319498896, "stop_lon": 106.6655458, "zone_id": 'SRP', "level_id": 'L0', "is_active": 1},
    'SUD': {"stop_id": 'SUD', "stop_code": '10067', "stop_name": 'STASIUN KCI SUDIRMAN', "stop_lat": -6.202308393, "stop_lon": 106.823334, "zone_id": 'SUD', "level_id": 'L0', "is_active": 1},
    'SUDB': {"stop_id": 'SUDB', "stop_code": '10082', "stop_name": 'STASIUN KCI SUDIRMAN BARU', "stop_lat": -6.200680452, "stop_lon": 106.8210886, "zone_id": 'SUDB', "level_id": 'L0', "is_active": 1},
    'SW': {"stop_id": 'SW', "stop_code": '10068', "stop_name": 'STASIUN KCI SAWAHBESAR', "stop_lat": -6.160363357, "stop_lon": 106.8276466, "zone_id": 'SW', "level_id": 'L0', "is_active": 1},
    'TB': {"stop_id": 'TB', "stop_code": '10069', "stop_name": 'STASIUN KCI TAMBUN', "stop_lat": -6.258551922, "stop_lon": 107.0558163, "zone_id": 'TB', "level_id": 'L0', "is_active": 1},
    'TEB': {"stop_id": 'TEB', "stop_code": '10070', "stop_name": 'STASIUN KCI TEBET', "stop_lat": -6.226237044, "stop_lon": 106.8584686, "zone_id": 'TEB', "level_id": 'L0', "is_active": 1},
    'TEJ': {"stop_id": 'TEJ', "stop_code": '10071', "stop_name": 'STASIUN KCI TENJO', "stop_lat": -6.327075216, "stop_lon": 106.4613482, "zone_id": 'TEJ', "level_id": 'L0', "is_active": 1},
    'TGS': {"stop_id": 'TGS', "stop_code": '10073', "stop_name": 'STASIUN KCI TIGARAKSA', "stop_lat": -6.328178972, "stop_lon": 106.4346166, "zone_id": 'TGS', "level_id": 'L0', "is_active": 1},
    'THB': {"stop_id": 'THB', "stop_code": '10072', "stop_name": 'STASIUN KCI TANAHABANG', "stop_lat": -6.185403437, "stop_lon": 106.8109823, "zone_id": 'THB', "level_id": 'L0', "is_active": 1},
    'THI': {"stop_id": 'THI', "stop_code": '10078', "stop_name": 'STASIUN KCI TANAHTINGGI', "stop_lat": -6.175260697, "stop_lon": 106.6454584, "zone_id": 'THI', "level_id": 'L0', "is_active": 1},
    'TKO': {"stop_id": 'TKO', "stop_code": '10074', "stop_name": 'STASIUN KCI TAMANKOTA', "stop_lat": -6.159373791, "stop_lon": 106.7587657, "zone_id": 'TKO', "level_id": 'L0', "is_active": 1},
    'TLM': {"stop_id": 'TLM', "stop_code": '10049', "stop_name": 'STASIUN KCI METLAND TELAGAMURNI', "stop_lat": -6.256882716, "stop_lon": 107.1084045, "zone_id": 'TLM', "level_id": 'L0', "is_active": 1},
    'TNG': {"stop_id": 'TNG', "stop_code": '10075', "stop_name": 'STASIUN KCI TANGERANG', "stop_lat": -6.176651797, "stop_lon": 106.6302066, "zone_id": 'TNG', "level_id": 'L0', "is_active": 1},
    'TNT': {"stop_id": 'TNT', "stop_code": '10076', "stop_name": 'STASIUN KCI TANJUNGBARAT', "stop_lat": -6.307602971, "stop_lon": 106.8387979, "zone_id": 'TNT', "level_id": 'L0', "is_active": 1},
    'TPK': {"stop_id": 'TPK', "stop_code": '10077', "stop_name": 'STASIUN KCI TANJUNGPRIUK', "stop_lat": -6.110665263, "stop_lon": 106.8812262, "zone_id": 'TPK', "level_id": 'L0', "is_active": 1},
    'UI': {"stop_id": 'UI', "stop_code": '10079', "stop_name": 'STASIUN KCI UNIV.INDONESIA', "stop_lat": -6.360444874, "stop_lon": 106.8317052, "zone_id": 'UI', "level_id": 'L0', "is_active": 1},
    'UP': {"stop_id": 'UP', "stop_code": '10080', "stop_name": 'STASIUN KCI UNIV.PANCASILA', "stop_lat": -6.338502657, "stop_lon": 106.8343745, "zone_id": 'UP', "level_id": 'L0', "is_active": 1},
}
STATION_ALIASES = {
    'BATU CEPER': 'BPR',
    'BATU-CEPER': 'BPR',
    'BATUCEPER': 'BPR',
    'BEKASI TIMUR': 'BKST',
    'BEKASI-TIMUR': 'BKST',
    'BOJONG GEDE': 'BJD',
    'BOJONG INDAH': 'BOI',
    'BOJONG-GEDE': 'BJD',
    'BOJONG-INDAH': 'BOI',
    'BOJONGGEDE': 'BJD',
    'BOJONGINDAH': 'BOI',
    'CAKUNG': 'CUK',
    'CAWANG': 'CW',
    'CIBINONG': 'CBN',
    'CIBITUNG': 'CIT',
    'CIKARANG': 'CKR',
    'CIKINI': 'CKI',
    'CIKOYA': 'CKY',
    'CILEBUT': 'CLT',
    'CILEJIT': 'CJT',
    'CISAUK': 'CSK',
    'CITAYAM': 'CTA',
    'CITERAS': 'CTR',
    'DARU': 'DAR',
    'DEPOK': 'DP',
    'DEPOK BARU': 'DPB',
    'DEPOK-BARU': 'DPB',
    'DEPOKBARU': 'DPB',
    'DUREN KALIBATA': 'DRN',
    'DUREN-KALIBATA': 'DRN',
    'DURENKALIBATA': 'DRN',
    'DURI': 'DU',
    'GANG SENTIONG': 'GST',
    'GANG-SENTIONG': 'GST',
    'GANGSENTIONG': 'GST',
    'GONDANGDIA': 'GDD',
    'GROGOL': 'GRG',
    'JAKARTA INTERNASIONAL STADIUM': 'JIS',
    'JAKARTA KOTA': 'JAKK',
    'JAKARTA KOTA (KOTA)': 'JAKK',
    'JAKARTA-KOTA': 'JAKK',
    'JAKARTAKOTA': 'JAKK',
    'JATAKE': 'JTK',
    'JATINEGARA': 'JNG',
    'JAYAKARTA': 'JAY',
    'JUANDA': 'JUA',
    'JURANG MANGU': 'JMU',
    'JURANG-MANGU': 'JMU',
    'JURANGMANGU': 'JMU',
    'KALIDERES': 'KDS',
    'KAMPUNG BANDAN': 'KPB',
    'KAMPUNG-BANDAN': 'KPB',
    'KAMPUNGBANDAN': 'KPB',
    'KARET': 'KAT',
    'KEBAYORAN': 'KBY',
    'KEMAYORAN': 'KMO',
    'KLENDER': 'KLD',
    'KLENDER BARU': 'KLDB',
    'KLENDER-BARU': 'KLDB',
    'KLENDERBARU': 'KLDB',
    'KRAMAT': 'KMT',
    'KRANJI': 'KRI',
    'LENTENG AGUNG': 'LNA',
    'LENTENG-AGUNG': 'LNA',
    'LENTENGAGUNG': 'LNA',
    'MAJA': 'MJ',
    'MANGGA BESAR': 'MGB',
    'MANGGA-BESAR': 'MGB',
    'MANGGABESAR': 'MGB',
    'MANGGARAI': 'MRI',
    'MATRAMAN': 'MTR',
    'METLAND TELAGAMURNI': 'TLM',
    'METLAND-TELAGAMURNI': 'TLM',
    'NAMBO': 'NMO',
    'PALMERAH': 'PLM',
    'PASAR MINGGU': 'PSM',
    'PASAR MINGGU BARU': 'PSMB',
    'PASAR MINGGU BDH': 'PSMB',
    'PASAR SENEN': 'PSE',
    'PASAR-MINGGU': 'PSM',
    'PASAR-MINGGU-BARU': 'PSMB',
    'PASAR-SENEN': 'PSE',
    'PASARMINGGU': 'PSM',
    'PASARMINGGUBARU': 'PSMB',
    'PASARSENEN': 'PSE',
    'PONDOK CINA': 'POC',
    'PONDOK JATI': 'POK',
    'PONDOK RAJEG': 'PDRG',
    'PONDOK RANJI': 'PDJ',
    'PONDOK-CINA': 'POC',
    'PONDOK-JATI': 'POK',
    'PONDOK-RAJEG': 'PDRG',
    'PONDOK-RANJI': 'PDJ',
    'PONDOKCINA': 'POC',
    'PONDOKJATI': 'POK',
    'PONDOKRAJEG': 'PDRG',
    'PONDOKRANJI': 'PDJ',
    'PORIS': 'PI',
    'RAJAWALI': 'RJW',
    'RANGKAS BITUNG': 'RK',
    'RANGKAS-BITUNG': 'RK',
    'RANGKASBITUNG': 'RK',
    'RAWA BUAYA': 'RW',
    'RAWA BUNTU': 'RU',
    'RAWA-BUAYA': 'RW',
    'RAWA-BUNTU': 'RU',
    'RAWABUAYA': 'RW',
    'RAWABUNTU': 'RU',
    'SAWAH BESAR': 'SW',
    'SAWAH-BESAR': 'SW',
    'SAWAHBESAR': 'SW',
    'SERPONG': 'SRP',
    'STASIUN KCI ANCOL': 'AC',
    'STASIUN KCI ANGKE': 'AK',
    'STASIUN KCI BATUCEPER': 'BPR',
    'STASIUN KCI BEKASI': 'BKS',
    'STASIUN KCI BEKASI TIMUR': 'BKST',
    'STASIUN KCI BOGOR': 'BOO',
    'STASIUN KCI BOJONGGEDE': 'BJD',
    'STASIUN KCI BOJONGINDAH': 'BOI',
    'STASIUN KCI BUARAN': 'BUA',
    'STASIUN KCI CAKUNG': 'CUK',
    'STASIUN KCI CAWANG': 'CW',
    'STASIUN KCI CIBINONG': 'CBN',
    'STASIUN KCI CIBITUNG': 'CIT',
    'STASIUN KCI CICAYUR': 'CC',
    'STASIUN KCI CIKARANG': 'CKR',
    'STASIUN KCI CIKINI': 'CKI',
    'STASIUN KCI CIKOYA': 'CKY',
    'STASIUN KCI CILEBUT': 'CLT',
    'STASIUN KCI CILEJIT': 'CJT',
    'STASIUN KCI CISAUK': 'CSK',
    'STASIUN KCI CITAYAM': 'CTA',
    'STASIUN KCI CITERAS': 'CTR',
    'STASIUN KCI DARU': 'DAR',
    'STASIUN KCI DEPOK': 'DP',
    'STASIUN KCI DEPOKBARU': 'DPB',
    'STASIUN KCI DURENKALIBATA': 'DRN',
    'STASIUN KCI DURI': 'DU',
    'STASIUN KCI GANGSENTIONG': 'GST',
    'STASIUN KCI GONDANGDIA': 'GDD',
    'STASIUN KCI GROGOL': 'GRG',
    'STASIUN KCI JAKARTA INTERNASIONAL STADIUM': 'JIS',
    'STASIUN KCI JAKARTAKOTA': 'JAKK',
    'STASIUN KCI JATAKE': 'JTK',
    'STASIUN KCI JATINEGARA': 'JNG',
    'STASIUN KCI JAYAKARTA': 'JAY',
    'STASIUN KCI JUANDA': 'JUA',
    'STASIUN KCI JURANGMANGU': 'JMU',
    'STASIUN KCI KALIDERES': 'KDS',
    'STASIUN KCI KAMPUNGBANDAN': 'KPB',
    'STASIUN KCI KARET': 'KAT',
    'STASIUN KCI KEBAYORAN': 'KBY',
    'STASIUN KCI KEMAYORAN': 'KMO',
    'STASIUN KCI KLENDER': 'KLD',
    'STASIUN KCI KLENDERBARU': 'KLDB',
    'STASIUN KCI KRAMAT': 'KMT',
    'STASIUN KCI KRANJI': 'KRI',
    'STASIUN KCI LENTENGAGUNG': 'LNA',
    'STASIUN KCI MAJA': 'MJ',
    'STASIUN KCI MANGGABESAR': 'MGB',
    'STASIUN KCI MANGGARAI': 'MRI',
    'STASIUN KCI MATRAMAN': 'MTR',
    'STASIUN KCI METLAND TELAGAMURNI': 'TLM',
    'STASIUN KCI NAMBO': 'NMO',
    'STASIUN KCI PALMERAH': 'PLM',
    'STASIUN KCI PARUNGPANJANG': 'PRP',
    'STASIUN KCI PASARMINGGU': 'PSM',
    'STASIUN KCI PASARMINGGUBARU': 'PSMB',
    'STASIUN KCI PASARSENEN': 'PSE',
    'STASIUN KCI PESING': 'PSG',
    'STASIUN KCI PONDOK RAJEG': 'PDRG',
    'STASIUN KCI PONDOKCINA': 'POC',
    'STASIUN KCI PONDOKJATI': 'POK',
    'STASIUN KCI PONDOKRANJI': 'PDJ',
    'STASIUN KCI PORIS': 'PI',
    'STASIUN KCI RAJAWALI': 'RJW',
    'STASIUN KCI RANGKAS BITUNG': 'RK',
    'STASIUN KCI RAWABUAYA': 'RW',
    'STASIUN KCI RAWABUNTU': 'RU',
    'STASIUN KCI SAWAHBESAR': 'SW',
    'STASIUN KCI SERPONG': 'SRP',
    'STASIUN KCI SUDIMARA': 'SDM',
    'STASIUN KCI SUDIRMAN': 'SUD',
    'STASIUN KCI SUDIRMAN BARU': 'SUDB',
    'STASIUN KCI TAMANKOTA': 'TKO',
    'STASIUN KCI TAMBUN': 'TB',
    'STASIUN KCI TANAHABANG': 'THB',
    'STASIUN KCI TANAHTINGGI': 'THI',
    'STASIUN KCI TANGERANG': 'TNG',
    'STASIUN KCI TANJUNGBARAT': 'TNT',
    'STASIUN KCI TANJUNGPRIUK': 'TPK',
    'STASIUN KCI TEBET': 'TEB',
    'STASIUN KCI TENJO': 'TEJ',
    'STASIUN KCI TIGARAKSA': 'TGS',
    'STASIUN KCI UNIV.INDONESIA': 'UI',
    'STASIUN KCI UNIV.PANCASILA': 'UP',
    'SUDIMARA': 'SDM',
    'SUDIRMAN': 'SUD',
    'SUDIRMAN BARU': 'SUDB',
    'SUDIRMAN-BARU': 'SUDB',
    'SUDIRMN BARU': 'SUDB',
    'TAMAN KOTA': 'TKO',
    'TAMAN-KOTA': 'TKO',
    'TAMANKOTA': 'TKO',
    'TAMBUN': 'TB',
    'TANAH ABANG': 'THB',
    'TANAH TINGGI': 'THI',
    'TANAH-ABANG': 'THB',
    'TANAH-TINGGI': 'THI',
    'TANAHABANG': 'THB',
    'TANAHTINGGI': 'THI',
    'TANGERANG': 'TNG',
    'TANJUNG BARAT': 'TNT',
    'TANJUNG PRIOK': 'TPK',
    'TANJUNG PRIUK': 'TPK',
    'TANJUNG-BARAT': 'TNT',
    'TANJUNG-PRIUK': 'TPK',
    'TANJUNGBARAT': 'TNT',
    'TANJUNGPRIOK': 'TPK',
    'TANJUNGPRIUK': 'TPK',
    'TEBET': 'TEB',
    'TENJO': 'TEJ',
    'TIGARAKSA': 'TGS',
    'UNIV INDONESIA': 'UI',
    'UNIV PANCASILA': 'UP',
    'UNIV.INDONESIA': 'UI',
    'UNIV.PANCASILA': 'UP',
    'UNIVERSITAS INDONESIA': 'UI',
    'UNIVERSITAS PANCASILA': 'UP',
}



# ===========================================================================
# Time helpers
# ===========================================================================
def fmt_hhmmss(seconds: int) -> str:
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def time_to_str(t):
    if t is None:
        return ""
    return t.strftime("%H:%M:%S")


def time_to_minutes(t):
    if t is None:
        return None
    return t.hour * 60 + t.minute + t.second / 60


def minutes_to_gtfs_str(minutes):
    h = int(minutes // 60)
    m = int(minutes % 60)
    s = int(round(minutes - int(minutes)) * 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def sanitize_trip_id(value):
    """Strip a trip_id down to characters GTFS tolerates."""
    if value is None:
        return ""
    s = str(value)
    s = re.sub(r"[^\x20-\x7E]", "", s).strip()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^A-Za-z0-9_\-]", "-", s)
    while "--" in s:
        s = s.replace("--", "-")
    return s.strip("-_")


def normalize_trip_times(stops):
    """Correct per-trip timings in minutes.

    For every non-express stop:
      - arrival <= departure within the stop
      - time never goes backwards between stops
      - across midnight the clock rolls into 24:xx / 25:xx instead of 00:xx
      - small backwards anomalies are nudged instead of jumping a day

    Returns copies of the stops carrying corrected 'arrival_min' /
    'departure_min' values.
    """
    adjusted = []
    prev_dep = None
    for s in stops:
        ns = dict(s)
        adjusted.append(ns)
        if s.get("express"):
            continue

        arr = time_to_minutes(s.get("arrival"))
        dep = time_to_minutes(s.get("departure"))
        if arr is None and dep is None:
            ns["arrival_min"] = None
            ns["departure_min"] = None
            continue
        if arr is None:
            arr = dep
        if dep is None:
            dep = arr

        base = 0
        if prev_dep is not None and arr + base < prev_dep:
            if (arr + 1440) - prev_dep <= 720:
                base = 1440
                while arr + base < prev_dep:
                    base += 1440
            else:
                arr = prev_dep

        arr_f = arr + base
        dep_f = dep + base
        if dep_f < arr_f:
            if dep <= 360 and arr_f >= 1200:
                dep_f += 1440
            else:
                dep_f = arr_f

        ns["arrival_min"] = int(round(arr_f))
        ns["departure_min"] = int(round(dep_f))
        prev_dep = dep_f

    return adjusted


# ===========================================================================
# Station activity
# ===========================================================================
def is_station_active(station_code):
    s = STATIONS.get(station_code)
    if s is None:
        return True
    return s.get("is_active", 1) != 0


def get_inactive_stations():
    return sorted(code for code, s in STATIONS.items() if s.get("is_active", 1) == 0)


def get_active_stations():
    return sorted(code for code, s in STATIONS.items() if s.get("is_active", 1) != 0)


def set_station_active(station_code, active):
    if station_code in STATIONS:
        STATIONS[station_code]["is_active"] = 1 if active else 0


def reset_station_activity():
    for s in STATIONS.values():
        s["is_active"] = 1


def add_station(station_code, stop_code="", stop_name="", stop_lat=0.0, stop_lon=0.0, zone_id=None):
    """Add a station to the database. Returns True if it was added."""
    if not station_code or station_code in STATIONS:
        return False
    STATIONS[station_code] = {
        "stop_id": station_code,
        "stop_code": stop_code or station_code,
        "stop_name": stop_name or station_code,
        "stop_lat": stop_lat,
        "stop_lon": stop_lon,
        "zone_id": zone_id or station_code,
        "level_id": "L0",
        "is_active": 1,
    }
    return True


def set_station_info(station_code, stop_code=None, stop_name=None):
    """Update stop_code / stop_name of an existing station."""
    if station_code not in STATIONS:
        return False
    s = STATIONS[station_code]
    if stop_code not in (None, ""):
        s["stop_code"] = str(stop_code).strip()
    if stop_name not in (None, ""):
        s["stop_name"] = str(stop_name).strip()
    return True


# ===========================================================================
# Station code resolution
# ===========================================================================
def resolve_station_code(code):
    """Map a timetable station label onto its database stop code."""
    if code in STATIONS:
        return code
    upper = code.upper().strip()
    if upper in STATIONS:
        return upper
    if upper in STATION_ALIASES:
        return STATION_ALIASES[upper]
    normalized = upper.replace(".", "").replace(" ", "").replace("-", "")
    for alias, target in STATION_ALIASES.items():
        alias_norm = alias.replace(".", "").replace(" ", "").replace("-", "")
        if normalized == alias_norm:
            return target
    for station_code in STATIONS:
        if station_code.upper() == upper:
            return station_code
    return code


def is_ls_value(val):
    if val is None:
        return False
    return str(val).strip() in LS_VALUES


def parse_time_value(val):
    if val is None or val == "" or val == "-":
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, str):
        val = val.strip()
        if val in ("", "-", "Ls", "Ls.", "LS"):
            return None
        try:
            parts = val.split(":")
            h, m = int(parts[0]), int(parts[1])
            s = int(parts[2]) if len(parts) > 2 else 0
            return time(h, m, s)
        except (ValueError, IndexError):
            return None
    return None


# ===========================================================================
# Excel (GAPEKA) parser
# ===========================================================================
TYPE_MARKERS = ("Ber", "Dat", "Ls")

_HEADER_LABELS = (
    "NO", "NO.", "No", "NO KA", "NOKA", "RELASI", "SF", "LOOP", "STASIUN",
    "KET", "KETERANGAN", "STASIUN/PERHENTIAN", "STASIUN/PERHENT",
)


def find_type_row(ws, start=1, stop=6):
    for r in range(start, stop + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val and str(val).strip() in TYPE_MARKERS:
                return r
    return None


def detect_sheet_type(ws):
    type_row = find_type_row(ws)
    if type_row is None or type_row != 2:
        return "type_a"
    return "type_b"


def find_ket_col(ws):
    for c in range(1, ws.max_column + 1):
        for r in range(1, 4):
            val = ws.cell(r, c).value
            if val and str(val).strip().upper() in ("KET", "KETERANGAN"):
                return c
    return None


def parse_ket_value(ket_val):
    """Map the KETERANGAN column onto a calendar service_id.

    Returns None when the trip is cancelled, which drops the trip.
    """
    if ket_val is None:
        return "AllDay"
    ket_str = str(ket_val).strip().upper()
    if not ket_str:
        return "AllDay"
    if "BATAL" in ket_str:
        return None
    if "WEEKDAY" in ket_str:
        return "Weekday"
    if "WEEKEND" in ket_str:
        return "Weekend"
    return "AllDay"


def _collect_station_columns(labels, types, skip_headers, col_offset=0):
    """Shared header scan: returns (stations_order, col_map).

    stations_order is [(code, first_column_index)]; col_map maps a code onto its
    Ber/Dat/Ls column positions. col_offset is the worksheet column that
    labels[0] came from, so a sliced scan still records absolute columns.
    """
    stations_order = []
    col_map = {}
    for i, (label, btype) in enumerate(zip(labels, types)):
        if label is None:
            continue
        code = str(label).strip()
        if code == "" or (skip_headers and code in _HEADER_LABELS):
            continue
        if code in SKIP_STATIONS:
            continue
        code = resolve_station_code(code)
        if code in SKIP_STATIONS:
            continue
        btype_str = str(btype).strip() if btype else ""
        if btype_str not in TYPE_MARKERS:
            continue
        col_idx = col_offset + i + 1
        entry = col_map.setdefault(code, {"ber_col": None, "dat_col": None, "ls_col": None})
        if btype_str == "Ber":
            entry["ber_col"] = col_idx
        elif btype_str == "Dat":
            entry["dat_col"] = col_idx
        else:
            entry["ls_col"] = col_idx
        if code not in [c for c, _ in stations_order]:
            stations_order.append((code, col_idx))
    return stations_order, col_map


def _stops_for_row(ws, row_idx, stations_order, col_map):
    """Build the stop list for one timetable row."""
    stops = []
    for code, _col_idx in stations_order:
        info = col_map.get(code)
        if not info:
            continue
        if info["ls_col"]:
            if is_ls_value(ws.cell(row_idx, info["ls_col"]).value):
                stops.append({"station": code, "arrival": None, "departure": None, "express": True})
                continue
        arrival_time = None
        departure_time = None
        if info["dat_col"]:
            dat_raw = ws.cell(row_idx, info["dat_col"]).value
            if is_ls_value(dat_raw):
                stops.append({"station": code, "arrival": None, "departure": None, "express": True})
                continue
            arrival_time = parse_time_value(dat_raw)
        if info["ber_col"]:
            departure_time = parse_time_value(ws.cell(row_idx, info["ber_col"]).value)
        if arrival_time or departure_time:
            stops.append({"station": code, "arrival": arrival_time, "departure": departure_time, "express": False})
    return stops


def parse_type_a(ws, sheet_name):
    type_row = find_type_row(ws)
    station_row = type_row - 1 if type_row else 3

    labels = [ws.cell(station_row, c).value for c in range(1, ws.max_column + 1)]
    types = [ws.cell(type_row, c).value for c in range(1, ws.max_column + 1)]
    stations_order, col_map = _collect_station_columns(labels, types, skip_headers=True)

    ket_col = find_ket_col(ws)
    data_start = type_row + 1

    trains = []
    for row_idx in range(data_start, ws.max_row + 1):
        no_ka = ws.cell(row_idx, 2).value
        if no_ka is None:
            continue
        no_ka_str = str(no_ka).strip()
        if no_ka_str == "" or no_ka_str == "NO KA":
            continue
        relasi = ws.cell(row_idx, 3).value
        relasi_str = str(relasi).strip() if relasi else ""
        service_id = parse_ket_value(ws.cell(row_idx, ket_col).value if ket_col else None)
        if service_id is None:
            continue
        stops = _stops_for_row(ws, row_idx, stations_order, col_map)
        if stops:
            trains.append({"no_ka": no_ka_str, "relasi": relasi_str, "service_id": service_id, "stops": stops})

    return stations_order, trains


def parse_type_b(ws, sheet_name):
    row1 = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    row2 = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)]
    stations_order, col_map = _collect_station_columns(row1, row2, skip_headers=True)

    no_col = noka_col = relasi_col = None
    for i, h in enumerate(row1):
        h_str = str(h).strip() if h else ""
        if h_str in ("NO", "NO.", "No"):
            no_col = i + 1
        elif h_str in ("NO KA", "NOKA"):
            noka_col = i + 1
        elif h_str == "RELASI":
            relasi_col = i + 1

    ket_col = find_ket_col(ws)
    ka_col = noka_col or 2

    trains = []
    for row_idx in range(3, ws.max_row + 1):
        no_ka = ws.cell(row_idx, ka_col).value
        if no_ka is None:
            continue
        no_ka_str = str(no_ka).strip()
        if no_ka_str == "" or no_ka_str.upper() in ("NO KA", "NOKA"):
            continue
        relasi = ws.cell(row_idx, relasi_col).value if relasi_col else ws.cell(row_idx, 3).value
        relasi_str = str(relasi).strip() if relasi else ""
        service_id = parse_ket_value(ws.cell(row_idx, ket_col).value if ket_col else None)
        if service_id is None:
            continue
        stops = _stops_for_row(ws, row_idx, stations_order, col_map)
        if stops:
            trains.append({"no_ka": no_ka_str, "relasi": relasi_str, "service_id": service_id, "stops": stops})

    return stations_order, trains


def parse_jakk_tpk(ws):
    """The Jakarta Kota - Tanjung Priok sheet packs two directions side by side."""
    row4 = [ws.cell(4, c).value for c in range(1, ws.max_column + 1)]
    row5 = [ws.cell(5, c).value for c in range(1, ws.max_column + 1)]

    # The sheet lists both directions side by side, and TPK appears twice: once
    # as the left direction's terminus (Dat) and once as the right direction's
    # origin (Ber). The split happens at the first TPK.
    split_at = None
    for i, label in enumerate(row4):
        if label is not None and str(label).strip() == "TPK":
            split_at = i
            break
    if split_at is None:
        split_at = len(row4)

    left_order, left_map = _collect_station_columns(row4[:split_at], row5[:split_at], skip_headers=True)
    right_order, right_map = _collect_station_columns(row4[split_at:], row5[split_at:], skip_headers=True, col_offset=split_at)

    ket_col = find_ket_col(ws)
    left_noka = left_relasi = right_noka = right_relasi = None
    for col in range(1, ws.max_column + 1):
        header_val = ws.cell(3, col).value
        if header_val is None:
            continue
        h = str(header_val).strip().upper()
        if col <= 12:
            if h == "NO KA":
                left_noka = col
            elif h == "RELASI":
                left_relasi = col
        else:
            if h == "NO KA":
                right_noka = col
            elif h == "RELASI":
                right_relasi = col

    def collect(noka_col, relasi_col, order, col_map):
        found = []
        for row_idx in range(6, ws.max_row + 1):
            no_ka = ws.cell(row_idx, noka_col).value if noka_col else None
            if not (no_ka and str(no_ka).strip()):
                continue
            relasi = ws.cell(row_idx, relasi_col).value if relasi_col else None
            service_id = parse_ket_value(ws.cell(row_idx, ket_col).value if ket_col else None)
            if service_id is None:
                continue
            stops = _stops_for_row(ws, row_idx, order, col_map)
            if stops:
                found.append({
                    "no_ka": str(no_ka).strip(),
                    "relasi": str(relasi).strip() if relasi else "",
                    "service_id": service_id,
                    "stops": stops,
                })
        return found

    left_trains = collect(left_noka, left_relasi, left_order, left_map)
    right_trains = collect(right_noka, right_relasi, right_order, right_map)
    return left_order, left_trains, right_order, right_trains


def parse_excel(filepath):
    """Parse a GAPEKA KCI workbook -> {sheet_key: {...}}."""
    if openpyxl is None:
        raise RuntimeError("Membaca --excel butuh openpyxl. Jalankan: pip install openpyxl")
    wb = openpyxl.load_workbook(filepath, data_only=True)
    results = {}
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if sheet_name == "JAKK-TPK":
                left_stations, left_trains, right_stations, right_trains = parse_jakk_tpk(ws)
                results["JAKK-TPK_left"] = {
                    "sheet_name": sheet_name,
                    "route_config": JAKK_TPK_MAP["left"],
                    "stations_order": left_stations,
                    "trains": left_trains,
                }
                results["JAKK-TPK_right"] = {
                    "sheet_name": sheet_name,
                    "route_config": JAKK_TPK_MAP["right"],
                    "stations_order": right_stations,
                    "trains": right_trains,
                }
            elif sheet_name in SHEET_ROUTE_MAP:
                if detect_sheet_type(ws) == "type_a":
                    stations, trains = parse_type_a(ws, sheet_name)
                else:
                    stations, trains = parse_type_b(ws, sheet_name)
                results[sheet_name] = {
                    "sheet_name": sheet_name,
                    "route_config": SHEET_ROUTE_MAP[sheet_name],
                    "stations_order": stations,
                    "trains": trains,
                }
    finally:
        wb.close()
    return results



# ===========================================================================
# PDF (public KCI schedule) parser
# ===========================================================================
TIME_RE = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")
NO_RE = re.compile(r"^\d{1,3}$")
KA_RE = re.compile(r"^\d{3,4}$")
RELASI_RE = re.compile(r"^[A-Za-z]{2,4}(?:-[A-Za-z]{2,4})+$")
CODE_RE = re.compile(r"^[A-Za-z]{2,4}$")

NOTE_BANNED = ("BATAL", "SABTU", "MINGGU", "LIBUR")

_HEADER_NOISE = {
    "NO", "NO.", "NOMOR", "KA", "RELASI", "KETERANGAN", "STASIUN",
    "STASIUN/PERHENTIAN", "STASIUN/PERHENT", "JADWAL", "COMMUTER",
    "LINE", "LIN", "TUJUAN", "TIMETABLE", "OF", "FOR", "TAP",
    "SF", "LOOP", "BER", "DAT", "LS",
}


def download_pdf(url, timeout=60):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _split_tables(word_boxes):
    """Locate the horizontal span of each timetable, bounded by NOMOR..KETERANGAN."""
    noms = sorted((w for w in word_boxes if w[4] == "NOMOR"), key=lambda w: w[0])
    kets = sorted((w for w in word_boxes if w[4] == "KETERANGAN"), key=lambda w: w[2])
    tables = []
    for n in noms:
        to_right = [k for k in kets if k[2] >= n[0]]
        if not to_right:
            continue
        k = min(to_right, key=lambda x: x[0])
        tables.append([n[0] - 18, k[2] + 18])
    if not tables:
        tables = [[min(w[0] for w in word_boxes), max(w[2] for w in word_boxes)]]

    merged = []
    for t in sorted(tables, key=lambda x: x[0]):
        if merged and t[0] < merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], t[1])
        else:
            merged.append(list(t))
    return merged


def _rows_by_y(word_boxes):
    rows = {}
    for w in word_boxes:
        rows.setdefault(round(w[1], 1), []).append(w)
    return rows


def _table_header(word_boxes):
    """Return (header_y, [(code, x0)]) for one timetable."""
    rows = _rows_by_y(word_boxes)
    ordered = sorted(rows)
    first_data_y = None
    for y in ordered:
        if len([w for w in rows[y] if TIME_RE.match(w[4])]) >= 2:
            first_data_y = y
            break
    if first_data_y is None:
        return None, []

    code_words = []
    for y in ordered:
        if y >= first_data_y:
            break
        row_codes = [w for w in rows[y] if CODE_RE.match(w[4]) and w[4].upper() not in _HEADER_NOISE]
        if len(row_codes) >= 2:
            code_words.extend(row_codes)
    if len(code_words) < 2:
        return None, []

    header_y = min(w[1] for w in code_words)
    code_words.sort(key=lambda w: (w[1], w[0]))
    stations = []
    seen = set()
    for w in code_words:
        code = resolve_station_code(w[4].strip())
        if code in SKIP_STATIONS or code in seen:
            continue
        seen.add(code)
        stations.append((code, w[0]))
    return header_y, stations


def _table_title(word_boxes, header_y):
    tokens = [w for w in word_boxes if w[1] < header_y - 2]
    tokens.sort(key=lambda w: (w[1], w[0]))
    text = " ".join(w[4] for w in tokens).split(" Timetable")[0]
    m = re.match(
        r"Jadwal\s+Commuter\s+Lin\w*\s*(?P<line>.+?)\s+tujuan\s+(?P<dest>.+)$",
        text, re.IGNORECASE,
    )
    if not m:
        return None
    return m.group("line").strip(), m.group("dest").strip()


def _row_groups(word_boxes, header_y):
    tokens = [w for w in word_boxes if w[1] > header_y + 1]
    tokens.sort(key=lambda w: (w[1], w[0]))
    groups = []
    for w in tokens:
        if groups and w[1] - groups[-1]["base"] <= 3.0:
            groups[-1]["words"].append(w)
        else:
            groups.append({"base": w[1], "words": [w]})
    return groups


def _parse_trains(word_boxes, header_y, station_cols):
    if not station_cols:
        return []
    first_stn_x = min(x0 for _, x0 in station_cols)
    last_stn_x = max(x0 for _, x0 in station_cols)
    groups = _row_groups(word_boxes, header_y)
    all_box = [w for w in word_boxes if w[1] > header_y + 1]

    trains = []
    for g in groups:
        words = g["words"]
        relasi_cands = [w for w in words if RELASI_RE.match(w[4]) and 40 <= w[0] < first_stn_x]
        if not relasi_cands:
            continue
        relasi = relasi_cands[0][4]

        digits = sorted((w for w in words if NO_RE.match(w[4]) and w[0] < first_stn_x), key=lambda x: x[0])
        no = digits[0][4] if digits else ""
        no_ka = ""
        for d in digits[1:]:
            if KA_RE.match(d[4]) and d[4] != no:
                no_ka = d[4]
                break
        if not no_ka:
            best = None
            for w in all_box:
                if not KA_RE.match(w[4]) or w[0] < 40 or w[0] >= first_stn_x:
                    continue
                dy = abs(w[1] - g["base"])
                if dy <= 22 and (best is None or dy < best[1]):
                    best = (w, dy)
            if best:
                no_ka = best[0][4]
        if not no_ka:
            no_ka = no

        note_zone = [w for w in words if w[0] > last_stn_x + 15]
        if note_zone:
            note = " ".join(w[4] for w in sorted(note_zone, key=lambda x: x[0])).upper()
            if any(b in note for b in NOTE_BANNED):
                continue

        stops = []
        for code, cx in station_cols:
            cand = [w for w in words if abs(w[0] - cx) <= 9 and w[2] < last_stn_x + 5]
            if not cand:
                continue
            cell = min(cand, key=lambda x: abs(x[0] - cx))[4]
            if cell in LS_VALUES:
                stops.append({"station": code, "arrival": None, "departure": None, "express": True})
            elif TIME_RE.match(cell):
                parts = cell.split(":")
                t = time(int(parts[0]), int(parts[1]))
                stops.append({"station": code, "arrival": t, "departure": t, "express": False})

        if stops:
            trains.append({"no_ka": no_ka, "relasi": relasi, "service_id": "AllDay", "stops": stops})
    return trains


def _map_route(line, dest):
    """Route a PDF timetable title onto a route-mapping key."""
    l = (line or "").upper()
    d = (dest or "").upper()
    if "BOGOR" in l:
        return "BOO-JAKK" if "JAKARTA" in d else "JAKK-BOO"
    if "BEKASI" in l:
        return "CKR-MRI-KPB" if "KAMPUNG" in d else "KPB-PSE-CKR"
    if "RANGKAS" in l:
        return "SRP-THB" if "TANAH" in d else "THB-SRP"
    if "TANGERANG" in l:
        return "TNG-DU" if "DURI" in d else "DU-TNG"
    if "PRIOK" in l or "TANJUNG" in l:
        return "JAKK-TPK_right" if "JAKARTA" in d else "JAKK-TPK_left"
    return None


def _route_config(key):
    if key.endswith("_left"):
        return JAKK_TPK_MAP["left"]
    if key.endswith("_right"):
        return JAKK_TPK_MAP["right"]
    return SHEET_ROUTE_MAP.get(key)


def parse_pdf(pdf_bytes):
    """Parse the public KCI schedule PDF -> {sheet_key: {...}}.

    A PDF only carries one time per station, so arrival == departure, and every
    trip is treated as all-day service. Trips whose notes say they do not run
    on Saturday, Sunday or public holidays are dropped.
    """
    if pymupdf is None:
        raise RuntimeError("Membaca --pdf butuh pymupdf. Jalankan: pip install pymupdf")
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    results = {}
    try:
        for page in doc:
            boxes = page.get_text("words")
            for x0, x1 in _split_tables(boxes):
                tbl = [w for w in boxes if x0 - 1 <= w[0] and w[2] <= x1 + 1]
                header_y, station_cols = _table_header(tbl)
                if not station_cols:
                    continue
                title = _table_title(tbl, header_y)
                if not title:
                    continue
                key = _map_route(title[0], title[1])
                if not key:
                    continue
                trains = _parse_trains(tbl, header_y, station_cols)
                if not trains:
                    continue
                if key not in results:
                    cfg = _route_config(key)
                    if cfg is None:
                        continue
                    sheet_name = "JAKK-TPK" if key.startswith("JAKK-TPK_") else key
                    results[key] = {
                        "sheet_name": sheet_name,
                        "route_config": cfg,
                        "stations_order": [(code, i + 1) for i, (code, _) in enumerate(station_cols)],
                        "trains": [],
                    }
                results[key]["trains"].extend(trains)
    finally:
        doc.close()
    return results


# ===========================================================================
# KCI feed generation
# ===========================================================================
CALENDAR_MAP = {
    "AllDay": {"monday": 1, "tuesday": 1, "wednesday": 1, "thursday": 1, "friday": 1, "saturday": 1, "sunday": 1},
    "Weekday": {"monday": 1, "tuesday": 1, "wednesday": 1, "thursday": 1, "friday": 1, "saturday": 0, "sunday": 0},
    "Weekend": {"monday": 0, "tuesday": 0, "wednesday": 0, "thursday": 0, "friday": 0, "saturday": 1, "sunday": 1},
    "Stop": {d: 0 for d in WEEKDAYS},
}


def generate_agency():
    a = AGENCY_INFO
    return "\n".join([
        AGENCY[1],
        f'{a["agency_id"]},{a["agency_name"]},{a["agency_url"]},'
        f'{a["agency_timezone"]},{a["agency_phone"]},{a["agency_lang"]}',
    ])


def generate_calendar(service_ids=None):
    if service_ids is None:
        service_ids = ["AllDay"]
    lines = [CALENDAR[1]]
    for sid in service_ids:
        cal = CALENDAR_MAP.get(sid, CALENDAR_MAP["AllDay"])
        lines.append(
            f'{FEED_END_DATE},{cal["friday"]},{cal["monday"]},{cal["saturday"]},{sid},'
            f'{FEED_START_DATE},{cal["sunday"]},{cal["thursday"]},{cal["tuesday"]},{cal["wednesday"]}'
        )
    return "\n".join(lines)


def generate_routes():
    lines = [ROUTES[1]]
    for r in ROUTE_LIST:
        lines.append(
            f'{r["agency_id"]},{r["route_desc"]},{r["route_id"]},'
            f'{r["route_long_name"]},{r["route_short_name"]},{r["route_type"]}'
        )
    return "\n".join(lines)


def generate_stops(used_stations=None):
    """stops.txt, honouring the is_active flag and any unused-station filter."""
    lines = [STOPS[1]]
    for code, s in sorted(STATIONS.items(), key=lambda x: x[1]["stop_code"]):
        if not is_station_active(code):
            continue
        if used_stations and code not in used_stations:
            continue
        lines.append(
            f'{s["level_id"]},1,,{s["stop_code"]},,{s["stop_id"]},'
            f'{s["stop_lat"]},{s["stop_lon"]},{s["stop_name"]},,{s["zone_id"]}'
        )
    return "\n".join(lines)


def generate_trips_and_stop_times(parsed_data, filter_services=None):
    """Build trips.txt / stop_times.txt, skipping trips left with <2 stops."""
    trips_lines = [TRIPS[1]]
    stop_times_lines = [STOP_TIMES[1]]

    used_stations = set()
    trip_counter = 0
    used_trip_ids = set()

    for sheet_data in parsed_data.values():
        route_id = sheet_data["route_config"]["route_id"]
        for train in sheet_data["trains"]:
            service_id = train.get("service_id", "AllDay")
            if filter_services and service_id not in filter_services:
                continue

            stops = normalize_trip_times(train["stops"])
            real_stops = [s for s in stops if not s.get("express") and is_station_active(s["station"])]
            if len(real_stops) < 2:
                continue

            base_trip_id = sanitize_trip_id(train["no_ka"])
            trip_id = base_trip_id
            n = 0
            while trip_id in used_trip_ids or not trip_id:
                n += 1
                trip_id = f"{base_trip_id}{chr(ord('A') + n - 1)}" if base_trip_id else f"trip{n}"
            used_trip_ids.add(trip_id)
            trip_counter += 1

            trips_lines.append(f'{route_id},{service_id},{trip_id},{train["relasi"]},{train["relasi"]},,')

            seq = 1
            for stop in stops:
                if stop.get("express") or not is_station_active(stop["station"]):
                    continue
                used_stations.add(stop["station"])

                arr_min = stop.get("arrival_min")
                dep_min = stop.get("departure_min")
                if dep_min is None:
                    dep_min = arr_min
                if arr_min is None:
                    arr_min = dep_min
                arrival = minutes_to_gtfs_str(arr_min) if arr_min is not None else ""
                departure = minutes_to_gtfs_str(dep_min) if dep_min is not None else ""

                stop_times_lines.append(f'{trip_id},,{departure},{stop["station"]},{arrival},{seq},,')
                seq += 1

    return "\n".join(trips_lines), "\n".join(stop_times_lines), used_stations, trip_counter


def generate_gtfs_zip(parsed_data, filter_services=None):
    """Build the whole KCI feed in memory. Returns (BytesIO, trip_count)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("agency.txt", generate_agency())

        all_services = {
            train.get("service_id", "AllDay")
            for sheet_data in parsed_data.values()
            for train in sheet_data.get("trains", [])
        }
        zf.writestr("calendar.txt", generate_calendar(sorted(all_services)))
        zf.writestr("routes.txt", generate_routes())

        trips_csv, stop_times_csv, used_stations, trip_count = generate_trips_and_stop_times(
            parsed_data, filter_services
        )
        zf.writestr("stops.txt", generate_stops(used_stations))
        zf.writestr("trips.txt", trips_csv)
        zf.writestr("stop_times.txt", stop_times_csv)
        zf.writestr("levels.txt", LEVELS_BODY)

    buf.seek(0)
    return buf, trip_count


def get_summary(parsed_data):
    summary = {"total_routes": 0, "total_trains": 0, "total_stops_entries": 0,
               "routes": [], "stations_used": set()}
    for sheet_data in parsed_data.values():
        trains = sheet_data["trains"]
        route_stations = set()
        for t in trains:
            for s in t["stops"]:
                if is_station_active(s["station"]):
                    route_stations.add(s["station"])
                    summary["stations_used"].add(s["station"])
        summary["routes"].append({
            "route_id": sheet_data["route_config"]["route_id"],
            "route_long_name": sheet_data["route_config"]["route_long_name"],
            "sheet": sheet_data["sheet_name"],
            "train_count": len(trains),
            "stations": sorted(route_stations),
        })
        summary["total_trains"] += len(trains)
        summary["total_stops_entries"] += sum(len(t["stops"]) for t in trains)

    summary["total_routes"] = len(summary["routes"])
    summary["stations_used"] = sorted(summary["stations_used"])
    return summary


# ===========================================================================
# Validation
# ===========================================================================
def collect_excel_codes(parsed_data):
    return {code for sheet_data in parsed_data.values() for code, _ in sheet_data["stations_order"]}


def validate_stop_ids(parsed_data):
    """Cross-check timetable station labels against the station database."""
    excel_codes = collect_excel_codes(parsed_data)
    db_codes = set(STATIONS.keys())

    field_mismatch = []
    for code in sorted(excel_codes & db_codes):
        s = STATIONS[code]
        if s["stop_id"] != code:
            field_mismatch.append({"code_excel": code, "field": "stop_id",
                                   "nilai_di_database": s["stop_id"], "harusnya": code})
        if s["zone_id"] != code:
            field_mismatch.append({"code_excel": code, "field": "zone_id",
                                   "nilai_di_database": s["zone_id"], "harusnya": code})

    return {
        "excel_count": len(excel_codes),
        "db_count": len(db_codes),
        "missing_in_db": sorted(excel_codes - db_codes),
        "not_in_excel": sorted(db_codes - excel_codes),
        "field_mismatch": field_mismatch,
    }


def _zip_column(zip_bytes, name, field):
    """Read one field out of a GTFS member as a set of non-empty values."""
    with gtfs_edit.open_zip(zip_bytes) as zf:
        if name not in zf.namelist():
            return set()
        raw = zf.read(name).decode("utf-8-sig")
    return {
        (row.get(field) or "").strip()
        for row in csv.DictReader(io.StringIO(raw))
        if (row.get(field) or "").strip()
    }


def validate_gtfs_zip(zip_buffer):
    """Report dangling stop_times references and stops nothing stops at."""
    data = zip_buffer.getvalue() if hasattr(zip_buffer, "getvalue") else zip_buffer
    stops_ids = _zip_column(data, "stops.txt", "stop_id")
    stop_time_refs = _zip_column(data, "stop_times.txt", "stop_id")
    return {
        "orphan_refs": sorted(stop_time_refs - stops_ids),
        "unused_stops": sorted(stops_ids - stop_time_refs),
        "stops": len(stops_ids),
        "stop_ids": stops_ids,
    }



# ===========================================================================
# Demo feed (no input file) -- keeps the tool runnable with no dependencies
# ===========================================================================
DEMO_AGENCY = ("agency.txt", ["agency_id", "agency_name", "agency_url", "agency_timezone"])
DEMO_STOPS = ("stops.txt", ["stop_id", "stop_name", "stop_lat", "stop_lon", "zone_id", "location_type"])
DEMO_ROUTES = ("routes.txt", ["route_id", "agency_id", "route_short_name", "route_long_name", "route_type"])
DEMO_TRIPS = ("trips.txt", ["route_id", "service_id", "trip_id", "trip_headsign", "direction_id"])
DEMO_STOP_TIMES = ("stop_times.txt", ["trip_id", "arrival_time", "departure_time", "stop_id", "stop_sequence"])
DEMO_CALENDAR = ("calendar.txt", ["service_id"] + WEEKDAYS + ["start_date", "end_date"])
DEMO_FEED_INFO = ("feed_info.txt", ["feed_publisher_name", "feed_publisher_url", "feed_lang",
                                    "feed_start_date", "feed_end_date"])

DEMO_FILES = [DEMO_AGENCY, DEMO_STOPS, DEMO_ROUTES, DEMO_TRIPS,
              DEMO_STOP_TIMES, DEMO_CALENDAR, DEMO_FEED_INFO]

DEMO_BASE_LAT = -6.2088
DEMO_BASE_LON = 106.8456
DEMO_SERVICE_DAYS = ("monday", "tuesday", "wednesday", "thursday", "friday")


def next_monday(today: date) -> date:
    return today + timedelta(days=(7 - today.weekday()) % 7 or 7)


def build_start_date(cfg: Config) -> date:
    if cfg.start_date:
        return date.fromisoformat(cfg.start_date)
    return next_monday(date.today())


def generate_demo(cfg: Config) -> dict:
    """Synthetic feed tables, keyed by filename."""
    import random

    rng = random.Random(cfg.seed)
    start = build_start_date(cfg)
    sdate = start.strftime("%Y%m%d")
    edate = (start + timedelta(days=29)).strftime("%Y%m%d")

    days = {d: "0" for d in WEEKDAYS}
    for d in DEMO_SERVICE_DAYS:
        days[d] = "1"

    stop_rows, route_rows, trip_rows, stop_time_rows = [], [], [], []

    for r in range(cfg.demo_routes):
        route_id = f"R{r + 1}"
        route_rows.append({
            "route_id": route_id,
            "agency_id": "DEMO",
            "route_short_name": route_id,
            "route_long_name": f"Demo Route {r + 1}",
            "route_type": 3,
        })
        lat = DEMO_BASE_LAT + r * 0.004
        lon = DEMO_BASE_LON + r * 0.003
        for s in range(cfg.demo_stops_per_route):
            stop_rows.append({
                "stop_id": f"{route_id}_ST_{s + 1}",
                "stop_name": f"Stop {r + 1}.{s + 1}",
                "stop_lat": round(lat + s * 0.0011, 6),
                "stop_lon": round(lon + s * 0.0013, 6),
                "zone_id": f"Z{r + 1}",
                "location_type": 0,
            })

    headway = max(300, 900 // max(cfg.demo_trips_per_day, 1))
    trip_seq = 0
    for r in range(cfg.demo_routes):
        route_id = f"R{r + 1}"
        for t in range(cfg.demo_trips_per_day):
            trip_seq += 1
            trip_id = f"{route_id}_T{trip_seq:03d}"
            trip_rows.append({
                "route_id": route_id,
                "service_id": "DEMO_WD",
                "trip_id": trip_id,
                "trip_headsign": f"Route {r + 1} - Trip {t + 1}",
                "direction_id": t % 2,
            })
            depart = 6 * 3600 + t * headway
            travel = rng.randint(120, 240)
            for s in range(cfg.demo_stops_per_route):
                stop_time_rows.append({
                    "trip_id": trip_id,
                    "arrival_time": fmt_hhmmss(depart),
                    "departure_time": fmt_hhmmss(depart + 30),
                    "stop_id": f"{route_id}_ST_{s + 1}",
                    "stop_sequence": s + 1,
                })
                depart += 30 + travel

    return {
        "agency.txt": [{
            "agency_id": "DEMO",
            "agency_name": "Demo Transit Authority",
            "agency_url": "https://example.com/transit",
            "agency_timezone": "Asia/Jakarta",
        }],
        "stops.txt": stop_rows,
        "routes.txt": route_rows,
        "trips.txt": trip_rows,
        "stop_times.txt": stop_time_rows,
        "calendar.txt": [{"service_id": "DEMO_WD", **days, "start_date": sdate, "end_date": edate}],
        "feed_info.txt": [{
            "feed_publisher_name": "Demo Transit Authority",
            "feed_publisher_url": "https://example.com/transit",
            "feed_lang": "id",
            "feed_start_date": sdate,
            "feed_end_date": edate,
        }],
    }


def _write_table(path: Path, header: list, rows: list) -> None:
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=header, lineterminator="\r\n")
        writer.writeheader()
        writer.writerows(rows)


def export_demo(tables: dict, out_dir: Path, as_zip: bool) -> Path | None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for filename, header in DEMO_FILES:
        _write_table(out_dir / filename, header, tables.get(filename, []))

    if not as_zip:
        return None
    zip_path = out_dir.parent / f"{out_dir.name}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, _header in DEMO_FILES:
            zf.write(out_dir / filename, filename)
    return zip_path


# ===========================================================================
# Output helpers
# ===========================================================================
def write_kci_feed(buf: io.BytesIO, out: str, as_zip: bool) -> Path:
    out_path = Path(out)
    if as_zip or out_path.suffix.lower() != ".zip":
        if not as_zip:
            out_path = out_path.with_suffix(".zip")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(buf.getvalue())
        return out_path
    out_dir = out_path
    out_dir.mkdir(parents=True, exist_ok=True)
    with gtfs_edit.open_zip(buf) as zf:
        for name in zf.namelist():
            (out_dir / name).write_bytes(zf.read(name))
    return out_dir


def summarize(cfg: Config, tables: dict) -> None:
    if not tables:
        print("Feed       : (kosong)")
        return
    for name in ("routes.txt", "stops.txt", "trips.txt", "stop_times.txt"):
        rows = tables.get(name)
        if rows is not None:
            print(f"{name[:-4].replace('_', '-').title():<11}: {len(rows)}")


def load_config(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


# ===========================================================================
# Subcommands -- stop activity inside an existing feed
# ===========================================================================
def _resolve_stops(args):
    """Resolve --stop codes, or fall back to the persisted inactive list."""
    if args.stop:
        return {resolve_station_code(s) for s in args.stop}
    return gtfs_edit.load_inactive_stops(args.state)


def cmd_disable(args) -> int:
    data = Path(args.feed).read_bytes()
    inactive = _resolve_stops(args) | {resolve_station_code(s) for s in args.stop}
    out_buf = gtfs_edit.rebuild_gtfs_without_inactive(data, inactive)
    if args.backfill:
        master = gtfs_edit.load_gtfs_master(args.master)
        if master:
            out_buf = gtfs_edit.backfill_stop_times_from_master(out_buf, master, inactive)
    written = write_kci_feed(out_buf, args.out, args.zip)
    gtfs_edit.save_inactive_stops(args.state, inactive)
    print(f"Nonaktif   : {len(inactive)} stop -> {sorted(inactive)}")
    print(f"Output     : {written.resolve()}")
    print(f"State      : {args.state}")
    return 0


def cmd_activate(args) -> int:
    data = Path(args.feed).read_bytes()
    inactive = _resolve_stops(args)
    master = gtfs_edit.load_gtfs_master(args.master)
    if master:
        out_buf = gtfs_edit.backfill_stop_times_from_master(data, master, inactive)
    else:
        print(f"Catatan    : tidak ada master di {args.master}, stop_times tidak di-backfill")
        out_buf = gtfs_edit.open_zip(data)
    out_buf = gtfs_edit.rebuild_gtfs_without_inactive(out_buf, set())
    written = write_kci_feed(out_buf, args.out, args.zip)
    remaining = gtfs_edit.load_inactive_stops(args.state) - inactive
    gtfs_edit.save_inactive_stops(args.state, remaining)
    print(f"Dihidupkan : {sorted(inactive)}")
    print(f"Sisa nonaktif: {sorted(remaining) or '()'}")
    print(f"Output     : {written.resolve()}")
    return 0


def cmd_stations(args) -> int:
    inactive = set(get_inactive_stations())
    for code in sorted(STATIONS):
        s = STATIONS[code]
        mark = "off" if code in inactive else "on "
        if args.active_only and code in inactive:
            continue
        print(f"{mark}  {code:<5} {s['stop_name']:<42} {s['stop_lat']},{s['stop_lon']}")
    print(f"\nTotal {len(STATIONS)} stasiun "
          f"({len(get_active_stations())} aktif, {len(inactive)} nonaktif)")
    return 0


def build_feed_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="generate_gtfs.py",
        description="Generate a GTFS feed for KCI Commuter (Jabodetabek), or a synthetic demo feed.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    src = parser.add_argument_group("timetable source (default: synthetic demo feed)")
    src.add_argument("--excel", default="", help="GAPEKA KCI workbook (.xlsx)")
    src.add_argument("--pdf", default="", help="local KCI schedule PDF")
    src.add_argument("--pdf-url", default="", help="URL of the KCI schedule PDF; default: the current Jabodetabek timetable")
    src.add_argument("--service", action="append", default=[],
                     help="only keep these service_ids (AllDay, Weekday, Weekend); repeatable")

    out = parser.add_argument_group("output")
    out.add_argument("--out", default="output", help="output folder or .zip path")
    out.add_argument("--zip", action="store_true", help="write a .zip of the feed")

    demo = parser.add_argument_group("demo feed (used when no source is given)")
    demo.add_argument("--routes", type=int, default=3, help="demo: number of routes")
    demo.add_argument("--stops-per-route", type=int, default=6, help="demo: stops per route")
    demo.add_argument("--trips-per-day", type=int, default=5, help="demo: departures per day per route")
    demo.add_argument("--seed", type=int, default=42, help="demo: random seed")
    demo.add_argument("--start-date", default="", help="demo: feed start date (YYYY-MM-DD)")
    demo.add_argument("--config", default="", help="optional JSON config file")
    return parser


def build_stop_parser(name: str, help_text: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog=f"generate_gtfs.py {name}", description=help_text,
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("--feed", required=True, help="input GTFS .zip")
    parser.add_argument("--stop", action="append", default=[],
                        help="stop_id to toggle; repeatable. Omit to use the persisted list")
    parser.add_argument("--out", default="output.zip", help="output .zip path")
    parser.add_argument("--zip", action="store_true", help="write a .zip of the feed")
    parser.add_argument("--state", default=gtfs_edit.INACTIVE_STOPS_DEFAULT,
                        help="JSON file holding the inactive stop list")
    parser.add_argument("--master", default=gtfs_edit.GTFS_MASTER_DEFAULT,
                        help="master feed used to restore stop_times on reactivation")
    if name == "disable-stops":
        parser.add_argument("--backfill", action="store_true",
                            help="also restore stop_times from master for the stops being disabled")
    return parser


def cmd_show_stops(args) -> int:
    data = Path(args.feed).read_bytes()
    stops = gtfs_edit.read_gtfs_stops(data)
    for s in stops:
        print(f"{s['stop_id']:<8} {s['stop_code']:<8} {s['stop_name']:<42} {s['stop_lat']},{s['stop_lon']}")
    print(f"\nTotal {len(stops)} stop di {args.feed}")
    return 0


def cmd_detect_orphans(args) -> int:
    data = Path(args.feed).read_bytes()
    orphans = gtfs_edit.find_unreferenced_stops(data)
    for stop_id in sorted(orphans):
        name = STATIONS.get(stop_id, {}).get("stop_name", "-")
        print(f"  {stop_id:<8} {name}")
    print(f"\nOrphan    : {len(orphans)} {sorted(orphans) if orphans else ''}")
    return 0


def cmd_classify(args) -> int:
    data = Path(args.feed).read_bytes()
    expected = set(STATIONS)
    result = gtfs_edit.classify_upload(data, expected)
    print(f"Feed penuh : {result['looks_full']}")
    print(f"Orphan     : {len(result['orphan_ids'])} {sorted(result['orphan_ids']) or ''}")
    missing = result["missing_stops"]
    print(f"Hilang     : {len(missing)} {missing or ''}")
    if not result["looks_full"]:
        print("\nPerhatian  : feed ini tidak penuh, jangan dipakai sebagai master.")
    return 0


def build_feed_inspect_parser(name: str, help_text: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog=f"generate_gtfs.py {name}", description=help_text)
    parser.add_argument("--feed", required=True, help="input GTFS .zip")
    return parser


def build_stations_parser(name: str, help_text: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="generate_gtfs.py stations", description=help_text)
    parser.add_argument("--active-only", action="store_true", help="hide inactive stations")
    return parser


SUBCOMMANDS = {
    "disable-stops": ("Disable stops in a feed and rebuild it", build_stop_parser, cmd_disable),
    "activate-stops": ("Re-enable stops, restoring stop_times from the master feed", build_stop_parser, cmd_activate),
    "show-stops": ("List the stops present in a feed", build_feed_inspect_parser, cmd_show_stops),
    "detect-orphans": ("List stops in stops.txt that stop_times.txt never uses",
                       build_feed_inspect_parser, cmd_detect_orphans),
    "classify": ("Check whether a feed is a full feed, before trusting it as master",
                 build_feed_inspect_parser, cmd_classify),
    "stations": ("List the station database", build_stations_parser, cmd_stations),
}


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    if argv and argv[0] in SUBCOMMANDS:
        help_text, builder, handler = SUBCOMMANDS[argv[0]]
        return handler(builder(argv[0], help_text).parse_args(argv[1:]))

    parser = build_feed_parser()
    args = parser.parse_args(argv)
    cfg = Config(
        excel=args.excel, pdf=args.pdf, pdf_url=args.pdf_url, out=args.out,
        as_zip=args.zip, filter_services=tuple(args.service), demo_routes=args.routes,
        demo_stops_per_route=args.stops_per_route, demo_trips_per_day=args.trips_per_day,
        seed=args.seed, start_date=args.start_date,
    )
    if args.config:
        cfg.__dict__.update(load_config(args.config))

    sources = [bool(cfg.excel), bool(cfg.pdf), bool(cfg.pdf_url)]
    if sum(sources) > 1:
        parser.error("pilih salah satu sumber: --excel, --pdf, atau --pdf-url")

    if not any(sources):
        tables = generate_demo(cfg)
        export_demo(tables, Path(cfg.out), cfg.as_zip)
        print("Mode       : demo feed (synthetic)")
        print(f"Agency     : {tables['agency.txt'][0]['agency_name']}")
        summarize(cfg, tables)
        print(f"Output     : {Path(cfg.out).resolve()}")
        if cfg.as_zip:
            print(f"ZIP        : {(Path(cfg.out).parent / (Path(cfg.out).name + '.zip')).resolve()}")
        return 0

    if cfg.excel:
        print(f"Reading    : {cfg.excel}")
        parsed = parse_excel(cfg.excel)
    elif cfg.pdf:
        print(f"Reading    : {cfg.pdf}")
        parsed = parse_pdf(Path(cfg.pdf).read_bytes())
    else:
        url = cfg.pdf_url or PDF_DEFAULT_URL
        print(f"Downloading: {url}")
        parsed = parse_pdf(download_pdf(url))

    if not parsed:
        print("Error      : tidak ada tabel jadwal yang bisa dibaca dari sumber tersebut.")
        return 1

    validation = validate_stop_ids(parsed)
    total_issues = len(validation["missing_in_db"]) + len(validation["field_mismatch"])
    if total_issues:
        print(f"Validasi   : {total_issues} masalah kecocokan stop id "
              f"({validation['excel_count']} kode di timetable, {validation['db_count']} di database)")
    else:
        print(f"Validasi   : {validation['excel_count']} kode di timetable cocok dengan database")

    buf, trip_count = generate_gtfs_zip(parsed, cfg.filter_services or None)
    written = write_kci_feed(buf, cfg.out, cfg.as_zip)

    summary = get_summary(parsed)
    checks = validate_gtfs_zip(buf)
    print(f"Routes     : {summary['total_routes']}")
    print(f"Trains     : {trip_count}")
    # Count from the feed, not from the timetable: a station that is only ever
    # passed express has no scheduled stop and is never written to stops.txt.
    print(f"Stops      : {checks['stops']} dipakai dari {len(STATIONS)}"
          f" ({', '.join(sorted(set(STATIONS) - checks['stop_ids'])) or 'semua dipakai'})")
    print(f"Orphan     : {checks['orphan_refs'] or '[]'}")
    print(f"Output     : {written.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

