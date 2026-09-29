"""Tests for generate_gtfs.py and gtfs_edit.py.

Self-contained: every fixture is built in memory, so the suite runs without a
timetable file or a network connection.
"""

import csv
import io
import os
import sys
import tempfile
import unittest
import zipfile
from datetime import time

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

import generate_gtfs as gg
import gtfs_edit


# ===========================================================================
# Fixtures
# ===========================================================================
STOPS_HDR = "level_id,location_type,parent_station,stop_code,stop_desc,stop_id,stop_lat,stop_lon,stop_name,wheelchair_boarding,zone_id"
TRIPS_HDR = "route_id,service_id,trip_id,trip_short_name,trip_headsign,direction_id,shape_id"
ST_HDR = "trip_id,pickup_type,departure_time,stop_id,arrival_time,stop_sequence,drop_off_type,departure_time_fixed"
AGENCY_HDR = "agency_id,agency_name,agency_url,agency_timezone,agency_phone,agency_lang"
ROUTES_HDR = "agency_id,route_desc,route_id,route_long_name,route_short_name,route_type"
CAL_HDR = "end_date,friday,monday,saturday,service_id,start_date,sunday,thursday,tuesday,wednesday"


def make_fixture_feed():
    """A 3-stop, 2-trip, 1-route feed used across the editing tests."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("agency.txt", AGENCY_HDR + "\n1,KCI,https://example.com,Asia/Jakarta,,id")
        zf.writestr("levels.txt", "level_id,level_index,level_name\nL0,0,Street")
        zf.writestr("routes.txt", ROUTES_HDR + "\n1,route one,1001,One,BO,0")
        zf.writestr("calendar.txt", CAL_HDR + "\n20261231,1,1,1,AllDay,20210816,1,1,1,1")
        zf.writestr("stops.txt", "\n".join([
            STOPS_HDR,
            "L0,1,,1,,THB,-6.185403437,106.8109823,STASIUN KCI TANAHABANG,,THB",
            "L0,1,,2,,SUD,-6.202308393,106.823334,STASIUN KCI SUDIRMAN,,SUD",
            "L0,1,,3,,PSE,-6.174532997,106.844259,STASIUN KCI PASARSENEN,,PSE",
        ]))
        zf.writestr("trips.txt", "\n".join([
            TRIPS_HDR,
            "1001,AllDay,T1,T1,One,,",
            "1001,AllDay,T2,T2,One,,",
        ]))
        st_lines = [ST_HDR]
        for trip in ("T1", "T2"):
            for seq, stop in enumerate(("THB", "SUD", "PSE"), start=1):
                hhmm = f"08:{seq * 2:02d}:00"
                st_lines.append(f"{trip},,{hhmm},{stop},{hhmm},{seq},,")
        zf.writestr("stop_times.txt", "\n".join(st_lines))
    buf.seek(0)
    return buf.getvalue()


def make_parsed_data():
    """Minimal parsed_data: one route, one trip over three stations."""
    def t(h, m):
        return time(h, m)

    return {
        "SHEET": {
            "sheet_name": "SHEET",
            "route_config": {"route_id": "1001", "route_long_name": "Test route"},
            "stations_order": [("THB", 1), ("SUD", 2), ("PSE", 3)],
            "trains": [
                {"no_ka": "1234", "relasi": "THB-PSE", "service_id": "AllDay", "stops": [
                    {"station": "THB", "arrival": t(6, 0), "departure": t(6, 0), "express": False},
                    {"station": "SUD", "arrival": t(6, 30), "departure": t(6, 31), "express": False},
                    {"station": "PSE", "arrival": t(7, 0), "departure": t(7, 0), "express": False},
                ]},
                {"no_ka": "1235", "relasi": "THB-PSE", "service_id": "AllDay", "stops": [
                    {"station": "THB", "arrival": t(8, 0), "departure": t(8, 0), "express": False},
                    {"station": "SUD", "arrival": t(8, 30), "departure": t(8, 31), "express": False},
                    {"station": "PSE", "arrival": t(9, 0), "departure": t(9, 0), "express": False},
                ]},
            ],
        }
    }


def rows_of(zip_bytes, name):
    with gtfs_edit.open_zip(zip_bytes) as zf:
        raw = zf.read(name).decode("utf-8-sig")
    return list(csv.DictReader(io.StringIO(raw)))


# ===========================================================================
# Station database
# ===========================================================================
class TestStationData(unittest.TestCase):
    def test_station_count(self):
        self.assertEqual(len(gg.STATIONS), 85)

    def test_every_station_is_well_formed(self):
        for code, s in gg.STATIONS.items():
            self.assertEqual(s["stop_id"], code, code)
            self.assertEqual(s["zone_id"], code, code)
            self.assertTrue(s["stop_name"].startswith("STASIUN KCI "), code)
            self.assertTrue(s["stop_code"], code)
            float(s["stop_lat"])
            float(s["stop_lon"])

    def test_stop_codes_are_unique(self):
        codes = [s["stop_code"] for s in gg.STATIONS.values()]
        self.assertEqual(len(codes), len(set(codes)))

    def test_jakarta_kota_uses_jakk(self):
        # "JAK" collides with other KCI products, so the code must be JAKK.
        self.assertIn("JAKK", gg.STATIONS)
        self.assertNotIn("JAK", gg.STATIONS)
        self.assertEqual(gg.resolve_station_code("JAKARTAKOTA"), "JAKK")

    def test_aliases_resolve(self):
        for alias, expected in [
            ("TANAH ABANG", "THB"), ("PASAR SENEN", "PSE"),
            ("DUREN KALIBATA", "DRN"), ("TANJUNG PRIOK", "TPK"),
        ]:
            self.assertEqual(gg.resolve_station_code(alias), expected, alias)

    def test_all_alias_targets_exist(self):
        for alias, target in gg.STATION_ALIASES.items():
            self.assertIn(target, gg.STATIONS, f"{alias} -> {target}")


class TestStationActivity(unittest.TestCase):
    def tearDown(self):
        gg.reset_station_activity()

    def test_toggle_and_reset(self):
        self.assertTrue(gg.is_station_active("THB"))
        gg.set_station_active("THB", False)
        self.assertFalse(gg.is_station_active("THB"))
        self.assertEqual(gg.get_inactive_stations(), ["THB"])
        self.assertNotIn("THB", gg.get_active_stations())
        gg.set_station_active("THB", True)
        self.assertTrue(gg.is_station_active("THB"))

    def test_add_and_update_station(self):
        self.addCleanup(gg.STATIONS.pop, "TESTX", None)
        self.assertTrue(gg.add_station("TESTX", stop_code="999", stop_name="STASIUN KCI TES"))
        self.assertFalse(gg.add_station("TESTX"))
        self.assertTrue(gg.set_station_info("TESTX", stop_code="998"))
        self.assertEqual(gg.STATIONS["TESTX"]["stop_code"], "998")


# ===========================================================================
# Feed generation
# ===========================================================================
class TestGenerate(unittest.TestCase):
    def setUp(self):
        gg.reset_station_activity()
        self.parsed = make_parsed_data()

    def tearDown(self):
        gg.reset_station_activity()

    def test_writes_the_seven_reference_files(self):
        buf, trips = gg.generate_gtfs_zip(self.parsed)
        with gtfs_edit.open_zip(buf) as zf:
            self.assertEqual(sorted(zf.namelist()), [
                "agency.txt", "calendar.txt", "levels.txt", "routes.txt",
                "stop_times.txt", "stops.txt", "trips.txt",
            ])
        self.assertEqual(trips, 2)

    def test_trip_ids_are_sanitised_and_unique(self):
        self.parsed["SHEET"]["trains"][0]["no_ka"] = "12 34/A"
        self.parsed["SHEET"]["trains"][1]["no_ka"] = "12 34/A"
        buf, trips = gg.generate_gtfs_zip(self.parsed)
        ids = [r["trip_id"] for r in rows_of(buf.getvalue(), "trips.txt")]
        self.assertEqual(trips, 2)
        self.assertEqual(len(ids), len(set(ids)))

    def test_times_are_monotonic(self):
        buf, _ = gg.generate_gtfs_zip(self.parsed)
        by_trip = {}
        for r in rows_of(buf.getvalue(), "stop_times.txt"):
            by_trip.setdefault(r["trip_id"], []).append(r)
        for tid, rws in by_trip.items():
            rws.sort(key=lambda x: int(x["stop_sequence"]))
            mins = [_mins(r["arrival_time"]) for r in rws]
            self.assertEqual(mins, sorted(mins), tid)

    def test_active_flag_removes_station_from_feed(self):
        gg.set_station_active("SUD", False)
        buf, _ = gg.generate_gtfs_zip(self.parsed)
        data = buf.getvalue()
        st_ids = {r["stop_id"] for r in rows_of(data, "stop_times.txt")}
        self.assertNotIn("SUD", st_ids)
        self.assertNotIn("SUD", {r["stop_id"] for r in rows_of(data, "stops.txt")})
        self.assertEqual(gg.validate_gtfs_zip(buf)["orphan_refs"], [])

    def test_trip_with_fewer_than_two_stops_is_dropped(self):
        gg.set_station_active("SUD", False)
        gg.set_station_active("PSE", False)
        buf, trips = gg.generate_gtfs_zip(self.parsed)
        self.assertEqual(trips, 0)
        self.assertEqual(rows_of(buf.getvalue(), "trips.txt"), [])

    def test_calendar_only_lists_services_in_use(self):
        self.parsed["SHEET"]["trains"][0]["service_id"] = "Weekday"
        buf, _ = gg.generate_gtfs_zip(self.parsed)
        sids = {r["service_id"] for r in rows_of(buf.getvalue(), "calendar.txt")}
        self.assertEqual(sids, {"AllDay", "Weekday"})

    def test_filter_services(self):
        self.parsed["SHEET"]["trains"][0]["service_id"] = "Weekday"
        _, trips = gg.generate_gtfs_zip(self.parsed, filter_services=("Weekday",))
        self.assertEqual(trips, 1)

    def test_stops_headers_match_reference_order(self):
        self.assertEqual(gg.STOPS[1].split(",")[5], "stop_id")
        self.assertEqual(gg.STOP_TIMES[1].split(",")[3], "stop_id")


class TestValidateStopIds(unittest.TestCase):
    def test_reports_unknown_station(self):
        parsed = make_parsed_data()
        parsed["SHEET"]["stations_order"].append(("ZZZ", 4))
        out = gg.validate_stop_ids(parsed)
        self.assertEqual(out["missing_in_db"], ["ZZZ"])

    def test_reports_mismatch(self):
        parsed = make_parsed_data()
        parsed["SHEET"]["stations_order"] = [("THB", 1), ("SUD", 2)]
        out = gg.validate_stop_ids(parsed)
        self.assertEqual(out["missing_in_db"], [])
        self.assertEqual(out["field_mismatch"], [])
        self.assertEqual(out["excel_count"], 2)


class TestTimeHelpers(unittest.TestCase):
    def test_minutes_round_trip(self):
        self.assertEqual(gg.minutes_to_gtfs_str(90), "01:30:00")
        self.assertEqual(gg.minutes_to_gtfs_str(0), "00:00:00")

    def test_midnight_rolls_forward_not_back(self):
        stops = [
            {"station": "A", "arrival": time(23, 50), "departure": time(23, 55), "express": False},
            {"station": "B", "arrival": time(0, 5), "departure": time(0, 6), "express": False},
        ]
        out = gg.normalize_trip_times(stops)
        self.assertEqual(out[1]["arrival_min"] // 60, 24)

    def test_sanitize_trip_id(self):
        self.assertEqual(gg.sanitize_trip_id("12 34"), "1234")
        self.assertEqual(gg.sanitize_trip_id("a/b:c"), "a-b-c")
        self.assertEqual(gg.sanitize_trip_id(""), "")


# ===========================================================================
# Demo feed
# ===========================================================================
class TestDemo(unittest.TestCase):
    def test_demo_feed_files(self):
        cfg = gg.Config(demo_routes=2, demo_stops_per_route=3, demo_trips_per_day=2)
        tables = gg.generate_demo(cfg)
        for name, _header in gg.DEMO_FILES:
            self.assertIn(name, tables)
        self.assertEqual(len(tables["routes.txt"]), 2)
        self.assertEqual(len(tables["stops.txt"]), 6)
        self.assertEqual(len(tables["trips.txt"]), 4)

    def test_demo_is_reproducible(self):
        cfg = gg.Config()
        self.assertEqual(gg.generate_demo(cfg), gg.generate_demo(gg.Config()))

    def test_export_zip(self):
        with tempfile.TemporaryDirectory() as td:
            cfg = gg.Config(out=os.path.join(td, "output"), as_zip=True)
            tables = gg.generate_demo(cfg)
            zip_path = gg.export_demo(tables, gg.Path(cfg.out), True)
            self.assertTrue(zip_path.exists())
            with zipfile.ZipFile(zip_path) as zf:
                self.assertEqual(sorted(zf.namelist()), sorted(n for n, _ in gg.DEMO_FILES))


def _mins(hhmmss):
    h, m, s = hhmmss.split(":")
    return int(h) * 60 + int(m) + int(s) / 60


# ===========================================================================
# gtfs_edit
# ===========================================================================
class TestRead(unittest.TestCase):
    def setUp(self):
        self.src = make_fixture_feed()

    def test_read_stops(self):
        stops = gtfs_edit.read_gtfs_stops(self.src)
        self.assertEqual([s["stop_id"] for s in stops], ["THB", "SUD", "PSE"])

    def test_read_trip_stop_ids(self):
        self.assertEqual(gtfs_edit.read_gtfs_trip_stop_ids(self.src), {"THB", "SUD", "PSE"})

    def test_count_rows(self):
        self.assertEqual(gtfs_edit.count_gtfs_rows(self.src, "stops.txt"), 3)
        self.assertEqual(gtfs_edit.count_gtfs_rows(self.src, "nope.txt"), 0)

    def test_no_orphans_in_full_feed(self):
        self.assertEqual(gtfs_edit.find_unreferenced_stops(self.src), set())


class TestRebuild(unittest.TestCase):
    def setUp(self):
        self.src = make_fixture_feed()

    def test_keeps_stop_row_but_drops_stop_times(self):
        out = gtfs_edit.rebuild_gtfs_without_inactive(self.src, {"SUD"})
        data = out.getvalue()
        self.assertIn("SUD", {s["stop_id"] for s in gtfs_edit.read_gtfs_stops(data)})
        self.assertNotIn("SUD", {r["stop_id"] for r in rows_of(data, "stop_times.txt")})

    def test_drops_trip_left_with_one_stop(self):
        out = gtfs_edit.rebuild_gtfs_without_inactive(self.src, {"SUD", "PSE"})
        data = out.getvalue()
        self.assertEqual(rows_of(data, "trips.txt"), [])
        self.assertEqual(rows_of(data, "stop_times.txt"), [])

    def test_output_has_no_orphans(self):
        out = gtfs_edit.rebuild_gtfs_without_inactive(self.src, {"SUD"})
        data = out.getvalue()
        st = {r["stop_id"] for r in rows_of(data, "stop_times.txt")}
        stops = {r["stop_id"] for r in rows_of(data, "stops.txt")}
        self.assertEqual(st - stops, set())

    def test_orphan_detection_finds_disabled_stops(self):
        out = gtfs_edit.rebuild_gtfs_without_inactive(self.src, {"SUD"})
        self.assertEqual(gtfs_edit.find_unreferenced_stops(out.getvalue()), {"SUD"})

    def test_missing_stops_tolerated(self):
        out = gtfs_edit.rebuild_gtfs_without_inactive(self.src, {"NOPE"})
        self.assertEqual(len(rows_of(out.getvalue(), "stop_times.txt")), 6)


class TestBackfill(unittest.TestCase):
    def test_restores_dropped_stop_times(self):
        src = make_fixture_feed()
        reduced = gtfs_edit.rebuild_gtfs_without_inactive(src, {"SUD"}).getvalue()
        self.assertNotIn("SUD", {r["stop_id"] for r in rows_of(reduced, "stop_times.txt")})

        restored = gtfs_edit.backfill_stop_times_from_master(reduced, src, {"SUD"})
        rows = rows_of(restored.getvalue(), "stop_times.txt")
        self.assertIn("SUD", {r["stop_id"] for r in rows})

        by_trip = {}
        for r in rows:
            by_trip.setdefault(r["trip_id"], []).append(int(r["stop_sequence"]))
        for tid, seqs in by_trip.items():
            self.assertEqual(seqs, sorted(seqs), tid)

    def test_no_duplicates(self):
        src = make_fixture_feed()
        restored = gtfs_edit.backfill_stop_times_from_master(src, src, {"SUD"})
        rows = rows_of(restored.getvalue(), "stop_times.txt")
        pairs = [(r["trip_id"], r["stop_id"]) for r in rows]
        self.assertEqual(len(pairs), len(set(pairs)))

    def test_no_op_without_master(self):
        src = make_fixture_feed()
        out = gtfs_edit.backfill_stop_times_from_master(src, None, {"SUD"})
        self.assertEqual(len(rows_of(out.getvalue(), "stop_times.txt")), 6)


class TestPersistence(unittest.TestCase):
    def test_inactive_roundtrip(self):
        with tempfile.TemporaryDirectory() as td:
            p = os.path.join(td, "inactive.json")
            self.assertEqual(gtfs_edit.load_inactive_stops(p), set())
            self.assertTrue(gtfs_edit.save_inactive_stops(p, {"SUD", " THB "}))
            self.assertEqual(gtfs_edit.load_inactive_stops(p), {"THB", "SUD"})

    def test_broken_state_file_is_ignored(self):
        with tempfile.TemporaryDirectory() as td:
            p = os.path.join(td, "inactive.json")
            with open(p, "w", encoding="utf-8") as f:
                f.write("{not json")
            self.assertEqual(gtfs_edit.load_inactive_stops(p), set())

    def test_master_roundtrip(self):
        src = make_fixture_feed()
        with tempfile.TemporaryDirectory() as td:
            p = os.path.join(td, "master.zip")
            self.assertIsNone(gtfs_edit.load_gtfs_master(p))
            self.assertTrue(gtfs_edit.save_gtfs_master(p, src))
            self.assertEqual(gtfs_edit.load_gtfs_master(p), src)


class TestApplyEdits(unittest.TestCase):
    def test_rename_updates_references(self):
        src = make_fixture_feed()
        out = gtfs_edit.apply_stop_edits(src, edits={"SUD": {"stop_id": "BB2"}})
        data = out.getvalue()
        ids = [s["stop_id"] for s in gtfs_edit.read_gtfs_stops(data)]
        self.assertIn("BB2", ids)
        self.assertNotIn("SUD", ids)
        self.assertIn("BB2", {r["stop_id"] for r in rows_of(data, "stop_times.txt")})

    def test_rename_name_only(self):
        src = make_fixture_feed()
        out = gtfs_edit.apply_stop_edits(src, edits={"SUD": {"stop_name": "STASIUN KCI SUDIRMAN EDIT"}})
        names = {s["stop_id"]: s["stop_name"] for s in gtfs_edit.read_gtfs_stops(out.getvalue())}
        self.assertEqual(names["SUD"], "STASIUN KCI SUDIRMAN EDIT")

    def test_add_stop(self):
        src = make_fixture_feed()
        out = gtfs_edit.apply_stop_edits(
            src, add_rows=[{"stop_id": "ZZZ", "stop_code": "9", "stop_name": "NEW"}]
        )
        stops = gtfs_edit.read_gtfs_stops(out.getvalue())
        self.assertEqual(len(stops), 4)
        self.assertEqual([s["stop_name"] for s in stops if s["stop_id"] == "ZZZ"], ["NEW"])


class TestClassifyUpload(unittest.TestCase):
    def test_full_feed(self):
        src = make_fixture_feed()
        info = gtfs_edit.classify_upload(src, {"THB", "SUD", "PSE"})
        self.assertTrue(info["looks_full"])
        self.assertEqual(info["missing_stops"], [])
        self.assertEqual(info["orphan_ids"], set())

    def test_missing_station_detected(self):
        src = make_fixture_feed()
        info = gtfs_edit.classify_upload(src, {"THB", "SUD", "PSE", "DDD"})
        self.assertEqual(info["missing_stops"], ["DDD"])


class TestCliStopSubcommands(unittest.TestCase):
    """The stop-management subcommands, driven through main()."""

    def _run(self, argv):
        buf = io.StringIO()
        old, old_err = sys.stdout, sys.stderr
        sys.stdout, sys.stderr = buf, io.StringIO()
        try:
            code = gg.main(argv)
        finally:
            sys.stdout, sys.stderr = old, old_err
        return code, buf.getvalue()

    def _paths(self, td):
        feed = os.path.join(td, "original.zip")
        with open(feed, "wb") as f:
            f.write(make_fixture_feed())
        return feed, os.path.join(td, "state.json"), os.path.join(td, "master.zip")

    def test_show_stops_lists_every_stop(self):
        with tempfile.TemporaryDirectory() as td:
            feed, _, _ = self._paths(td)
            code, out = self._run(["show-stops", "--feed", feed])
            self.assertEqual(code, 0)
            for stop in ("THB", "SUD", "PSE"):
                self.assertIn(stop, out)
            self.assertIn("Total 3 stop", out)

    def test_detect_orphans_on_clean_feed(self):
        with tempfile.TemporaryDirectory() as td:
            feed, _, _ = self._paths(td)
            _, out = self._run(["detect-orphans", "--feed", feed])
            self.assertIn("Orphan    : 0", out)

    def test_classify_flags_a_partial_feed(self):
        # The fixture holds 3 stops against the 85 in the database, so the
        # feed must be rejected before it can overwrite the master.
        with tempfile.TemporaryDirectory() as td:
            feed, _, _ = self._paths(td)
            _, out = self._run(["classify", "--feed", feed])
            self.assertIn("Feed penuh : False", out)
            self.assertIn("jangan dipakai sebagai master", out)

    def test_disable_then_detect_orphans(self):
        with tempfile.TemporaryDirectory() as td:
            feed, state, _ = self._paths(td)
            out_zip = os.path.join(td, "disabled.zip")
            code, _ = self._run([
                "disable-stops", "--feed", feed, "--stop", "SUD",
                "--out", out_zip, "--zip", "--state", state,
            ])
            self.assertEqual(code, 0)
            _, out = self._run(["detect-orphans", "--feed", out_zip])
            self.assertIn("SUD", out)
            self.assertIn("Orphan    : 1", out)
            self.assertEqual(gtfs_edit.load_inactive_stops(state), {"SUD"})

    def test_activate_restores_content(self):
        with tempfile.TemporaryDirectory() as td:
            feed, state, master = self._paths(td)
            with open(master, "wb") as f:
                f.write(make_fixture_feed())
            disabled = os.path.join(td, "disabled.zip")
            self._run(["disable-stops", "--feed", feed, "--stop", "SUD",
                       "--out", disabled, "--zip", "--state", state])
            back = os.path.join(td, "back.zip")
            code, _ = self._run(["activate-stops", "--feed", disabled,
                                 "--out", back, "--zip", "--state", state,
                                 "--master", master])
            self.assertEqual(code, 0)
            self.assertEqual(gtfs_edit.load_inactive_stops(state), set())
            # Content must match the master again, ignoring the trailing newline
            # that the CSV writer adds to every row it rewrites.
            with open(back, "rb") as f:
                restored = f.read()
            self.assertEqual(gtfs_edit.find_unreferenced_stops(restored), set())
            with open(master, "rb") as f:
                original = f.read()
            for name in ("stops.txt", "stop_times.txt", "trips.txt", "routes.txt"):
                self.assertEqual(
                    rows_of(restored, name), rows_of(original, name), name
                )

    def test_stations_lists_the_database(self):
        _, out = self._run(["stations"])
        self.assertIn("JAKK", out)
        self.assertIn("Total 85 stasiun", out)

    def test_unknown_subcommand_is_not_swallowed(self):
        with self.assertRaises(SystemExit):
            self._run(["show-stops"])


if __name__ == "__main__":
    unittest.main()
