"""gtfs_edit -- read an existing GTFS feed and rebuild it with stops toggled.

Standard library only. Operates on any GTFS .zip, not just KCI feeds.

The disable/activate cycle is lossless in both directions:

* ``stops.txt`` keeps every row, so a disabled stop stays visible in the feed.
* ``stop_times.txt`` loses the rows of disabled stops; trips left with fewer
  than 2 stops are dropped, since GTFS needs at least an origin and a
  destination.
* Re-activating a stop restores its ``stop_times`` rows from a master feed,
  keeping the original ordering and times.
"""

import csv
import io
import json
import os
import zipfile
from datetime import datetime

INACTIVE_STOPS_DEFAULT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                      "inactive_stops.json")
GTFS_MASTER_DEFAULT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   "gtfs_master.zip")
GTFS_MASTER_PREV_DEFAULT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                        "gtfs_master_prev.zip")

PASS_THROUGH = ("stops.txt", "stop_times.txt", "trips.txt", "routes.txt", "calendar.txt")
ALWAYS_PRESENT = ("agency.txt", "levels.txt")


# ===========================================================================
# Persisted state
# ===========================================================================
def load_inactive_stops(path=INACTIVE_STOPS_DEFAULT):
    """Read the inactive stop_id list from JSON. Missing or broken file -> empty set."""
    if not path or not os.path.exists(path):
        return set()
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        ids = data.get("inactive_stop_ids", []) if isinstance(data, dict) else data
        return {str(x).strip() for x in ids if str(x).strip()}
    except (OSError, ValueError, TypeError):
        return set()


def save_inactive_stops(path, stop_ids):
    """Write the inactive stop_id list. Never raises; returns True on success."""
    try:
        ids = sorted({str(x).strip() for x in stop_ids if str(x).strip()})
        payload = {
            "inactive_stop_ids": ids,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        return True
    except OSError:
        return False


def load_gtfs_master(path=GTFS_MASTER_DEFAULT):
    """Read the master feed bytes, or None if unavailable.

    The master is the source of the original stop_times when re-activating:
    a generated feed is the master minus the currently disabled stops.
    """
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, "rb") as f:
            return f.read()
    except OSError:
        return None


def save_gtfs_master(path, data):
    """Write master feed bytes, replacing any previous master."""
    if not data:
        return False
    try:
        with open(path, "wb") as f:
            f.write(data)
        return True
    except OSError:
        return False


# ===========================================================================
# Low-level zip / csv plumbing
# ===========================================================================
def as_zip(zip_buffer):
    """Normalise bytes into a seekable BytesIO."""
    if isinstance(zip_buffer, (bytes, bytearray, memoryview)):
        buf = io.BytesIO(zip_buffer)
        buf.seek(0)
        return buf
    return zip_buffer


def open_zip(zip_buffer):
    """Open a GTFS feed given as bytes or a file-like object."""
    if isinstance(zip_buffer, (bytes, bytearray, memoryview)):
        return zipfile.ZipFile(io.BytesIO(zip_buffer))
    return zipfile.ZipFile(zip_buffer)


def read_csv_rows(zf, name):
    """Return (fieldnames, rows) for a member, or (None, []) when absent."""
    if name not in zf.namelist():
        return None, []
    raw = zf.read(name).decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(raw))
    return reader.fieldnames, [dict(r) for r in reader]


def write_csv_rows(rows, fieldnames):
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=fieldnames, lineterminator="\n")
    writer.writeheader()
    for r in rows:
        writer.writerow(r)
    return out.getvalue().encode("utf-8")


def count_gtfs_rows(zip_buffer, name):
    """Data row count (header excluded) of a member. 0 when absent."""
    if not zip_buffer:
        return 0
    try:
        with open_zip(zip_buffer) as zf:
            if name not in zf.namelist():
                return 0
            raw = zf.read(name).decode("utf-8-sig")
            return max(0, len(raw.strip().splitlines()) - 1)
    except (OSError, ValueError, KeyError):
        return 0


# ===========================================================================
# Reading
# ===========================================================================
def read_gtfs_stops(zip_buffer):
    """stops.txt rows, excluding the location_type 2/3/4 entries.

    Returns dicts with stop_id, stop_name, stop_code, stop_lat, stop_lon.
    """
    stops = []
    with open_zip(zip_buffer) as zf:
        if "stops.txt" not in zf.namelist():
            raise ValueError("File GTFS tidak memiliki stops.txt.")
        raw = zf.read("stops.txt").decode("utf-8-sig")
        for row in csv.DictReader(io.StringIO(raw)):
            stop_id = (row.get("stop_id") or "").strip()
            if not stop_id:
                continue
            if (row.get("location_type") or "").strip() in ("2", "3", "4"):
                continue
            stops.append({
                "stop_id": stop_id,
                "stop_name": (row.get("stop_name") or "").strip(),
                "stop_code": (row.get("stop_code") or "").strip(),
                "stop_lat": (row.get("stop_lat") or "").strip(),
                "stop_lon": (row.get("stop_lon") or "").strip(),
            })
    return stops


def read_gtfs_trip_stop_ids(zip_buffer):
    """Set of stop_id values actually used by stop_times.txt."""
    stop_ids = set()
    with open_zip(zip_buffer) as zf:
        if "stop_times.txt" not in zf.namelist():
            return stop_ids
        raw = zf.read("stop_times.txt").decode("utf-8-sig")
        for row in csv.DictReader(io.StringIO(raw)):
            sid = (row.get("stop_id") or "").strip()
            if sid:
                stop_ids.add(sid)
    return stop_ids


def find_unreferenced_stops(zip_buffer, exclude_location_types=("2", "3", "4")):
    """stops.txt stop_id values that stop_times.txt never references.

    After a rebuild these are exactly the disabled stops, which lets the
    disabled set be recovered from the feed alone. location_type 2/3/4 are
    ignored because they are entrances and platforms.
    """
    with open_zip(zip_buffer) as zf:
        names = set(zf.namelist())
        if "stops.txt" not in names or "stop_times.txt" not in names:
            return set()
        referenced = {
            (row.get("stop_id") or "").strip()
            for row in csv.DictReader(io.StringIO(zf.read("stop_times.txt").decode("utf-8-sig")))
            if (row.get("stop_id") or "").strip()
        }
        all_stops = set()
        for row in csv.DictReader(io.StringIO(zf.read("stops.txt").decode("utf-8-sig"))):
            sid = (row.get("stop_id") or "").strip()
            if not sid or (row.get("location_type") or "").strip() in exclude_location_types:
                continue
            all_stops.add(sid)
    return all_stops - referenced


def classify_upload(uploaded_bytes, expected_stop_ids):
    """Check whether an uploaded feed is a full feed, to protect the master.

    Orphan detection alone only catches stops present in stops.txt but missing
    from stop_times.txt. A partial upload that drops a stop from stops.txt as
    well slips past that check. This closes the gap, treating the upload as a
    full feed only once it holds at least half the expected stops.

    Returns looks_full, missing_stops and orphan_ids.
    """
    missing = []
    orphan_ids = set()
    try:
        present = {s["stop_id"] for s in read_gtfs_stops(uploaded_bytes)}
    except Exception:  # noqa: BLE001
        present = set()
    expected = {str(sid).strip() for sid in (expected_stop_ids or set()) if str(sid).strip()}
    looks_full = len(present) >= max(1, len(expected) // 2)
    if looks_full:
        missing = sorted(expected - present)
    try:
        orphan_ids = find_unreferenced_stops(uploaded_bytes)
    except Exception:  # noqa: BLE001
        orphan_ids = set()
    return {"looks_full": looks_full, "missing_stops": missing, "orphan_ids": orphan_ids}


# ===========================================================================
# Rebuilding
# ===========================================================================
def rebuild_gtfs_without_inactive(zip_buffer, inactive_stop_ids):
    """Rebuild a feed with the given stops disabled. Returns a BytesIO.

    - stops.txt: every row is kept, so disabled stops remain listed.
    - stop_times.txt: rows for disabled stops are dropped, and any trip left
      with fewer than 2 stops is dropped with them.
    - trips.txt: only trips that still have stop_times survive.
    - routes.txt / calendar.txt: entries no longer referenced are dropped.
    - Every other member passes through unchanged.
    """
    inactive = {str(s).strip() for s in (inactive_stop_ids or set())}

    with open_zip(zip_buffer) as zf:
        names = sorted(zf.namelist())
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as oz:
            stops_f, stops = read_csv_rows(zf, "stops.txt")
            st_f, st_rows = read_csv_rows(zf, "stop_times.txt")
            trips_f, trips = read_csv_rows(zf, "trips.txt")
            routes_f, routes = read_csv_rows(zf, "routes.txt")
            cal_f, cal = read_csv_rows(zf, "calendar.txt")

            oz.writestr("stops.txt", write_csv_rows(stops, stops_f) if stops_f else "")

            kept_trip_ids = set()
            if st_f is not None:
                kept = [r for r in st_rows if (r.get("stop_id") or "").strip() not in inactive]
                counts = {}
                for r in kept:
                    tid = r.get("trip_id")
                    if tid:
                        counts[tid] = counts.get(tid, 0) + 1
                kept_trip_ids = {tid for tid, n in counts.items() if n >= 2}
                kept = [r for r in kept if r.get("trip_id") in kept_trip_ids]
                oz.writestr("stop_times.txt", write_csv_rows(kept, st_f))
            else:
                oz.writestr("stop_times.txt", "")

            active_routes, active_services = set(), set()
            if trips_f is not None:
                kept_trips = [r for r in trips if r.get("trip_id") in kept_trip_ids]
                active_routes = {r.get("route_id") for r in kept_trips if r.get("route_id")}
                active_services = {r.get("service_id") for r in kept_trips if r.get("service_id")}
                oz.writestr("trips.txt", write_csv_rows(kept_trips, trips_f))
            else:
                oz.writestr("trips.txt", "")

            if routes_f is not None:
                oz.writestr("routes.txt",
                            write_csv_rows([r for r in routes if r.get("route_id") in active_routes], routes_f))
            if cal_f is not None:
                oz.writestr("calendar.txt",
                            write_csv_rows([r for r in cal if r.get("service_id") in active_services], cal_f))

            for n in names:
                if n in PASS_THROUGH:
                    continue
                oz.writestr(n, zf.read(n))
            for n in ALWAYS_PRESENT:
                if n not in names:
                    oz.writestr(n, "")

        out_buf.seek(0)
        return out_buf


def backfill_stop_times_from_master(upload_bytes, master_bytes, stop_ids):
    """Restore stop_times rows for re-activated stops from the master feed.

    Re-activating a stop whose stop_times rows were dropped during disabling
    needs those rows back, taken from the master snapshot taken before the
    master was last overwritten.

    - Only trip_id values present in the upload are touched, so the result stays
      consistent with the stops that were left in place.
    - Existing (trip_id, stop_id) pairs are never duplicated.
    - Rows are re-sorted per trip by stop_sequence, which keeps the feed valid
      and the file readable.
    - Headers, columns and all other members follow the upload.
    """
    stop_ids = {str(s).strip() for s in (stop_ids or set())}
    if not stop_ids or not master_bytes:
        return as_zip(upload_bytes)
    try:
        with open_zip(upload_bytes) as zf:
            names = sorted(zf.namelist())
            if "stop_times.txt" not in names or "trips.txt" not in names:
                return as_zip(upload_bytes)
            st_f, st = read_csv_rows(zf, "stop_times.txt")
            if st_f is None:
                return as_zip(upload_bytes)
            _, trips = read_csv_rows(zf, "trips.txt")
            upload_trips = {r.get("trip_id") for r in trips if r.get("trip_id")}
            if not upload_trips:
                return as_zip(upload_bytes)
    except (OSError, ValueError, KeyError):
        return as_zip(upload_bytes)

    existing = {(r.get("trip_id"), (r.get("stop_id") or "").strip()) for r in st}
    try:
        with open_zip(master_bytes) as mf:
            _mst_f, mst = read_csv_rows(mf, "stop_times.txt")
    except (OSError, ValueError, KeyError):
        mst = []
    if not mst:
        return as_zip(upload_bytes)

    added_by_trip = {}
    added = 0
    for r in mst:
        tid = r.get("trip_id")
        sid = (r.get("stop_id") or "").strip()
        if sid in stop_ids and tid in upload_trips and (tid, sid) not in existing:
            added_by_trip.setdefault(tid, []).append({f: r.get(f, "") for f in st_f})
            existing.add((tid, sid))
            added += 1
    if not added:
        return as_zip(upload_bytes)

    def _seq(r):
        try:
            return int(r.get("stop_sequence") or 0)
        except (TypeError, ValueError):
            return 0

    order, st_by_trip = [], {}
    for r in st:
        tid = r.get("trip_id") or ""
        if tid not in st_by_trip:
            order.append(tid)
            st_by_trip[tid] = []
        st_by_trip[tid].append(r)
    out = []
    for tid in order:
        rows = st_by_trip[tid] + added_by_trip.get(tid, [])
        rows.sort(key=_seq)
        out.extend(rows)

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as oz:
        with open_zip(upload_bytes) as zf:
            for n in names:
                oz.writestr(n, write_csv_rows(out, st_f) if n == "stop_times.txt" else zf.read(n))
    out_buf.seek(0)
    return out_buf


def apply_stop_edits(zip_buffer, edits=None, add_rows=None):
    """Edit stops.txt (and follow references in stop_times.txt). Returns BytesIO.

    - edits: {old_stop_id: {"stop_id": new_id, "stop_code": ..., "stop_name": ...}}.
      Only the supplied fields change. Renaming a stop_id also rewrites the
      matching stop_times.txt references. A rename to "" is ignored.
    - add_rows: rows for new stops, keyed by stops.txt column name, appended to
      stops.txt.

    Every other member passes through unchanged.
    """
    edits = edits or {}
    add_rows = add_rows or []

    with open_zip(zip_buffer) as zf:
        names = sorted(zf.namelist())
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as oz:
            stops_f, stops = read_csv_rows(zf, "stops.txt")
            st_f, st_rows = read_csv_rows(zf, "stop_times.txt")

            rename = {}
            if stops_f is not None:
                existing_ids = {r.get("stop_id", "").strip() for r in stops}
                for old_id, changes in edits.items():
                    new_id = str(changes.get("stop_id", "") or old_id).strip()
                    if new_id and new_id != old_id and new_id not in existing_ids and new_id not in rename.values():
                        rename[old_id] = new_id

                for r in stops:
                    old_id = (r.get("stop_id") or "").strip()
                    ch = edits.get(old_id, {})
                    for field in ("stop_id", "stop_code", "stop_name"):
                        if field in ch and ch[field] not in (None, ""):
                            r[field] = str(ch[field]).strip()
                    if old_id in rename and (r.get("stop_id") or "") == old_id:
                        r["stop_id"] = rename[old_id]

                existing_ids = {r.get("stop_id", "").strip() for r in stops}
                for nr in add_rows:
                    sid = str(nr.get("stop_id", "")).strip()
                    if not sid or sid in existing_ids:
                        continue
                    row = {}
                    for f in stops_f:
                        row[f] = str(nr.get(f, "") or "").strip() if nr.get(f) not in (None, "") else ""
                    row["stop_id"] = sid
                    row["location_type"] = row.get("location_type") or "1"
                    stops.append(row)
                    existing_ids.add(sid)

                oz.writestr("stops.txt", write_csv_rows(stops, stops_f))
            else:
                oz.writestr("stops.txt", "")

            if st_f is not None:
                new_rows = []
                for r in st_rows:
                    r = dict(r)
                    if rename:
                        sid = (r.get("stop_id") or "").strip()
                        if sid in rename:
                            r["stop_id"] = rename[sid]
                    new_rows.append(r)
                oz.writestr("stop_times.txt", write_csv_rows(new_rows, st_f))
            else:
                oz.writestr("stop_times.txt", "")

            for n in names:
                if n in ("stops.txt", "stop_times.txt"):
                    continue
                oz.writestr(n, zf.read(n))
            for n in ALWAYS_PRESENT:
                if n not in names:
                    oz.writestr(n, "")

        out_buf.seek(0)
        return out_buf
