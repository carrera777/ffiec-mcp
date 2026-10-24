"""Local SQLite store for the FFIEC server.

Holds, in <download folder>/.ffiec_cache/cache.sqlite:
- filings:  a catalog of downloaded facsimile files (when each was fetched, and whether it is stale),
- missing:  filings FFIEC reported as not found, so resumed calls skip them,
- lookups:  reporting-period lists and Panels of Reporters, so restarts do not refetch them,
- checks:   when each period was last checked for new or amended filings,
- bulk:     Call Reports imported from FFIEC bulk-download ZIPs, one compressed row per bank and period.

Everything here is local to one user's machine. Nothing is shared or uploaded.
"""
from __future__ import annotations

import csv
import json
import re
import sqlite3
import threading
import time
import zlib
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterator

SCHEMA = """
CREATE TABLE IF NOT EXISTS filings (
    path TEXT PRIMARY KEY, period TEXT, fi_id TEXT, id_type TEXT, fmt TEXT, rssd TEXT,
    fetched_at REAL, stale INTEGER DEFAULT 0);
CREATE INDEX IF NOT EXISTS filings_period ON filings(period);
CREATE TABLE IF NOT EXISTS missing (
    path TEXT PRIMARY KEY, period TEXT, rssd TEXT, reason TEXT, at REAL);
CREATE TABLE IF NOT EXISTS lookups (key TEXT PRIMARY KEY, data TEXT, fetched_at REAL);
CREATE TABLE IF NOT EXISTS checks (period TEXT PRIMARY KEY, checked_at REAL, failed_at REAL);
CREATE TABLE IF NOT EXISTS bulk (
    period TEXT, rssd TEXT, as_of REAL, data BLOB, PRIMARY KEY (period, rssd));
CREATE TABLE IF NOT EXISTS bulk_panel (period TEXT PRIMARY KEY, data TEXT);
CREATE TABLE IF NOT EXISTS bulk_imports (source TEXT PRIMARY KEY, periods TEXT, banks INTEGER, imported_at REAL);
"""

try:
    from zoneinfo import ZoneInfo
    EASTERN: Any = ZoneInfo("America/New_York")
except Exception:  # no tz database (some Windows Pythons): EST, which errs toward treating copies as stale
    EASTERN = timezone(timedelta(hours=-5))

MDRM_RE = re.compile(r"^[A-Z]{4}[A-Z0-9]{4}$")


def eastern_ts(s: Any) -> float | None:
    """FFIEC timestamps (Eastern time, several formats) -> epoch seconds."""
    s = str(s or "").strip()
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
                "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[:26].split(".")[0], fmt).replace(tzinfo=EASTERN).timestamp()
        except ValueError:
            continue
    return None


class Store:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self.lock = threading.RLock()
        self.db = sqlite3.connect(str(path), check_same_thread=False, timeout=30)
        self.db.executescript(SCHEMA)
        self.db.commit()

    def _q(self, sql: str, args: tuple = ()) -> list[tuple]:
        with self.lock:
            return self.db.execute(sql, args).fetchall()

    def _w(self, sql: str, args: tuple = ()) -> int:
        with self.lock:
            n = self.db.execute(sql, args).rowcount
            self.db.commit()
            return n

    # ------------------------------------------------------------ filing catalog

    def is_fresh(self, path: Path, period: str, fi_id: str, id_type: str, fmt: str, rssd: str | None) -> bool:
        """True if `path` exists and has not been marked stale. Files cached by older versions are registered
        with their modification time, so the next amendment check covers them."""
        if not path.exists():
            return False
        row = self._q("SELECT stale FROM filings WHERE path=?", (str(path),))
        if not row:
            self.record(path, period, fi_id, id_type, fmt, rssd, path.stat().st_mtime)
            return True
        return not row[0][0]

    def record(self, path: Path, period: str, fi_id: str, id_type: str, fmt: str, rssd: str | None,
               fetched_at: float | None = None) -> None:
        self._w("INSERT OR REPLACE INTO filings VALUES (?,?,?,?,?,?,?,0)",
                (str(path), period, str(fi_id), id_type, fmt, rssd, fetched_at or time.time()))
        self._w("DELETE FROM missing WHERE path=?", (str(path),))

    # ------------------------------------------------------------ not-found memory

    def missing(self, path: Path, ttl: float) -> str | None:
        row = self._q("SELECT reason, at FROM missing WHERE path=?", (str(path),))
        return row[0][0] if row and time.time() - row[0][1] < ttl else None

    def set_missing(self, path: Path, period: str, rssd: str | None, reason: str) -> None:
        self._w("INSERT OR REPLACE INTO missing VALUES (?,?,?,?,?)", (str(path), period, rssd, reason, time.time()))

    # ------------------------------------------------------------ persisted lookups

    def lookup(self, key: str, ttl: float) -> Any | None:
        row = self._q("SELECT data, fetched_at FROM lookups WHERE key=?", (key,))
        return json.loads(row[0][0]) if row and time.time() - row[0][1] < ttl else None

    def put_lookup(self, key: str, data: Any) -> None:
        self._w("INSERT OR REPLACE INTO lookups VALUES (?,?,?)", (key, json.dumps(data), time.time()))

    def panels(self) -> Iterator[list[dict]]:
        """Every cached panel, newest period first (for ID lookups that must not cost an API call)."""
        rows = self._q("SELECT key, data FROM lookups WHERE key LIKE 'panel:%'")
        rows += [("panel:" + p, d) for p, d in self._q("SELECT period, data FROM bulk_panel")]
        key = lambda r: (r[0][-4:], r[0][-10:-8], r[0][-7:-5])  # noqa: E731  panel:MM/DD/YYYY
        for _, data in sorted(rows, key=key, reverse=True):
            yield json.loads(data)

    def bulk_panel(self, period: str) -> list[dict] | None:
        row = self._q("SELECT data FROM bulk_panel WHERE period=?", (period,))
        return json.loads(row[0][0]) if row else None

    # ------------------------------------------------------------ amendment checks

    def check_since(self, period: str, ttl: float, backoff: float = 3600) -> float | None:
        """If `period` is due for an amendment check, the epoch time to ask FFIEC about; else None."""
        now = time.time()
        row = self._q("SELECT checked_at, failed_at FROM checks WHERE period=?", (period,))
        if row:
            checked, failed = row[0]
            if checked and now - checked < ttl:
                return None
            if failed and now - failed < backoff:
                return None
        oldest = self._q(
            "SELECT MIN(t) FROM (SELECT MIN(fetched_at) t FROM filings WHERE period=? AND stale=0 "
            "UNION ALL SELECT MIN(as_of) FROM bulk WHERE period=? UNION ALL SELECT MIN(at) FROM missing WHERE period=?)",
            (period, period, period))[0][0]
        if oldest is None:
            return None  # nothing cached for this period, nothing to check
        return min(oldest, row[0][0]) if row and row[0][0] else oldest

    def check_done(self, period: str, failed: bool = False) -> None:
        now = time.time()
        if failed:
            self._w("INSERT INTO checks VALUES (?,NULL,?) ON CONFLICT(period) DO UPDATE SET failed_at=?",
                    (period, now, now))
        else:
            self._w("INSERT OR REPLACE INTO checks VALUES (?,?,NULL)", (period, now))

    def invalidate(self, period: str, submitted: dict[str, float]) -> dict[str, int]:
        """Mark cached copies older than each bank's latest submission as stale. `submitted` maps RSSD to
        its submission time (epoch). Also forgets not-found results for banks that have now filed."""
        out = {"files": 0, "bulk": 0, "missing": 0}
        with self.lock:
            for path, rssd, fetched in self.db.execute(
                    "SELECT path, rssd, fetched_at FROM filings WHERE period=? AND stale=0", (period,)).fetchall():
                if rssd in submitted and submitted[rssd] > fetched:
                    self.db.execute("UPDATE filings SET stale=1 WHERE path=?", (path,))
                    out["files"] += 1
            for rssd, as_of in self.db.execute("SELECT rssd, as_of FROM bulk WHERE period=?", (period,)).fetchall():
                if rssd in submitted and submitted[rssd] > as_of:
                    self.db.execute("DELETE FROM bulk WHERE period=? AND rssd=?", (period, rssd))
                    out["bulk"] += 1
            for path, rssd in self.db.execute("SELECT path, rssd FROM missing WHERE period=?", (period,)).fetchall():
                if rssd in submitted:
                    self.db.execute("DELETE FROM missing WHERE path=?", (path,))
                    out["missing"] += 1
            self.db.commit()
        return out

    # ------------------------------------------------------------ bulk-imported Call Reports

    def bulk_rows(self, period: str, rssd: str) -> list[tuple[str, str, str]] | None:
        row = self._q("SELECT data FROM bulk WHERE period=? AND rssd=?", (period, str(rssd)))
        return json.loads(zlib.decompress(row[0][0])) if row else None

    def has_bulk(self, period: str, rssd: str) -> bool:
        return bool(self._q("SELECT 1 FROM bulk WHERE period=? AND rssd=?", (period, str(rssd))))

    def status(self) -> dict:
        per: dict[str, dict] = {}
        for p, fmt, n, stale in self._q("SELECT period, fmt, COUNT(*), SUM(stale) FROM filings GROUP BY period, fmt"):
            d = per.setdefault(p, {})
            d[fmt] = n
            if stale:
                d["stale"] = d.get("stale", 0) + stale
        for p, n in self._q("SELECT period, COUNT(*) FROM bulk GROUP BY period"):
            per.setdefault(p, {})["bulk_banks"] = n
        for p, t in self._q("SELECT period, checked_at FROM checks WHERE checked_at IS NOT NULL"):
            per.setdefault(p, {})["last_amendment_check"] = datetime.fromtimestamp(t).strftime("%Y-%m-%d %H:%M")
        key = lambda p: (p[6:], p[:2], p[3:5])  # noqa: E731
        return {"database": str(self.path), "database_mb": round(self.path.stat().st_size / 1e6, 1),
                "periods": {p: per[p] for p in sorted(per, key=key, reverse=True)},
                "bulk_imports": [{"source": s, "periods": json.loads(ps), "banks": b,
                                  "imported": datetime.fromtimestamp(t).strftime("%Y-%m-%d %H:%M")}
                                 for s, ps, b, t in self._q("SELECT * FROM bulk_imports ORDER BY imported_at DESC")]}


# ---------------------------------------------------------------- bulk import

def _period_from_name(name: str) -> str | None:
    m = re.search(r"(0[369]|12)(30|31)(\d{4})", Path(name).name)
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else None


def _norm_header(h: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", h.upper())


def _sources(src: Path) -> list[tuple[str, float, Any]]:
    """(name, timestamp, opener) for each .txt file in a bulk ZIP or an extracted folder."""
    import zipfile
    out = []
    if src.is_dir():
        for f in sorted(src.rglob("*.txt")):
            out.append((f.name, f.stat().st_mtime, (lambda f=f: f.read_bytes())))
    else:
        z = zipfile.ZipFile(src)
        for info in z.infolist():
            if info.filename.lower().endswith(".txt"):
                ts = datetime(*info.date_time).replace(tzinfo=EASTERN).timestamp()
                out.append((info.filename, ts, (lambda n=info.filename: z.read(n))))
    return out


def _read_table(raw: bytes) -> tuple[list[str], Iterator[list[str]]]:
    text = raw.decode("utf-8-sig", errors="replace")
    rows = csv.reader(text.splitlines(), delimiter="\t")
    header = [h.strip().strip('"') for h in next(rows, [])]
    return header, rows


POR_FIELDS = {  # normalized bulk POR header -> Panel of Reporters key used by the API
    "IDRSSD": "ID_RSSD", "FDICCERTIFICATENUMBER": "FDICCertNumber", "OCCCHARTERNUMBER": "OCCChartNumber",
    "OTSDOCKETNUMBER": "OTSDockNumber", "PRIMARYABAROUTINGNUMBER": "PrimaryABARoutNumber",
    "FINANCIALINSTITUTIONNAME": "Name", "FINANCIALINSTITUTIONADDRESS": "Address",
    "FINANCIALINSTITUTIONCITY": "City", "FINANCIALINSTITUTIONSTATE": "State",
    "FINANCIALINSTITUTIONZIPCODE": "ZIP", "FINANCIALINSTITUTIONFILINGTYPE": "FilingType",
    "LASTDATETIMESUBMISSIONUPDATEDON": "LastUpdate",
}


def import_bulk(store: Store, src: Path, time_budget: float) -> dict:
    """Load an FFIEC 'Call Reports -- Single Period' tab-delimited bulk ZIP (or its extracted folder).

    Staged in a scratch database so a large ZIP can be loaded across several calls: each call parses
    files until `time_budget` seconds pass, then returns status "in_progress".
    """
    t0 = time.monotonic()
    files = _sources(src)
    if not files:
        raise ValueError(f"No .txt files found in {src}. Expected the tab-delimited Call Report bulk ZIP.")
    stamp = f"{src.resolve()}|{src.stat().st_size}|{src.stat().st_mtime}"
    stage_path = store.path.with_name("bulk_stage.sqlite")
    stage = sqlite3.connect(str(stage_path))
    stage.executescript("""
        CREATE TABLE IF NOT EXISTS meta (stamp TEXT);
        CREATE TABLE IF NOT EXISTS done (name TEXT PRIMARY KEY);
        CREATE TABLE IF NOT EXISTS por (period TEXT, rssd TEXT, data TEXT, as_of REAL);
        CREATE TABLE IF NOT EXISTS stage (period TEXT, rssd TEXT, items TEXT);
        CREATE TABLE IF NOT EXISTS file_ts (period TEXT PRIMARY KEY, ts REAL);""")
    if stage.execute("SELECT stamp FROM meta").fetchone() != (stamp,):
        stage.executescript("DELETE FROM meta; DELETE FROM done; DELETE FROM por; DELETE FROM stage; DELETE FROM file_ts;")
        stage.execute("INSERT INTO meta VALUES (?)", (stamp,))
        stage.commit()
    done = {r[0] for r in stage.execute("SELECT name FROM done")}
    skipped = []
    todo = [f for f in files if f[0] not in done]
    for i, (name, ts, read) in enumerate(todo):
        if i and time.monotonic() - t0 > time_budget:
            stage.close()
            return {"status": "in_progress", "files_done": len(files) - len(todo) + i, "files_total": len(files),
                    "next_step": "Call ffiec_import_bulk again with the same path; it resumes where it stopped."}
        period = _period_from_name(name)
        base = Path(name).name
        if not period:
            skipped.append(base)
        elif "POR" in base.upper().split():
            header, rows = _read_table(read())
            keys = [POR_FIELDS.get(_norm_header(h), h) for h in header]
            batch = []
            for row in rows:
                rec = {k: v.strip() for k, v in zip(keys, row)}
                rssd = rec.get("ID_RSSD", "")
                if not rssd.isdigit():
                    continue
                for k in ("ID_RSSD", "FDICCertNumber", "OCCChartNumber", "OTSDockNumber", "PrimaryABARoutNumber"):
                    if rec.get(k, "").isdigit():
                        rec[k] = int(rec[k])
                rec["HasFiledForReportingPeriod"] = True
                batch.append((period, rssd, json.dumps(rec), eastern_ts(rec.get("LastUpdate"))))
            stage.executemany("INSERT INTO por VALUES (?,?,?,?)", batch)
        elif re.search(r"schedule\s+\S", base, re.I):
            sched = re.search(r"schedule\s+([A-Z0-9-]+)", base, re.I).group(1).upper()
            header, rows = _read_table(read())
            norm = [_norm_header(h) for h in header]
            if "IDRSSD" not in norm:
                skipped.append(base)
            else:
                id_col = norm.index("IDRSSD")
                cols = [(j, h) for j, h in enumerate(norm) if MDRM_RE.match(h)]
                batch = []
                for row in rows:
                    rssd = row[id_col].strip().strip('"') if id_col < len(row) else ""
                    if not rssd.isdigit():
                        continue  # the second header row holds item descriptions
                    items = []
                    for j, code in cols:
                        if j < len(row):
                            v = row[j].strip().strip('"')
                            if v != "":
                                items.append((code, v, sched))
                    batch.append((period, rssd, json.dumps(items, separators=(",", ":"))))
                stage.executemany("INSERT INTO stage VALUES (?,?,?)", batch)
            stage.execute("INSERT INTO file_ts VALUES (?,?) ON CONFLICT(period) DO UPDATE SET ts=MIN(ts, excluded.ts)",
                          (period, ts))
        else:
            skipped.append(base)
        stage.execute("INSERT INTO done VALUES (?)", (name,))
        stage.commit()

    if todo and time.monotonic() - t0 > time_budget / 2:
        stage.close()  # leave the final step (about as long as parsing) to a fresh call
        return {"status": "in_progress", "files_done": len(files), "files_total": len(files),
                "next_step": "All files read. Call ffiec_import_bulk again with the same path to finish."}

    # Every file parsed: build one compressed row per bank and period.
    stage.execute("CREATE INDEX IF NOT EXISTS stage_bank ON stage(period, rssd)")
    file_ts = dict(stage.execute("SELECT period, ts FROM file_ts"))
    as_of = {(p, r): t for p, r, t in stage.execute("SELECT period, rssd, as_of FROM por") if t}
    periods: dict[str, int] = {}
    with store.lock:
        cur, bank, items = None, None, []

        def flush() -> None:
            if bank is None:
                return
            t = as_of.get((cur, bank)) or file_ts.get(cur) or time.time()
            store.db.execute("INSERT OR REPLACE INTO bulk VALUES (?,?,?,?)",
                             (cur, bank, t, zlib.compress(json.dumps(items, separators=(",", ":")).encode(), 1)))
            periods[cur] = periods.get(cur, 0) + 1

        seen: set[str] = set()
        for p, rssd, chunk in stage.execute("SELECT period, rssd, items FROM stage ORDER BY period, rssd, rowid"):
            if (p, rssd) != (cur, bank):
                flush()
                cur, bank, items, seen = p, rssd, [], set()
            for code, v, sched in json.loads(chunk):
                if code not in seen:  # an item repeated across schedule files is kept once
                    seen.add(code)
                    items.append((code, v, sched))
        flush()
        for p in {r[0] for r in stage.execute("SELECT DISTINCT period FROM por")}:
            panel = [json.loads(d) for (d,) in stage.execute("SELECT data FROM por WHERE period=?", (p,))]
            store.db.execute("INSERT OR REPLACE INTO bulk_panel VALUES (?,?)", (p, json.dumps(panel)))
        store.db.execute("INSERT OR REPLACE INTO bulk_imports VALUES (?,?,?,?)",
                         (src.name, json.dumps(sorted(periods)), sum(periods.values()), time.time()))
        store.db.commit()
    stage.close()
    stage_path.unlink(missing_ok=True)
    return {"status": "complete", "periods": periods, "files": len(files), "skipped_files": skipped}
