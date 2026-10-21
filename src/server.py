# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "mcp>=1.10,<2",
#   "httpx>=0.27",
#   "openpyxl>=3.1",
# ]
# ///
"""FFIEC Central Data Repository (CDR) MCP server.

Wraps the CDR Public Web Service REST API (SIS611 v1.10):
https://cdr.ffiec.gov/public/Files/SIS611_-_Retrieve_Public_Data_via_Web_Service.pdf

Read-only. Requires a CDR PWS account (username + security token, renewed every 90 days).
"""
from __future__ import annotations

import base64
import csv
import gzip
import io
import json
import os
import re
import time
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

import httpx
from mcp.server.fastmcp import FastMCP

BASE_URL = "https://ffieccdr.azure-api.us/public/"
USER_ID = os.environ.get("FFIEC_USER_ID", "").strip()
TOKEN = re.sub(r"^\s*Bearer\s+", "", os.environ.get("FFIEC_TOKEN", "")).strip()


def _documents_dir() -> Path:
    """The user's real Documents folder (honors OneDrive redirection on Windows)."""
    if os.name == "nt":
        try:
            import ctypes
            buf = ctypes.create_unicode_buffer(260)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf) == 0 and buf.value:
                return Path(buf.value)
        except Exception:
            pass
    return Path.home() / "Documents"


def _resolve_dir(raw: str | None) -> Path:
    """Expand ${DOCUMENTS}/${HOME}/${DESKTOP}/${DOWNLOADS}, env vars and ~; fall back to Documents/FFIEC."""
    default = _documents_dir() / "FFIEC"
    if not raw or not raw.strip():
        return default
    s = raw.strip()
    for var, val in {"DOCUMENTS": _documents_dir(), "HOME": Path.home(), "DESKTOP": Path.home() / "Desktop",
                     "DOWNLOADS": Path.home() / "Downloads"}.items():
        s = s.replace("${" + var + "}", str(val))
    s = os.path.expandvars(os.path.expanduser(s))
    if "${" in s or not Path(s).is_absolute():
        return default
    return Path(s)


DOWNLOAD_DIR = _resolve_dir(os.environ.get("FFIEC_DOWNLOAD_DIR"))
HOURLY_LIMIT = 2400      # SIS611 3.1: about 2,500 calls/hour; keep a margin
MIN_INTERVAL = 0.2       # small courtesy gap between calls
TIME_BUDGET = float(os.environ.get("FFIEC_TIME_BUDGET", "40"))  # seconds per tool call before pausing

IdType = Literal["ID_RSSD", "FDICCertNumber", "OCCChartNumber", "OTSDockNumber"]

mcp = FastMCP("ffiec")
_last_call = 0.0
_call_times: list[float] = []
_missing: dict[str, str] = {}  # cache path -> reason, so resumed calls skip filings that do not exist
_panel_cache: dict[str, list[dict]] = {}
_labels: dict[str, str] | None = None
LABELS_BUNDLED = Path(__file__).with_name("mdrm_labels.json.gz")
MDRM_URL = "https://www.federalreserve.gov/apps/mdrm/pdf/MDRM.zip"


class FFIECError(Exception):
    pass


# ---------------------------------------------------------------- HTTP layer

def _norm_date(d: str) -> str:
    """Accept YYYY-MM-DD, MM/DD/YYYY, YYYYMMDD or e.g. 2025Q1 and return MM/DD/YYYY."""
    d = d.strip()
    m = re.fullmatch(r"(\d{4})\s*-?\s*Q([1-4])", d, re.I)
    if m:
        y, q = m.group(1), int(m.group(2))
        return {1: f"03/31/{y}", 2: f"06/30/{y}", 3: f"09/30/{y}", 4: f"12/31/{y}"}[q]
    m = re.fullmatch(r"(\d{4})-?(\d{2})-?(\d{2})", d)
    if m:
        return f"{m.group(2)}/{m.group(3)}/{m.group(1)}"
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", d)
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"
    raise FFIECError(f"Unrecognized date '{d}'. Use MM/DD/YYYY, YYYY-MM-DD, or e.g. 2025Q1.")


def _call(method: str, **headers: str) -> httpx.Response:
    global _last_call
    if not USER_ID or not TOKEN:
        raise FFIECError(
            "FFIEC credentials are not configured. Set the PWS username and security token "
            "in the extension settings (create them at https://cdr.ffiec.gov/public/PWS/)."
        )
    now = time.monotonic()
    _call_times[:] = [t for t in _call_times if now - t < 3600]
    if len(_call_times) >= HOURLY_LIMIT:
        raise FFIECError(f"Paused to stay under FFIEC's hourly limit ({HOURLY_LIMIT} calls in the last hour). "
                         f"Try again in {int(3600 - (now - _call_times[0])) // 60 + 1} minutes; cached data still works.")
    wait = MIN_INTERVAL - (now - _last_call)
    if wait > 0:
        time.sleep(wait)
    _call_times.append(time.monotonic())
    hdrs = {"UserID": USER_ID, "Authentication": f"Bearer {TOKEN}", "Accept": "application/json"}
    hdrs.update({k: v for k, v in headers.items() if v is not None})
    try:
        r = httpx.get(BASE_URL + method, headers=hdrs, timeout=120)
    finally:
        _last_call = time.monotonic()
    if r.status_code == 204:
        raise FFIECError("No content: facsimile not found or unavailable for that institution/period.")
    if r.status_code >= 400:
        msg = r.text.strip()[:500]
        hint = {
            401: " Check the username and that the token has not expired (tokens last 90 days).",
            403: " Hourly call quota (~2,500) exhausted; wait for it to replenish.",
        }.get(r.status_code, "")
        if "Security Token" in msg or "5002" in msg:
            hint = (" Regenerate the REST API token in the PWS portal, paste it in full into the extension "
                    "settings, confirm the username matches, then fully restart Claude Desktop.")
        raise FFIECError(f"FFIEC {method} returned HTTP {r.status_code}: {msg}{hint}")
    return r


def _date_key(d: str) -> tuple[int, int, int]:
    m, dd, y = (int(x) for x in d.split("/"))
    return (y, m, dd)


def _json(r: httpx.Response) -> Any:
    try:
        return r.json()
    except ValueError:
        raise FFIECError(f"Unexpected non-JSON response: {r.text[:300]}")


def _decode_file(r: httpx.Response) -> bytes:
    """Facsimile responses are a JSON-serialized byte array (base64 string or int list)."""
    ctype = r.headers.get("content-type", "")
    if "json" not in ctype and not r.content[:1] in (b'"', b"[", b"{"):
        return r.content
    data = _json(r)
    if isinstance(data, dict):
        data = next((v for k, v in data.items() if k.lower() in ("facsimilefile", "xbrlfile")), data)
    if isinstance(data, list) and len(data) == 1 and isinstance(data[0], str):
        data = data[0]
    if isinstance(data, str):
        return base64.b64decode(data)
    if isinstance(data, list) and all(isinstance(x, int) for x in data[:100]):
        return bytes(data)
    raise FFIECError(f"Could not decode file payload (type {type(data).__name__}).")


# ---------------------------------------------------------------- parsers

def _parse_sdf(text: str) -> list[dict]:
    """SDF = semicolon-delimited: Call Date;Bank RSSD Identifier;MDRM #;Value;Last Update;
    Short Definition;Call Schedule;Line Number."""
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return []
    delim = ";" if lines[0].count(";") >= lines[0].count("\t") else "\t"
    rows = list(csv.reader(lines, delimiter=delim))
    header = [h.strip().strip('"') for h in rows[0]]
    idx = {h.lower(): i for i, h in enumerate(header)}

    def col(row: list[str], *names: str) -> str:
        for n in names:
            for h, i in idx.items():
                if h.startswith(n) and i < len(row):
                    return row[i].strip().strip('"')
        return ""

    out = []
    for row in rows[1:]:
        mdrm = col(row, "mdrm")
        if not mdrm:
            continue
        out.append({
            "mdrm": mdrm,
            "value": col(row, "value"),
            "description": col(row, "short definition", "description"),
            "schedule": col(row, "call schedule", "schedule"),
            "line": col(row, "line"),
        })
    return out


def _parse_xbrl(content: bytes) -> list[dict]:
    root = ET.fromstring(content)
    contexts = {}
    for ctx in root.iter():
        if ctx.tag.endswith("}context"):
            inst = next((e.text for e in ctx.iter() if e.tag.endswith("}instant") or e.tag.endswith("}endDate")), "")
            dur = any(e.tag.endswith("}startDate") for e in ctx.iter())
            contexts[ctx.get("id")] = (inst, "duration" if dur else "instant")
    facts = []
    for el in root.iter():
        ctx = el.get("contextRef")
        if ctx is None:
            continue
        ns, _, local = el.tag[1:].partition("}") if el.tag.startswith("{") else ("", "", el.tag)
        facts.append({
            "concept": local,
            "value": (el.text or "").strip() if el.get("{http://www.w3.org/2001/XMLSchema-instance}nil") != "true" else None,
            "date": contexts.get(ctx, ("", ""))[0],
            "period_type": contexts.get(ctx, ("", ""))[1],
            "unit": el.get("unitRef"),
            "decimals": el.get("decimals"),
        })
    return facts


def _num(v: Any) -> float | None:
    try:
        return float(str(v).replace(",", "").rstrip("%"))
    except (TypeError, ValueError):
        return None


def _filter_items(items: list[dict], key: str, codes: list[str] | None, search: str | None,
                  schedule: str | None = None) -> list[dict]:
    if codes:
        wanted = {c.upper() for c in codes}
        suffix = {c for c in wanted if len(c) == 4}  # bare 4-char item numbers match any prefix
        items = [i for i in items if i[key].upper() in wanted or i[key].upper()[-4:] in suffix]
    if schedule:
        norm = lambda x: re.sub(r"[\s\-,]|SCHEDULE|PART", "", x.upper())
        s = norm(schedule)
        items = [i for i in items if norm(i.get("schedule", "")) == s]
    if search:
        q = search.lower()
        items = [i for i in items if q in json.dumps(i).lower()]
    return items


# ---------------------------------------------------------------- facsimile fetch + cache

def _cache_path(period: str, fi_id: str, id_type: str, fmt: str) -> Path:
    period = _norm_date(period)
    ymd = f"{period[6:]}{period[:2]}{period[3:5]}"
    ext = {"PDF": "pdf", "XBRL": "xml", "SDF": "txt", "UBPR": "xml"}[fmt]
    tag = "UBPR" if fmt == "UBPR" else "Call"
    return DOWNLOAD_DIR / f"{tag}_{id_type}_{fi_id}_{ymd}.{ext}"


def _is_ready(period: str, fi_id: str, id_type: str, fmt: str) -> bool:
    """True if the filing is cached on disk or already known not to exist (no API call needed)."""
    path = _cache_path(period, fi_id, id_type, fmt)
    return path.exists() or str(path) in _missing


def _fetch_facsimile(period: str, fi_id: str, id_type: str, fmt: str, refresh: bool = False) -> Path:
    period = _norm_date(period)
    path = _cache_path(period, fi_id, id_type, fmt)
    if path.exists() and not refresh:
        return path
    if str(path) in _missing and not refresh:
        raise FFIECError(_missing[str(path)])
    try:
        if fmt == "UBPR":
            r = _call("RetrieveUBPRXBRLFacsimile", reportingPeriodEndDate=period, fiIdType=id_type, fiId=str(fi_id))
        else:
            r = _call("RetrieveFacsimile", dataSeries="Call", reportingPeriodEndDate=period,
                      fiIdType=id_type, fiId=str(fi_id), facsimileFormat=fmt)
    except FFIECError as e:
        if "No content" in str(e) or "not found" in str(e).lower():
            _missing[str(path)] = str(e)
        raise
    data = _decode_file(r)
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def _panel(period: str, refresh: bool = False) -> list[dict]:
    period = _norm_date(period)
    if refresh or period not in _panel_cache:
        rows = _json(_call("RetrievePanelOfReporters", dataSeries="Call", reportingPeriodEndDate=period))
        _panel_cache[period] = [{k: v.strip() if isinstance(v, str) else v for k, v in b.items()} for b in rows]
    return _panel_cache[period]


def _latest_period() -> str:
    periods = _json(_call("RetrieveReportingPeriods", dataSeries="Call"))
    return _norm_date(max(periods, key=_date_key))


def _label_cache_path() -> Path:
    return DOWNLOAD_DIR / ".ffiec_cache" / "mdrm_labels.json"


def _get_labels() -> dict[str, str]:
    """MDRM item names keyed by full code (e.g. RIAD4340, UBPRE013). A refreshed copy wins over the bundled one."""
    global _labels
    if _labels is None:
        cache = _label_cache_path()
        try:
            if cache.exists():
                _labels = json.loads(cache.read_text(encoding="utf-8"))
            else:
                _labels = json.loads(gzip.decompress(LABELS_BUNDLED.read_bytes()).decode("utf-8"))
        except Exception:
            _labels = {}
    return _labels


def _refresh_labels() -> int:
    global _labels
    r = httpx.get(MDRM_URL, timeout=120, follow_redirects=True)
    r.raise_for_status()
    z = zipfile.ZipFile(io.BytesIO(r.content))
    name = next(n for n in z.namelist() if n.lower().endswith(".csv"))
    text = z.read(name).decode("latin-1").splitlines()
    best: dict[str, tuple[str, str]] = {}
    for row in csv.DictReader(text[1:]):
        mn = row.get("Mnemonic", "")
        if not (mn.startswith(("RC", "RI", "UBP")) or mn == "TEXT"):
            continue
        k = mn + row.get("Item Code", "")
        if k not in best or row.get("End Date", "") > best[k][1]:
            best[k] = (" ".join(row.get("Item Name", "").split()), row.get("End Date", ""))
    _labels = {k: v[0] for k, v in sorted(best.items())}
    cache = _label_cache_path()
    cache.parent.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps(_labels, separators=(",", ":")), encoding="utf-8")
    return len(_labels)


def _label(code: str) -> str | None:
    return _get_labels().get(code.upper())


# UBPR carries these Call Report capital ratios as decimal fractions rounded to 2 places (0.11 = 11%),
# unlike UBPR's own ratios, which are in percent units (4.25 = 4.25%).
UBPR_FRACTION = {
    "UBPR7204": "Decimal fraction rounded to 2 places (0.11 = 11%). For a precise percent use UBPRD486.",
    "UBPR7205": "Decimal fraction rounded to 2 places (0.14 = 14%). For a precise figure use the Call Report RCFA7205/RCOA7205.",
    "UBPR7206": "Decimal fraction rounded to 2 places (0.12 = 12%). For a precise percent use UBPRD487.",
}


def _err(e: Exception) -> dict:
    return {"error": str(e)}


# ---------------------------------------------------------------- tools

@mcp.tool()
def ffiec_reporting_periods(series: Literal["Call", "UBPR"] = "Call") -> dict:
    """List reporting period end dates available in the FFIEC CDR.

    series: "Call" for Call Reports (FFIEC 031/041/051) or "UBPR" for Uniform Bank Performance Reports.
    Returns dates newest first (MM/DD/YYYY).
    """
    try:
        if series == "UBPR":
            periods = _json(_call("RetrieveUBPRReportingPeriods"))
        else:
            periods = _json(_call("RetrieveReportingPeriods", dataSeries="Call"))
        periods = [_norm_date(p) for p in sorted(periods, key=_date_key, reverse=True)]
        return {"series": series, "count": len(periods), "latest": periods[0] if periods else None,
                "periods": periods}
    except Exception as e:
        return _err(e)


@mcp.tool()
def ffiec_find_bank(query: str = "", state: str | None = None, city: str | None = None,
                    period: str | None = None, filed_only: bool = False, limit: int = 25) -> dict:
    """Search the Call Report Panel of Reporters (all banks expected to file) for a period.

    Use this to resolve a bank name to its ID_RSSD / FDIC cert / OCC charter before pulling reports.
    query: case-insensitive substring of the bank name, or an exact ID (RSSD, FDIC cert, ABA).
    state: 2-letter state code. city: city name substring.
    period: reporting period (MM/DD/YYYY, YYYY-MM-DD, or 2025Q1). Defaults to the latest period.
    filed_only: only banks that have already filed for the period.
    FilingType is the form: 031 (foreign offices), 041 (domestic only), 051 (small domestic, <$5B).
    """
    try:
        period = _norm_date(period) if period else _latest_period()
        rows = _panel(period)
        q = query.strip().lower()
        out = []
        for b in rows:
            ids = {str(b.get(k)) for k in ("ID_RSSD", "FDICCertNumber", "OCCChartNumber", "PrimaryABARoutNumber")}
            if q and q not in str(b.get("Name", "")).lower() and q not in ids:
                continue
            if state and str(b.get("State", "")).upper() != state.upper():
                continue
            if city and city.lower() not in str(b.get("City", "")).lower():
                continue
            if filed_only and not b.get("HasFiledForReportingPeriod"):
                continue
            out.append(b)
        return {"period": period, "matches": len(out), "panel_size": len(rows), "results": out[:limit]}
    except Exception as e:
        return _err(e)


@mcp.tool()
def ffiec_recent_filers(period: str, since: str, include_times: bool = True) -> dict:
    """List institutions that filed (or amended) a Call Report for `period` after `since`.

    period: reporting period end date. since: date or datetime, e.g. "04/15/2025" or "04/15/2025 9:00 PM".
    include_times: return each filer's submission timestamp (RetrieveFilersSubmissionDateTime);
    otherwise just ID_RSSDs (RetrieveFilersSinceDate).
    """
    try:
        period = _norm_date(period)
        since = since.strip()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", since):
            since = _norm_date(since)
        method = "RetrieveFilersSubmissionDateTime" if include_times else "RetrieveFilersSinceDate"
        data = _json(_call(method, dataSeries="Call", reportingPeriodEndDate=period, lastUpdateDateTime=since))
        if include_times:
            from datetime import datetime
            ts = lambda f: datetime.strptime(f.get("DateTime", "1/1/1900 12:00:00 AM"), "%m/%d/%Y %I:%M:%S %p")
            data = sorted(data, key=ts, reverse=True)
        return {"period": period, "since": since, "count": len(data), "filers": data}
    except Exception as e:
        return _err(e)


@mcp.tool()
def ffiec_call_report(fi_id: str, period: str, id_type: IdType = "ID_RSSD",
                      items: list[str] | None = None, schedule: str | None = None,
                      search: str | None = None, limit: int = 200, refresh: bool = False) -> dict:
    """Pull a bank's Call Report data (parsed from the SDF facsimile) for one period.

    fi_id: institution identifier (default ID_RSSD; set id_type for FDIC cert etc.).
    items: MDRM codes to return, e.g. ["RCFD2170", "RIAD4340"]; a bare 4-character item number
      like "2170" matches any prefix (RCFD/RCON/RIAD...).
    schedule: filter to one schedule, e.g. "RC", "RI", "RC-C part I" / "RCCI", "RCRII" (hyphens optional).
    search: substring filter across code/description, e.g. "deposits".
    Dollar amounts are in thousands of USD. The raw file is cached in the download folder.
    """
    try:
        path = _fetch_facsimile(period, fi_id, id_type, "SDF", refresh)
        rows = _parse_sdf(path.read_text(encoding="utf-8-sig", errors="replace"))
        hits = _filter_items(rows, "mdrm", items, search, schedule)
        schedules = sorted({r["schedule"] for r in rows if r["schedule"]})
        return {"fi_id": fi_id, "id_type": id_type, "period": _norm_date(period), "units": "USD thousands",
                "total_items": len(rows), "matched": len(hits), "returned": min(len(hits), limit),
                "schedules": schedules, "file": str(path), "items": hits[:limit]}
    except Exception as e:
        return _err(e)


# Key Call Report items: label -> candidate MDRMs (consolidated first, then domestic).
KEY_ITEMS = {
    "total_assets": ["RCFD2170", "RCON2170"],
    "total_loans_and_leases": ["RCFD2122", "RCON2122"],
    "allowance_for_credit_losses": ["RCFD3123", "RCON3123"],
    "total_deposits": ["RCON2200"],
    "total_liabilities": ["RCFD2948", "RCON2948"],
    "total_equity_capital": ["RCFD3210", "RCON3210"],
    "total_interest_income": ["RIAD4107"],
    "total_interest_expense": ["RIAD4073"],
    "net_interest_income": ["RIAD4074"],
    "noninterest_income": ["RIAD4079"],
    "noninterest_expense": ["RIAD4093"],
    "provision_for_credit_losses": ["RIADJJ33", "RIAD4230"],
    "net_income": ["RIAD4340"],
    "tier1_leverage_ratio_pct": ["RCFA7204", "RCOA7204"],
    "tier1_capital_ratio_pct": ["RCFA7206", "RCOA7206"],
    "total_capital_ratio_pct": ["RCFA7205", "RCOA7205"],
}


def _summary(fi_id: str, period: str, id_type: str) -> dict:
    path = _fetch_facsimile(period, fi_id, id_type, "SDF")
    by_code = {r["mdrm"].upper(): r["value"] for r in _parse_sdf(path.read_text(encoding="utf-8-sig", errors="replace"))}
    out: dict[str, Any] = {}
    for label, codes in KEY_ITEMS.items():
        for c in codes:
            if c in by_code and by_code[c] != "":
                out[label] = _num(by_code[c]) if _num(by_code[c]) is not None else by_code[c]
                break
        else:
            out[label] = None
    ta, eq, ni = out.get("total_assets"), out.get("total_equity_capital"), out.get("net_income")
    loans, deps = out.get("total_loans_and_leases"), out.get("total_deposits")
    if isinstance(ta, float) and isinstance(eq, float) and ta:
        out["equity_to_assets_pct"] = round(100 * eq / ta, 2)
    if isinstance(loans, float) and isinstance(deps, float) and deps:
        out["loans_to_deposits_pct"] = round(100 * loans / deps, 2)
    if isinstance(ni, float) and isinstance(ta, float) and ta:
        q = int(_norm_date(period)[:2]) // 3
        out["roa_annualized_pct_approx"] = round(100 * ni * (4 / q) / ta, 2)
    return out


@mcp.tool()
def ffiec_bank_summary(fi_ids: list[str], period: str | None = None, id_type: IdType = "ID_RSSD") -> dict:
    """Headline balance sheet, income, and capital figures for one or more banks (up to 25).

    Pulls each bank's Call Report and extracts total assets, loans, deposits, equity, NII,
    net income (year-to-date), provision, and regulatory capital ratios, plus simple derived ratios.
    Dollar amounts in USD thousands; income items are year-to-date as reported.
    ROA is approximated from YTD net income annualized over period-end assets.
    period defaults to the latest Call Report period.
    """
    try:
        period = _norm_date(period) if period else _latest_period()
        results = {}
        for fid in fi_ids[:25]:
            try:
                results[str(fid)] = _summary(str(fid), period, id_type)
            except Exception as e:
                results[str(fid)] = _err(e)
        return {"period": period, "id_type": id_type, "units": "USD thousands; ratios in %", "banks": results}
    except Exception as e:
        return _err(e)


@mcp.tool()
def ffiec_bank_trend(fi_id: str, periods: list[str], items: list[str] | None = None,
                     id_type: IdType = "ID_RSSD") -> dict:
    """Time series of Call Report items for one bank across several periods (up to 20).

    periods: e.g. ["2024Q1", "2024Q2", "2024Q3", "2024Q4"].
    items: MDRM codes (e.g. ["RCON2200", "RIAD4340"]); omit to use the headline set from ffiec_bank_summary.
    Each period costs one API call the first time (cached afterward).
    """
    try:
        series: dict[str, dict] = {}
        t0 = time.monotonic()
        for n, p in enumerate(periods[:20]):
            pd = _norm_date(p)
            if n and time.monotonic() - t0 > TIME_BUDGET and not _is_ready(pd, fi_id, id_type, "SDF"):
                return {"fi_id": fi_id, "id_type": id_type, "units": "USD thousands", "series": series,
                        "status": "in_progress", "remaining_periods": [_norm_date(x) for x in periods[n:20]],
                        "next_step": "Call ffiec_bank_trend again with the same arguments; fetched periods are cached."}
            try:
                if items:
                    path = _fetch_facsimile(pd, fi_id, id_type, "SDF")
                    rows = _filter_items(_parse_sdf(path.read_text(encoding="utf-8-sig", errors="replace")),
                                         "mdrm", items, None)
                    series[pd] = {r["mdrm"]: _num(r["value"]) if _num(r["value"]) is not None else r["value"]
                                  for r in rows}
                else:
                    series[pd] = _summary(fi_id, pd, id_type)
            except Exception as e:
                series[pd] = _err(e)
        return {"fi_id": fi_id, "id_type": id_type, "units": "USD thousands", "series": series}
    except Exception as e:
        return _err(e)


@mcp.tool()
def ffiec_download_call_report(fi_id: str, period: str, format: Literal["PDF", "XBRL", "SDF"] = "PDF",
                               id_type: IdType = "ID_RSSD", refresh: bool = False) -> dict:
    """Download a Call Report facsimile file to disk and return its path.

    format: "PDF" (printable report), "XBRL" (instance document), or "SDF" (semicolon-delimited text).
    For XBRL, also returns a count and first facts parsed from the instance.
    """
    try:
        path = _fetch_facsimile(period, fi_id, id_type, format, refresh)
        out: dict[str, Any] = {"file": str(path), "bytes": path.stat().st_size, "format": format}
        if format == "XBRL":
            facts = _parse_xbrl(path.read_bytes())
            out.update(units="USD (whole dollars) for monetary facts", fact_count=len(facts), sample=facts[:20])
        return out
    except Exception as e:
        return _err(e)


@mcp.tool()
def ffiec_ubpr(fi_id: str, period: str, id_type: IdType = "ID_RSSD", concepts: list[str] | None = None,
               search: str | None = None, all_dates: bool = False, limit: int = 200,
               refresh: bool = False) -> dict:
    """Pull a bank's Uniform Bank Performance Report (UBPR) XBRL data for one period, with item names.

    Each fact has concept code, name (from the Federal Reserve MDRM dictionary), value, date, unit.
    concepts: exact codes, e.g. ["UBPRE013"]. search: case-insensitive match on code OR name,
      e.g. "net income", "tier 1", "efficiency". Use ffiec_lookup_items to explore names first.
    all_dates: the UBPR file also carries prior-period comparatives; default returns only the requested date.
    UBPR income ratios are YTD and annualized; UBPR dollar income items are YTD. Most UBPR ratios are in
    percent units (4.25 = 4.25%), but UBPR7204/7205/7206 are decimal fractions (see unit_note).
    """
    try:
        path = _fetch_facsimile(period, fi_id, id_type, "UBPR", refresh)
        facts = _parse_xbrl(path.read_bytes())
        iso = datetime.strptime(_norm_date(period), "%m/%d/%Y").strftime("%Y-%m-%d")
        if not all_dates:
            facts = [f for f in facts if f["date"] == iso]
        for f in facts:
            f["name"] = _label(f["concept"])
            if f["concept"].upper() in UBPR_FRACTION:
                f["unit_note"] = UBPR_FRACTION[f["concept"].upper()]
        hits = _filter_items(facts, "concept", concepts, search)
        return {"fi_id": fi_id, "id_type": id_type, "period": _norm_date(period),
                "units": "USD (whole dollars) for monetary facts; ratios as reported", "total_facts": len(facts),
                "matched": len(hits), "returned": min(len(hits), limit), "file": str(path),
                "facts": hits[:limit]}
    except Exception as e:
        return _err(e)


@mcp.tool()
def ffiec_lookup_items(query: str, prefix: str | None = None, limit: int = 30, refresh: bool = False) -> dict:
    """Look up Call Report and UBPR item names (Federal Reserve MDRM dictionary), in both directions.

    query: a full code ("UBPRE013", "RIAD4340"), a bare item number ("2170" matches RCFD2170, RCON2170, ...),
      or keywords ("return on assets", "tier 1 leverage"); all keywords must appear in the name.
    prefix: restrict to a series, e.g. "UBPR", "RIAD", "RCON", "RCFD".
    refresh: re-download the latest dictionary from federalreserve.gov (about 7.5 MB) before searching.
    """
    try:
        note = None
        if refresh:
            note = f"Refreshed {_refresh_labels()} names from {MDRM_URL}"
        labels = _get_labels()
        q = query.strip().upper()
        if re.fullmatch(r"[A-Z]{4}[A-Z0-9]{4}", q):
            hits = [(q, labels[q])] if q in labels else []
        elif re.fullmatch(r"[A-Z0-9]{4}", q):
            hits = [(k, v) for k, v in labels.items() if k[4:] == q]
        else:
            for phrase, repl in sorted(LOOKUP_SYNONYMS.items(), key=lambda x: -len(x[0])):
                q = re.sub(rf"\b{re.escape(phrase)}\b", repl, q)
            words = [w for w in re.split(r"[\s/]+", q) if w]
            hits = [(k, v) for k, v in labels.items() if all(w in v.upper() for w in words)]
            if not hits and len(words) > 1:
                scored = [(sum(w in v.upper() for w in words), k, v) for k, v in labels.items()]
                best = max((sc for sc, _, _ in scored), default=0)
                if best >= max(1, len(words) - 1):
                    hits = [(k, v) for sc, k, v in scored if sc == best]
                    note = f"No item contains every keyword; showing closest matches ({best} of {len(words)} words)."
        if prefix:
            hits = [(k, v) for k, v in hits if k.startswith(prefix.upper())]
        hits.sort(key=lambda kv: _prefix_rank(kv[0]))
        out = {"query": query, "matches": len(hits), "results": [{"code": k, "name": v} for k, v in hits[:limit]]}
        if note:
            out["note"] = note
        return out
    except Exception as e:
        return _err(e)


# Finance shorthand -> words used in MDRM item names
LOOKUP_SYNONYMS = {
    "RETURN ON ASSETS": "NET INCOME AVERAGE ASSETS", "ROA": "NET INCOME AVERAGE ASSETS",
    "RETURN ON EQUITY": "NET INCOME AVERAGE EQUITY", "ROE": "NET INCOME AVERAGE EQUITY",
    "NET INTEREST MARGIN": "NET INTEREST INCOME AVERAGE EARNING ASSETS",
    "NIM": "NET INTEREST INCOME AVERAGE EARNING ASSETS",
    "NET CHARGE-OFFS": "NET LOSS", "NET CHARGEOFFS": "NET LOSS", "NCO": "NET LOSS",
    "NONPERFORMING": "NONCURRENT", "NPL": "NONCURRENT", "CET1": "COMMON EQUITY TIER 1",
    "LOAN TO DEPOSIT": "LOANS DEPOSITS", "LOANS TO DEPOSITS": "LOANS DEPOSITS",
}
PREFIX_ORDER = ["UBPR", "RCFD", "RCON", "RIAD", "RCFA", "RCOA", "RCFN", "RCFW", "RCOW", "TEXT"]


def _prefix_rank(code: str) -> tuple:
    return (PREFIX_ORDER.index(code[:4]) if code[:4] in PREFIX_ORDER else len(PREFIX_ORDER), code)


def _qtr_range(start: str, end: str) -> list[str]:
    s, e = _norm_date(start), _norm_date(end)
    y, q = int(s[6:]), int(s[:2]) // 3
    ey, eq = int(e[6:]), int(e[:2]) // 3
    out = []
    while (y, q) <= (ey, eq):
        out.append(_norm_date(f"{y}Q{q}"))
        q += 1
        if q == 5:
            y, q = y + 1, 1
    return out


def _sched_norm(x: str) -> str:
    return re.sub(r"[\s\-,]|SCHEDULE|PART", "", str(x).upper())


@mcp.tool()
def ffiec_export_excel(fi_ids: list[str], start: str | None = None, end: str | None = None,
                       periods: list[str] | None = None, id_type: IdType = "ID_RSSD",
                       detail_schedules: list[str] | None = None, detail_items: list[str] | None = None,
                       ubpr_concepts: list[str] | None = None, output_path: str | None = None) -> dict:
    """Export Call Report financials for one or more banks to a formatted Excel workbook.

    Income statement items in the Call Report are YEAR-TO-DATE. Each bank sheet shows the balance
    sheet, the income statement YTD as filed, and the income statement as DISCRETE QUARTERS derived by
    live Excel formulas (current YTD less prior-quarter YTD; Q1 = YTD), plus key ratios and capital.
    Prior quarters needed for the derivation are fetched automatically as grey helper columns.

    fi_ids: up to 10 banks. Period range: start/end (e.g. "2024Q1" to "2026Q2"), or an explicit
      `periods` list; default is the latest 8 quarters. Max 20 quarters.
    detail_schedules: add a per-bank Detail sheet with every item in these schedules, e.g. ["RC-C part I",
      "RI-B part I"], or ["ALL"]. detail_items: specific MDRM codes for the Detail sheet. Detail rows are
      flagged YTD vs period end from the filing's own XBRL and get the same discrete-quarter formulas.
    ubpr_concepts: optional UBPR codes (e.g. ["UBPRE013"]) added as published on a UBPR sheet with names.
    output_path: .xlsx path; defaults to the download folder.
    Cost: about one API call per bank per quarter (cached), plus one per bank for Detail and one per
    bank-quarter for UBPR. Large exports download in stages: if the result has status "in_progress",
    call this tool again with exactly the same arguments until status is "complete".
    """
    try:
        import excel_export as xx
        if periods:
            req = sorted({_norm_date(p) for p in periods}, key=_date_key)
        else:
            end_p = _norm_date(end) if end else _latest_period()
            req = _qtr_range(start, end_p) if start else _qtr_range("2001Q1", end_p)[-8:]
        if not req:
            return {"error": "Empty period range."}
        if len(req) > 20:
            return {"error": f"{len(req)} quarters requested; the limit is 20 per export."}
        fi_ids = [str(f) for f in fi_ids][:10]
        helpers = {xx.prior_quarter(p) for p in req if xx.prior_quarter(p)} - set(req)
        columns = sorted(set(req) | helpers, key=_date_key)

        panel = {}
        try:
            panel = {str(b.get(id_type)): b for b in _panel(req[-1])}
        except Exception:
            pass

        want_detail = bool(detail_schedules or detail_items)
        # Download phase: fetch uncached filings until the time budget is used, then ask to be called again.
        jobs = [(p, fid, "SDF") for fid in fi_ids for p in columns]
        if want_detail:
            jobs += [(req[-1], fid, "XBRL") for fid in fi_ids]
        if ubpr_concepts:
            jobs += [(p, fid, "UBPR") for fid in fi_ids for p in req]
        pending = [j for j in jobs if not _is_ready(j[0], j[1], id_type, j[2])]
        t0 = time.monotonic()
        for i, (p, fid, fmt) in enumerate(pending):
            if i and time.monotonic() - t0 > TIME_BUDGET:
                left = len(pending) - i
                return {"status": "in_progress", "downloaded_this_call": i, "remaining_downloads": left,
                        "total_downloads": len(jobs),
                        "next_step": "Call ffiec_export_excel again with exactly the same arguments. Downloaded "
                                     "filings are cached, so it resumes where it stopped; the workbook is written "
                                     "once every filing is in hand.",
                        "estimated_calls_remaining": -(-left * 2 // max(1, int(TIME_BUDGET)))}
            try:
                _fetch_facsimile(p, fid, id_type, fmt)
            except FFIECError as e:
                if "hourly limit" in str(e):
                    return {"status": "paused", "error": str(e)}
            except Exception:
                pass

        banks, problems = [], []
        for fid in fi_ids:
            pb = panel.get(fid, {})
            bank = {"rssd": pb.get("ID_RSSD", fid), "name": pb.get("Name") or f"{id_type} {fid}",
                    "cert": pb.get("FDICCertNumber"), "form": pb.get("FilingType"), "city": pb.get("City"),
                    "state": pb.get("State"), "periods": {}, "desc": {}, "missing": {}}
            for p in columns:
                try:
                    path = _fetch_facsimile(p, fid, id_type, "SDF")
                    rows = _parse_sdf(path.read_text(encoding="utf-8-sig", errors="replace"))
                    bank["periods"][p] = {r["mdrm"].upper(): r["value"] for r in rows}
                    for r in rows:
                        bank["desc"][r["mdrm"].upper()] = (r["description"], r["schedule"], r["line"])
                except Exception as e:
                    bank["missing"][p] = str(e)
                    problems.append(f"{fid} {p}: {e}")
            if not bank["periods"]:
                problems.append(f"{fid}: no data for any period; skipped")
                continue
            if want_detail:
                latest = req[-1] if req[-1] in bank["periods"] else max(bank["periods"], key=_date_key)
                try:
                    xpath = _fetch_facsimile(latest, fid, id_type, "XBRL")
                    bank["basis"] = {f["concept"].upper(): f["period_type"] for f in _parse_xbrl(xpath.read_bytes())}
                except Exception as e:
                    problems.append(f"{fid}: XBRL basis flags unavailable ({e}); RIAD prefix rule used instead")
            banks.append(bank)
        if not banks:
            return {"error": "No data retrieved.", "problems": problems}

        select = None
        if want_detail:
            scheds = {_sched_norm(x) for x in (detail_schedules or [])}
            items = {i.upper() for i in (detail_items or [])}
            bare = {i for i in items if len(i) == 4}

            def select(code: str, sched: str) -> bool:
                return (("ALL" in scheds and not code.startswith("TEXT")) or _sched_norm(sched) in scheds
                        or code in items or code[-4:] in bare)

        ubpr = None
        if ubpr_concepts:
            wanted = {c.upper() for c in ubpr_concepts}
            ubpr = {}
            for b in banks:
                for p in req:
                    try:
                        fid_u = next(f for f in fi_ids if str(panel.get(f, {}).get("ID_RSSD", f)) == str(b["rssd"]))
                        facts = _parse_xbrl(_fetch_facsimile(p, fid_u, id_type, "UBPR").read_bytes())
                        iso = datetime.strptime(p, "%m/%d/%Y").strftime("%Y-%m-%d")
                        ubpr.setdefault(str(b["rssd"]), {})[p] = {
                            f["concept"].upper(): (f["value"], f["unit"]) for f in facts
                            if f["concept"].upper() in wanted and f["date"] == iso}
                    except Exception as e:
                        problems.append(f"UBPR {b['rssd']} {p}: {e}")

        if output_path:
            out = Path(os.path.expanduser(output_path))
        else:
            tag = str(banks[0]["rssd"]) if len(banks) == 1 else f"{len(banks)}banks"
            span = f"{xx.qlabel(req[0])}-{xx.qlabel(req[-1])}".replace(" ", "")
            out = DOWNLOAD_DIR / f"FFIEC_Call_{tag}_{span}.xlsx"
        out.parent.mkdir(parents=True, exist_ok=True)
        try:
            info = xx.build_workbook(str(out), banks, columns, helpers, select, _label, ubpr, UBPR_FRACTION)
        except PermissionError:
            out = out.with_name(out.stem + datetime.now().strftime("_%H%M%S") + out.suffix)
            info = xx.build_workbook(str(out), banks, columns, helpers, select, _label, ubpr, UBPR_FRACTION)
        return {"status": "complete", "file": str(out), "banks": [b["name"] for b in banks],
                "quarters": [xx.qlabel(p) for p in req],
                "helper_quarters": [xx.qlabel(p) for p in sorted(helpers, key=_date_key)],
                "sheets": info["sheets"], "detail_rows": info["detail_rows"], "problems": problems,
                "note": "Income statement shown YTD as filed and as discrete quarters (formula-derived)."}
    except Exception as e:
        return _err(e)


# ---------------------------------------------------------------- Federal Reserve holding company data (Y-9)
_fry9_store = None


def _fry9():
    """Lazy import of the Y-9 module and the store over the local BHCF folder (FFIEC_FRY9_DIR or <download dir>/FRY9)."""
    global _fry9_store
    import fry9
    raw = os.environ.get("FFIEC_FRY9_DIR", "").strip()
    folder = _resolve_dir(raw) if raw and "${user_config" not in raw else DOWNLOAD_DIR / "FRY9"
    if _fry9_store is None or _fry9_store.folder != folder:
        _fry9_store = fry9.Store(folder)
    return fry9, _fry9_store


def _fry9_pick(holding_company: str, period: str | None):
    """Resolve a name or RSSD to (rssd, None) or (None, error dict with candidates)."""
    fry9, store = _fry9()
    p = fry9.norm_period(period) if period else None
    rssd, cands = fry9.resolve(store, holding_company, p)
    if rssd:
        return rssd, None
    if not cands:
        return None, {"error": f"No holding company matched {holding_company!r}. Try fry9_find with fewer words."}
    return None, {"error": f"{holding_company!r} matches several holding companies; call again with an RSSD.", "candidates": cands}


FRY9_NOTES = ("Dollars in thousands. Y-9C is consolidated and quarterly (about $3B of assets and up), income items are year to date as filed. "
              "Y-9SP is parent-only and semiannual (June and December) with the holding company's consolidated assets reported separately. "
              "Y-9C ratios here are computed from the filed items; the Fed's own ratios (capital) are taken as filed.")


@mcp.tool()
def fry9_status() -> dict:
    """Show which Federal Reserve holding company (FR Y-9C / Y-9SP) quarterly files are loaded, whether newer ones should exist, and how to add them."""
    try:
        fry9, store = _fry9()
        have = store.periods()
        exp = fry9.expected_latest()
        return {"folder": str(store.folder), "periods_loaded": [fry9.fmt_period(p) for p in have],
                "latest_expected_by_now": fry9.fmt_period(exp) if exp else None,
                "newer_file_expected": bool(exp and (not have or exp > have[-1])),
                "how_to_add": f"Open {fry9.NIC_URL} in a browser (it blocks scripts), pick the year, click BHCFYYYYMMDD.ZIP, "
                              "then call fry9_import_file with the downloaded path. The Y-9 filing deadline is 45 days after quarter end; "
                              "June and December files include the semiannual Y-9SP filers.",
                "notes": FRY9_NOTES}
    except Exception as e:
        return _err(e)


@mcp.tool()
def fry9_import_file(path: str) -> dict:
    """Validate a downloaded NIC file (BHCFYYYYMMDD.ZIP) and copy it into the holding company data folder."""
    try:
        fry9, store = _fry9()
        info = fry9.import_zip(Path(path), store.folder)
        store._cache.pop(info["period"].replace("-", ""), None)
        return {**info, "periods_loaded": [fry9.fmt_period(p) for p in store.periods()]}
    except Exception as e:
        return _err(e)


@mcp.tool()
def fry9_find(query: str = "", state: str | None = None, period: str | None = None, limit: int = 25) -> dict:
    """Search Federal Reserve holding companies by name words and optional state (two letters). Returns RSSD, city, filer type (Y-9C, Y-9SP, Y-9LP) and assets, largest first. Period defaults to the latest loaded file."""
    try:
        fry9, store = _fry9()
        have = store.periods()
        if not have:
            return {"error": f"No BHCF files in {store.folder}. Call fry9_status for how to add them."}
        p = fry9.norm_period(period) if period else have[-1]
        hits = fry9.search(store, query, state, p, max(1, min(limit, 200)))
        return {"period": fry9.fmt_period(p), "count": len(hits), "results": hits,
                "note": "assets = consolidated total assets (Y-9C) or the holding company's consolidated assets as reported on the Y-9SP"}
    except Exception as e:
        return _err(e)


def _fry9_values(fry9, row: dict) -> tuple[str, dict]:
    ft = fry9.filer_type(row)
    h = fry9.headline(row)
    labs = fry9.labels_for(ft)
    out = {k: {"label": labs.get(k, k.replace("_", " ")), "value": v} for k, v in h.items()}
    return ft, out


@mcp.tool()
def fry9_summary(holding_company: str, period: str | None = None) -> dict:
    """Headline Federal Reserve filing for one holding company (name or RSSD) and period (default: latest loaded file where it filed).
    Y-9C filers get balance sheet, income (year to date), capital ratios, asset quality and computed ratios; Y-9SP filers get parent-only statements plus consolidated assets."""
    try:
        fry9, store = _fry9()
        rssd, bad = _fry9_pick(holding_company, period)
        if bad:
            return bad
        periods = store.periods()
        use = [fry9.norm_period(period)] if period else list(reversed(periods))
        for p in use:
            if p not in periods:
                return {"error": f"No file for {fry9.fmt_period(p)}. Loaded: {', '.join(fry9.fmt_period(x) for x in periods)}"}
            row = store.row(rssd, p)
            if row and fry9.filer_type(row) != "none":
                ft, vals = _fry9_values(fry9, row)
                return {"entity": fry9.entity(row), "period": fry9.fmt_period(p), "form": ft, "items": vals,
                        "months_ytd": int(p[4:6]), "notes": FRY9_NOTES}
        return {"error": f"RSSD {rssd} has no filing in the loaded files", "periods": [fry9.fmt_period(x) for x in periods]}
    except Exception as e:
        return _err(e)


@mcp.tool()
def fry9_trend(holding_company: str, periods: list[str] | None = None, discrete: bool = True) -> dict:
    """Multi-period series of the headline Federal Reserve items for one holding company across the loaded files.
    Income items are year to date as filed; with discrete=True each flow item also gets a single-quarter (Y-9C) or single-half (Y-9SP) value derived by differencing within the year."""
    try:
        fry9, store = _fry9()
        rssd, bad = _fry9_pick(holding_company, None)
        if bad:
            return bad
        have = store.periods()
        use = sorted(fry9.norm_period(p) for p in periods) if periods else have
        rows: dict[str, dict] = {}
        ent = None
        for p in use:
            if p not in have:
                continue
            row = store.row(rssd, p)
            if not row or fry9.filer_type(row) == "none":
                continue
            ent = ent or fry9.entity(row)
            ft = fry9.filer_type(row)
            rows[p] = {"form": ft, "ytd": fry9.headline(row)}
        out = []
        for p in sorted(rows):
            ft, ytd = rows[p]["form"], rows[p]["ytd"]
            rec: dict[str, Any] = {"period": fry9.fmt_period(p), "form": ft, "values": ytd}
            if discrete:
                step, first = (3, 3) if ft == "Y-9C" else (6, 6)
                month = int(p[4:6])
                prev = f"{p[:4]}{month - step:02d}{'30' if (month - step) in (6, 9) else '31'}"
                d: dict[str, Any] = {}
                for k in fry9.flow_keys(ft):
                    if k not in ytd:
                        continue
                    if month == first:
                        d[k] = ytd[k]
                    elif prev in rows and rows[prev]["form"] == ft and k in rows[prev]["ytd"]:
                        d[k] = ytd[k] - rows[prev]["ytd"][k]
                if d:
                    rec["discrete_period"] = d
            out.append(rec)
        return {"entity": ent, "periods": out, "notes": FRY9_NOTES,
                "discrete_note": "Single-period values need the prior filing in the same year; they are omitted when it is not loaded or the form changed."}
    except Exception as e:
        return _err(e)


@mcp.tool()
def fry9_items(holding_company: str, period: str | None = None, codes: list[str] | None = None, search: str | None = None,
               include_zero: bool = False, limit: int = 80) -> dict:
    """Raw Federal Reserve line items (MDRM codes such as BHCK2170, BHSP3210) for one holding company and period, with names.
    Filter by codes or a name search (e.g. 'nonaccrual', 'tier 1'). Zero items are hidden unless include_zero is true."""
    try:
        fry9, store = _fry9()
        rssd, bad = _fry9_pick(holding_company, period)
        if bad:
            return bad
        have = store.periods()
        p = fry9.norm_period(period) if period else next((x for x in reversed(have) if store.has(x, rssd)), None)
        if not p or p not in have:
            return {"error": "No matching file for that holding company and period"}
        row = store.row(rssd, p)
        if not row:
            return {"error": f"RSSD {rssd} did not file for {fry9.fmt_period(p)}"}
        want = {c.upper() for c in codes} if codes else None
        pat = re.compile(search, re.I) if search else None
        items = []
        for code, val in row.items():
            if code.startswith("RSSD") or code.startswith("TEXT"):
                continue
            lab = fry9.label(code) or ""
            if want and code not in want:
                continue
            if pat and not pat.search(lab + " " + code):
                continue
            num = fry9._num(val)
            if num == 0 and not include_zero and not want:
                continue
            items.append({"code": code, "name": lab, "value": num if num is not None else val})
        items.sort(key=lambda x: x["code"])
        return {"entity": fry9.entity(row), "period": fry9.fmt_period(p), "count": len(items), "items": items[:limit],
                "truncated": len(items) > limit, "notes": "Dollars in thousands; income items are year to date."}
    except Exception as e:
        return _err(e)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
