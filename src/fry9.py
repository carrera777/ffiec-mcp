"""Federal Reserve holding company data (FR Y-9C, FR Y-9LP, FR Y-9SP) from the FFIEC NIC quarterly BHCF files.

Pure standard library (no MCP or network code) so it can be tested offline. The NIC download page sits behind a
browser check that blocks scripts, so the quarterly files (BHCFYYYYMMDD.ZIP, caret-delimited, one row per
holding company) are placed in a local folder by hand or with fry9_import_file; this module only reads them.

Form coverage: Y-9C is consolidated and quarterly (about $3 billion of assets and up), Y-9LP is parent-only for
Y-9C filers, Y-9SP is parent-only and semiannual (June and December) for smaller holding companies. Y-9SP filers
also report the consolidated assets of the holding company (BHSP8519). All dollar items are in thousands.
"""
from __future__ import annotations

import gzip
import json
import re
import shutil
import zipfile
from datetime import date, timedelta
from pathlib import Path
from typing import Any

FILE_RE = re.compile(r"^BHCF(\d{8})\.zip$", re.I)
NIC_URL = "https://www.ffiec.gov/npw/FinancialReport/FinancialDataDownload"
LABELS_PATH = Path(__file__).with_name("bhc_labels.json.gz")

# ---------------------------------------------------------------- headline definitions
# (key, label, MDRM code or list of codes to sum, kind)  kind: bs = point in time, flow = year-to-date
Y9C_ITEMS: list[tuple[str, str, Any, str]] = [
    ("total_assets", "Total assets", "BHCK2170", "bs"),
    ("loans", "Loans and leases, net of unearned income", "BHCK2122", "bs"),
    ("allowance", "Allowance for loan and lease losses", "BHCK3123", "bs"),
    ("goodwill", "Goodwill", "BHCK3163", "bs"),
    ("deposits_noninterest", "Noninterest-bearing deposits", ["BHDM6631", "BHFN6631"], "bs"),
    ("deposits_interest", "Interest-bearing deposits", ["BHDM6636", "BHFN6636"], "bs"),
    ("total_liabilities", "Total liabilities", "BHCK2948", "bs"),
    ("total_equity", "Total equity capital", "BHCK3210", "bs"),
    ("interest_income", "Total interest income", "BHCK4107", "flow"),
    ("interest_expense", "Total interest expense", "BHCK4073", "flow"),
    ("nii", "Net interest income", "BHCK4074", "flow"),
    ("provision", "Provision for loan and lease losses", "BHCK4230", "flow"),
    ("nonint_income", "Total noninterest income", "BHCK4079", "flow"),
    ("nonint_expense", "Total noninterest expense", "BHCK4093", "flow"),
    ("pretax_income", "Income before income taxes", "BHCK4301", "flow"),
    ("net_income", "Net income", "BHCK4340", "flow"),
    ("chargeoffs", "Charge-offs on loans and leases", "BHCK4635", "flow"),
    ("recoveries", "Recoveries on loans and leases", "BHCK4605", "flow"),
    ("nonaccrual", "Loans and leases on nonaccrual", "BHCK1403", "bs"),
    ("past_due_90", "Past due 90+ days and still accruing", "BHCK1407", "bs"),
    ("past_due_30_89", "Past due 30-89 days and still accruing", "BHCK1406", "bs"),
    ("avg_assets_qtr", "Quarterly average total assets", "BHCK3368", "bs"),
    ("cet1_ratio", "Common equity tier 1 ratio (%)", "BHCAP793", "ratio"),
    ("tier1_rbc_ratio", "Tier 1 risk-based capital ratio (%)", "BHCA7206", "ratio"),
    ("total_rbc_ratio", "Total risk-based capital ratio (%)", "BHCA7205", "ratio"),
    ("tier1_leverage_ratio", "Tier 1 leverage ratio (%)", "BHCA7204", "ratio"),
    ("rwa", "Risk-weighted assets", "BHCAA223", "bs"),
    ("tier1_capital", "Tier 1 capital", "BHCA8274", "bs"),
]
Y9SP_ITEMS: list[tuple[str, str, Any, str]] = [
    ("consolidated_assets", "Consolidated assets of the holding company", "BHSP8519", "bs"),
    ("parent_assets", "Parent-only total assets", "BHSP2170", "bs"),
    ("parent_equity", "Parent-only total equity capital", "BHSP3210", "bs"),
    ("investment_in_bank_subs", "Parent investment in bank subsidiaries", "BHSP3239", "bs"),
    ("parent_lt_borrowings", "Parent long-term borrowings", "BHSP3151", "bs"),
    ("parent_st_borrowings", "Parent borrowings due in one year or less", "BHSP2724", "bs"),
    ("parent_cash", "Parent cash and balances due from depository institutions", "BHSP0010", "bs"),
    ("dividends_from_bank_subs", "Dividend income from bank subsidiaries", "BHSP0508", "flow"),
    ("parent_total_income", "Parent total interest and noninterest income", "BHSP4000", "flow"),
    ("parent_interest_expense", "Parent interest expense", "BHSP4073", "flow"),
    ("parent_pretax_income", "Parent income before income taxes and securities gains", "BHSP4250", "flow"),
    ("parent_net_income", "Parent-only net income (includes equity in undistributed income)", "BHSP4340", "flow"),
    ("dividends_declared", "Cash dividends declared to shareholders", "BHSP3158", "flow"),
]
Y9LP_ITEMS: list[tuple[str, str, Any, str]] = [
    ("parent_assets", "Parent-only total assets", "BHCP2170", "bs"),
    ("parent_equity", "Parent-only total equity capital", "BHCP3210", "bs"),
    ("parent_lt_borrowings", "Parent long-term borrowings", "BHCP3151", "bs"),
]

_labels: dict[str, str] | None = None


def label(code: str) -> str | None:
    global _labels
    if _labels is None:
        try:
            _labels = json.loads(gzip.decompress(LABELS_PATH.read_bytes()).decode("utf-8"))
        except Exception:
            _labels = {}
    return _labels.get(code.upper())


def _num(v: str | None) -> float | None:
    if v is None:
        return None
    v = v.strip()
    if not v:
        return None
    try:
        f = float(v)
    except ValueError:
        return None
    return int(f) if f == int(f) and abs(f) < 1e15 else f


def fmt_period(p: str) -> str:
    return f"{p[:4]}-{p[4:6]}-{p[6:]}"


def norm_period(p: str) -> str:
    """Accept 20260630, 2026-06-30, 6/30/2026 or 06/30/2026."""
    p = p.strip()
    m = re.fullmatch(r"(\d{4})-?(\d{2})-?(\d{2})", p)
    if m:
        return "".join(m.groups())
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", p)
    if m:
        return f"{m.group(3)}{int(m.group(1)):02d}{int(m.group(2)):02d}"
    raise ValueError(f"Unrecognized period {p!r}; use YYYY-MM-DD or M/D/YYYY")


class Store:
    """Reads BHCF zips from a folder; rows are parsed lazily and cached per file."""

    def __init__(self, folder: Path):
        self.folder = Path(folder)
        self._cache: dict[str, tuple[dict[str, int], dict[str, str]]] = {}

    # -- files
    def files(self) -> dict[str, Path]:
        out: dict[str, Path] = {}
        if self.folder.is_dir():
            for p in self.folder.iterdir():
                m = FILE_RE.match(p.name)
                if m and p.is_file():
                    out[m.group(1)] = p
        return dict(sorted(out.items()))

    def periods(self) -> list[str]:
        return list(self.files())

    def _load(self, period: str) -> tuple[dict[str, int], dict[str, str]]:
        if period not in self._cache:
            path = self.files().get(period)
            if path is None:
                raise FileNotFoundError(f"No BHCF{period}.ZIP in {self.folder}. Available: {', '.join(self.periods()) or 'none'}")
            z = zipfile.ZipFile(path)
            name = next(n for n in z.namelist() if n.lower().endswith(".txt"))
            lines = z.read(name).decode("latin-1").splitlines()
            hdr = lines[0].split("^")
            idx = {h: i for i, h in enumerate(hdr)}
            rows: dict[str, str] = {}
            for ln in lines[1:]:
                k = ln.split("^", 1)[0].strip()
                if k:
                    rows[k] = ln
            self._cache[period] = (idx, rows)
        return self._cache[period]

    def row(self, rssd: str, period: str) -> dict[str, str] | None:
        idx, rows = self._load(period)
        ln = rows.get(str(rssd))
        if ln is None:
            return None
        f = ln.split("^")
        return {h: f[i] for h, i in idx.items() if i < len(f) and f[i].strip() != ""}

    def all_rssd(self, period: str) -> list[str]:
        return list(self._load(period)[1])

    def has(self, period: str, rssd: str) -> bool:
        return str(rssd) in self._load(period)[1]


def filer_type(row: dict[str, str] | None) -> str:
    if not row:
        return "none"
    if _num(row.get("BHCK2170")) is not None:
        return "Y-9C"
    if _num(row.get("BHSP2170")) is not None:
        return "Y-9SP"
    if _num(row.get("BHCP2170")) is not None:
        return "Y-9LP"
    return "none"


def entity(row: dict[str, str]) -> dict:
    return {"rssd": row.get("RSSD9001"), "name": row.get("RSSD9017"), "city": row.get("RSSD9130"),
            "state": row.get("RSSD9200"), "filer": filer_type(row)}


def _val(row: dict[str, str], code: Any) -> float | None:
    if isinstance(code, list):
        vals = [_num(row.get(c)) for c in code]
        vals = [v for v in vals if v is not None]
        return sum(vals) if vals else None
    return _num(row.get(code))


def headline(row: dict[str, str]) -> dict[str, Any]:
    """Headline items plus derived ratios for one filing. Flow items are year to date as filed."""
    ft = filer_type(row)
    items = Y9C_ITEMS if ft == "Y-9C" else Y9SP_ITEMS if ft == "Y-9SP" else Y9LP_ITEMS
    out: dict[str, Any] = {}
    for key, _lab, code, _kind in items:
        v = _val(row, code)
        if v is not None:
            out[key] = v
    if ft == "Y-9C":
        d = derive_y9c(out, row.get("RSSD9999", ""))
        out.update(d)
    elif ft == "Y-9SP":
        e, b = out.get("parent_equity"), out.get("investment_in_bank_subs")
        if e:
            if b is not None:
                out["double_leverage"] = round(b / e, 3)
            if out.get("parent_lt_borrowings") is not None:
                out["parent_debt_to_equity"] = round(out["parent_lt_borrowings"] / e, 3)
    return out


def derive_y9c(h: dict[str, Any], period: str) -> dict[str, float]:
    d: dict[str, float] = {}
    months = int(period[4:6]) if len(period) == 8 else 12
    ann = 12 / months if months else 1
    dep = (h.get("deposits_noninterest") or 0) + (h.get("deposits_interest") or 0)
    if dep:
        d["total_deposits"] = dep
        if h.get("loans") is not None:
            d["loans_to_deposits"] = round(h["loans"] / dep, 5)
        if h.get("deposits_noninterest") is not None:
            d["noninterest_deposit_share"] = round(h["deposits_noninterest"] / dep, 5)
    rev = (h.get("nii") or 0) + (h.get("nonint_income") or 0)
    if rev and h.get("nonint_expense") is not None:
        d["efficiency_ratio"] = round(h["nonint_expense"] / rev, 5)
    if h.get("loans"):
        if h.get("allowance") is not None:
            d["allowance_to_loans"] = round(h["allowance"] / h["loans"], 5)
        npl = (h.get("nonaccrual") or 0) + (h.get("past_due_90") or 0)
        if h.get("nonaccrual") is not None:
            d["npl_to_loans"] = round(npl / h["loans"], 5)
        if h.get("chargeoffs") is not None and h.get("recoveries") is not None:
            d["nco_annualized_to_loans"] = round((h["chargeoffs"] - h["recoveries"]) * ann / h["loans"], 5)
    if h.get("net_income") is not None:
        if h.get("avg_assets_qtr"):
            d["roa_annualized"] = round(h["net_income"] * ann / h["avg_assets_qtr"], 5)
        if h.get("total_equity"):
            d["roe_annualized_on_period_end_equity"] = round(h["net_income"] * ann / h["total_equity"], 5)
    if h.get("total_equity") is not None and h.get("goodwill") is not None and h.get("total_assets"):
        d["equity_to_assets"] = round(h["total_equity"] / h["total_assets"], 5)
    return d


def flow_keys(ft: str) -> list[str]:
    items = Y9C_ITEMS if ft == "Y-9C" else Y9SP_ITEMS if ft == "Y-9SP" else Y9LP_ITEMS
    return [k for k, _l, _c, kind in items if kind == "flow"]


def labels_for(ft: str) -> dict[str, str]:
    items = Y9C_ITEMS if ft == "Y-9C" else Y9SP_ITEMS if ft == "Y-9SP" else Y9LP_ITEMS
    return {k: lab for k, lab, _c, _kind in items}


def resolve(store: Store, query: str, period: str | None = None) -> tuple[str | None, list[dict]]:
    """Return (rssd, candidates). A numeric query is an RSSD; otherwise a name search that must be unambiguous."""
    q = query.strip()
    periods = store.periods()
    if not periods:
        raise FileNotFoundError(f"No BHCF files in {store.folder}")
    p = period or periods[-1]
    if re.fullmatch(r"\d{3,10}", q):
        return q, []
    hits = search(store, q, None, p, 10000)
    exact = [h for h in hits if (h["name"] or "").upper() == q.upper()]
    if len(exact) == 1:
        return exact[0]["rssd"], []
    if len(hits) == 1:
        return hits[0]["rssd"], []
    return None, hits[:10]


def search(store: Store, query: str, state: str | None, period: str, limit: int) -> list[dict]:
    tokens = [t for t in re.split(r"\W+", query.upper()) if t]
    out = []
    for rssd in store.all_rssd(period):
        row = store.row(rssd, period)
        if not row:
            continue
        name = (row.get("RSSD9017") or "").upper()
        if tokens and not all(t in name for t in tokens):
            continue
        if state and (row.get("RSSD9200") or "").upper() != state.upper():
            continue
        e = entity(row)
        e["assets"] = _num(row.get("BHCK2170")) if e["filer"] == "Y-9C" else _num(row.get("BHSP8519"))
        out.append(e)
    out.sort(key=lambda x: -(x["assets"] or 0))
    return out[:limit]


def expected_latest(today: date | None = None) -> str:
    """Most recent quarter end that is at least 45 days old (the Y-9 filing deadline)."""
    today = today or date.today()
    for y in (today.year, today.year - 1):
        for md in ((12, 31), (9, 30), (6, 30), (3, 31)):
            qe = date(y, *md)
            if qe + timedelta(days=45) <= today:
                return qe.strftime("%Y%m%d")
    return ""


def import_zip(src: Path, folder: Path) -> dict:
    src = Path(src)
    m = FILE_RE.match(src.name)
    if not m:
        raise ValueError("File name must look like BHCF20260630.ZIP")
    z = zipfile.ZipFile(src)
    name = next((n for n in z.namelist() if n.lower().endswith(".txt")), None)
    if not name:
        raise ValueError("Zip has no .txt data file")
    with z.open(name) as f:
        first = f.readline().decode("latin-1")
    if not first.startswith("RSSD9001^"):
        raise ValueError("Unexpected header; this does not look like a BHCF file")
    folder.mkdir(parents=True, exist_ok=True)
    dest = folder / f"BHCF{m.group(1)}.ZIP"
    if src.resolve() != dest.resolve():
        shutil.copy2(src, dest)
    return {"period": fmt_period(m.group(1)), "saved_to": str(dest), "size": dest.stat().st_size}
