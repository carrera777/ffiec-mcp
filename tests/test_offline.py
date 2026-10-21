"""Offline tests: every tool runs against a mocked FFIEC API. No credentials or network needed.

Run: python tests/test_offline.py
"""
import asyncio
import base64
import json
import os
import sys
import tempfile
from pathlib import Path

TMP = Path(tempfile.mkdtemp(prefix="ffiec_test_"))
os.environ.update(FFIEC_USER_ID="tester", FFIEC_TOKEN="Bearer test-token", FFIEC_DOWNLOAD_DIR=str(TMP))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import httpx  # noqa: E402
import server  # noqa: E402

server.MIN_INTERVAL = 0
HEADER = '"Call Date";"Bank RSSD Identifier";"MDRM #";"Value";"Last Update";"Short Definition";"Call Schedule";"Line Number"\n'


def sdf(period: str, ni_ytd: int, nii_ytd: int) -> str:
    d = period[6:] + period[:2] + period[3:5]
    rows = [("RCON2170", 1000000, "Total assets", "RCRII", "11"), ("RCON3300", 1000000, "Total L&E", "RC", "29"),
            ("RCONB528", 600000, "Loans HFI", "RC", "4b"), ("RCON2200", 800000, "Deposits", "RC", "13a"),
            ("RCON3210", 100000, "Total bank equity", "RC", "27a"), ("RIAD4107", nii_ytd + 100, "Total interest income", "RI", "1h"),
            ("RIAD4073", 100, "Total interest expense", "RI", "2e"), ("RIAD4074", nii_ytd, "Net interest income", "RI", "3"),
            ("RIAD4340", ni_ytd, "Net income", "RIA", "4"), ("RIAD4150", 250, "FTE employees", "RI", "M5"),
            ("RCOA7204", "10.5%", "Leverage ratio", "RCRI", "31")]
    return HEADER + "".join(f'{d};480228;"{c}";"{v}";"{d}";"{t}";"{s}";"{l}"\n' for c, v, t, s, l in rows)


XBRL = b'''<xbrl xmlns="http://www.xbrl.org/2003/instance" xmlns:uc="u">
<context id="CI_1_2026-06-30"><entity/><period><instant>2026-06-30</instant></period></context>
<context id="CD_1_2026-06-30"><entity/><period><startDate>2026-01-01</startDate><endDate>2026-06-30</endDate></period></context>
<uc:UBPRE013 contextRef="CI_1_2026-06-30" unitRef="PURE" decimals="2">1.05</uc:UBPRE013>
<uc:RIAD4340 contextRef="CD_1_2026-06-30" unitRef="USD" decimals="0">3000000</uc:RIAD4340>
<uc:RIAD4150 contextRef="CI_1_2026-06-30" unitRef="NON-MONETARY" decimals="0">250</uc:RIAD4150>
<uc:UBPR7204 contextRef="CI_1_2026-06-30" unitRef="PURE" decimals="2">0.11</uc:UBPR7204></xbrl>'''
YTD = {"03/31/2026": (1000, 5000), "06/30/2026": (3000, 10500)}
calls = []


def fake_get(url, headers, timeout, **kw):
    m = url.rsplit("/", 1)[1]
    calls.append((m, headers))
    req = httpx.Request("GET", url)
    J = lambda o: httpx.Response(200, json=o, request=req)  # noqa: E731
    if m in ("RetrieveReportingPeriods", "RetrieveUBPRReportingPeriods"):
        return J(["12/31/2025", "3/31/2026", "6/30/2026"])  # FFIEC does not zero-pad
    if m == "RetrievePanelOfReporters":
        return J([{"ID_RSSD": 480228, "FDICCertNumber": 3510, "Name": "TEST BANK, NATIONAL ASSOCIATION    ",
                   "State": "NC", "City": "CHARLOTTE  ", "FilingType": "041", "HasFiledForReportingPeriod": True}])
    if m == "RetrieveFilersSubmissionDateTime":
        return J([{"ID_RSSD": 1, "DateTime": "7/1/2026 6:17:33 PM"}, {"ID_RSSD": 2, "DateTime": "7/9/2026 11:42:45 AM"}])
    if m == "RetrieveFacsimile":
        p = headers["reportingPeriodEndDate"]
        if headers["fiId"] == "999" or p not in YTD:
            return httpx.Response(204, request=req)
        fmt = headers["facsimileFormat"]
        body = sdf(p, *YTD[p]).encode() if fmt == "SDF" else XBRL if fmt == "XBRL" else b"%PDF-1.4"
        return J(base64.b64encode(body).decode()) if fmt != "PDF" else J(list(body))
    if m == "RetrieveUBPRXBRLFacsimile":
        return J(base64.b64encode(XBRL).decode())
    if m == "RetrieveFilersSinceDate":
        return httpx.Response(500, json={"Message": "Error Code 5002: Security Token is invalid."}, request=req)
    raise AssertionError(m)


server.httpx.get = fake_get
failures = []


def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  {detail}"))
    if not cond:
        failures.append(name)


async def call(name, args):
    r = await server.mcp.call_tool(name, args)
    return r[1].get("result", r[1]) if isinstance(r, tuple) else json.loads(r[0].text)


async def main():
    r = await call("ffiec_reporting_periods", {})
    check("periods sorted, zero-padded", r["latest"] == "06/30/2026" and r["periods"][-1] == "12/31/2025", r)
    r = await call("ffiec_find_bank", {"query": "test bank"})
    check("find bank strips padding", r["results"][0]["City"] == "CHARLOTTE", r)
    r = await call("ffiec_recent_filers", {"period": "2026Q2", "since": "2026-07-01"})
    check("recent filers newest first", r["filers"][0]["ID_RSSD"] == 2, r)
    r = await call("ffiec_recent_filers", {"period": "2026Q2", "since": "07/01/2026", "include_times": False})
    check("token error gets a hint", "Regenerate" in r.get("error", ""), r)
    r = await call("ffiec_call_report", {"fi_id": "480228", "period": "2026Q2", "items": ["2170"]})
    check("call report item filter", r["items"][0]["mdrm"] == "RCON2170", r)
    r = await call("ffiec_call_report", {"fi_id": "480228", "period": "2026Q2", "schedule": "RC-R part II"})
    check("schedule filter ignores hyphens", [i["mdrm"] for i in r["items"]] == ["RCON2170"], r)
    r = await call("ffiec_bank_summary", {"fi_ids": ["480228", "999"], "period": "2026Q2"})
    b = r["banks"]["480228"]
    check("summary parses % ratios", b["tier1_leverage_ratio_pct"] == 10.5, b)
    check("summary reports missing bank", "error" in r["banks"]["999"], r)
    r = await call("ffiec_ubpr", {"fi_id": "480228", "period": "2026Q2", "search": "average assets"})
    check("ubpr names attached and searchable", r["facts"] and r["facts"][0]["concept"] == "UBPRE013", r)
    r = await call("ffiec_lookup_items", {"query": "ROA", "prefix": "UBPR"})
    check("lookup synonyms", any(x["code"] == "UBPRE013" for x in r["results"]), r)
    r = await call("ffiec_lookup_items", {"query": "2170", "limit": 3})
    check("lookup bare item number", r["results"][0]["code"] == "UBPR2170", r)

    out = TMP / "export.xlsx"
    r = await call("ffiec_export_excel", {"fi_ids": ["480228"], "periods": ["2026Q2"], "detail_items": ["RIAD4340", "RIAD4150"],
                                          "ubpr_concepts": ["UBPRE013"], "output_path": str(out)})
    check("export ran", r.get("file") == str(out) and r["helper_quarters"] == ["Q1 2026"], r)
    from openpyxl import load_workbook
    wb = load_workbook(out)
    ws = wb[r["sheets"][1]]
    found = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        label, basis = row[1].value, row[3].value
        if label == "Net income attributable to bank":
            found[basis] = (row[4].value, row[5].value)
    check("YTD net income as filed", found.get("YTD") == (1000, 3000), found)
    q = found.get("Quarter", ("", ""))
    check("discrete Q1 = YTD", str(q[0]).startswith("=IF(ISNUMBER(E"), q)
    check("discrete Q2 = YTD less Q1 YTD", "F" in str(q[1]) and "-E" in str(q[1]), q)
    det = wb[[s for s in wb.sheetnames if s.endswith("Detail")][0]]
    bases = {row[2].value: row[5].value for row in det.iter_rows(min_row=6) if row[2].value}
    check("XBRL basis flags (employees are period end)", bases.get("RIAD4150") == "Period end" and bases.get("RIAD4340") == "YTD", bases)
    check("notes, data, UBPR sheets", {"Notes", "Data", "UBPR"} <= set(wb.sheetnames), wb.sheetnames)

    # resumable export: a zero time budget forces one download per call
    server.TIME_BUDGET = 0
    YTD["12/31/2025"] = (4000, 20000)
    YTD["09/30/2025"] = (3000, 15000)
    args = {"fi_ids": ["480228"], "periods": ["2025Q4"], "output_path": str(TMP / "resume.xlsx")}
    statuses = []
    for _ in range(5):
        r = await call("ffiec_export_excel", args)
        statuses.append(r.get("status"))
        if r.get("status") == "complete":
            break
    check("export resumes across calls", statuses == ["in_progress", "complete"], statuses)
    server.TIME_BUDGET = 40
    r = await call("ffiec_export_excel", {"fi_ids": ["999"], "periods": ["2026Q2"], "output_path": str(TMP / "none.xlsx")})
    check("missing filings do not loop", "error" in r and "No data" in r["error"], r)
    n_before = len(calls)
    await call("ffiec_export_excel", {"fi_ids": ["999"], "periods": ["2026Q2"], "output_path": str(TMP / "none.xlsx")})
    check("missing filings remembered (no repeat API calls)", len(calls) == n_before, len(calls) - n_before)

    r = await call("ffiec_ubpr", {"fi_id": "480228", "period": "2026Q2", "concepts": ["UBPR7204"]})
    check("UBPR fraction ratio flagged", r["facts"] and "fraction" in r["facts"][0].get("unit_note", ""), r)

    docs = server._documents_dir()
    check("dir: unexpanded ${DOCUMENTS} default", server._resolve_dir("${DOCUMENTS}/FFIEC") == docs / "FFIEC")
    check("dir: empty -> Documents/FFIEC", server._resolve_dir("") == docs / "FFIEC")
    check("dir: relative -> Documents/FFIEC", server._resolve_dir("FFIEC") == docs / "FFIEC")
    check("dir: unknown variable -> default", server._resolve_dir("${FOO}/x") == docs / "FFIEC")
    check("dir: absolute kept", server._resolve_dir(str(TMP)) == TMP)

    h = calls[0][1]
    check("auth headers", h["Authentication"] == "Bearer test-token" and h["UserID"] == "tester", h)
    print(f"\n{len(failures)} failure(s)")
    sys.exit(1 if failures else 0)


asyncio.run(main())
