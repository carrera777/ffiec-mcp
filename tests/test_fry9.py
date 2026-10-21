"""Offline tests for the Federal Reserve holding company (Y-9) tools. No network or credentials needed.

Run: python tests/test_fry9.py
Uses a small synthetic BHCF file set. If FFIEC_FRY9_REAL points to a folder of real NIC files, it also
checks Glacier Bancorp (Y-9C) and R Corp Financial (Y-9SP) against the 6/30/2026 filings and the Q2 2026
earnings call.
"""
import io
import os
import sys
import tempfile
import zipfile
from pathlib import Path

TMP = Path(tempfile.mkdtemp(prefix="fry9_test_"))
os.environ.update(FFIEC_USER_ID="tester", FFIEC_TOKEN="Bearer test-token", FFIEC_DOWNLOAD_DIR=str(TMP))
FRY_DIR = TMP / "FRY9"
os.environ["FFIEC_FRY9_DIR"] = str(FRY_DIR)
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fry9  # noqa: E402
import server  # noqa: E402

HDR = ["RSSD9001", "RSSD9999", "RSSD9017", "RSSD9130", "RSSD9200", "BHCK2170", "BHCK2122", "BHCK3123", "BHDM6631", "BHDM6636",
       "BHCK3210", "BHCK4074", "BHCK4079", "BHCK4093", "BHCK4230", "BHCK4340", "BHCK4635", "BHCK4605", "BHCK1403", "BHCK1407",
       "BHCK3368", "BHCAP793", "BHSP2170", "BHSP8519", "BHSP3210", "BHSP3239", "BHSP3151", "BHSP4340"]


def make(period: str, rows: list[dict]) -> bytes:
    lines = ["^".join(HDR)]
    for r in rows:
        lines.append("^".join(str(r.get(h, "")) for h in HDR))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(f"BHCF{period}.txt", "\n".join(lines))
    return buf.getvalue()


def c_row(rssd, name, period, ni, nii, ta=1000000):
    return dict(RSSD9001=rssd, RSSD9999=period, RSSD9017=name, RSSD9130="HELENA", RSSD9200="MT", BHCK2170=ta, BHCK2122=700000,
                BHCK3123=8000, BHDM6631=250000, BHDM6636=600000, BHCK3210=90000, BHCK4074=nii, BHCK4079=nii // 10,
                BHCK4093=int(nii * 0.6), BHCK4230=1000, BHCK4340=ni, BHCK4635=300, BHCK4605=100, BHCK1403=2000, BHCK1407=500,
                BHCK3368=ta, BHCAP793=12.5)


def sp_row(rssd, name, period, ni):
    return dict(RSSD9001=rssd, RSSD9999=period, RSSD9017=name, RSSD9130="ROUND ROCK", RSSD9200="TX", BHSP2170=94000, BHSP8519=1000000,
                BHSP3210=64000, BHSP3239=92000, BHSP3151=30000, BHSP4340=ni)


def setup():
    FRY_DIR.mkdir(parents=True, exist_ok=True)
    files = {
        "20250331": [c_row("111", "ALPHA BANCORP, INC.", "20250331", 100, 400)],
        "20250630": [c_row("111", "ALPHA BANCORP, INC.", "20250630", 230, 820), sp_row("222", "SMALL CORP FINANCIAL", "20250630", 3000)],
        "20251231": [c_row("111", "ALPHA BANCORP, INC.", "20251231", 500, 1700), sp_row("222", "SMALL CORP FINANCIAL", "20251231", 5500)],
    }
    for p, rows in files.items():
        (FRY_DIR / f"BHCF{p}.ZIP").write_bytes(make(p, rows))


def check(cond, msg):
    if not cond:
        raise AssertionError(msg)
    print("ok:", msg)


def run():
    setup()
    st = server.fry9_status()
    check(st["periods_loaded"] == ["2025-03-31", "2025-06-30", "2025-12-31"], "status lists loaded periods")

    f = server.fry9_find("alpha")
    check(f["count"] == 1 and f["results"][0]["filer"] == "Y-9C", "find by name returns a Y-9C filer")
    f2 = server.fry9_find("", state="TX")
    check(f2["results"][0]["rssd"] == "222" and f2["results"][0]["filer"] == "Y-9SP" and f2["results"][0]["assets"] == 1000000,
          "find by state returns the Y-9SP filer with consolidated assets")

    s = server.fry9_summary("Alpha Bancorp")
    v = s["items"]
    check(s["form"] == "Y-9C" and v["total_assets"]["value"] == 1000000, "summary by name uses latest Y-9C")
    check(v["total_deposits"]["value"] == 850000 and abs(v["loans_to_deposits"]["value"] - 700000 / 850000) < 1e-3, "deposits sum and loans/deposits")
    check(abs(v["efficiency_ratio"]["value"] - 1020 / 1870) < 1e-3, "efficiency ratio from YTD NII and noninterest items")
    check(abs(v["nco_annualized_to_loans"]["value"] - 200 * 1 / 700000) < 1e-5, "net charge-off ratio annualized (full-year factor 1)")

    t = server.fry9_trend("111")
    d = {r["period"]: r.get("discrete_period", {}) for r in t["periods"]}
    check(d["2025-03-31"]["net_income"] == 100 and d["2025-06-30"]["net_income"] == 130,
          "discrete quarter net income by differencing within the year")
    check("net_income" not in d["2025-12-31"], "Dec quarter omitted (not guessed) when the September file is not loaded")

    sp = server.fry9_trend("Small Corp")
    d2 = {r["period"]: r.get("discrete_period", {}) for r in sp["periods"]}
    check(d2["2025-06-30"]["parent_net_income"] == 3000 and d2["2025-12-31"]["parent_net_income"] == 2500, "semiannual discrete half-year for Y-9SP")
    sps = server.fry9_summary("222")
    check(sps["form"] == "Y-9SP" and sps["items"]["consolidated_assets"]["value"] == 1000000 and "double_leverage" in sps["items"],
          "Y-9SP summary has consolidated assets and parent leverage")

    amb = server.fry9_summary("c")
    check("candidates" in amb or "error" in amb, "ambiguous or empty name returns guidance instead of guessing")

    it = server.fry9_items("111", search="nonaccrual|tier")
    check(it["count"] >= 1 and all("code" in x for x in it["items"]), "items search by name")
    it2 = server.fry9_items("111", codes=["BHCK2170"])
    check(it2["items"][0]["value"] == 1000000, "items by code")

    src = TMP / "drop" / "BHCF20260331.ZIP"
    src.parent.mkdir()
    src.write_bytes(make("20260331", [c_row("111", "ALPHA BANCORP, INC.", "20260331", 120, 450)]))
    imp = server.fry9_import_file(str(src))
    check(imp["period"] == "2026-03-31" and "2026-03-31" in imp["periods_loaded"], "import validates and copies a file")
    bad = server.fry9_import_file(str(TMP / "nothing.zip"))
    check("error" in bad, "import rejects a wrong file name")
    print("synthetic tests passed")


def real():
    d = Path(os.environ.get("FFIEC_FRY9_REAL", ""))
    if not os.environ.get("FFIEC_FRY9_REAL") or not (d / "BHCF20260630.ZIP").exists():
        print("real data not present; skipped")
        return
    os.environ["FFIEC_FRY9_DIR"] = str(d)
    g = server.fry9_summary("Glacier Bancorp")
    gv = g["items"]
    check(g["period"] == "2026-06-30" and g["form"] == "Y-9C", "real: Glacier latest is the 6/30/2026 Y-9C")
    check(gv["total_assets"]["value"] == 31599001 and gv["total_equity"]["value"] == 4313022 and gv["total_deposits"]["value"] == 24654117,
          "real: assets, equity, deposits at 6/30/2026")
    gt = server.fry9_trend("2003975")
    q = {r["period"]: r.get("discrete_period", {}) for r in gt["periods"]}
    check(q["2026-06-30"]["net_income"] == 97862, "real: Q2 2026 net income 97,862 matches the $97.9M on the earnings call")
    check(q["2025-12-31"]["net_income"] == 239028 - 175249, "real: Q4 2025 discrete net income")
    check(len(gt["periods"]) == 6, "real: Glacier has all six quarters")
    r = server.fry9_summary("R Corp Financial")
    check(r["form"] == "Y-9SP" and r["items"]["consolidated_assets"]["value"] == 1033883, "real: R Corp Financial Y-9SP with consolidated assets")
    rt = server.fry9_trend("4478027")
    check([p["period"] for p in rt["periods"]] == ["2025-06-30", "2025-12-31", "2026-06-30"], "real: R Corp files semiannually (3 of 6 files)")
    rd = {p["period"]: p.get("discrete_period", {}) for p in rt["periods"]}
    check(rd["2025-12-31"]["parent_net_income"] == 5533 - 2962, "real: R Corp second-half 2025 parent net income by differencing")
    f = server.fry9_find("", state="TX", limit=5)
    check(f["count"] == 5, "real: state filter returns Texas holding companies")
    print("real-data tests passed")


if __name__ == "__main__":
    run()
    real()
