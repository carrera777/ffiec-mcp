"""Build src/bhc_labels.json.gz (holding company MDRM names) from the Fed MDRM.zip.
Usage: python -I build_bhc_labels.py <MDRM_CSV.csv> <out.json.gz>
Keeps BHC-series mnemonics (BHCK, BHCP, BHSP, BHCA, BHCB, BHCT, BHCM, BHDM, BHFN, BHOD, BHCS, BHBC, BHCX ...) and RSSD ids."""
import csv, gzip, json, sys, collections

src, out = sys.argv[1], sys.argv[2]
best = {}
with open(src, encoding="latin-1", newline="") as f:
    text = f.read().splitlines()
# first line may be a title row; find the header
start = next(i for i, ln in enumerate(text[:5]) if "Mnemonic" in ln)
for row in csv.DictReader(text[start:]):
    mn = (row.get("Mnemonic") or "").strip()
    ic = (row.get("Item Code") or "").strip()
    if not mn or not ic:
        continue
    if not (mn.startswith("BH") or mn.startswith("FH") or mn == "RSSD"):
        continue
    code = mn + ic
    end = (row.get("End Date") or "").strip()
    name = " ".join((row.get("Item Name") or "").split())
    if code not in best or end > best[code][1]:
        best[code] = (name, end)
labels = {k: v[0] for k, v in sorted(best.items())}
with gzip.open(out, "wt", encoding="utf-8") as g:
    json.dump(labels, g, separators=(",", ":"))
print(len(labels), collections.Counter(k[:4] for k in labels).most_common(14))
for k in ("BHCK2170", "BHCK4340", "BHCK3210", "BHCKA223", "BHCK4074", "BHCK4079", "BHCP2170", "BHSP2170", "BHCK2122", "BHCK2200", "RSSD9017"):
    print(k, labels.get(k))
