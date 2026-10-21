"""Dev helper: list non-empty items for one RSSD in a BHCF zip with MDRM labels. Usage: python -I explore_bhcf.py <zip> <rssd> [regex]"""
import zipfile, gzip, json, sys, re, os
z, rssd = sys.argv[1], sys.argv[2]
pat = re.compile(sys.argv[3], re.I) if len(sys.argv) > 3 else None
labels = json.load(gzip.open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src", "bhc_labels.json.gz"), "rt"))
zf = zipfile.ZipFile(z)
lines = zf.read(zf.namelist()[0]).decode("latin-1").splitlines()
hdr = lines[0].split("^")
for ln in lines[1:]:
    f = ln.split("^")
    if f[0] == rssd:
        for h, v in zip(hdr, f):
            if v.strip() and (not pat or pat.search(labels.get(h, "") + " " + h)):
                print(h, "|", labels.get(h, ""), "|", v)
        break
