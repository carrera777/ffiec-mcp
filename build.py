"""Package the extension into dist/ffiec.mcpb (a zip of manifest.json and src/)."""
import json
import zipfile
from pathlib import Path

ROOT = Path(__file__).parent
FILES = ["manifest.json", "src/server.py", "src/excel_export.py", "src/mdrm_labels.json.gz", "src/fry9.py", "src/bhc_labels.json.gz"]


def main() -> None:
    version = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))["version"]
    out = ROOT / "dist" / "ffiec.mcpb"
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in FILES:
            z.write(ROOT / f, f)
    print(f"Built {out} (version {version}, {out.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
