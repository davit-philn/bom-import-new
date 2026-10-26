"""Kiểm thử: nạp file mẫu -> xuất template -> nạp lại phải ra cùng dữ liệu nhập; (tùy chọn) LibreOffice tính lại.

Chạy:  python3 tests/test_roundtrip.py
"""
import csv
import glob
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import exporter  # noqa: E402
import importer  # noqa: E402

SAMPLES = sorted(glob.glob(str(ROOT / "tests" / "samples" / "*.xlsm")))


def norm(state):
    """Chỉ giữ dữ liệu nhập; bỏ khác biệt vô hại (khoảng trắng, None)."""
    def clean(x):
        if isinstance(x, dict):
            return {k: clean(v) for k, v in x.items() if v not in (None, "", [])}
        if isinstance(x, list):
            return [clean(v) for v in x]
        return x
    s = clean(state)
    s.pop("packaging", None)
    return s


def main():
    failed = 0
    for f in SAMPLES:
        name = Path(f).name[:8]
        src = importer.import_xlsm(Path(f).read_bytes())
        out = exporter.export_xlsm(src)
        back = importer.import_xlsm(out)
        a, b = norm(src), norm(back)
        for key in sorted(set(a) | set(b)):
            if a.get(key) != b.get(key):
                failed += 1
                print("LỆCH", name, key)
                if isinstance(a.get(key), list):
                    for i, (x, y) in enumerate(zip(a[key], b.get(key, []))):
                        if x != y:
                            print("   dòng", i + 1, x, "!=", y)
                            break
                else:
                    print("   ", a.get(key), "!=", b.get(key))
        print("OK" if not failed else "…", name, "round-trip", {k: len(v) for k, v in src.items() if isinstance(v, list)})
        # LibreOffice: tính lại các cột công thức Phần II và so với web
        if shutil.which("soffice"):
            tmp = Path(tempfile.mkdtemp())
            (tmp / "x.xlsm").write_bytes(out)
            subprocess.run(["soffice", "--headless", "--convert-to",
                            "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1",
                            "--outdir", str(tmp), str(tmp / "x.xlsm")], capture_output=True, timeout=240)
            rows = list(csv.reader(open(tmp / "x-BOM_Phan II.csv", encoding="utf-8")))
            print("   LO tính lại Phần II:", len(rows), "dòng; ô A13 =", rows[12][0], "| E6 =", rows[5][4])
            shutil.rmtree(tmp, ignore_errors=True)
    print("KẾT QUẢ:", "ĐẠT" if not failed else "%d LỆCH" % failed)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
