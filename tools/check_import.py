"""Kiểm tra web đọc/ghi được bao nhiêu file BOM thật: nạp -> xuất -> nạp lại, báo lệch.

Chạy:  python tools/check_import.py "D:\\boho\\BOM"      -> bom_check_report.md
"""
import json
import sys
import time
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import exporter  # noqa: E402
import importer  # noqa: E402


def clean(x):
    if isinstance(x, dict):
        return {k: clean(v) for k, v in x.items() if v not in (None, "", [])}
    if isinstance(x, list):
        return [clean(v) for v in x]
    return x


def main(folder):
    files = sorted(p for p in Path(folder).rglob("*.xls*") if not p.name.startswith("~$"))
    lines, bad = [], 0
    for i, p in enumerate(files, 1):
        t0 = time.time()
        try:
            a = clean(importer.import_xlsm(p.read_bytes()))
            b = clean(importer.import_xlsm(exporter.export_xlsm(a)))
            a.pop("packaging", None); b.pop("packaging", None)
            diff = [k for k in set(a) | set(b) if a.get(k) != b.get(k)]
            n = sum(len(a.get(k, [])) for k in ("partsA", "partsB", "partsC"))
            status = "OK" if not diff else "LỆCH " + ",".join(sorted(diff))
        except Exception as e:  # noqa: BLE001
            status = "LỖI " + repr(e)[:150]
            n = 0
            traceback.print_exc(limit=1)
        bad += status != "OK"
        lines.append("- %s | %s / %s | %d chi tiết | %.1fs" % (status, p.parent.name, p.name, n, time.time() - t0))
        print("[%d/%d] %s %s" % (i, len(files), status, p.name))
    head = "# Kiểm tra nạp/xuất BOM\n\nTổng %d file, có vấn đề: %d\n\n" % (len(files), bad)
    Path("bom_check_report.md").write_text(head + "\n".join(sorted(lines, key=lambda s: s.startswith("- OK"))), encoding="utf-8")
    print("\nXong. Gửi file bom_check_report.md cho Claude.")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit('Cách dùng: python tools/check_import.py "đường-dẫn-folder-BOM"')
    main(sys.argv[1])
