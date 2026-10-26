"""Tạo data/lists.json (danh sách gợi ý + danh mục vật tư) từ các file BOM mẫu.

Chạy:  python3 tools/build_lists.py tests/samples/*.xlsm
"""
import json
import sys
import warnings
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import importer  # noqa: E402

OUT = Path(__file__).resolve().parent.parent / "data" / "lists.json"
STANDARDS = ["XK_CHUẨN 1", "XK_CHUẨN 2", "XK_CHUẨN 3", "XK_CHUẨN 4-ĐI CONT", "XK_CHUẨN 5-ĐI AIR",
             "XK_CHUẨN 5-AIR PALLET", "TIÊU CHUẨN DỰ ÁN", "TC_CUA", "TC_KH"]
EDGES = ["2d2r", "2d1r", "2d0r", "1d2r", "0d2r", "1d1r", "1d0r", "0d1r"]
FACES = [f"{m}m{d}d{r}r" for m in (2, 1, 0) for d in (2, 1, 0) for r in (2, 1, 0)]


def main(paths):
    materials, edge_codes, catalog = set(), set(), {}
    for p in paths:
        raw = Path(p).read_bytes()
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(p, data_only=True)
        ws = wb["BOM_Phan II"]
        for col, bucket in ((30, materials), (31, edge_codes)):  # AD, AE (danh sách mẫu)
            for r in range(2, 700):
                v = ws.cell(r, col).value
                if isinstance(v, str) and v.strip() and v.strip() not in ("Vật tư điển hình", "Dán cạnh", "_"):
                    bucket.add(v.strip())
        st = importer.import_xlsm(raw)
        for key in ("partsA", "partsB", "partsC"):
            for row in st[key]:
                if row.get("material"):
                    materials.add(row["material"])
                if row.get("edgeCode"):
                    edge_codes.add(row["edgeCode"])
        for rows in st["supplies"].values():
            for row in rows:
                code = row.get("code")
                if code and row.get("name"):
                    catalog.setdefault(code, {k: row[k] for k in ("name", "unit", "spec", "material", "supplier") if k in row})
    data = {
        "materials": sorted(materials, key=str.lower),
        "edgeCodes": sorted(edge_codes, key=str.lower),
        "supplyCatalog": dict(sorted(catalog.items())),
        "edges": EDGES, "faces": FACES, "standards": STANDARDS,
        "delivery": ["HOÀN THIỆN ĐI MODULE", "GIAO RỜI"],
        "structures": ["Module", "Giao rời", "Theo module và giao rời", "Giao rời và Module"],
        "doorTypes": ["CỬA CHỐNG CHÁY", "CỬA THƯỜNG"],
        "units": ["Md", "Sf", "Con", "Cụm", "Hệ", "Thanh", "Tấm", "M2", "M3"],
    }
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print("lists:", {k: len(v) for k, v in data.items()})


if __name__ == "__main__":
    main(sys.argv[1:])
