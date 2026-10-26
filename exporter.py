"""Điền dữ liệu từ web vào template .xlsm bằng cách sửa trực tiếp XML.

Giữ nguyên macro VBA, hình ảnh, pivot, data validation và công thức của biểu mẫu;
chỉ ghi các ô NHẬP TAY và để Excel tự tính lại khi mở file (fullCalcOnLoad).
"""
import io
import re
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl.formula.translate import Translator

TEMPLATE = Path(__file__).parent / "templates" / "BOM_TEMPLATE.xlsm"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_RNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
Q = lambda t: "{%s}%s" % (NS, t)

SHEET_I, SHEET_II, SHEET_III, SHEET_IV = "BOM_Phan I", "BOM_Phan II", "BOM_Phan III", "BOM_Phan IV"

# Cột nhập tay Phần II (key trong state -> cột Excel)
PART_COLS = {"stt": "C", "name": "D", "thick": "E", "width": "F", "length": "G", "qty": "H",
             "material": "I", "edgeCode": "J", "edges": "K", "finishCode": "L",
             "faces": "M", "primeFaces": "N", "delivery": "O", "pldg": "X"}
PART_NUMERIC = {"thick", "width", "length", "qty", "pldg"}
PART_SECTIONS = [("partsA", "VẬT TƯ GỖ, VẬT TƯ TẤM"), ("partsB", "VẬT TƯ FOAM SOFA"),
                 ("partsC", "VẬT TƯ VẢI, DA, SIMILI")]

SUPPLY_COLS = {"code": "D", "name": "E", "unit": "F", "spec": "G", "material": "H",
               "supplier": "I", "qtyFactory": "J", "qtySite": "K", "qtySub": "L",
               "note": "M", "image": "N"}
SUPPLY_EXTRA = {"thick": "Q", "width": "R", "length": "S", "qty": "T", "pldg": "U"}
SUPPLY_NUMERIC = {"qtyFactory", "qtySite", "qtySub", "thick", "width", "length", "qty"}
SUPPLY_SECTIONS = [("metal", "VẬT TƯ KIM LOẠI"), ("metalWork", "GIA CÔNG KIM LOẠI"),
                   ("aux", "VẬT TƯ PHỤ"), ("accessories", "PHỤ KIỆN"),
                   ("subcon", "THẦU PHỤ GIA CÔNG")]


class ExportError(ValueError):
    pass


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))


class Sheet:
    def __init__(self, xml_bytes):
        self.tree = etree.fromstring(xml_bytes)
        self.data = self.tree.find(Q("sheetData"))
        self.rows = {int(r.get("r")): r for r in self.data.findall(Q("row"))}
        self._shared_expanded = False

    # ---- cấu trúc ----
    def _row(self, n, create=True):
        row = self.rows.get(n)
        if row is None and create:
            row = etree.Element(Q("row"), r=str(n))
            later = [k for k in self.rows if k > n]
            if later:
                self.rows[min(later)].addprevious(row)
            else:
                self.data.append(row)
            self.rows[n] = row
        return row

    def cell(self, ref, create=True):
        col, n = split_ref(ref)
        row = self._row(n, create)
        if row is None:
            return None
        ci = col_index(col)
        for c in row.findall(Q("c")):
            cc, _ = split_ref(c.get("r"))
            if cc == col:
                return c
            if col_index(cc) > ci:
                if not create:
                    return None
                new = etree.Element(Q("c"), r=ref)
                c.addprevious(new)
                return new
        if not create:
            return None
        new = etree.SubElement(row, Q("c"), r=ref)
        return new

    def has_formula(self, ref):
        c = self.cell(ref, create=False)
        return c is not None and c.find(Q("f")) is not None

    def is_array(self, c):
        f = c.find(Q("f"))
        return f is not None and f.get("t") == "array"

    # ---- đọc ----
    def text(self, ref, sst):
        c = self.cell(ref, create=False)
        if c is None:
            return None
        t = c.get("t")
        if t == "inlineStr":
            return "".join(c.itertext())
        v = c.find(Q("v"))
        if v is None or v.text is None:
            return None
        if t == "s":
            return sst[int(v.text)]
        return v.text

    # ---- ghi ----
    def expand_shared_formulas(self):
        """Biến shared formula thành công thức riêng để có thể ghi đè từng ô an toàn."""
        if self._shared_expanded:
            return
        masters = {}
        for c in self.data.iter(Q("c")):
            f = c.find(Q("f"))
            if f is not None and f.get("t") == "shared" and f.text:
                masters[f.get("si")] = (c.get("r"), f.text)
        for c in self.data.iter(Q("c")):
            f = c.find(Q("f"))
            if f is None or f.get("t") != "shared":
                continue
            origin, text = masters[f.get("si")]
            if not f.text:
                f.text = Translator("=" + text, origin=origin).translate_formula(c.get("r"))[1:]
            for a in ("t", "si", "ref"):
                f.attrib.pop(a, None)
        self._shared_expanded = True

    def clear_constant(self, ref):
        """Xóa giá trị hằng; giữ nguyên ô có công thức."""
        c = self.cell(ref, create=False)
        if c is None or c.find(Q("f")) is not None:
            return
        for ch in list(c):
            c.remove(ch)
        c.attrib.pop("t", None)

    def set_value(self, ref, value, force=False):
        """Ghi giá trị (ghi đè công thức thường; bỏ qua array formula)."""
        c = self.cell(ref)
        if self.is_array(c):
            return
        for ch in list(c):
            c.remove(ch)
        c.attrib.pop("t", None)
        if value is None or value == "":
            return
        if isinstance(value, bool):
            value = int(value)
        if isinstance(value, (int, float)):
            etree.SubElement(c, Q("v")).text = repr(value) if isinstance(value, float) else str(value)
            return
        c.set("t", "inlineStr")
        is_ = etree.SubElement(c, Q("is"))
        t = etree.SubElement(is_, Q("t"))
        t.text = str(value)
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def find_row(self, col, pred, start=1, end=100000):
        for n in sorted(self.rows):
            if n < start or n > end:
                continue
            c = self.cell("%s%d" % (col, n), create=False)
            if c is None:
                continue
            v = self._plain(c)
            if v is not None and pred(v):
                return n
        return None

    def _plain(self, c):
        t = c.get("t")
        if t == "inlineStr":
            return "".join(c.itertext())
        v = c.find(Q("v"))
        if v is None or v.text is None:
            return None
        if t == "s":
            return SST[int(v.text)] if SST else v.text
        return v.text

    def tobytes(self):
        return etree.tostring(self.tree, xml_declaration=True, encoding="UTF-8", standalone=True)


SST = []


def _load_sst(raw):
    global SST
    SST = []
    if not raw:
        return
    root = etree.fromstring(raw)
    for si in root.findall(Q("si")):
        SST.append("".join(t.text or "" for t in si.iter(Q("t"))))


def _num(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        f = float(str(v).replace(",", "."))
    except ValueError:
        return v
    return int(f) if f == int(f) else f


def _section_bounds(sheet, names, col="D"):
    """Trả về {tên: (hàng đầu, hàng cuối)} dùng được để ghi dữ liệu (không gồm hàng tiêu đề)."""
    hits = []
    for key, title in names:
        n = sheet.find_row(col, lambda v, t=title: v.strip().upper() == t)
        if n is None:
            raise ExportError("Template thiếu mục '%s' trong %s" % (title, "Phần II/III"))
        hits.append((n, key))
    hits.sort()
    return hits


def _fill_parts(sh, state):
    sh.expand_shared_formulas()
    hits = _section_bounds(sh, PART_SECTIONS)
    total = sh.find_row("D", lambda v: v.strip().lower().startswith("tổng cộng"), start=hits[-1][0])
    if total is None:
        raise ExportError("Template Phần II thiếu dòng 'Tổng cộng'")
    ends = [h[0] - 1 for h in hits[1:]] + [total - 1]
    for (start, key), end in zip(hits, ends):
        rows = state.get(key) or []
        cap = end - start
        if len(rows) > cap:
            raise ExportError("Phần II mục %s có %d dòng, template chỉ chứa %d dòng" % (key, len(rows), cap))
        for i in range(cap):
            r = start + 1 + i
            data = rows[i] if i < len(rows) else None
            for field, col in PART_COLS.items():
                ref = "%s%d" % (col, r)
                if data is None:
                    sh.clear_constant(ref)
                    continue
                v = data.get(field)
                if field in PART_NUMERIC:
                    v = _num(v)
                if v in (None, "") and sh.has_formula(ref):
                    # để nguyên công thức mẫu (vd STT tự sinh)
                    continue
                sh.set_value(ref, v)
            if data is not None and key != "partsA" and data.get("unit"):
                sh.set_value("Q%d" % r, data["unit"])
    return total


def _fill_header(sh, state, total_row):
    h = state.get("header") or {}
    pairs = {"E4": "project", "P4": "order", "E5": "name", "P5": "size", "E6": "code",
             "E7": "muc", "P7": "structure", "E8": "material", "P6": "qty"}
    for ref, key in pairs.items():
        sh.set_value(ref, h.get(key) or None)
    sh.set_value("Y3", _num(h.get("slsp")))
    sh.set_value("Q7", h.get("doorType") or None)
    codes = [c for c in (h.get("finishCodes") or []) if c][:7]
    if sh.has_formula("P8"):
        for i in range(7):
            sh.set_value("V%d" % (2 + i), codes[i] if i < len(codes) else None)
    else:
        sh.set_value("P8", ", ".join(codes) or None)
    sign = state.get("sign") or {}
    date_row = sh.find_row("P", lambda v: v.strip().startswith("Ngày"), start=total_row)
    if date_row:
        sh.set_value("P%d" % date_row, sign.get("dateText") or None)
    k = sh.find_row("G", lambda v: "duyệt" in v.lower(), start=total_row)
    if k:
        sh.set_value("G%d" % (k + 4), sign.get("approver") or None)
        sh.set_value("P%d" % (k + 4), sign.get("author") or None)


def _fill_supplies(sh, state):
    sh.expand_shared_formulas()
    hits = _section_bounds(sh, SUPPLY_SECTIONS)
    last = hits[-1][0]
    end_last = last
    # hàng cuối của mục cuối: hàng liền trước hàng đầu tiên không còn công thức cờ IN (cột P)
    while sh.has_formula("P%d" % (end_last + 1)):
        end_last += 1
    ends = [h[0] - 1 for h in hits[1:]] + [end_last]
    # Phần I-nhóm chỉ cạnh nằm trước mục đầu, do pivot tự tính -> không đụng tới
    for (start, key), end in zip(hits, ends):
        rows = state.get("supplies", {}).get(key) or []
        cap = end - start
        if len(rows) > cap:
            raise ExportError("Phần III mục %s có %d dòng, template chỉ chứa %d dòng" % (key, len(rows), cap))
        cols = dict(SUPPLY_COLS)
        if key == "metalWork":
            cols.update(SUPPLY_EXTRA)
        for i in range(cap):
            r = start + 1 + i
            data = rows[i] if i < len(rows) else None
            for field, col in cols.items():
                ref = "%s%d" % (col, r)
                if data is None:
                    sh.clear_constant(ref)
                    continue
                v = data.get(field)
                if field in SUPPLY_NUMERIC:
                    v = _num(v)
                if v in (None, "") and sh.has_formula(ref):
                    continue
                sh.set_value(ref, v)


def _fill_phieu(sh, state):
    p = state.get("phieu") or {}
    hist = p.get("history") or []
    for i in range(6):
        r = 11 + i
        row = hist[i] if i < len(hist) else {}
        sh.set_value("C%d" % r, row.get("no") or (i + 1 if row else None))
        sh.set_value("D%d" % r, row.get("basis") or None)
        sh.set_value("E%d" % r, row.get("content") or None)
        sh.set_value("G%d" % r, row.get("date") or None)
        sh.set_value("H%d" % r, row.get("by") or None)
    sign = state.get("sign") or {}
    sh.set_value("G19", sign.get("dateText") or None)
    sh.set_value("D24", sign.get("approver") or None)
    sh.set_value("G24", sign.get("author") or None)


def _fill_packaging(sh, state):
    std = (state.get("packaging") or {}).get("standard")
    if std:
        sh.set_value("O13", std)
    # xóa "Phân lớp" mẫu còn sót
    for n in sorted(sh.rows):
        if n >= 13:
            sh.clear_constant("Z%d" % n)


def strip_cached(sheet):
    """Xóa giá trị cache cũ của công thức (và vùng spill) để không còn số liệu mẫu cũ.

    Excel tính lại khi mở (fullCalcOnLoad); viewer không tính lại sẽ thấy ô trống
    thay vì số liệu sai của file mẫu.
    """
    spill = []
    for c in sheet.data.iter(Q("c")):
        f = c.find(Q("f"))
        if f is None:
            continue
        v = c.find(Q("v"))
        if v is not None:
            c.remove(v)
        c.attrib.pop("t", None)
        if f.get("t") == "array" and f.get("ref") and ":" in f.get("ref"):
            spill.append((f.get("ref"), c.get("r")))
    for ref, master in spill:
        a, b = ref.split(":")
        (c1, r1), (c2, r2) = split_ref(a), split_ref(b)
        for r in range(r1, r2 + 1):
            for ci in range(col_index(c1), col_index(c2) + 1):
                col = ""
                n = ci
                while n:
                    n, rem = divmod(n - 1, 26)
                    col = chr(65 + rem) + col
                ref_c = "%s%d" % (col, r)
                if ref_c == master:
                    continue
                cell = sheet.cell(ref_c, create=False)
                if cell is not None and cell.find(Q("f")) is None:
                    for ch in list(cell):
                        cell.remove(ch)
                    cell.attrib.pop("t", None)


def _fill_edge_pivot(sh, state):
    """Điền sẵn danh sách pivot 'Lọc chỉ cạnh' (Q13:Q22); Excel sẽ làm mới lại khi mở file."""
    codes = set()
    for key in ("partsA", "partsB", "partsC"):
        for row in state.get(key) or []:
            if (row.get("edgeCode") or "").strip():
                codes.add(row["edgeCode"].strip())
    ordered = sorted(codes, key=str.lower)
    start = 12  # Q12 = tiêu đề pivot, dữ liệu từ Q13
    for i in range(10):
        ref = "Q%d" % (start + 1 + i)
        if i < len(ordered):
            sh.set_value(ref, ordered[i])
        else:
            sh.clear_constant(ref)
    # phiên bản mẫu có dòng "(blank)" cuối pivot; không ghi thêm


def export_xlsm(state, template=TEMPLATE):
    zin = zipfile.ZipFile(template)
    names = zin.namelist()
    files = {n: zin.read(n) for n in names}
    _load_sst(files.get("xl/sharedStrings.xml"))

    wb = etree.fromstring(files["xl/workbook.xml"])
    rels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
    rid_target = {r.get("Id"): r.get("Target") for r in rels}
    sheet_file = {}
    for s in wb.find(Q("sheets")):
        sheet_file[s.get("name")] = "xl/" + rid_target[s.get("{%s}id" % RNS)]

    sheets = {n: Sheet(files[sheet_file[n]]) for n in (SHEET_I, SHEET_II, SHEET_III, SHEET_IV)}
    total = _fill_parts(sheets[SHEET_II], state)
    _fill_header(sheets[SHEET_II], state, total)
    _fill_supplies(sheets[SHEET_III], state)
    _fill_edge_pivot(sheets[SHEET_III], state)
    _fill_phieu(sheets[SHEET_I], state)
    _fill_packaging(sheets[SHEET_IV], state)
    for name, path in sheet_file.items():
        sh = sheets.get(name) or Sheet(files[path])
        strip_cached(sh)
        files[path] = sh.tobytes()

    # bắt Excel tính lại + làm mới pivot, bỏ calcChain cũ
    calc = wb.find(Q("calcPr"))
    if calc is None:
        calc = etree.SubElement(wb, Q("calcPr"))
    calc.set("fullCalcOnLoad", "1")
    files["xl/workbook.xml"] = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)

    if "xl/calcChain.xml" in files:
        del files["xl/calcChain.xml"]
        for r in list(rels):
            if r.get("Target") == "calcChain.xml":
                rels.remove(r)
        files["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        ct = etree.fromstring(files["[Content_Types].xml"])
        for o in list(ct):
            if o.get("PartName") == "/xl/calcChain.xml":
                ct.remove(o)
        files["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    for n in list(files):
        if n.startswith("xl/pivotCache/pivotCacheDefinition"):
            root = etree.fromstring(files[n])
            root.set("refreshOnLoad", "1")
            files[n] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            if n in files:
                zout.writestr(n, files[n])
    return out.getvalue()
