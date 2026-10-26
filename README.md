# Khai BOM (BM02-QT03-NM) trên web

Nhập BOM theo đúng biểu mẫu Excel đang dùng (Phần I–V), xuất ra file `.xlsm` giống mẫu.

## Chạy
```bash
pip install -r requirements.txt
python3 server.py 8000      # mở http://localhost:8000
```

## Cách dùng
- **Phần I**: thông tin sản phẩm (nhập một lần), mã hoàn thiện, lịch sử thay đổi, ký duyệt.
- **Phần II**: lưới nhập chi tiết gỗ/tấm (A), foam (B), vải/da (C). Dán nhiều dòng từ Excel được.
  Cột xám tự tính (ĐVT, khối lượng tinh, chỉ cạnh, diện tích sơn, khối lượng).
- **Phần III**: chỉ cạnh tự tổng hợp; kim loại, vật tư phụ, phụ kiện, thầu phụ nhập tay (gõ mã `NVL…` sẽ tự điền tên/ĐVT).
- **Phần IV/V**: xem trước kiện hàng (kể cả xếp lớp giao rời) và vật tư đóng gói.
- **Mở file Excel…** nạp BOM `.xlsm` cũ vào web; **Xuất Excel** điền dữ liệu vào `templates/BOM_TEMPLATE.xlsm`
  (giữ nguyên macro, hình, pivot, công thức; Excel tự tính lại khi mở).

## Cấu trúc
| File | Vai trò |
|---|---|
| `static/calc.js` | công thức biểu mẫu (port từ Excel + macro xếp lớp) |
| `static/app.js` | giao diện |
| `exporter.py` | điền dữ liệu vào template `.xlsm` (sửa XML trực tiếp) |
| `importer.py` | đọc `.xlsm` hiện có thành dữ liệu web |
| `server.py` | máy chủ HTTP nhỏ |

## Kiểm thử
```bash
node tests/test_calc.js          # công thức web = giá trị Excel trong 3 file mẫu
python3 tests/test_roundtrip.py  # nạp -> xuất -> nạp lại phải khớp
```

## Giới hạn đã biết
- Template là bản `TB-404`. Hai file mẫu cũ (DC113, DH115b) cộng phụ trội chỉ cạnh khác (+40mm/cạnh) nên số "Chỉ cạnh" lệch khi nạp vào template này.
- Phần IV/V trong Excel dùng hàm `LET/FILTER/VSTACK` (Excel 365). Web chỉ xem trước; bảng tính cuối do Excel làm khi mở file.
  File xuất chưa chứa giá trị tính sẵn của công thức: **mở bằng Excel và lưu lại** trước khi dùng tool import BRAVO đọc file.
- Kết quả xếp lớp "giao rời" chưa được ghi ngược vào file Excel (chạy macro `ToiUu_GiaoRoi_CotB` trong Excel nếu cần).
- Chưa import trực tiếp vào BRAVO (bước sau).
