# Phần mềm nghiệp vụ BHYT - Giám định XML1

Giao diện dạng **menu hiện đại (sidebar)** bên trái, nội dung từng chức năng hiển
thị bên phải, được thiết kế theo **Modern Enterprise Design System** (dùng skill
`pyqt6-ui-designer`): nền sáng, sidebar màu xám-xanh nhạt với thanh chỉ báo mục
đang chọn bên trái, thẻ nội dung (card) bo góc, màu chủ đạo xanh dương đậm
**#003d9b**. Gồm 8 chức năng:

1. 🔎 Tra cứu dữ liệu XML1 từ CSDL SQLite
2. 🗂 Tách file Excel theo MA_CSKCB
3. 📊 Tổng hợp trừ chi phí theo MA_CSKCB & Mã chuyên đề
4. 💾 Lưu hồ sơ đã trừ vào CSDL & Kiểm tra trùng
5. 🧩 Định nghĩa chuyên đề (điều kiện SQL)
6. ▶️ Chạy chuyên đề theo kỳ
7. 📁 Ghép Folder Excel -> Nhiều Sheet
8. 📑 Ghép Nhiều Sheet -> 1 Sheet

## 1. Cài đặt

```
pip install -r requirements.txt
```

Giữ 3 file sau **cùng 1 thư mục**: `app_tra_cuu_xml1.py`, `requirements.txt`,
và `DB.sqlite` (CSDL dùng chung, đã tạo sẵn 2 bảng rỗng — xem mục 4, 5). Nếu
thiếu file `DB.sqlite`, phần mềm sẽ tự tạo lại khi dùng lần đầu, nên không bắt
buộc phải có sẵn.

Chạy chương trình:

```
python app_tra_cuu_xml1.py
```

---

## 2. Chức năng "Tra cứu dữ liệu XML1"

Bạn có sẵn danh sách `XML1_ID` (hoặc `XML1_ID` + `ID_CP`) trong 1 file Excel/CSV,
phần mềm đọc ngược vào CSDL SQLite nguồn **do bạn tự chọn đường dẫn** (dữ liệu
gốc của bạn, khác với `DB.sqlite` dùng chung của phần mềm) để lấy đầy đủ các cột:

```
XML1_ID, MA_BN, MA_LK, HO_TEN, MA_THE, MA_BENH, NGAY_VAO, NGAY_RA, LOAI_CP,
ID_CP, NGAY_Y_LENH, MA_CP, TEN_CP, SO_DANG_KY, SL_DC, DON_GIA_DC, TYLE_TT_DC,
MUC_HUONG_DC, LY_DO_TC, MA_LY_DO_TC, MA_CSKCB, KY_QT, T_BHTT_DTL,
MA_CHUYEN_DE, CONG_VAN
```

- Các cột **SL_DC, DON_GIA_DC, TYLE_TT_DC, MUC_HUONG_DC, LY_DO_TC, MA_LY_DO_TC**
  không có trong CSDL nguồn nên luôn để **trống**, bạn tự nhập tay sau.
- Hai cột **MA_CHUYEN_DE, CONG_VAN** lấy **trực tiếp từ file đầu vào** của bạn
  (không tra cứu CSDL).
- Nếu chỉ có XML1_ID (không có ID_CP): lấy **tất cả các dòng chi phí** của
  XML1_ID đó. Nếu có cả XML1_ID + ID_CP: đối chiếu chính xác từng dòng chi phí.

## 3. Chức năng "Tách Excel theo MA_CSKCB"

Đọc 1 file Excel, dựa vào cột `MA_CSKCB` (hoặc cột khác bạn chọn), tách thành
nhiều file Excel riêng — mỗi giá trị một file — lưu vào thư mục bạn chỉ định.
Dòng có giá trị trống sẽ gom vào file `KHONG_XAC_DINH.xlsx`.

## 4. Chức năng "Tổng hợp trừ chi phí theo MA_CSKCB & Mã chuyên đề"

**Mục đích:** với dữ liệu XML1 có nhiều mã chuyên đề ở mỗi MA_CSKCB, `XML1_ID`
là khóa của 1 hồ sơ (1 hồ sơ có thể có nhiều dòng chi phí), cột `T_BHTT` là số
tiền dùng để trừ. Bạn cần biết: theo từng MA_CSKCB, từng Mã chuyên đề — trừ bao
nhiêu **hồ sơ** và bao nhiêu **chi phí**.

Phần mềm tự quy đổi **Loại hồ sơ** (Ngoại trú/Nội trú) từ mã loại KCB:

| Mã   | Loại hồ sơ  |
|------|-------------|
| 1, 2 | Ngoại trú   |
| 3, 9 | Nội trú     |

### Cách dùng

1. Chọn file Excel → "Đọc cột của file" → chọn cột XML1_ID, MA_CSKCB, Mã
   chuyên đề, mã loại KCB (tuỳ chọn), T_BHTT.
2. **Cột đánh dấu KHÔNG trừ** (tuỳ chọn): ghi bất kỳ ký tự nào vào dòng không
   muốn trừ trong file của bạn, để trống các dòng còn lại, rồi chọn đúng cột
   đó. Nếu để "-- Không dùng --", toàn bộ dữ liệu sẽ được tính là bị trừ.
3. Bấm **"Tính tổng hợp"** → xem tóm tắt kết quả ở khung log.
4. Bấm **"Xuất kết quả tổng hợp ra Excel"** để lưu file 3 sheet:
   - `Tong_Hop_Chuyen_De`: `MA_CSKCB, MA_CHUYEN_DE, SO_HO_SO, TONG_CHI_PHI`.
   - `Tong_Hop_Loai_Ho_So`: Loại hồ sơ tách thành CỘT riêng (dạng rộng), ví dụ
     `SO_HO_SO_NGOAI_TRU, TONG_CHI_PHI_NGOAI_TRU, SO_HO_SO_NOI_TRU,
     TONG_CHI_PHI_NOI_TRU, TONG_SO_HO_SO, TONG_CHI_PHI`.
   - `Chi_Tiet`: dữ liệu chi tiết từng dòng kèm `TRANG_THAI_TRU` (Trừ/Không
     trừ) và `T_BHTT_SO` (số tiền đã chuẩn hoá).

> Muốn lưu các hồ sơ đã trừ vào CSDL để kiểm tra trùng ở đợt sau, dùng file
> `Chi_Tiet` vừa xuất ra ở chức năng mục 5 bên dưới.

## 5. Chức năng "Lưu hồ sơ đã trừ vào CSDL & Kiểm tra trùng"

Trang riêng, tải trực tiếp 1 file Excel (ví dụ sheet `Chi_Tiet` ở mục 4) để lưu
vào bảng **HO_SO_DA_TRU** trong CSDL dùng chung `DB.sqlite`, phục vụ đối chiếu
về sau. Khi có đợt xử lý mới — kể cả với **mã chuyên đề mới** — dùng chức năng
này để kiểm tra xem các hồ sơ có bị **trừ trùng** với các lần trước hay không.

Dữ liệu lưu gồm:
- **5 cột lõi** (khai báo qua combo): `MA_CSKCB, XML1_ID, ID_CP, MA_CHUYEN_DE, CONG_VAN`.
- **21 cột bổ sung** (tự động lấy theo đúng tên cột nếu file có sẵn, KHÔNG cần
  chọn thủ công): `MA_BN, HO_TEN, MA_THE, MA_BENH, MA_BENH_KHAC, NGAY_VAO,
  NGAY_RA, LOAI_CP, MA_CP, TEN_CP, SO_DANG_KY, SL_DC, DON_GIA_DC, TYLE_TT_DC,
  MUC_HUONG_DC, LY_DO_TC, MA_LY_DO_TC, KY_QT, SO_LUONG_DN, DON_GIA_DN,
  TYLE_TT_DN, MUC_HUONG_DN`.
- Thời điểm lưu `NGAY_LUU`.

Khoá kiểm tra trùng: **(XML1_ID, ID_CP, Mã chuyên đề)**.

**Cách dùng:** khai báo cột (mục 1) → mục 2 chỉ hiển thị đường dẫn CSDL cố
định (không cần chọn) → **"Kiểm tra trùng trong CSDL"** để xem trước hồ sơ nào
đã từng lưu → **"Lưu vào CSDL"** để ghi các dòng mới (dòng trùng khoá tự động
bị bỏ qua, không tạo bản ghi trùng dù bấm lưu nhiều lần).

## 6. Định nghĩa chuyên đề & Chạy theo kỳ (điều kiện SQL)

**Đây là hệ quy tắc giám định DUY NHẤT của phần mềm** (thay cho hệ quy tắc
theo danh sách mã bệnh ở các bản trước — đã gộp lại vì cả hai đều quy về việc
đặt 1 điều kiện SQL dựa trên các cột dữ liệu, nên không cần 2 hệ song song).
Phù hợp với mọi loại chuyên đề giám định, kể cả những chuyên đề không liên
quan đến mã bệnh — ví dụ chuyên đề "kế thừa thuốc vượt số lượng thực kê":

```sql
SO_LUONG_BV > 2 AND MA_CP in (
    '05C.224.8','05C.224.121','05C.222.18','05C.223','05C.158.3',
    'HD.224.121','HD.224.6','HD.158.3','HD.222.18','HD.150'
)
```

Gồm 2 trang: **định nghĩa** chuyên đề (lưu điều kiện SQL) và **chạy** chuyên đề
đó theo từng kỳ.

### 6a. Trang "🧩 Định nghĩa chuyên đề (SQL)"

Mỗi chuyên đề gồm:

| Trường | Bắt buộc | Ý nghĩa |
|---|---|---|
| `MA_CHUYEN_DE` | Có | Mã định danh chuyên đề (khoá, ví dụ `KETHUA_THUOC`) |
| `TEN_CHUYEN_DE` | Không | Tên hiển thị |
| `TEN_SHEET` | Không | Tên sheet khi xuất Excel (mặc định dùng MA_CHUYEN_DE) |
| `DANH_SACH_COT` | Không | Danh sách cột cần lấy, cách nhau dấu phẩy; để trống = lấy tất cả (`*`) |
| `NOI_DUNG_CANH_BAO` | Không | Text sẽ tự điền vào cột `NOI_DUNG_CANH_BAO` của kết quả |
| `DIEU_KIEN_SQL` | Có | Điều kiện SQL (phần sau `WHERE ... AND (`), **không cần viết điều kiện KY_QT** — phần mềm tự thêm khi chạy |

**Cách nhập:** điền form (có khung nhập nhiều dòng cho `DIEU_KIEN_SQL`) → bấm
"Lưu chuyên đề". Lưu trùng `MA_CHUYEN_DE` sẽ tự động cập nhật (không tạo bản
ghi trùng). Bảng "Danh sách chuyên đề hiện có" cho xem/xoá.

> **Chỉ nhập tay từng chuyên đề** — không hỗ trợ nạp hàng loạt bằng Excel
> (phiên bản trước có, nhưng do số dòng cấu hình mỗi chuyên đề thường không
> nhiều nên đã bỏ để giao diện gọn hơn).
>
> Lưu vào bảng `DINH_NGHIA_CHUYEN_DE` trong `DB.sqlite` (CSDL dùng chung).

### 6b. Trang "▶️ Chạy chuyên đề theo kỳ"

Chạy các chuyên đề đã định nghĩa ở trang 6a trên dữ liệu thật, **chỉ cần chọn
chuyên đề và kỳ (KY_QT/tháng)** — không cần viết lại script Python cho từng
chuyên đề.

**Cách dùng:**

1. **Nguồn dữ liệu**: chọn file CSDL SQLite (ví dụ `xml123.sqlite`) → chọn
   bảng → bấm "Đọc cột & tải danh sách kỳ". Phần mềm tự đọc danh sách cột và
   truy vấn `SELECT DISTINCT` trên cột `KY_QT` để đổ vào ô "Kỳ quyết toán /
   Tháng" dạng chọn nhanh (vẫn gõ tay được nếu kỳ chưa có trong danh sách).
2. **Chọn chuyên đề**: bấm "Tải lại danh sách chuyên đề", giữ Ctrl để chọn
   nhiều chuyên đề cùng lúc (chạy gộp nhiều chuyên đề trong 1 lần).
3. Bấm **"Chạy"**. Với mỗi chuyên đề, phần mềm build câu lệnh:

   ```sql
   SELECT <DANH_SACH_COT hoặc *> FROM "<bảng>"
   WHERE "<cột KY_QT>" = ? AND (<DIEU_KIEN_SQL của chuyên đề>)
   ```

   (giá trị kỳ được truyền qua tham số `?`, an toàn khỏi SQL injection), sau
   đó tự thêm cột `NOI_DUNG_CANH_BAO` theo nội dung đã khai báo cho chuyên đề
   đó.
4. Bấm **"Xuất kết quả ra Excel"**: 1 file, **mỗi chuyên đề 1 sheet** (đặt tên
   theo `TEN_SHEET`), đúng theo cách tổ chức file kiểu `sheets[sheet_name] =
   DataOutput` mà bạn đang dùng.

## 7. Chức năng "Ghép Folder Excel -> Nhiều Sheet"

**Mục đích:** Khi có 1 thư mục chứa hàng chục / hàng trăm file Excel (`.xlsx`, `.xls`, `.xlsm`), chức năng này đọc tất cả file và gom thành **1 file Excel duy nhất**, trong đó mỗi file (hoặc mỗi sheet) trở thành 1 sheet riêng.

### Cách dùng
1. Bấm **"Chọn thư mục..."** → chọn thư mục chứa các file Excel cần ghép.
2. Chọn **Chế độ lấy Sheet**:
   - *Chỉ lấy Sheet đầu tiên của mỗi file (Tên Sheet = Tên file)*: Thích hợp khi mỗi file là 1 bảng dữ liệu độc lập.
   - *Lấy tất cả Sheet của mỗi file (Tên Sheet = Tên file_Tên sheet)*: Thích hợp khi các file con chứa nhiều sheet nghiệp vụ khác nhau.
3. Chọn đường dẫn file Excel kết quả (.xlsx).
4. Bấm **"Bắt đầu ghép các file Excel"** → Hệ thống xử lý tự động chuẩn hóa tên sheet (bỏ ký tự cấm, cắt ngắn tối đa 31 ký tự, chống trùng tên) và xuất ra file.

## 8. Chức năng "Ghép Nhiều Sheet -> 1 Sheet"

**Mục đích:** Khi có 1 file Excel chứa nhiều Sheet có cùng hoặc khác cấu trúc cột, chức năng này cho phép gộp tất cả các dòng dữ liệu từ các sheet được chọn thành **1 sheet duy nhất** (tự động ghép các cột trùng tên và giữ nguyên các cột riêng).

### Cách dùng
1. Bấm **"Chọn file Excel..."** → Hệ thống tự động đọc và hiển thị toàn bộ danh sách sheet có trong file.
2. Đánh dấu chọn / bỏ chọn các sheet cần ghép (có sẵn nút "Chọn tất cả" và "Bỏ chọn tất cả").
3. Tùy chọn:
   - **Thêm cột ghi rõ tên Sheet nguồn**: Thêm 1 cột (mặc định tên là `TEN_SHEET`) vào đầu mỗi dòng để biết dòng đó xuất phát từ sheet nào.
   - **Bỏ qua các sheet không có dòng dữ liệu nào**: Tự động lọc bỏ các sheet trống.
4. Bấm **"Bắt đầu ghép các sheet"** → Xem trước 500 dòng kết quả trực tiếp trên bảng xem trước.
5. Bấm **"Xuất kết quả ra file Excel..."** để lưu file Excel tổng hợp hoàn chỉnh.

## 9. Ghi chú kỹ thuật

- Các thao tác đọc CSDL, tra cứu, tách file, tổng hợp, lưu/kiểm tra trùng và
  chạy chuyên đề đều chạy trên luồng riêng (QThread) nên giao diện không bị
  treo khi xử lý file lớn.
- Đối chiếu XML1_ID/ID_CP dùng `pandas.merge` (join toàn bộ dữ liệu một lần)
  thay vì truy vấn từng dòng, xử lý nhanh hơn với file lớn.
- **CSDL dùng chung `DB.sqlite`**: gồm 2 bảng — `HO_SO_DA_TRU` (mục 5) và
  `DINH_NGHIA_CHUYEN_DE` (mục 6). Đường dẫn tính tự động và cố định qua hàm
  `get_app_dir()` (biến `DB_PATH` ở đầu file `app_tra_cuu_xml1.py`): cùng thư
  mục với file `.py` khi chạy script, hoặc cùng thư mục với file `.exe` khi đã
  đóng gói (kiểm tra qua `sys.frozen`). Muốn đổi tên/vị trí file CSDL, chỉnh
  biến `DB_PATH` này — mọi chức năng liên quan sẽ tự dùng theo.
- CSDL nguồn dùng ở mục 2 (Tra cứu XML1) và CSDL ngoài chọn ở mục 6 (Chạy
  chuyên đề theo kỳ) là dữ liệu **của riêng bạn**, do bạn tự chọn đường dẫn
  mỗi lần — khác với `DB.sqlite` (CSDL nội bộ, cố định, do phần mềm quản lý).
- Mục 6 build câu lệnh SQL bằng cách nối trực tiếp `DIEU_KIEN_SQL` (do bạn tự
  khai báo, được xem là đáng tin cậy vì do chính người quản trị nhập) vào mệnh
  đề `WHERE`, riêng **giá trị kỳ (KY_QT) luôn được truyền qua tham số `?`**
  (parameterized query) chứ không nối chuỗi trực tiếp, tránh lỗi cú pháp/SQL
  injection từ giá trị chọn ở dropdown.
- Kết quả mục 6 giữ nguyên cột theo `DANH_SACH_COT` hoặc toàn bộ cột của bảng
  nguồn (không ép theo khuôn cột cố định), vì mỗi chuyên đề SQL có thể cần bộ
  cột khác nhau — khác với mục 2 (Tra cứu XML1) vốn luôn xuất đúng 25 cột
  chuẩn `OUTPUT_COLUMNS`.
- Muốn đổi danh sách cột đầu ra hoặc mapping mã loại KCB, chỉnh các biến
  `OUTPUT_COLUMNS`, `COLUMNS_NOT_IN_DB` và hàm `classify_loai_ho_so()` ở đầu
  file `app_tra_cuu_xml1.py`.

## 10. Giao diện (Design System)

Giao diện được viết theo skill **`pyqt6-ui-designer`** (Modern Enterprise
Design System — xem `references/design_tokens.md`, `qss_patterns.md`,
`component_library.md` trong skill gốc):

- **Sidebar**: nền `#e9edff` (COLOR_SURFACE), viền phải 1px, mỗi mục là 1
  `QPushButton` cao 44px với thuộc tính động `active` — khi đang chọn sẽ có
  nền xanh nhạt `#dae2ff` + thanh chỉ báo 3px bên trái màu xanh đậm
  `#003d9b` (COLOR_PRIMARY), khớp đúng pattern "left-anchored 3px active
  indicator" của skill.
- **Màu chủ đạo**: `#003d9b` (Corporate Blue, COLOR_PRIMARY) cho nút chính,
  tiêu đề trang, viền ô nhập khi focus.
- **Card**: mọi `QGroupBox` được style lại làm thẻ nội dung — nền trắng, viền
  `1px #c3c6d6`, bo góc `8px` (RADIUS_LG), tiêu đề màu primary.
- **Khung ghi chú** (`#noteLabel`): nền xanh nhạt `#dae2ff`, chữ xám-xanh, bo
  góc `6px` (RADIUS_MD).
- **Khung log** (`QTextEdit`): nền tối `#1d3054` (COLOR_INVERSE_SURFACE), chữ
  sáng, font monospace. Riêng khung nhập liệu dài (`#sqlInputArea`, ví dụ ô
  nhập Điều kiện SQL) được override lại thành nền sáng để phân biệt với khung
  log.
- **Typography**: tiêu đề dùng font "Hanken Grotesk" (có fallback Segoe UI),
  nội dung dùng "Inter" (có fallback Segoe UI) — nếu máy không cài 2 font này,
  Qt tự thay bằng font hệ thống tương đương, không lỗi.
- Toàn bộ token màu/khoảng cách/bo góc được khai báo dưới dạng hằng số Python
  (`COLOR_*`, `SPACING_*`, `RADIUS_*`) ngay phía trên biến `APP_STYLE` trong
  `app_tra_cuu_xml1.py` — sửa 1 hằng số sẽ áp dụng lại cho toàn bộ giao diện.
