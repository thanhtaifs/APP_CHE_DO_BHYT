# Phần mềm nghiệp vụ BHYT - Giám định XML1

Giao diện dạng **menu hiện đại (sidebar)** bên trái, nội dung từng chức năng hiển
thị bên phải, được thiết kế theo **Modern Enterprise Design System** (dùng skill
`pyqt6-ui-designer`): nền sáng, sidebar màu xám-xanh nhạt với thanh chỉ báo mục
đang chọn bên trái, thẻ nội dung (card) bo góc, màu chủ đạo xanh dương đậm
**#003d9b**. Gồm 6 chức năng:

1. 🔎 Tra cứu dữ liệu XML1 từ CSDL SQLite
2. 🗂 Tách file Excel theo MA_CSKCB
3. 📊 Tổng hợp trừ chi phí theo MA_CSKCB & Mã chuyên đề
4. 💾 Lưu hồ sơ đã trừ vào CSDL & Kiểm tra trùng
5. 🛠 Quản lý quy tắc giám định
6. 🚨 Kiểm tra hồ sơ theo quy tắc

## 1. Cài đặt

```
pip install -r requirements.txt
```

Giữ 3 file sau **cùng 1 thư mục**: `app_tra_cuu_xml1.py`, `requirements.txt`,
và `DB.sqlite` (CSDL dùng chung, đã tạo sẵn 2 bảng rỗng — xem mục 5 và 6/7). Nếu
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
gốc của bạn, ví dụ "xml123...", khác với `DB.sqlite` dùng chung của phần mềm)
để lấy đầy đủ các cột:

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
nhiêu **hồ sơ** và bao nhiêu **chi phí**, đồng thời có thể loại trừ một số dòng
theo ý muốn (không tính vào kết quả trừ).

Phần mềm tự quy đổi **Loại hồ sơ** (Ngoại trú/Nội trú) từ mã loại KCB:

| Mã   | Loại hồ sơ  |
|------|-------------|
| 1, 2 | Ngoại trú   |
| 3, 9 | Nội trú     |

### Cách dùng

1. Chọn file Excel → "Đọc cột của file" → chọn cột XML1_ID, MA_CSKCB, Mã
   chuyên đề, mã loại KCB (tuỳ chọn), T_BHTT.
2. **Cột đánh dấu KHÔNG trừ** (tuỳ chọn): nếu file của bạn đã có sẵn 1 cột
   dùng để đánh dấu các dòng không muốn trừ (ghi bất kỳ ký tự nào, để trống
   các dòng còn lại), chọn đúng cột đó. Nếu để "-- Không dùng --", toàn bộ dữ
   liệu sẽ được tính là bị trừ.
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

## 6. Chức năng "Quản lý quy tắc giám định"

**Mục đích:** định nghĩa quy tắc "Mã chi phí (MA_CP) chỉ được chỉ định khi chẩn
đoán bệnh thuộc 1 danh sách mã bệnh (ICD) cho phép", theo từng Mã chuyên đề, áp
dụng cho **nhiều chuyên đề khác nhau cùng lúc**. Quy tắc được dùng ở chức năng
mục 7 để rà soát hồ sơ.

**Ví dụ:** MA_CP `40.67` chỉ được dùng khi `MA_BENH` hoặc 1 trong các
`MA_BENH_KHAC` (nhiều mã cách nhau bởi dấu `;`) thuộc nhóm **BONG** (Bỏng
nắng): `L55, L55.0, L55.1, L55.2, L55.8`. Hồ sơ dùng MA_CP `40.67` nhưng không
có mã bệnh nào trong nhóm này → bị coi là chỉ định sử dụng SAI quy định.

### 1. Nạp danh mục quy tắc bằng file Excel

Cách duy nhất để thêm quy tắc là **import file Excel** đúng cấu trúc cột quy
định, mỗi dòng là 1 mã bệnh được phép cho 1 tổ hợp (Mã chuyên đề, MA_CP):

| Cột | Bắt buộc | Ý nghĩa |
|---|---|---|
| `MA_CHUYEN_DE` | Có | Mã chuyên đề áp dụng quy tắc |
| `TEN_CHUYEN_DE` | Không | Tên chuyên đề (để hiển thị) |
| `MA_CP` | Có | Mã chi phí áp dụng quy tắc |
| `TEN_CP` | Không | Tên chi phí (để hiển thị) |
| `MA_BENH` | Có | 1 mã bệnh (ICD) được phép chỉ định |
| `TEN_BENH` | Không | Tên bệnh (để hiển thị) |
| `NHOM_BENH` | Không | Nhóm bệnh (ví dụ `BONG`) |

Chọn file → "Đọc cột của file" → khai báo cột (kể cả khi tên cột trong file
khác với tên chuẩn ở trên) → bấm **"Nạp / Cập nhật danh mục quy tắc vào
CSDL"**. Quy tắc trùng khoá (Mã chuyên đề + MA_CP + MA_BENH) sẽ tự động **cập
nhật** (ghi đè Tên bệnh/Nhóm bệnh mới) thay vì tạo bản ghi trùng — nạp lại
file đã sửa để cập nhật quy tắc, không cần thao tác gì thêm. Có thể nạp nhiều
file khác nhau cho **nhiều chuyên đề khác nhau**.

### 2. Danh sách quy tắc hiện có (xem / xoá)

Bảng hiển thị toàn bộ quy tắc đang có trong CSDL. Bấm "Tải lại danh sách" để
làm mới; chọn 1 dòng rồi bấm **"Xoá dòng đã chọn"** để xoá hẳn quy tắc đó khỏi
CSDL (có hỏi xác nhận). Muốn sửa nội dung 1 quy tắc, chỉnh lại trong file Excel
rồi nạp lại (mục 1) — quy tắc trùng khoá sẽ tự cập nhật.

## 7. Chức năng "Kiểm tra hồ sơ theo quy tắc"

Dùng quy tắc đã định nghĩa ở mục 6 để rà soát hàng loạt hồ sơ, tự động tìm ra
các dòng chỉ định SAI, kèm mã và nội dung lý do từ chối.

### Nguồn dữ liệu: chỉ đọc từ CSDL SQLite

Chức năng này **chỉ đọc dữ liệu từ 1 CSDL SQLite ngoài** (ví dụ
`xml123.sqlite` — CSDL nguồn XML1 của riêng bạn, khác với `DB.sqlite` nội bộ
của phần mềm), không hỗ trợ Excel. Chọn đường dẫn tới file CSDL → chọn bảng dữ
liệu → "Đọc cột dữ liệu" → khai báo cột: `MA_CP`, `MA_BENH`, `MA_BENH_KHAC`
(nhiều mã cách nhau bởi `;`), `MA_CHUYEN_DE` (tuỳ chọn). Tuỳ chọn thêm:

- **Mã lý do từ chối (MA_LY_DO_TC)**: mặc định `CHOT_3`, có thể đổi.
- **Nội dung lý do (LY_DO_TC)**: mẫu câu tự động điền, dùng được các chỗ giữ
  chỗ `{MA_CP} {MA_BENH} {MA_BENH_KHAC} {NHOM_BENH} {DANH_SACH_MA_BENH}`.

### Cách xác định dòng sai: dùng 1 câu lệnh SQL duy nhất

Bấm **"Kiểm tra hồ sơ"**. Thay vì lặp qua từng dòng bằng Python, phần mềm
**gắn (ATTACH DATABASE) CSDL quy tắc `DB.sqlite` vào cùng kết nối** với CSDL
nguồn, rồi chạy 1 câu lệnh SQL duy nhất kiểu:

```sql
SELECT * FROM "<bảng_hồ_sơ>" s
WHERE EXISTS (
    SELECT 1 FROM rulesdb.QUY_TAC_BENH r WHERE r.MA_CP = s.MA_CP [AND r.MA_CHUYEN_DE = s.MA_CHUYEN_DE]
)
AND NOT EXISTS (
    SELECT 1 FROM rulesdb.QUY_TAC_BENH r
    WHERE r.MA_CP = s.MA_CP [AND r.MA_CHUYEN_DE = s.MA_CHUYEN_DE]
      AND (r.MA_BENH = s.MA_BENH OR instr(';'||s.MA_BENH_KHAC||';', ';'||r.MA_BENH||';') > 0)
)
```

Cách này lấy được đúng tập hồ sơ chỉ định sai ngay trong CSDL (nhanh với dữ
liệu lớn, không cần tải hết dữ liệu vào bộ nhớ để lặp), phù hợp với yêu cầu
"lấy dữ liệu theo nguyên tắc bằng SQL". Nếu có khai báo cột Mã chuyên đề, quy
tắc được khớp đúng theo (Mã chuyên đề, MA_CP); nếu không khai báo, quy tắc
được khớp theo MA_CP (gộp từ mọi chuyên đề có quy tắc cho MA_CP đó) — nhờ vậy
hồ sơ thuộc **nhiều chuyên đề khác nhau** vẫn được xử lý đúng trong cùng 1 lần
kiểm tra. Sau khi có tập hồ sơ sai (thường nhỏ hơn nhiều so với toàn bộ dữ
liệu), phần mềm mới dùng Python để dựng nội dung câu `LY_DO_TC` cho từng dòng.

> Phiên bản này bỏ tuỳ chọn "khớp theo tiền tố mã bệnh" (có ở bản trước) để
> giữ câu lệnh SQL đơn giản, rõ ràng — chỉ khớp đúng chính xác mã bệnh đã khai
> báo trong quy tắc.

### Kết xuất — chuẩn hoá đúng cấu trúc cột đầu ra Excel

Bấm **"Xuất dữ liệu SAI ra Excel"** để lưu toàn bộ dòng vi phạm. File xuất ra
**luôn đủ 25 cột theo đúng thứ tự chuẩn** (giống hệt cấu trúc cột của chức
năng "Tra cứu dữ liệu XML1" ở mục 2):

```
XML1_ID, MA_BN, MA_LK, HO_TEN, MA_THE, MA_BENH,
NGAY_VAO, NGAY_RA, LOAI_CP, ID_CP, NGAY_Y_LENH,
MA_CP, TEN_CP, SO_DANG_KY,
SL_DC, DON_GIA_DC, TYLE_TT_DC, MUC_HUONG_DC,
LY_DO_TC, MA_LY_DO_TC,
MA_CSKCB, KY_QT, T_BHTT_DTL,
MA_CHUYEN_DE, CONG_VAN
```

- Cột nào nguồn dữ liệu (Excel hoặc bảng CSDL ngoài) **có sẵn đúng tên** sẽ
  được lấy giá trị tương ứng; cột nào nguồn không có sẽ để **trống**.
- Riêng **`MA_LY_DO_TC`** và **`LY_DO_TC`** luôn được điền bằng giá trị vừa
  tính từ quy tắc (không lấy từ nguồn, kể cả khi nguồn có sẵn 2 cột này).


## 8. Ghi chú kỹ thuật

- Các thao tác đọc CSDL, tra cứu, tách file, tổng hợp, lưu/kiểm tra trùng và
  kiểm tra quy tắc đều chạy trên luồng riêng (QThread) nên giao diện không bị
  treo khi xử lý file lớn.
- Đối chiếu XML1_ID/ID_CP dùng `pandas.merge` (join toàn bộ dữ liệu một lần)
  thay vì truy vấn từng dòng, xử lý nhanh hơn với file lớn.
- **CSDL dùng chung `DB.sqlite`**: gồm 2 bảng — `HO_SO_DA_TRU` (mục 5) và
  `QUY_TAC_BENH` (mục 6, 7). Đường dẫn tính tự động và cố định qua hàm
  `get_app_dir()` (biến `DB_PATH` ở đầu file `app_tra_cuu_xml1.py`): cùng thư
  mục với file `.py` khi chạy script, hoặc cùng thư mục với file `.exe` khi đã
  đóng gói (kiểm tra qua `sys.frozen`). Muốn đổi tên/vị trí file CSDL, chỉnh
  biến `DB_PATH` này — mọi chức năng liên quan sẽ tự dùng theo.
- CSDL nguồn dùng ở mục 2 (Tra cứu XML1) và CSDL ngoài chọn ở mục 7 (Kiểm tra
  hồ sơ theo quy tắc) là dữ liệu **của riêng bạn**, do bạn tự chọn đường dẫn
  mỗi lần — khác với `DB.sqlite` (CSDL nội bộ, cố định, do phần mềm quản lý).
- Mục 7 dùng `sqlite3` `ATTACH DATABASE` để gộp 2 CSDL (nguồn + `DB.sqlite`)
  vào cùng 1 kết nối, cho phép JOIN/EXISTS trực tiếp bằng SQL giữa 2 file
  `.sqlite` khác nhau mà không cần nạp dữ liệu ra ngoài trước.
- Kết quả xuất Excel của mục 7 (dữ liệu chỉ định sai) dùng chung danh sách cột
  `OUTPUT_COLUMNS` với chức năng Tra cứu XML1 (mục 2), để đảm bảo cấu trúc file
  đồng nhất trong toàn phần mềm.
- Muốn đổi danh sách cột đầu ra hoặc mapping mã loại KCB, chỉnh các biến
  `OUTPUT_COLUMNS`, `COLUMNS_NOT_IN_DB` và hàm `classify_loai_ho_so()` ở đầu
  file `app_tra_cuu_xml1.py`.

## 9. Giao diện (Design System)

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
  sáng, font monospace — vẫn giữ phong cách "console" tương phản cao để dễ
  đọc log xử lý dù tổng thể giao diện đã chuyển sang nền sáng.
- **Typography**: tiêu đề dùng font "Hanken Grotesk" (có fallback Segoe UI),
  nội dung dùng "Inter" (có fallback Segoe UI) — nếu máy không cài 2 font này,
  Qt tự thay bằng font hệ thống tương đương, không lỗi.
- Toàn bộ token màu/khoảng cách/bo góc được khai báo dưới dạng hằng số Python
  (`COLOR_*`, `SPACING_*`, `RADIUS_*`) ngay phía trên biến `APP_STYLE` trong
  `app_tra_cuu_xml1.py` — sửa 1 hằng số sẽ áp dụng lại cho toàn bộ giao diện.

