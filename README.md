# Phần mềm nghiệp vụ BHYT - Giám định XML1

Giao diện dạng **menu hiện đại (sidebar)** bên trái, nội dung từng chức năng hiển
thị bên phải, phối màu chủ đạo **#0066a3** (xanh đậm) và **#008acd** (xanh sáng).
Gồm 5 chức năng:

1. 🔎 Tra cứu dữ liệu XML1 từ CSDL SQLite
2. 🗂 Tách file Excel theo MA_CSKCB
3. 📊 Tổng hợp trừ chi phí theo MA_CSKCB & Mã chuyên đề
4. 💾 Lưu hồ sơ đã trừ vào CSDL & Kiểm tra trùng
5. ⚖️ Quy tắc giám định theo chuyên đề

## 1. Cài đặt

```
pip install -r requirements.txt
```

Giữ 3 file sau **cùng 1 thư mục**: `app_tra_cuu_xml1.py`, `requirements.txt`,
và `DB.sqlite` (CSDL dùng chung, đã tạo sẵn 2 bảng rỗng — xem mục 5 và 6). Nếu
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

## 6. Chức năng "Quy tắc giám định theo chuyên đề"

**Mục đích:** định nghĩa quy tắc "Mã chi phí (MA_CP) chỉ được chỉ định khi chẩn
đoán bệnh thuộc 1 danh sách mã bệnh (ICD) cho phép", theo từng Mã chuyên đề, áp
dụng cho **nhiều chuyên đề khác nhau cùng lúc**. Sau đó dùng quy tắc để rà soát
hàng loạt hồ sơ, tự động tìm ra các dòng chỉ định SAI, kèm mã và nội dung lý do
từ chối.

**Ví dụ:** MA_CP `40.67` chỉ được dùng khi `MA_BENH` hoặc 1 trong các
`MA_BENH_KHAC` (nhiều mã cách nhau bởi dấu `;`) thuộc nhóm **BONG** (Bỏng
nắng): `L55, L55.0, L55.1, L55.2, L55.8`. Hồ sơ dùng MA_CP `40.67` nhưng không
có mã bệnh nào trong nhóm này → bị coi là chỉ định sử dụng SAI quy định.

### Bước 1 — Nạp danh mục quy tắc bằng file Excel

Chuẩn bị file Excel, mỗi dòng là 1 mã bệnh được phép cho 1 tổ hợp (Mã chuyên
đề, MA_CP): cột `MA_CHUYEN_DE, TEN_CHUYEN_DE, MA_CP, TEN_CP, MA_BENH,
TEN_BENH, NHOM_BENH`. Chọn file → "Đọc cột của file" → khai báo cột → bấm
**"Nạp / Cập nhật danh mục quy tắc vào CSDL"**.

### Bước 1b — Import nhanh (dán danh sách, không cần file Excel)

Dùng khi bạn có sẵn 1 bảng mã bệnh copy từ Word/Excel như đúng ví dụ bạn đưa:

```
1.        L55       Bỏng nắng            BONG
2.        L55.0     Bỏng nắng độ một     BONG
3.        L55.1     Bỏng nắng độ hai     BONG
4.        L55.2     Bỏng nắng độ ba      BONG
5.        L55.8     Bỏng nắng khác       BONG
```

Nhập **Mã chuyên đề** và **MA_CP** áp dụng chung (ví dụ `40.67`), dán nguyên
văn bảng trên vào ô văn bản, bấm **"Import danh sách đã dán vào CSDL quy
tắc"**. Phần mềm tự tách cột (theo Tab hoặc nhiều khoảng trắng), tự bỏ qua số
thứ tự (`1.`, `2.`...) ở đầu mỗi dòng, và ghi vào CSDL — không cần chuẩn bị
file Excel riêng cho việc này.

> Cả Bước 1 và Bước 1b đều ghi vào cùng 1 bảng **QUY_TAC_BENH** trong
> `DB.sqlite`. Nạp/import trùng khoá (Mã chuyên đề + MA_CP + MA_BENH) sẽ tự
> động CẬP NHẬT thay vì tạo bản ghi trùng, và bạn có thể lặp lại Bước 1b nhiều
> lần cho **nhiều chuyên đề khác nhau** (mỗi lần đổi Mã chuyên đề + MA_CP rồi
> dán danh sách mã bệnh tương ứng).

### Bước 2 — Kiểm tra hồ sơ theo quy tắc

Chọn **nguồn dữ liệu hồ sơ cần kiểm tra**, 1 trong 2 cách:

- **File Excel**: như các chức năng khác.
- **CSDL SQLite ngoài**: chọn đường dẫn tới 1 file CSDL SQLite của riêng bạn
  (ví dụ CSDL nguồn XML1 bạn đang có ở nơi khác) → chọn bảng dữ liệu → phần
  mềm đọc trực tiếp từ bảng đó, không cần xuất ra Excel trước.

Sau đó khai báo cột: `MA_CP`, `MA_BENH`, `MA_BENH_KHAC` (nhiều mã cách nhau bởi
`;`, tự tách), `MA_CHUYEN_DE` (tuỳ chọn). Tuỳ chọn thêm:

- **Khớp theo tiền tố mã bệnh** (mặc định TẮT): bật lên thì mã `L55` trong quy
  tắc sẽ khớp cả `L55.9` trong hồ sơ dù không có sẵn trong danh sách; tắt thì
  chỉ khớp đúng chính xác.
- **Mã lý do từ chối (MA_LY_DO_TC)**: mặc định `CHOT_3`, có thể đổi.
- **Nội dung lý do (LY_DO_TC)**: mẫu câu tự động điền, dùng được các chỗ giữ
  chỗ `{MA_CP} {MA_BENH} {MA_BENH_KHAC} {NHOM_BENH} {DANH_SACH_MA_BENH}`.

Bấm **"Kiểm tra hồ sơ"**. Phần mềm chỉ xét những dòng có MA_CP đã được định
nghĩa quy tắc; nếu có khai báo cột Mã chuyên đề, ưu tiên khớp đúng theo (Mã
chuyên đề, MA_CP), nếu không khớp đúng thì vẫn dùng danh sách mã bệnh cho phép
chung của MA_CP đó (gộp từ mọi chuyên đề có quy tắc cho MA_CP này) — nhờ vậy hồ
sơ thuộc **nhiều chuyên đề khác nhau trong cùng 1 lần kiểm tra** vẫn được xử lý
đúng theo quy tắc riêng của từng chuyên đề.

Bấm **"Xuất dữ liệu SAI ra Excel"** để lưu toàn bộ dòng vi phạm, kèm 2 cột mới
`MA_LY_DO_TC` và `LY_DO_TC` đã tự động điền.

## 7. Ghi chú kỹ thuật

- Các thao tác đọc CSDL, tra cứu, tách file, tổng hợp, lưu/kiểm tra trùng và
  kiểm tra quy tắc đều chạy trên luồng riêng (QThread) nên giao diện không bị
  treo khi xử lý file lớn.
- Đối chiếu XML1_ID/ID_CP dùng `pandas.merge` (join toàn bộ dữ liệu một lần)
  thay vì truy vấn từng dòng, xử lý nhanh hơn với file lớn.
- **CSDL dùng chung `DB.sqlite`**: gồm 2 bảng — `HO_SO_DA_TRU` (mục 5) và
  `QUY_TAC_BENH` (mục 6). Đường dẫn tính tự động và cố định qua hàm
  `get_app_dir()` (biến `DB_PATH` ở đầu file `app_tra_cuu_xml1.py`): cùng thư
  mục với file `.py` khi chạy script, hoặc cùng thư mục với file `.exe` khi đã
  đóng gói (kiểm tra qua `sys.frozen`). Muốn đổi tên/vị trí file CSDL, chỉnh
  biến `DB_PATH` này — cả 2 chức năng sẽ tự dùng theo.
- CSDL nguồn dùng ở mục 2 (Tra cứu XML1) và CSDL ngoài chọn ở mục 6 Bước 2 là
  dữ liệu **của riêng bạn**, do bạn tự chọn đường dẫn mỗi lần — khác với
  `DB.sqlite` (CSDL nội bộ, cố định, do phần mềm quản lý).
- Muốn đổi danh sách cột đầu ra hoặc mapping mã loại KCB, chỉnh các biến
  `OUTPUT_COLUMNS`, `COLUMNS_NOT_IN_DB` và hàm `classify_loai_ho_so()` ở đầu
  file `app_tra_cuu_xml1.py`. Muốn đổi 2 màu chủ đạo, chỉnh biến `APP_STYLE`
  ở cuối file (tìm các mã màu `#0066a3` và `#008acd`).
