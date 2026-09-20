# -*- coding: utf-8 -*-
"""
PHẦN MỀM NGHIỆP VỤ BHYT - GIÁM ĐỊNH XML1
Giao diện menu hiện đại (sidebar) với 3 chức năng:
    1. Tra cứu dữ liệu XML1 từ CSDL SQLite
    2. Tách file Excel theo MA_CSKCB
    3. Tổng hợp trừ chi phí theo MA_CSKCB & Mã chuyên đề

Yêu cầu thư viện:
    pip install PyQt6 pandas openpyxl

Chạy:
    python app_tra_cuu_xml1.py
"""

import os
import re
import sys
import sqlite3
import traceback
import unicodedata

import pandas as pd

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt6.QtGui import QFont, QDesktopServices
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QStackedWidget, QListWidget,
    QListWidgetItem, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel,
    QLineEdit, QPushButton, QComboBox, QFileDialog, QTableWidget,
    QTableWidgetItem, QTextEdit, QProgressBar, QMessageBox, QGroupBox,
    QCheckBox, QFrame, QSizePolicy
)

NONE_OPTION = "-- Không dùng --"


def get_app_dir() -> str:
    """
    Thư mục chứa ứng dụng: nếu chạy dạng file .py thì là thư mục chứa file
    này; nếu đã đóng gói thành .exe (PyInstaller...) thì là thư mục chứa file
    .exe. Dùng để đặt CSDL "hồ sơ đã trừ" cố định cạnh chương trình, không cần
    người dùng chọn đường dẫn mỗi lần.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# CSDL SQLite DUY NHẤT cho toàn bộ phần mềm - đường dẫn CỐ ĐỊNH, gắn kèm dự án/gói
# cài đặt. Gồm nhiều bảng: HO_SO_DA_TRU (hồ sơ đã trừ), QUY_TAC_BENH (quy tắc giám
# định theo chuyên đề)...
DB_PATH = os.path.join(get_app_dir(), "DB.sqlite")

# ============================================================
# CẤU HÌNH DANH SÁCH CỘT NGHIỆP VỤ (CHỨC NĂNG TRA CỨU XML1)
# ============================================================

OUTPUT_COLUMNS = [
    "XML1_ID", "MA_BN", "MA_LK", "HO_TEN", "MA_THE", "MA_BENH",
    "NGAY_VAO", "NGAY_RA", "LOAI_CP", "ID_CP", "NGAY_Y_LENH",
    "MA_CP", "TEN_CP", "SO_DANG_KY",
    "SL_DC", "DON_GIA_DC", "TYLE_TT_DC", "MUC_HUONG_DC",
    "LY_DO_TC", "MA_LY_DO_TC",
    "MA_CSKCB", "KY_QT", "T_BHTT_DTL",
    "MA_CHUYEN_DE", "CONG_VAN",
]

COLUMNS_NOT_IN_DB = [
    "SL_DC", "DON_GIA_DC", "TYLE_TT_DC", "MUC_HUONG_DC",
    "LY_DO_TC", "MA_LY_DO_TC",
]

INPUT_ONLY_COLUMNS = ["MA_CHUYEN_DE", "CONG_VAN"]

DB_NEEDED_COLUMNS = [
    c for c in OUTPUT_COLUMNS
    if c not in COLUMNS_NOT_IN_DB and c not in INPUT_ONLY_COLUMNS
]

# --- Cột dùng cho CSDL "hồ sơ đã trừ" (chức năng Lưu hồ sơ đã trừ & Kiểm tra trùng) ---
# Cột lõi: bắt buộc/khai báo qua combo, dùng làm khoá đối chiếu trùng.
DEDUCTED_CORE_COLUMNS = ["MA_CSKCB", "XML1_ID", "ID_CP", "MA_CHUYEN_DE", "CONG_VAN"]
# Cột bổ sung: tự động lấy theo đúng tên cột nếu file có sẵn (không cần chọn thủ công),
# lưu kèm để phục vụ đối chiếu/tra soát dữ liệu trùng lặp cho đầy đủ thông tin hồ sơ.
DEDUCTED_EXTRA_COLUMNS = [
    "MA_BN", "HO_TEN", "MA_THE", "MA_BENH", "MA_BENH_KHAC", "NGAY_VAO", "NGAY_RA",
    "LOAI_CP", "MA_CP", "TEN_CP", "SO_DANG_KY",
    "SL_DC", "DON_GIA_DC", "TYLE_TT_DC", "MUC_HUONG_DC",
    "LY_DO_TC", "MA_LY_DO_TC", "KY_QT",
    "SO_LUONG_DN", "DON_GIA_DN", "TYLE_TT_DN", "MUC_HUONG_DN",
]
DEDUCTED_ALL_COLUMNS = DEDUCTED_CORE_COLUMNS + DEDUCTED_EXTRA_COLUMNS


# ============================================================
# HÀM TIỆN ÍCH DÙNG CHUNG
# ============================================================

def sanitize_filename(name: str) -> str:
    if name is None:
        name = "KHONG_XAC_DINH"
    name = str(name).strip()
    if name == "" or name.lower() == "nan":
        name = "KHONG_XAC_DINH"
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name[:150]


def read_table_any(path: str) -> pd.DataFrame:
    """Đọc file Excel (.xlsx/.xls) hoặc CSV thành DataFrame, dữ liệu dạng chuỗi."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        df = pd.read_excel(path, dtype=str)
    elif ext == ".csv":
        try:
            df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(path, dtype=str, encoding="cp1258")
    else:
        raise ValueError(f"Không hỗ trợ định dạng file: {ext}")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def parse_money(x) -> float:
    """Chuyển giá trị tiền tệ dạng chuỗi (có thể có dấu , hoặc .) sang float."""
    if x is None:
        return 0.0
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return 0.0
    s = s.replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts[-1]) == 3 and len(parts) > 1:
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def is_note_marked(v) -> bool:
    """
    Xác định 1 ô ở cột Ghi chú có được coi là 'đã đánh dấu' (không trừ) hay không.
    Bao quát nhiều kiểu giá trị có thể gặp khi đọc từ Excel: ô trống/NaN, chuỗi
    rỗng, giá trị boolean TRUE/FALSE (kể cả dạng checkbox), số 0/1, và văn bản
    ghi chú tự do (ví dụ 'x', 'sai chỉ định'...).
    """
    if v is None:
        return False
    try:
        if pd.isna(v):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(v, bool):
        return v is True
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "nat", "<na>"):
        return False
    if s.lower() in ("false", "0", "0.0"):
        return False
    return True


def format_money(v) -> str:
    try:
        return f"{float(v):,.0f}".replace(",", ".")
    except (ValueError, TypeError):
        return str(v)


def ascii_slug(text: str) -> str:
    """Chuyển văn bản (có dấu tiếng Việt) thành hậu tố cột an toàn, VIẾT_HOA_CÓ_GẠCH_DƯỚI."""
    if text is None:
        text = ""
    text = str(text).strip()
    if text == "":
        return "CHUA_XAC_DINH"
    norm = unicodedata.normalize("NFKD", text)
    no_marks = "".join(c for c in norm if not unicodedata.combining(c))
    no_marks = no_marks.replace("Đ", "D").replace("đ", "d")
    slug = re.sub(r"[^A-Za-z0-9]+", "_", no_marks).strip("_").upper()
    return slug or "KHAC"


def classify_loai_ho_so(code) -> str:
    """Quy đổi mã loại KCB -> loại hồ sơ: 1,2 = Ngoại trú; 3,9 = Nội trú."""
    if code is None:
        return ""
    s = str(code).strip()
    if s == "" or s.lower() == "nan":
        return ""
    try:
        n = int(float(s))
        s_norm = str(n)
    except (ValueError, TypeError):
        s_norm = s
    if s_norm in ("1", "2"):
        return "Ngoại trú"
    if s_norm in ("3", "9"):
        return "Nội trú"
    return "Khác/Không xác định"


def normalize_code(v) -> str:
    """Chuẩn hoá 1 mã (bệnh, chi phí, chuyên đề...): bỏ khoảng trắng thừa, viết hoa."""
    if v is None:
        return ""
    s = str(v).strip().upper()
    return "" if s == "" or s.lower() == "nan" else s


def split_multi_codes(v) -> list:
    """Tách chuỗi nhiều mã bệnh cách nhau bởi dấu ';' (hoặc xuống dòng) thành danh sách mã đã chuẩn hoá."""
    if v is None:
        return []
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return []
    parts = re.split(r"[;\n]+", s)
    return [normalize_code(p) for p in parts if normalize_code(p) != ""]


def code_matches(row_code: str, allowed_code: str, prefix_match: bool) -> bool:
    """So khớp 1 mã bệnh của hồ sơ với 1 mã bệnh cho phép trong quy tắc."""
    if row_code == allowed_code:
        return True
    if prefix_match and row_code and allowed_code:
        if row_code.startswith(allowed_code) or allowed_code.startswith(row_code):
            return True
    return False


DEFAULT_LY_DO_TC_TEMPLATE = (
    "MA_CP {MA_CP} không phù hợp chẩn đoán (MA_BENH={MA_BENH}; "
    "MA_BENH_KHAC={MA_BENH_KHAC}); chỉ áp dụng nhóm {NHOM_BENH}: {DANH_SACH_MA_BENH}"
)


def ensure_rules_table(conn):
    """Đảm bảo bảng QUY_TAC_BENH tồn tại trong CSDL dùng chung (DB_PATH)."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS QUY_TAC_BENH (
            MA_CHUYEN_DE TEXT,
            TEN_CHUYEN_DE TEXT,
            MA_CP TEXT,
            TEN_CP TEXT,
            MA_BENH TEXT,
            TEN_BENH TEXT,
            NHOM_BENH TEXT,
            UNIQUE(MA_CHUYEN_DE, MA_CP, MA_BENH)
        )
    """)
    conn.commit()


def upsert_rules(conn, rules_df: pd.DataFrame) -> int:
    """
    Ghi (thêm mới hoặc cập nhật nếu trùng khoá Mã chuyên đề+MA_CP+MA_BENH) danh
    sách quy tắc vào bảng QUY_TAC_BENH. Trả về tổng số quy tắc trong CSDL sau khi ghi.
    """
    ensure_rules_table(conn)
    cur = conn.cursor()
    rows = list(rules_df[
        ["MA_CHUYEN_DE", "TEN_CHUYEN_DE", "MA_CP", "TEN_CP", "MA_BENH", "TEN_BENH", "NHOM_BENH"]
    ].itertuples(index=False, name=None))
    cur.executemany(
        "INSERT OR REPLACE INTO QUY_TAC_BENH "
        "(MA_CHUYEN_DE, TEN_CHUYEN_DE, MA_CP, TEN_CP, MA_BENH, TEN_BENH, NHOM_BENH) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        rows
    )
    conn.commit()
    cur.execute("SELECT COUNT(*) FROM QUY_TAC_BENH")
    return cur.fetchone()[0]


_STT_PATTERN = re.compile(r"^\d+[.\)]?$")


def parse_pasted_benh_lines(text: str) -> list:
    """
    Phân tích văn bản dán nhiều dòng (copy từ Word/Excel, các cột cách nhau bởi
    Tab hoặc nhiều khoảng trắng) thành danh sách dict {MA_BENH, TEN_BENH, NHOM_BENH}.
    Tự động bỏ qua cột số thứ tự (STT) nếu có ở đầu dòng.
    Ví dụ dòng hợp lệ: "1.        L55    Bỏng nắng    BONG"
    """
    result = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if "\t" in line:
            tokens = [t.strip() for t in line.split("\t")]
        else:
            tokens = [t.strip() for t in re.split(r"\s{2,}", line)]
        tokens = [t for t in tokens if t != ""]
        if not tokens:
            continue
        if _STT_PATTERN.match(tokens[0]):
            tokens = tokens[1:]
        if not tokens:
            continue
        ma_benh = normalize_code(tokens[0])
        ten_benh = tokens[1] if len(tokens) > 1 else ""
        nhom_benh = normalize_code(tokens[2]) if len(tokens) > 2 else ""
        if ma_benh:
            result.append({"MA_BENH": ma_benh, "TEN_BENH": ten_benh, "NHOM_BENH": nhom_benh})
    return result


def fill_table_widget(table: QTableWidget, df: pd.DataFrame, columns, max_rows=1000):
    table.setColumnCount(len(columns))
    table.setHorizontalHeaderLabels(columns)
    table.setRowCount(min(len(df), max_rows))
    for i, (_, row) in enumerate(df.head(max_rows).iterrows()):
        for j, col in enumerate(columns):
            val = row[col] if col in df.columns else ""
            item = QTableWidgetItem("" if pd.isna(val) else str(val))
            table.setItem(i, j, item)
    table.resizeColumnsToContents()


# ============================================================
# WORKER: TRA CỨU DỮ LIỆU TỪ SQLITE
# ============================================================

class LookupWorker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished_ok = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, db_path, table_name, input_path,
                 col_xml1_id, col_id_cp,
                 col_ma_chuyen_de=None, col_cong_van=None):
        super().__init__()
        self.db_path = db_path
        self.table_name = table_name
        self.input_path = input_path
        self.col_xml1_id = col_xml1_id
        self.col_id_cp = col_id_cp
        self.col_ma_chuyen_de = col_ma_chuyen_de
        self.col_cong_van = col_cong_van

    def run(self):
        try:
            self.log.emit("Đang đọc file dữ liệu đầu vào...")
            input_df = read_table_any(self.input_path)
            self.progress.emit(15)

            if self.col_xml1_id not in input_df.columns:
                raise ValueError(f"Không tìm thấy cột '{self.col_xml1_id}' trong file đầu vào.")

            input_df["_XML1_ID_KEY"] = input_df[self.col_xml1_id].astype(str).str.strip()
            merge_keys = ["_XML1_ID_KEY"]

            use_id_cp = bool(self.col_id_cp) and self.col_id_cp != NONE_OPTION
            if use_id_cp:
                if self.col_id_cp not in input_df.columns:
                    raise ValueError(f"Không tìm thấy cột '{self.col_id_cp}' trong file đầu vào.")
                input_df["_ID_CP_KEY"] = input_df[self.col_id_cp].astype(str).str.strip()
                merge_keys.append("_ID_CP_KEY")

            use_ma_cd = bool(self.col_ma_chuyen_de) and self.col_ma_chuyen_de != NONE_OPTION
            if use_ma_cd:
                if self.col_ma_chuyen_de not in input_df.columns:
                    raise ValueError(f"Không tìm thấy cột '{self.col_ma_chuyen_de}' trong file đầu vào.")
                input_df["MA_CHUYEN_DE"] = input_df[self.col_ma_chuyen_de]
            else:
                input_df["MA_CHUYEN_DE"] = ""

            use_cong_van = bool(self.col_cong_van) and self.col_cong_van != NONE_OPTION
            if use_cong_van:
                if self.col_cong_van not in input_df.columns:
                    raise ValueError(f"Không tìm thấy cột '{self.col_cong_van}' trong file đầu vào.")
                input_df["CONG_VAN"] = input_df[self.col_cong_van]
            else:
                input_df["CONG_VAN"] = ""

            self.log.emit(f"Đã đọc {len(input_df)} dòng từ file đầu vào.")
            self.log.emit("Đang kết nối CSDL SQLite và kiểm tra cấu trúc bảng...")

            conn = sqlite3.connect(self.db_path)
            try:
                cur = conn.cursor()
                cur.execute(f'PRAGMA table_info("{self.table_name}")')
                table_cols = [row[1] for row in cur.fetchall()]

                cols_available = [c for c in DB_NEEDED_COLUMNS if c in table_cols]
                cols_missing_in_db = [c for c in DB_NEEDED_COLUMNS if c not in table_cols]
                if cols_missing_in_db:
                    self.log.emit(
                        "Cảnh báo: các cột sau không có trong bảng CSDL đã chọn, "
                        f"sẽ để trống trong kết quả: {', '.join(cols_missing_in_db)}"
                    )

                if "XML1_ID" not in cols_available:
                    raise ValueError("Bảng CSDL đã chọn không có cột XML1_ID, không thể tra cứu.")

                select_cols = ", ".join(f'"{c}"' for c in cols_available)
                sql = f'SELECT {select_cols} FROM "{self.table_name}"'
                self.log.emit("Đang đọc dữ liệu từ CSDL (có thể mất một chút thời gian)...")
                db_df = pd.read_sql_query(sql, conn)
            finally:
                conn.close()

            self.progress.emit(45)

            db_df["_XML1_ID_KEY"] = db_df["XML1_ID"].astype(str).str.strip()
            if use_id_cp:
                if "ID_CP" not in db_df.columns:
                    raise ValueError("Bảng CSDL không có cột ID_CP nên không thể tra cứu theo ID_CP.")
                db_df["_ID_CP_KEY"] = db_df["ID_CP"].astype(str).str.strip()

            self.progress.emit(60)
            self.log.emit("Đang đối chiếu (merge) dữ liệu...")

            merged = input_df.merge(db_df, on=merge_keys, how="left", suffixes=("_INPUT", ""))

            self.progress.emit(80)

            for c in OUTPUT_COLUMNS:
                if c not in merged.columns:
                    merged[c] = ""
            for c in COLUMNS_NOT_IN_DB:
                merged[c] = ""

            result = merged[OUTPUT_COLUMNS].copy().fillna("")

            not_found = merged["MA_BN"].isna().sum() if "MA_BN" in merged.columns else 0
            self.log.emit(
                f"Hoàn tất đối chiếu. Tổng {len(result)} dòng kết quả. "
                f"Số dòng KHÔNG tìm thấy trong CSDL: {not_found}."
            )

            self.progress.emit(100)
            self.finished_ok.emit(result)

        except Exception as e:
            self.log.emit("LỖI: " + str(e))
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(e))


# ============================================================
# WORKER: TÁCH FILE EXCEL THEO MÃ CSKCB
# ============================================================

class SplitWorker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished_ok = pyqtSignal(int)
    failed = pyqtSignal(str)

    def __init__(self, input_path, split_column, output_folder):
        super().__init__()
        self.input_path = input_path
        self.split_column = split_column
        self.output_folder = output_folder

    def run(self):
        try:
            self.log.emit("Đang đọc file Excel nguồn...")
            df = read_table_any(self.input_path)
            self.progress.emit(20)

            if self.split_column not in df.columns:
                raise ValueError(f"Không tìm thấy cột '{self.split_column}' trong file.")

            os.makedirs(self.output_folder, exist_ok=True)

            groups = list(df.groupby(df[self.split_column].fillna("KHONG_XAC_DINH")))
            total = len(groups)
            self.log.emit(f"Tìm thấy {total} giá trị khác nhau của '{self.split_column}'.")

            for idx, (value, sub_df) in enumerate(groups, start=1):
                fname = sanitize_filename(value) + ".xlsx"
                fpath = os.path.join(self.output_folder, fname)
                sub_df.to_excel(fpath, index=False)
                self.log.emit(f"  -> Đã tạo: {fname} ({len(sub_df)} dòng)")
                pct = 20 + int(idx / max(total, 1) * 75)
                self.progress.emit(pct)

            self.progress.emit(100)
            self.log.emit("Hoàn tất tách file.")
            self.finished_ok.emit(total)

        except Exception as e:
            self.log.emit("LỖI: " + str(e))
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(e))


# ============================================================
# WORKER: BƯỚC 2 - TÍNH TỔNG HỢP TRỪ CHI PHÍ THEO CHUYÊN ĐỀ
# ============================================================

class SummaryExcludeWorker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished_ok = pyqtSignal(object, object, object)  # detail, summary_cd, summary_cd_loai
    failed = pyqtSignal(str)

    def __init__(self, input_path,
                 col_xml1id, col_ma_cskcb, col_chuyen_de,
                 col_loai_kcb, loai_kcb_is_text,
                 col_tbhtt, col_ghi_chu):
        super().__init__()
        self.input_path = input_path
        self.col_xml1id = col_xml1id
        self.col_ma_cskcb = col_ma_cskcb
        self.col_chuyen_de = col_chuyen_de
        self.col_loai_kcb = col_loai_kcb
        self.loai_kcb_is_text = loai_kcb_is_text
        self.col_tbhtt = col_tbhtt
        self.col_ghi_chu = col_ghi_chu

    def run(self):
        try:
            self.log.emit("Đang đọc file Excel nguồn...")
            df = read_table_any(self.input_path)
            self.progress.emit(15)

            for c, label in [(self.col_xml1id, "XML1_ID"),
                              (self.col_ma_cskcb, "MA_CSKCB"),
                              (self.col_chuyen_de, "Mã chuyên đề"),
                              (self.col_tbhtt, "T_BHTT")]:
                if c not in df.columns:
                    raise ValueError(f"Không tìm thấy cột '{c}' (dùng cho {label}) trong file.")

            if self.col_loai_kcb and self.col_loai_kcb != NONE_OPTION:
                if self.col_loai_kcb not in df.columns:
                    raise ValueError(f"Không tìm thấy cột '{self.col_loai_kcb}' trong file.")
                if self.loai_kcb_is_text:
                    df["LOAI_HO_SO"] = df[self.col_loai_kcb].astype(str).str.strip()
                else:
                    df["LOAI_HO_SO"] = df[self.col_loai_kcb].apply(classify_loai_ho_so)
            else:
                df["LOAI_HO_SO"] = ""

            self.progress.emit(30)
            self.log.emit("Đang xử lý số tiền T_BHTT...")
            df["_TBHTT_NUM"] = df[self.col_tbhtt].apply(parse_money)

            use_ghi_chu = bool(self.col_ghi_chu) and self.col_ghi_chu != NONE_OPTION
            if use_ghi_chu:
                if self.col_ghi_chu not in df.columns:
                    raise ValueError(f"Không tìm thấy cột '{self.col_ghi_chu}' trong file.")
                sample_vals = df[self.col_ghi_chu].dropna().astype(str).str.strip()
                sample_vals = sample_vals[sample_vals != ""].unique()[:8]
                self.log.emit(
                    "Một vài giá trị mẫu trong cột Ghi chú không trừ: "
                    + (", ".join(f"'{v}'" for v in sample_vals) if len(sample_vals) else "(không có giá trị nào khác rỗng)")
                )
                df["_KHONG_TRU"] = df[self.col_ghi_chu].apply(is_note_marked)
                self.log.emit(
                    f"Số dòng có ghi chú (không trừ): {int(df['_KHONG_TRU'].sum())} / {len(df)}"
                )
            else:
                df["_KHONG_TRU"] = False
                self.log.emit(
                    "Không chọn cột Ghi chú không trừ -> mặc định TẤT CẢ các dòng đều được tính là bị trừ."
                )

            df["TRANG_THAI_TRU"] = df["_KHONG_TRU"].map({True: "Không trừ", False: "Trừ"})

            self.progress.emit(50)
            self.log.emit("Đang tổng hợp theo MA_CSKCB và Mã chuyên đề...")

            included = df[~df["_KHONG_TRU"]].copy()
            if included.empty:
                self.log.emit(
                    "CẢNH BÁO: sau khi loại các dòng đã đánh dấu, KHÔNG còn dòng nào để tính "
                    "trừ. Vui lòng kiểm tra lại cột Ghi chú không trừ đã chọn đúng chưa "
                    "(có thể đang chọn nhầm cột luôn có giá trị ở mọi dòng)."
                )

            summary_cd = (
                included.groupby([self.col_ma_cskcb, self.col_chuyen_de])
                .agg(SO_HO_SO=(self.col_xml1id, "nunique"),
                     TONG_CHI_PHI=("_TBHTT_NUM", "sum"))
                .reset_index()
                .rename(columns={self.col_ma_cskcb: "MA_CSKCB", self.col_chuyen_de: "MA_CHUYEN_DE"})
            )
            summary_cd = summary_cd.sort_values(["MA_CSKCB", "MA_CHUYEN_DE"]).reset_index(drop=True)

            self.progress.emit(70)
            self.log.emit("Đang tổng hợp chi tiết theo Loại hồ sơ (Ngoại trú/Nội trú)...")

            included["_LOAI_PIVOT"] = included["LOAI_HO_SO"].replace("", "Chưa xác định")

            long_cd_loai = (
                included.groupby([self.col_ma_cskcb, self.col_chuyen_de, "_LOAI_PIVOT"])
                .agg(SO_HO_SO=(self.col_xml1id, "nunique"),
                     TONG_CHI_PHI=("_TBHTT_NUM", "sum"))
                .reset_index()
            )

            # Thứ tự cột mong muốn: Ngoại trú, Nội trú, rồi các loại khác (nếu có)
            uu_tien = ["Ngoại trú", "Nội trú"]
            cac_loai = list(dict.fromkeys(long_cd_loai["_LOAI_PIVOT"].tolist()))
            thu_tu_loai = [x for x in uu_tien if x in cac_loai] + [x for x in cac_loai if x not in uu_tien]

            if long_cd_loai.empty:
                summary_cd_loai = summary_cd[["MA_CSKCB", "MA_CHUYEN_DE"]].copy()
                summary_cd_loai["TONG_SO_HO_SO"] = 0
                summary_cd_loai["TONG_CHI_PHI"] = 0.0
            else:
                pivot_so = long_cd_loai.pivot_table(
                    index=[self.col_ma_cskcb, self.col_chuyen_de],
                    columns="_LOAI_PIVOT", values="SO_HO_SO", fill_value=0
                )
                pivot_cp = long_cd_loai.pivot_table(
                    index=[self.col_ma_cskcb, self.col_chuyen_de],
                    columns="_LOAI_PIVOT", values="TONG_CHI_PHI", fill_value=0.0
                )

                summary_cd_loai = pd.DataFrame(index=pivot_so.index)
                for loai in thu_tu_loai:
                    hau_to = ascii_slug(loai)
                    summary_cd_loai[f"SO_HO_SO_{hau_to}"] = (
                        pivot_so[loai] if loai in pivot_so.columns else 0
                    )
                    summary_cd_loai[f"TONG_CHI_PHI_{hau_to}"] = (
                        pivot_cp[loai] if loai in pivot_cp.columns else 0.0
                    )
                summary_cd_loai = summary_cd_loai.reset_index().rename(
                    columns={self.col_ma_cskcb: "MA_CSKCB", self.col_chuyen_de: "MA_CHUYEN_DE"}
                )

                so_cols = [c for c in summary_cd_loai.columns if c.startswith("SO_HO_SO_")]
                cp_cols = [c for c in summary_cd_loai.columns if c.startswith("TONG_CHI_PHI_")]
                summary_cd_loai["TONG_SO_HO_SO"] = summary_cd_loai[so_cols].sum(axis=1)
                summary_cd_loai["TONG_CHI_PHI"] = summary_cd_loai[cp_cols].sum(axis=1)

            summary_cd_loai = summary_cd_loai.sort_values(
                ["MA_CSKCB", "MA_CHUYEN_DE"]
            ).reset_index(drop=True)

            self.progress.emit(90)

            detail = df.drop(columns=["_KHONG_TRU"]).rename(columns={"_TBHTT_NUM": "T_BHTT_SO"})

            total_ho_so = included[self.col_xml1id].nunique()
            total_chi_phi = included["_TBHTT_NUM"].sum()
            self.log.emit(
                f"Hoàn tất. Tổng số hồ sơ bị trừ: {total_ho_so} | "
                f"Tổng chi phí bị trừ: {format_money(total_chi_phi)}"
            )

            self.progress.emit(100)
            self.finished_ok.emit(detail, summary_cd, summary_cd_loai)

        except Exception as e:
            self.log.emit("LỖI: " + str(e))
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(e))


# ============================================================
# WORKER: LƯU HỒ SƠ ĐÃ TRỪ VÀO CSDL SQLITE & KIỂM TRA TRÙNG
# ============================================================

class DeductedDbWorker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished_check = pyqtSignal(object, int, int)   # trung_df, so_trung, so_kiem_tra
    finished_save = pyqtSignal(int, int, int)        # so_da_luu, so_trung_bo_qua, tong_so
    failed = pyqtSignal(str)

    def __init__(self, db_path, input_path,
                 col_ma_cskcb, col_xml1, col_id_cp, col_chuyen_de, col_cong_van,
                 col_trang_thai, gia_tri_tru, mode):
        super().__init__()
        self.db_path = db_path
        self.input_path = input_path
        self.col_ma_cskcb = col_ma_cskcb
        self.col_xml1 = col_xml1
        self.col_id_cp = col_id_cp
        self.col_chuyen_de = col_chuyen_de
        self.col_cong_van = col_cong_van
        self.col_trang_thai = col_trang_thai
        self.gia_tri_tru = gia_tri_tru
        self.mode = mode  # "check" hoặc "save"

    @staticmethod
    def _ensure_table(conn):
        all_cols = DEDUCTED_ALL_COLUMNS + ["NGAY_LUU"]
        col_defs = ",\n                ".join(f"{c} TEXT" for c in all_cols)
        conn.execute(f"""
            CREATE TABLE IF NOT EXISTS HO_SO_DA_TRU (
                {col_defs},
                UNIQUE(XML1_ID, ID_CP, MA_CHUYEN_DE)
            )
        """)
        conn.commit()
        # Nếu CSDL được tạo từ phiên bản cũ (thiếu cột mới), tự động bổ sung cột.
        cur = conn.execute("PRAGMA table_info(HO_SO_DA_TRU)")
        existing_cols = {row[1] for row in cur.fetchall()}
        for c in all_cols:
            if c not in existing_cols:
                conn.execute(f"ALTER TABLE HO_SO_DA_TRU ADD COLUMN {c} TEXT")
        conn.commit()

    def _load_records(self):
        self.log.emit("Đang đọc file Excel...")
        df = read_table_any(self.input_path)
        self.progress.emit(10)

        for c, label in [(self.col_ma_cskcb, "MA_CSKCB"),
                          (self.col_xml1, "XML1_ID"),
                          (self.col_chuyen_de, "Mã chuyên đề")]:
            if c not in df.columns:
                raise ValueError(f"Không tìm thấy cột '{c}' (dùng cho {label}) trong file.")

        if self.col_trang_thai and self.col_trang_thai != NONE_OPTION:
            if self.col_trang_thai not in df.columns:
                raise ValueError(f"Không tìm thấy cột '{self.col_trang_thai}' trong file.")
            before_n = len(df)
            gia_tri = (self.gia_tri_tru or "").strip()
            mask = df[self.col_trang_thai].astype(str).str.strip().str.casefold() == gia_tri.casefold()
            df = df[mask].copy()
            self.log.emit(
                f"Đã lọc theo cột trạng thái '{self.col_trang_thai}' = '{gia_tri}': "
                f"giữ lại {len(df)}/{before_n} dòng."
            )

        def get_col(df_, colname):
            if colname and colname != NONE_OPTION and colname in df_.columns:
                return df_[colname].fillna("").astype(str).str.strip()
            return pd.Series([""] * len(df_), index=df_.index)

        data = {
            "MA_CSKCB": get_col(df, self.col_ma_cskcb),
            "XML1_ID": get_col(df, self.col_xml1),
            "ID_CP": get_col(df, self.col_id_cp),
            "MA_CHUYEN_DE": get_col(df, self.col_chuyen_de),
            "CONG_VAN": get_col(df, self.col_cong_van),
        }

        found_extra, missing_extra = [], []
        for c in DEDUCTED_EXTRA_COLUMNS:
            if c in df.columns:
                data[c] = df[c].fillna("").astype(str).str.strip()
                found_extra.append(c)
            else:
                data[c] = pd.Series([""] * len(df), index=df.index)
                missing_extra.append(c)

        records = pd.DataFrame(data)[DEDUCTED_ALL_COLUMNS]
        self.progress.emit(25)

        if found_extra:
            self.log.emit(
                f"Đã tự động lấy thêm {len(found_extra)} cột dữ liệu bổ sung từ file: "
                + ", ".join(found_extra)
            )
        if missing_extra:
            self.log.emit(
                "Các cột bổ sung KHÔNG có trong file (sẽ để trống): " + ", ".join(missing_extra)
            )

        return records

    def run(self):
        try:
            records = self._load_records()
            if records.empty:
                raise ValueError("Không có dòng dữ liệu nào để xử lý (có thể do bộ lọc trạng thái).")

            self.log.emit(f"Đã chuẩn bị {len(records)} dòng dữ liệu từ file.")

            conn = sqlite3.connect(self.db_path)
            try:
                self._ensure_table(conn)
                self.progress.emit(35)

                if self.mode == "check":
                    self.log.emit("Đang đọc dữ liệu đã lưu trong CSDL để đối chiếu...")
                    select_cols_sql = ", ".join(DEDUCTED_ALL_COLUMNS + ["NGAY_LUU"])
                    db_df = pd.read_sql_query(
                        f"SELECT {select_cols_sql} FROM HO_SO_DA_TRU", conn
                    )
                    self.progress.emit(70)
                    for c in ["XML1_ID", "ID_CP", "MA_CHUYEN_DE"]:
                        db_df[c] = db_df[c].fillna("").astype(str).str.strip()

                    merged = records.merge(
                        db_df, on=["XML1_ID", "ID_CP", "MA_CHUYEN_DE"],
                        how="inner", suffixes=("", "_DA_LUU")
                    )
                    self.progress.emit(100)
                    self.log.emit(
                        f"Tìm thấy {len(merged)}/{len(records)} dòng trùng với dữ liệu đã lưu trước đó."
                    )
                    self.finished_check.emit(merged, len(merged), len(records))

                else:  # save
                    before_dedup = len(records)
                    records = records.drop_duplicates(
                        subset=["XML1_ID", "ID_CP", "MA_CHUYEN_DE"]
                    )
                    intra_dup = before_dedup - len(records)
                    if intra_dup > 0:
                        self.log.emit(
                            f"Cảnh báo: có {intra_dup} dòng trùng lặp ngay trong lô dữ liệu "
                            "hiện tại, chỉ giữ 1 dòng cho mỗi tổ hợp (XML1_ID, ID_CP, Mã chuyên đề)."
                        )
                    self.progress.emit(50)

                    cur = conn.cursor()
                    cur.execute("SELECT COUNT(*) FROM HO_SO_DA_TRU")
                    count_before = cur.fetchone()[0]

                    now = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                    all_insert_cols = DEDUCTED_ALL_COLUMNS + ["NGAY_LUU"]
                    col_list_sql = ", ".join(all_insert_cols)
                    placeholders = ", ".join(["?"] * len(all_insert_cols))
                    rows = [
                        tuple(getattr(r, c) for c in DEDUCTED_ALL_COLUMNS) + (now,)
                        for r in records.itertuples(index=False)
                    ]
                    self.log.emit(f"Đang ghi {len(rows)} dòng vào CSDL (bỏ qua tự động nếu trùng)...")
                    cur.executemany(
                        f"INSERT OR IGNORE INTO HO_SO_DA_TRU ({col_list_sql}) VALUES ({placeholders})",
                        rows
                    )
                    conn.commit()
                    self.progress.emit(90)

                    cur.execute("SELECT COUNT(*) FROM HO_SO_DA_TRU")
                    count_after = cur.fetchone()[0]
                    so_da_luu = count_after - count_before
                    so_trung_bo_qua = before_dedup - so_da_luu
                    self.progress.emit(100)
                    self.log.emit(
                        f"Đã lưu mới {so_da_luu} dòng. Bỏ qua {so_trung_bo_qua} dòng vì đã "
                        f"tồn tại trong CSDL hoặc trùng ngay trong lô dữ liệu hiện tại."
                    )
                    self.finished_save.emit(so_da_luu, so_trung_bo_qua, before_dedup)
            finally:
                conn.close()

        except Exception as e:
            self.log.emit("LỖI: " + str(e))
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(e))


# ============================================================
# TRANG 1: TRA CỨU DỮ LIỆU XML1
# ============================================================

class LookupPage(QWidget):
    def __init__(self):
        super().__init__()
        self.result_df = None
        self.input_df_columns = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        layout.addWidget(page_title("Tra cứu dữ liệu XML1 từ CSDL SQLite"))

        db_group = QGroupBox("1. Cơ sở dữ liệu SQLite nguồn")
        db_form = QFormLayout()

        self.db_path_edit = QLineEdit()
        db_browse_btn = QPushButton("Chọn file .sqlite/.db...")
        db_browse_btn.clicked.connect(self.browse_db)
        db_row = QHBoxLayout()
        db_row.addWidget(self.db_path_edit)
        db_row.addWidget(db_browse_btn)
        db_form.addRow("File CSDL:", db_row)

        self.table_combo = QComboBox()
        db_form.addRow("Bảng dữ liệu:", self.table_combo)

        db_group.setLayout(db_form)
        layout.addWidget(db_group)

        in_group = QGroupBox("2. File dữ liệu đầu vào (chứa cột XML1_ID)")
        in_form = QFormLayout()

        self.input_path_edit = QLineEdit()
        input_browse_btn = QPushButton("Chọn file Excel/CSV...")
        input_browse_btn.clicked.connect(self.browse_input)
        in_row = QHBoxLayout()
        in_row.addWidget(self.input_path_edit)
        in_row.addWidget(input_browse_btn)
        in_form.addRow("File đầu vào:", in_row)

        self.load_cols_btn = QPushButton("Đọc cột của file")
        self.load_cols_btn.clicked.connect(self.load_input_columns)
        in_form.addRow("", self.load_cols_btn)

        self.col_xml1_combo = QComboBox()
        in_form.addRow("Cột XML1_ID:", self.col_xml1_combo)

        self.col_idcp_combo = QComboBox()
        in_form.addRow("Cột ID_CP (nếu có):", self.col_idcp_combo)

        self.col_ma_cd_combo = QComboBox()
        in_form.addRow("Cột Mã chuyên đề (nếu có):", self.col_ma_cd_combo)

        self.col_cong_van_combo = QComboBox()
        in_form.addRow("Cột Công văn (nếu có):", self.col_cong_van_combo)

        in_group.setLayout(in_form)
        layout.addWidget(in_group)

        action_row = QHBoxLayout()
        self.run_btn = QPushButton("Thực hiện tra cứu")
        self.run_btn.setObjectName("primaryBtn")
        self.run_btn.clicked.connect(self.run_lookup)
        self.export_btn = QPushButton("Xuất kết quả ra Excel")
        self.export_btn.clicked.connect(self.export_result)
        self.export_btn.setEnabled(False)
        action_row.addWidget(self.run_btn)
        action_row.addWidget(self.export_btn)
        layout.addLayout(action_row)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        self.table_preview = QTableWidget()
        self.table_preview.setColumnCount(len(OUTPUT_COLUMNS))
        self.table_preview.setHorizontalHeaderLabels(OUTPUT_COLUMNS)
        layout.addWidget(self.table_preview, stretch=3)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumHeight(130)
        layout.addWidget(self.log_box, stretch=1)

    def append_log(self, text):
        self.log_box.append(text)

    def browse_db(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file CSDL SQLite", "",
            "SQLite Database (*.sqlite *.sqlite3 *.db);;Tất cả file (*.*)"
        )
        if not path:
            return
        self.db_path_edit.setText(path)
        self.load_tables(path)

    def load_tables(self, db_path):
        self.table_combo.clear()
        try:
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
            tables = [r[0] for r in cur.fetchall()]
            conn.close()
            self.table_combo.addItems(tables)
            self.append_log(f"Đã tìm thấy {len(tables)} bảng trong CSDL.")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể đọc danh sách bảng:\n{e}")

    def browse_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file dữ liệu đầu vào", "",
            "Excel/CSV (*.xlsx *.xls *.csv);;Tất cả file (*.*)"
        )
        if not path:
            return
        self.input_path_edit.setText(path)
        self.load_input_columns()

    def load_input_columns(self):
        path = self.input_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file đầu vào trước.")
            return
        try:
            df = read_table_any(path)
            self.input_df_columns = list(df.columns)

            self.col_xml1_combo.clear()
            self.col_xml1_combo.addItems(self.input_df_columns)
            if "XML1_ID" in self.input_df_columns:
                self.col_xml1_combo.setCurrentText("XML1_ID")

            self.col_idcp_combo.clear()
            self.col_idcp_combo.addItem(NONE_OPTION)
            self.col_idcp_combo.addItems(self.input_df_columns)
            if "ID_CP" in self.input_df_columns:
                self.col_idcp_combo.setCurrentText("ID_CP")

            self.col_ma_cd_combo.clear()
            self.col_ma_cd_combo.addItem(NONE_OPTION)
            self.col_ma_cd_combo.addItems(self.input_df_columns)
            for guess in ("MA_CHUYEN_DE", "MA_CHUYEN DE", "CHUYEN_DE", "MA CHUYEN DE"):
                if guess in self.input_df_columns:
                    self.col_ma_cd_combo.setCurrentText(guess)
                    break

            self.col_cong_van_combo.clear()
            self.col_cong_van_combo.addItem(NONE_OPTION)
            self.col_cong_van_combo.addItems(self.input_df_columns)
            for guess in ("CONG_VAN", "SO_CONG_VAN", "CONG VAN", "SO CONG VAN"):
                if guess in self.input_df_columns:
                    self.col_cong_van_combo.setCurrentText(guess)
                    break

            self.append_log(f"Đã đọc {len(self.input_df_columns)} cột từ file đầu vào.")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không đọc được file đầu vào:\n{e}")

    def run_lookup(self):
        db_path = self.db_path_edit.text().strip()
        table_name = self.table_combo.currentText().strip()
        input_path = self.input_path_edit.text().strip()
        col_xml1 = self.col_xml1_combo.currentText().strip()
        col_idcp = self.col_idcp_combo.currentText().strip()
        col_ma_cd = self.col_ma_cd_combo.currentText().strip()
        col_cong_van = self.col_cong_van_combo.currentText().strip()

        if not db_path or not os.path.isfile(db_path):
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file CSDL SQLite hợp lệ.")
            return
        if not table_name:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn bảng dữ liệu.")
            return
        if not input_path or not os.path.isfile(input_path):
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file dữ liệu đầu vào hợp lệ.")
            return
        if not col_xml1:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn cột XML1_ID.")
            return

        self.run_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()

        self.worker = LookupWorker(
            db_path, table_name, input_path, col_xml1, col_idcp,
            col_ma_chuyen_de=col_ma_cd, col_cong_van=col_cong_van
        )
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished_ok.connect(self.on_lookup_done)
        self.worker.failed.connect(self.on_lookup_failed)
        self.worker.start()

    def on_lookup_done(self, df):
        self.result_df = df
        self.run_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        fill_table_widget(self.table_preview, df, OUTPUT_COLUMNS, max_rows=500)
        if len(df) > 500:
            self.append_log(f"(Chỉ hiển thị 500/{len(df)} dòng trong bảng xem trước.)")
        QMessageBox.information(self, "Hoàn tất", f"Tra cứu xong: {len(df)} dòng kết quả.")

    def on_lookup_failed(self, msg):
        self.run_btn.setEnabled(True)
        QMessageBox.critical(self, "Lỗi", msg)

    def export_result(self):
        if self.result_df is None or self.result_df.empty:
            QMessageBox.warning(self, "Không có dữ liệu", "Chưa có kết quả để xuất.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file kết quả", "ket_qua_tra_cuu_xml1.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            self.result_df.to_excel(path, index=False)
            QMessageBox.information(self, "Thành công", f"Đã lưu file:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể lưu file:\n{e}")


# ============================================================
# TRANG 2: TÁCH FILE EXCEL THEO MÃ CSKCB
# ============================================================

class SplitPage(QWidget):
    def __init__(self):
        super().__init__()
        self.input_columns = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        layout.addWidget(page_title("Tách file Excel theo MA_CSKCB"))

        group = QGroupBox("Tách 1 file Excel thành nhiều file theo cột MA_CSKCB")
        form = QFormLayout()

        self.input_path_edit = QLineEdit()
        browse_in_btn = QPushButton("Chọn file Excel...")
        browse_in_btn.clicked.connect(self.browse_input)
        row1 = QHBoxLayout()
        row1.addWidget(self.input_path_edit)
        row1.addWidget(browse_in_btn)
        form.addRow("File Excel nguồn:", row1)

        self.read_cols_btn = QPushButton("Đọc cột của file")
        self.read_cols_btn.clicked.connect(self.load_columns)
        form.addRow("", self.read_cols_btn)

        self.split_col_combo = QComboBox()
        form.addRow("Cột dùng để tách (MA_CSKCB):", self.split_col_combo)

        self.output_folder_edit = QLineEdit()
        browse_out_btn = QPushButton("Chọn thư mục lưu...")
        browse_out_btn.clicked.connect(self.browse_output_folder)
        row2 = QHBoxLayout()
        row2.addWidget(self.output_folder_edit)
        row2.addWidget(browse_out_btn)
        form.addRow("Thư mục đích:", row2)

        group.setLayout(form)
        layout.addWidget(group)

        self.split_btn = QPushButton("Bắt đầu tách file")
        self.split_btn.setObjectName("primaryBtn")
        self.split_btn.clicked.connect(self.run_split)
        layout.addWidget(self.split_btn)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box, stretch=1)

        layout.addStretch()

    def append_log(self, text):
        self.log_box.append(text)

    def browse_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel nguồn", "",
            "Excel (*.xlsx *.xls);;Tất cả file (*.*)"
        )
        if not path:
            return
        self.input_path_edit.setText(path)
        self.load_columns()

    def load_columns(self):
        path = self.input_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file Excel trước.")
            return
        try:
            df = read_table_any(path)
            self.input_columns = list(df.columns)
            self.split_col_combo.clear()
            self.split_col_combo.addItems(self.input_columns)
            if "MA_CSKCB" in self.input_columns:
                self.split_col_combo.setCurrentText("MA_CSKCB")
            self.append_log(f"Đã đọc {len(self.input_columns)} cột từ file.")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không đọc được file:\n{e}")

    def browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục đích")
        if folder:
            self.output_folder_edit.setText(folder)

    def run_split(self):
        input_path = self.input_path_edit.text().strip()
        split_col = self.split_col_combo.currentText().strip()
        out_folder = self.output_folder_edit.text().strip()

        if not input_path or not os.path.isfile(input_path):
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file Excel nguồn hợp lệ.")
            return
        if not split_col:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn cột để tách (MA_CSKCB).")
            return
        if not out_folder:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn thư mục đích.")
            return

        self.split_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()

        self.worker = SplitWorker(input_path, split_col, out_folder)
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished_ok.connect(self.on_split_done)
        self.worker.failed.connect(self.on_split_failed)
        self.worker.start()

    def on_split_done(self, count):
        self.split_btn.setEnabled(True)
        QMessageBox.information(self, "Hoàn tất", f"Đã tách thành {count} file.")

    def on_split_failed(self, msg):
        self.split_btn.setEnabled(True)
        QMessageBox.critical(self, "Lỗi", msg)


# ============================================================
# TRANG 3: TỔNG HỢP TRỪ CHI PHÍ THEO MA_CSKCB & MÃ CHUYÊN ĐỀ
# ============================================================

class ExcludeSummaryPage(QWidget):
    def __init__(self):
        super().__init__()
        self.input_columns = []
        self.detail_df = None
        self.summary_cd_df = None
        self.summary_cd_loai_df = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        layout.addWidget(page_title("Tổng hợp trừ chi phí theo MA_CSKCB & Mã chuyên đề"))
        note = QLabel(
            "Tính số hồ sơ và chi phí bị trừ theo từng MA_CSKCB và Mã chuyên đề. "
            "XML1_ID được coi là khóa của 1 hồ sơ (có thể có nhiều dòng chi phí)."
        )
        note.setWordWrap(True)
        note.setObjectName("noteLabel")
        layout.addWidget(note)

        # --- Nhóm chọn file & cột dữ liệu ---
        in_group = QGroupBox("File dữ liệu & khai báo cột")
        in_form = QFormLayout()

        self.input_path_edit = QLineEdit()
        browse_in_btn = QPushButton("Chọn file Excel...")
        browse_in_btn.clicked.connect(self.browse_input)
        row1 = QHBoxLayout()
        row1.addWidget(self.input_path_edit)
        row1.addWidget(browse_in_btn)
        in_form.addRow("File Excel:", row1)

        self.read_cols_btn = QPushButton("Đọc cột của file")
        self.read_cols_btn.clicked.connect(self.load_columns)
        in_form.addRow("", self.read_cols_btn)

        self.col_xml1_combo = QComboBox()
        in_form.addRow("Cột XML1_ID (khóa hồ sơ):", self.col_xml1_combo)

        self.col_ma_cskcb_combo = QComboBox()
        in_form.addRow("Cột MA_CSKCB:", self.col_ma_cskcb_combo)

        self.col_chuyen_de_combo = QComboBox()
        in_form.addRow("Cột Mã chuyên đề:", self.col_chuyen_de_combo)

        loai_row = QHBoxLayout()
        self.col_loai_kcb_combo = QComboBox()
        loai_row.addWidget(self.col_loai_kcb_combo)
        self.loai_is_text_check = QCheckBox("Cột này đã ghi sẵn chữ Ngoại trú/Nội trú (không cần quy đổi mã)")
        loai_row.addWidget(self.loai_is_text_check)
        in_form.addRow("Cột mã loại KCB\n(1,2=Ngoại trú; 3,9=Nội trú):", loai_row)

        self.col_tbhtt_combo = QComboBox()
        in_form.addRow("Cột T_BHTT (số tiền để trừ):", self.col_tbhtt_combo)

        self.col_ghi_chu_combo = QComboBox()
        in_form.addRow("Cột đánh dấu KHÔNG trừ\n(nếu file đã có):", self.col_ghi_chu_combo)

        ghi_chu_hint = QLabel(
            "💡 Trong file Excel của bạn, chọn cột nào đang được dùng để đánh dấu các "
            "dòng KHÔNG bị trừ ở trên (ví dụ cột tên GHI_CHU_KHONG_TRU). Chỉ cần ghi bất "
            "kỳ ký tự nào (ví dụ 'x') vào cột đó ở dòng không muốn trừ, để trống các dòng "
            "còn lại. Nếu file chưa có cột này, để \"-- Không dùng --\" thì toàn bộ dữ "
            "liệu sẽ được tính là bị trừ."
        )
        ghi_chu_hint.setWordWrap(True)
        ghi_chu_hint.setObjectName("noteLabel")
        in_form.addRow("", ghi_chu_hint)

        in_group.setLayout(in_form)
        layout.addWidget(in_group)

        btn_row = QHBoxLayout()
        self.summary_btn = QPushButton("Tính tổng hợp")
        self.summary_btn.setObjectName("primaryBtn")
        self.summary_btn.clicked.connect(self.run_summary)
        self.export_summary_btn = QPushButton("Xuất kết quả tổng hợp ra Excel (3 sheet)")
        self.export_summary_btn.clicked.connect(self.export_summary)
        self.export_summary_btn.setEnabled(False)
        btn_row.addWidget(self.summary_btn)
        btn_row.addWidget(self.export_summary_btn)
        layout.addLayout(btn_row)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        layout.addWidget(QLabel("Chi tiết các bước đang xử lý:"))
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box, stretch=1)

    def append_log(self, text):
        self.log_box.append(text)

    def browse_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel", "",
            "Excel/CSV (*.xlsx *.xls *.csv);;Tất cả file (*.*)"
        )
        if not path:
            return
        self.input_path_edit.setText(path)
        self.load_columns()

    def load_columns(self):
        path = self.input_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file Excel trước.")
            return
        try:
            df = read_table_any(path)
            self.input_columns = list(df.columns)

            def fill_combo(combo, guesses, allow_none=False):
                combo.clear()
                if allow_none:
                    combo.addItem(NONE_OPTION)
                combo.addItems(self.input_columns)
                for g in guesses:
                    if g in self.input_columns:
                        combo.setCurrentText(g)
                        return g
                return None

            fill_combo(self.col_xml1_combo, ["XML1_ID"])
            fill_combo(self.col_ma_cskcb_combo, ["MA_CSKCB"])
            fill_combo(self.col_chuyen_de_combo, ["MA_CHUYEN_DE", "CHUYEN_DE"])
            loai_col_chosen = fill_combo(
                self.col_loai_kcb_combo,
                ["LOAI_HO_SO", "MA_LOAI_KCB", "LOAI_KCB"], allow_none=True
            )
            # Nếu cột chọn được chính là LOAI_HO_SO (đã tạo sẵn ở Bước 1, đã là
            # chữ "Ngoại trú"/"Nội trú") thì tự tích chọn "đã là chữ", tránh
            # quy đổi nhầm lần 2 làm sai/rỗng dữ liệu.
            self.loai_is_text_check.setChecked(loai_col_chosen == "LOAI_HO_SO")
            fill_combo(self.col_tbhtt_combo, ["T_BHTT", "T_BHTT_DTL"])
            fill_combo(self.col_ghi_chu_combo,
                       ["GHI_CHU_KHONG_TRU", "GHI_CHU"], allow_none=True)

            self.append_log(f"Đã đọc {len(self.input_columns)} cột từ file.")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không đọc được file:\n{e}")

    def _validate_common(self):
        input_path = self.input_path_edit.text().strip()
        if not input_path or not os.path.isfile(input_path):
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file Excel hợp lệ.")
            return None
        col_xml1 = self.col_xml1_combo.currentText().strip()
        col_ma_cskcb = self.col_ma_cskcb_combo.currentText().strip()
        col_chuyen_de = self.col_chuyen_de_combo.currentText().strip()
        if not col_xml1 or not col_ma_cskcb or not col_chuyen_de:
            QMessageBox.warning(
                self, "Thiếu thông tin",
                "Vui lòng chọn đủ cột XML1_ID, MA_CSKCB và Mã chuyên đề."
            )
            return None
        return input_path, col_xml1, col_ma_cskcb, col_chuyen_de

    def run_summary(self):
        common = self._validate_common()
        if common is None:
            return
        input_path, col_xml1, col_ma_cskcb, col_chuyen_de = common

        col_tbhtt = self.col_tbhtt_combo.currentText().strip()
        if not col_tbhtt:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn cột T_BHTT.")
            return

        col_loai_kcb = self.col_loai_kcb_combo.currentText().strip()
        loai_is_text = self.loai_is_text_check.isChecked()
        col_ghi_chu = self.col_ghi_chu_combo.currentText().strip()

        if (not col_ghi_chu) or col_ghi_chu == NONE_OPTION:
            reply = QMessageBox.question(
                self, "Xác nhận",
                "Bạn chưa chọn cột đánh dấu KHÔNG trừ -> toàn bộ dữ liệu sẽ được "
                "tính là BỊ TRỪ (không có ngoại lệ). Tiếp tục?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        self.summary_btn.setEnabled(False)
        self.export_summary_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()

        self.summary_worker = SummaryExcludeWorker(
            input_path, col_xml1, col_ma_cskcb, col_chuyen_de,
            col_loai_kcb, loai_is_text, col_tbhtt, col_ghi_chu
        )
        self.summary_worker.log.connect(self.append_log)
        self.summary_worker.progress.connect(self.progress.setValue)
        self.summary_worker.finished_ok.connect(self.on_summary_done)
        self.summary_worker.failed.connect(self.on_summary_failed)
        self.summary_worker.start()

    def on_summary_done(self, detail_df, summary_cd_df, summary_cd_loai_df):
        self.detail_df = detail_df
        self.summary_cd_df = summary_cd_df
        self.summary_cd_loai_df = summary_cd_loai_df

        self.summary_btn.setEnabled(True)
        self.export_summary_btn.setEnabled(True)

        so_ho_so = int(summary_cd_df["SO_HO_SO"].sum()) if not summary_cd_df.empty else 0
        tong_chi_phi = summary_cd_df["TONG_CHI_PHI"].sum() if not summary_cd_df.empty else 0
        self.append_log(
            f"KẾT QUẢ: tổng {so_ho_so} hồ sơ bị trừ, tổng chi phí bị trừ "
            f"{format_money(tong_chi_phi)}, trên {len(summary_cd_df)} tổ hợp "
            f"MA_CSKCB & Mã chuyên đề. Bấm 'Xuất kết quả tổng hợp ra Excel' để "
            f"xem chi tiết đầy đủ."
        )

        QMessageBox.information(
            self, "Hoàn tất",
            "Đã tính xong tổng hợp trừ chi phí.\n\n"
            "Mẹo: để lưu các hồ sơ đã trừ vào CSDL và kiểm tra trùng với các lần "
            "trước, dùng chức năng '💾 Lưu hồ sơ đã trừ & Kiểm tra trùng' ở menu bên trái "
            "(tải trực tiếp file Excel bạn vừa xuất, ví dụ sheet Chi_Tiet)."
        )

    def on_summary_failed(self, msg):
        self.summary_btn.setEnabled(True)
        QMessageBox.critical(self, "Lỗi", msg)

    def export_summary(self):
        if self.summary_cd_df is None or self.summary_cd_loai_df is None or self.detail_df is None:
            QMessageBox.warning(self, "Không có dữ liệu", "Chưa có kết quả để xuất.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file tổng hợp", "tong_hop_tru_chi_phi.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                self.summary_cd_df.to_excel(writer, sheet_name="Tong_Hop_Chuyen_De", index=False)
                self.summary_cd_loai_df.to_excel(writer, sheet_name="Tong_Hop_Loai_Ho_So", index=False)
                self.detail_df.to_excel(writer, sheet_name="Chi_Tiet", index=False)
            QMessageBox.information(self, "Thành công", f"Đã lưu file:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể lưu file:\n{e}")


# ============================================================
# WORKER: NẠP/CẬP NHẬT DANH MỤC QUY TẮC GIÁM ĐỊNH THEO CHUYÊN ĐỀ
# ============================================================

class RuleLoadWorker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished_ok = pyqtSignal(int, int)  # số dòng nạp từ file, tổng số quy tắc trong CSDL sau khi nạp
    failed = pyqtSignal(str)

    def __init__(self, input_path, col_ma_cd, col_ten_cd, col_ma_cp, col_ten_cp,
                 col_ma_benh, col_ten_benh, col_nhom_benh):
        super().__init__()
        self.input_path = input_path
        self.col_ma_cd = col_ma_cd
        self.col_ten_cd = col_ten_cd
        self.col_ma_cp = col_ma_cp
        self.col_ten_cp = col_ten_cp
        self.col_ma_benh = col_ma_benh
        self.col_ten_benh = col_ten_benh
        self.col_nhom_benh = col_nhom_benh

    def run(self):
        try:
            self.log.emit("Đang đọc file Excel danh mục quy tắc...")
            df = read_table_any(self.input_path)
            self.progress.emit(15)

            for c, label in [(self.col_ma_cd, "Mã chuyên đề"),
                              (self.col_ma_cp, "MA_CP"),
                              (self.col_ma_benh, "MA_BENH")]:
                if c not in df.columns:
                    raise ValueError(f"Không tìm thấy cột '{c}' (dùng cho {label}) trong file.")

            def get_col(colname):
                if colname and colname != NONE_OPTION and colname in df.columns:
                    return df[colname].fillna("").astype(str).str.strip()
                return pd.Series([""] * len(df), index=df.index)

            rules = pd.DataFrame({
                "MA_CHUYEN_DE": get_col(self.col_ma_cd).str.upper(),
                "TEN_CHUYEN_DE": get_col(self.col_ten_cd),
                "MA_CP": get_col(self.col_ma_cp).str.upper(),
                "TEN_CP": get_col(self.col_ten_cp),
                "MA_BENH": get_col(self.col_ma_benh).str.upper(),
                "TEN_BENH": get_col(self.col_ten_benh),
                "NHOM_BENH": get_col(self.col_nhom_benh).str.upper(),
            })
            before_n = len(rules)
            rules = rules[
                (rules["MA_CHUYEN_DE"] != "") & (rules["MA_CP"] != "") & (rules["MA_BENH"] != "")
            ]
            rules = rules.drop_duplicates(subset=["MA_CHUYEN_DE", "MA_CP", "MA_BENH"])
            if len(rules) < before_n:
                self.log.emit(
                    f"Đã bỏ qua {before_n - len(rules)} dòng thiếu Mã chuyên đề/MA_CP/MA_BENH "
                    "hoặc trùng lặp ngay trong file."
                )
            self.progress.emit(35)

            if rules.empty:
                raise ValueError("Không có dòng quy tắc hợp lệ nào để nạp.")

            conn = sqlite3.connect(DB_PATH)
            try:
                self.progress.emit(55)
                self.log.emit(f"Đang nạp/cập nhật {len(rules)} dòng quy tắc vào CSDL...")
                total_after = upsert_rules(conn, rules)
                self.progress.emit(90)
            finally:
                conn.close()

            self.progress.emit(100)
            self.log.emit(
                f"Hoàn tất. Đã nạp {len(rules)} dòng quy tắc từ file. "
                f"Tổng số quy tắc hiện có trong CSDL: {total_after}."
            )
            self.finished_ok.emit(len(rules), total_after)

        except Exception as e:
            self.log.emit("LỖI: " + str(e))
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(e))


# ============================================================
# WORKER: THÊM NHANH QUY TẮC (DÁN DANH SÁCH MÃ BỆNH, KHÔNG CẦN FILE EXCEL)
# ============================================================

class QuickAddRuleWorker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished_ok = pyqtSignal(int, int)  # số dòng thêm, tổng số quy tắc trong CSDL sau khi thêm
    failed = pyqtSignal(str)

    def __init__(self, ma_chuyen_de, ten_chuyen_de, ma_cp, ten_cp, nhom_benh_override, parsed_lines):
        super().__init__()
        self.ma_chuyen_de = ma_chuyen_de
        self.ten_chuyen_de = ten_chuyen_de
        self.ma_cp = ma_cp
        self.ten_cp = ten_cp
        self.nhom_benh_override = nhom_benh_override
        self.parsed_lines = parsed_lines  # list[dict{MA_BENH, TEN_BENH, NHOM_BENH}]

    def run(self):
        try:
            if not self.parsed_lines:
                raise ValueError(
                    "Không phân tích được dòng mã bệnh nào từ văn bản đã dán. "
                    "Mỗi dòng cần có ít nhất 1 mã bệnh (cách các cột bởi Tab hoặc "
                    "nhiều khoảng trắng)."
                )
            self.progress.emit(20)

            rules = pd.DataFrame([
                {
                    "MA_CHUYEN_DE": normalize_code(self.ma_chuyen_de),
                    "TEN_CHUYEN_DE": self.ten_chuyen_de,
                    "MA_CP": normalize_code(self.ma_cp),
                    "TEN_CP": self.ten_cp,
                    "MA_BENH": item["MA_BENH"],
                    "TEN_BENH": item["TEN_BENH"],
                    "NHOM_BENH": normalize_code(self.nhom_benh_override) or item["NHOM_BENH"],
                }
                for item in self.parsed_lines
            ])
            rules = rules.drop_duplicates(subset=["MA_CHUYEN_DE", "MA_CP", "MA_BENH"])
            self.progress.emit(50)

            self.log.emit(f"Đang thêm {len(rules)} dòng quy tắc (Mã chuyên đề={self.ma_chuyen_de}, MA_CP={self.ma_cp})...")

            conn = sqlite3.connect(DB_PATH)
            try:
                total_after = upsert_rules(conn, rules)
            finally:
                conn.close()

            self.progress.emit(100)
            self.log.emit(
                f"Hoàn tất. Đã thêm {len(rules)} dòng quy tắc. "
                f"Tổng số quy tắc hiện có trong CSDL: {total_after}."
            )
            self.finished_ok.emit(len(rules), total_after)

        except Exception as e:
            self.log.emit("LỖI: " + str(e))
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(e))


# ============================================================
# WORKER: KIỂM TRA HỒ SƠ THEO QUY TẮC GIÁM ĐỊNH CHUYÊN ĐỀ
# ============================================================

class RuleCheckWorker(QThread):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    # violations_df, tổng số dòng, số dòng có quy tắc áp dụng, số dòng sai
    finished_ok = pyqtSignal(object, int, int, int)
    failed = pyqtSignal(str)

    def __init__(self, source_type, input_path, src_db_path, src_table_name,
                 col_ma_cp, col_ma_benh, col_ma_benh_khac,
                 col_ma_chuyen_de, prefix_match, ma_ly_do_tc, ly_do_tc_template):
        super().__init__()
        self.source_type = source_type  # "excel" hoặc "sqlite"
        self.input_path = input_path
        self.src_db_path = src_db_path
        self.src_table_name = src_table_name
        self.col_ma_cp = col_ma_cp
        self.col_ma_benh = col_ma_benh
        self.col_ma_benh_khac = col_ma_benh_khac
        self.col_ma_chuyen_de = col_ma_chuyen_de
        self.prefix_match = prefix_match
        self.ma_ly_do_tc = ma_ly_do_tc
        self.ly_do_tc_template = ly_do_tc_template

    @staticmethod
    def _get_val(row, colname):
        if colname and colname != NONE_OPTION:
            v = row.get(colname)
            return "" if pd.isna(v) else str(v).strip()
        return ""

    def run(self):
        try:
            if self.source_type == "sqlite":
                self.log.emit(f"Đang đọc bảng '{self.src_table_name}' từ CSDL SQLite ngoài...")
                src_conn = sqlite3.connect(self.src_db_path)
                try:
                    df = pd.read_sql_query(f'SELECT * FROM "{self.src_table_name}"', src_conn)
                finally:
                    src_conn.close()
            else:
                self.log.emit("Đang đọc file Excel hồ sơ cần kiểm tra...")
                df = read_table_any(self.input_path)
            self.progress.emit(10)

            if self.col_ma_cp not in df.columns:
                raise ValueError(f"Không tìm thấy cột '{self.col_ma_cp}' (dùng cho MA_CP) trong file.")
            use_ma_benh = bool(self.col_ma_benh) and self.col_ma_benh != NONE_OPTION
            use_ma_benh_khac = bool(self.col_ma_benh_khac) and self.col_ma_benh_khac != NONE_OPTION
            if not use_ma_benh and not use_ma_benh_khac:
                raise ValueError("Cần chọn ít nhất 1 trong 2 cột MA_BENH hoặc MA_BENH_KHAC để kiểm tra.")
            if use_ma_benh and self.col_ma_benh not in df.columns:
                raise ValueError(f"Không tìm thấy cột '{self.col_ma_benh}' trong file.")
            if use_ma_benh_khac and self.col_ma_benh_khac not in df.columns:
                raise ValueError(f"Không tìm thấy cột '{self.col_ma_benh_khac}' trong file.")
            use_ma_cd = bool(self.col_ma_chuyen_de) and self.col_ma_chuyen_de != NONE_OPTION
            if use_ma_cd and self.col_ma_chuyen_de not in df.columns:
                raise ValueError(f"Không tìm thấy cột '{self.col_ma_chuyen_de}' trong file.")

            self.log.emit("Đang nạp danh mục quy tắc từ CSDL...")
            conn = sqlite3.connect(DB_PATH)
            try:
                ensure_rules_table(conn)
                rules_df = pd.read_sql_query("SELECT * FROM QUY_TAC_BENH", conn)
            finally:
                conn.close()

            if rules_df.empty:
                raise ValueError(
                    "Chưa có quy tắc nào trong danh mục. Vui lòng nạp danh mục quy tắc ở "
                    "mục 1 trước khi kiểm tra."
                )
            self.progress.emit(25)

            rules_by_cd_cp = {}
            rules_by_cp = {}
            for _, r in rules_df.iterrows():
                cd, cp, benh, nhom = r["MA_CHUYEN_DE"], r["MA_CP"], r["MA_BENH"], r["NHOM_BENH"]
                entry_cd = rules_by_cd_cp.setdefault((cd, cp), {"codes": set(), "nhom": set()})
                entry_cd["codes"].add(benh)
                if nhom:
                    entry_cd["nhom"].add(nhom)
                entry_cp = rules_by_cp.setdefault(cp, {"codes": set(), "nhom": set()})
                entry_cp["codes"].add(benh)
                if nhom:
                    entry_cp["nhom"].add(nhom)

            self.log.emit(
                f"Đã nạp {len(rules_df)} dòng quy tắc, gồm {len(rules_by_cp)} MA_CP khác nhau."
            )

            violations = []
            total = len(df)
            so_co_quy_tac = 0
            row_dicts = df.to_dict("records")

            for i, row in enumerate(row_dicts):
                ma_cp = normalize_code(self._get_val(row, self.col_ma_cp))
                if ma_cp == "":
                    continue
                ma_chuyen_de = normalize_code(self._get_val(row, self.col_ma_chuyen_de)) if use_ma_cd else ""

                rule_entry = None
                if ma_chuyen_de and (ma_chuyen_de, ma_cp) in rules_by_cd_cp:
                    rule_entry = rules_by_cd_cp[(ma_chuyen_de, ma_cp)]
                elif ma_cp in rules_by_cp:
                    rule_entry = rules_by_cp[ma_cp]

                if rule_entry is None:
                    continue
                so_co_quy_tac += 1

                ma_benh = normalize_code(self._get_val(row, self.col_ma_benh)) if use_ma_benh else ""
                ma_benh_khac_raw = self._get_val(row, self.col_ma_benh_khac) if use_ma_benh_khac else ""
                ma_benh_khac_list = split_multi_codes(ma_benh_khac_raw)

                row_codes = ([ma_benh] if ma_benh else []) + ma_benh_khac_list
                allowed = rule_entry["codes"]

                matched = any(
                    code_matches(rc, ac, self.prefix_match)
                    for rc in row_codes for ac in allowed
                )

                if not matched:
                    out_row = dict(row)
                    out_row["MA_LY_DO_TC"] = self.ma_ly_do_tc
                    nhom_txt = ", ".join(sorted(rule_entry["nhom"])) if rule_entry["nhom"] else "(không xác định)"
                    codes_txt = ", ".join(sorted(allowed))
                    try:
                        ly_do = self.ly_do_tc_template.format(
                            MA_CP=ma_cp,
                            MA_BENH=ma_benh or "(trống)",
                            MA_BENH_KHAC=(";".join(ma_benh_khac_list) or "(trống)"),
                            NHOM_BENH=nhom_txt,
                            DANH_SACH_MA_BENH=codes_txt,
                        )
                    except (KeyError, IndexError):
                        ly_do = DEFAULT_LY_DO_TC_TEMPLATE.format(
                            MA_CP=ma_cp,
                            MA_BENH=ma_benh or "(trống)",
                            MA_BENH_KHAC=(";".join(ma_benh_khac_list) or "(trống)"),
                            NHOM_BENH=nhom_txt,
                            DANH_SACH_MA_BENH=codes_txt,
                        )
                    out_row["LY_DO_TC"] = ly_do
                    violations.append(out_row)

                if i % 500 == 0:
                    self.progress.emit(25 + int(i / max(total, 1) * 65))

            violations_df = pd.DataFrame(violations)
            self.progress.emit(100)
            self.log.emit(
                f"Đã kiểm tra {total} dòng, {so_co_quy_tac} dòng có quy tắc áp dụng theo MA_CP, "
                f"phát hiện {len(violations_df)} dòng chỉ định SAI quy định."
            )
            self.finished_ok.emit(violations_df, total, so_co_quy_tac, len(violations_df))

        except Exception as e:
            self.log.emit("LỖI: " + str(e))
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(e))


# ============================================================
# TRANG 4: LƯU HỒ SƠ ĐÃ TRỪ VÀO CSDL & KIỂM TRA TRÙNG
# ============================================================

class SaveDeductedPage(QWidget):
    def __init__(self):
        super().__init__()
        self.input_columns = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        layout.addWidget(page_title("Lưu hồ sơ đã trừ vào CSDL & Kiểm tra trùng"))
        note = QLabel(
            "Tải file Excel chứa danh sách hồ sơ/dòng chi phí ĐÃ TRỪ (ví dụ sheet "
            "Chi_Tiet xuất ra từ chức năng Tổng hợp trừ chi phí, hoặc file bạn tự "
            "chuẩn bị) để lưu vào 1 CSDL SQLite dùng chung, phục vụ đối chiếu về sau. "
            "Khi có đợt xử lý mới (kể cả với mã chuyên đề mới), dùng chức năng này để "
            "kiểm tra xem các hồ sơ có bị TRỪ TRÙNG với các lần trước hay không trước "
            "khi lưu chính thức.\n\n"
            "Ngoài 5 cột khai báo bên dưới, phần mềm tự động lấy thêm (nếu file có sẵn "
            "đúng tên cột, không cần chọn thủ công): MA_BN, HO_TEN, MA_THE, MA_BENH, "
            "MA_BENH_KHAC, NGAY_VAO, NGAY_RA, LOAI_CP, MA_CP, TEN_CP, SO_DANG_KY, SL_DC, "
            "DON_GIA_DC, TYLE_TT_DC, MUC_HUONG_DC, LY_DO_TC, MA_LY_DO_TC, KY_QT, "
            "SO_LUONG_DN, DON_GIA_DN, TYLE_TT_DN, MUC_HUONG_DN — để hồ sơ lưu trong CSDL "
            "đầy đủ thông tin phục vụ tra soát, đối chiếu trùng lặp."
        )
        note.setWordWrap(True)
        note.setObjectName("noteLabel")
        layout.addWidget(note)

        # --- Nhóm 1: File & khai báo cột ---
        in_group = QGroupBox("1. File Excel hồ sơ đã trừ & khai báo cột")
        in_form = QFormLayout()

        self.input_path_edit = QLineEdit()
        browse_in_btn = QPushButton("Chọn file Excel...")
        browse_in_btn.clicked.connect(self.browse_input)
        row1 = QHBoxLayout()
        row1.addWidget(self.input_path_edit)
        row1.addWidget(browse_in_btn)
        in_form.addRow("File Excel:", row1)

        self.read_cols_btn = QPushButton("Đọc cột của file")
        self.read_cols_btn.clicked.connect(self.load_columns)
        in_form.addRow("", self.read_cols_btn)

        self.col_ma_cskcb_combo = QComboBox()
        in_form.addRow("Cột MA_CSKCB:", self.col_ma_cskcb_combo)

        self.col_xml1_combo = QComboBox()
        in_form.addRow("Cột XML1_ID:", self.col_xml1_combo)

        self.col_id_cp_combo = QComboBox()
        in_form.addRow("Cột ID_CP (nếu có):", self.col_id_cp_combo)

        self.col_chuyen_de_combo = QComboBox()
        in_form.addRow("Cột Mã chuyên đề:", self.col_chuyen_de_combo)

        self.col_cong_van_combo = QComboBox()
        in_form.addRow("Cột Công văn (nếu có):", self.col_cong_van_combo)

        trang_thai_row = QHBoxLayout()
        self.col_trang_thai_combo = QComboBox()
        trang_thai_row.addWidget(self.col_trang_thai_combo)
        trang_thai_row.addWidget(QLabel("Giá trị coi là ĐÃ TRỪ:"))
        self.gia_tri_tru_edit = QLineEdit("Trừ")
        self.gia_tri_tru_edit.setMaximumWidth(120)
        trang_thai_row.addWidget(self.gia_tri_tru_edit)
        in_form.addRow(
            "Cột trạng thái (tuỳ chọn, dùng để\nchỉ lấy các dòng đã trừ):",
            trang_thai_row
        )

        in_group.setLayout(in_form)
        layout.addWidget(in_group)

        # --- Nhóm 2: CSDL lưu trữ (đường dẫn cố định, gắn kèm dự án) ---
        db_group = QGroupBox("2. CSDL lưu trữ (dùng chung 1 file DB.sqlite)")
        db_layout = QVBoxLayout()
        db_info = QLabel(
            f"📁 Dữ liệu được lưu vào bảng HO_SO_DA_TRU trong CSDL DÙNG CHUNG của "
            f"toàn phần mềm (gắn cố định kèm theo phần mềm, không cần chọn đường dẫn "
            f"mỗi lần dùng):\n{DB_PATH}"
        )
        db_info.setWordWrap(True)
        db_layout.addWidget(db_info)
        open_folder_btn = QPushButton("Mở thư mục chứa CSDL")
        open_folder_btn.clicked.connect(self.open_db_folder)
        db_layout.addWidget(open_folder_btn)
        db_group.setLayout(db_layout)
        layout.addWidget(db_group)

        # --- Nút hành động ---
        btn_row = QHBoxLayout()
        self.check_dup_btn = QPushButton("Kiểm tra trùng trong CSDL")
        self.check_dup_btn.clicked.connect(self.run_check_duplicates)
        self.save_btn = QPushButton("Lưu vào CSDL")
        self.save_btn.setObjectName("primaryBtn")
        self.save_btn.clicked.connect(self.run_save_deducted)
        btn_row.addWidget(self.check_dup_btn)
        btn_row.addWidget(self.save_btn)
        layout.addLayout(btn_row)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        layout.addWidget(QLabel("Kết quả kiểm tra trùng (nếu có):"))
        self.table_trung = QTableWidget()
        layout.addWidget(self.table_trung, stretch=2)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumHeight(140)
        layout.addWidget(self.log_box, stretch=1)

    def append_log(self, text):
        self.log_box.append(text)

    def browse_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel hồ sơ đã trừ", "",
            "Excel/CSV (*.xlsx *.xls *.csv);;Tất cả file (*.*)"
        )
        if not path:
            return
        self.input_path_edit.setText(path)
        self.load_columns()

    def load_columns(self):
        path = self.input_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file Excel trước.")
            return
        try:
            df = read_table_any(path)
            self.input_columns = list(df.columns)

            def fill_combo(combo, guesses, allow_none=False):
                combo.clear()
                if allow_none:
                    combo.addItem(NONE_OPTION)
                combo.addItems(self.input_columns)
                for g in guesses:
                    if g in self.input_columns:
                        combo.setCurrentText(g)
                        return g
                return None

            fill_combo(self.col_ma_cskcb_combo, ["MA_CSKCB"])
            fill_combo(self.col_xml1_combo, ["XML1_ID"])
            fill_combo(self.col_id_cp_combo, ["ID_CP"], allow_none=True)
            fill_combo(self.col_chuyen_de_combo, ["MA_CHUYEN_DE", "CHUYEN_DE"])
            fill_combo(self.col_cong_van_combo, ["CONG_VAN", "SO_CONG_VAN"], allow_none=True)
            fill_combo(self.col_trang_thai_combo, ["TRANG_THAI_TRU"], allow_none=True)

            self.append_log(f"Đã đọc {len(self.input_columns)} cột, {len(df)} dòng từ file.")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không đọc được file:\n{e}")

    def open_db_folder(self):
        folder = os.path.dirname(DB_PATH)
        os.makedirs(folder, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(folder))

    def _build_worker(self, mode):
        input_path = self.input_path_edit.text().strip()
        if not input_path or not os.path.isfile(input_path):
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file Excel hợp lệ.")
            return None
        db_path = DB_PATH
        col_ma_cskcb = self.col_ma_cskcb_combo.currentText().strip()
        col_xml1 = self.col_xml1_combo.currentText().strip()
        col_chuyen_de = self.col_chuyen_de_combo.currentText().strip()
        if not col_ma_cskcb or not col_xml1 or not col_chuyen_de:
            QMessageBox.warning(
                self, "Thiếu thông tin",
                "Vui lòng chọn đủ cột MA_CSKCB, XML1_ID và Mã chuyên đề."
            )
            return None
        col_id_cp = self.col_id_cp_combo.currentText().strip()
        col_cong_van = self.col_cong_van_combo.currentText().strip()
        col_trang_thai = self.col_trang_thai_combo.currentText().strip()
        gia_tri_tru = self.gia_tri_tru_edit.text().strip()

        if col_trang_thai and col_trang_thai != NONE_OPTION and not gia_tri_tru:
            QMessageBox.warning(
                self, "Thiếu thông tin",
                "Bạn đã chọn cột trạng thái, vui lòng nhập giá trị coi là ĐÃ TRỪ (ví dụ: Trừ)."
            )
            return None

        return DeductedDbWorker(
            db_path, input_path, col_ma_cskcb, col_xml1, col_id_cp,
            col_chuyen_de, col_cong_van, col_trang_thai, gia_tri_tru, mode
        )

    def run_check_duplicates(self):
        worker = self._build_worker("check")
        if worker is None:
            return
        self.check_dup_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()

        self.worker = worker
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished_check.connect(self.on_check_done)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def on_check_done(self, trung_df, so_trung, so_kiem_tra):
        self.check_dup_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        cols = ["MA_CSKCB", "XML1_ID", "ID_CP", "MA_CHUYEN_DE", "MA_BN", "HO_TEN",
                "CONG_VAN", "NGAY_LUU"]
        display_cols = [c for c in cols if c in trung_df.columns]
        fill_table_widget(self.table_trung, trung_df, display_cols, max_rows=500)
        if so_trung == 0:
            QMessageBox.information(
                self, "Kết quả kiểm tra",
                f"Không phát hiện trùng trong {so_kiem_tra} dòng đã kiểm tra."
            )
        else:
            QMessageBox.warning(
                self, "Phát hiện trùng",
                f"Có {so_trung}/{so_kiem_tra} dòng đã tồn tại trong CSDL "
                "(có thể đã bị trừ ở lần xử lý trước đó). Xem chi tiết ở bảng bên dưới."
            )

    def run_save_deducted(self):
        worker = self._build_worker("save")
        if worker is None:
            return
        reply = QMessageBox.question(
            self, "Xác nhận",
            f"Lưu dữ liệu từ file:\n{self.input_path_edit.text().strip()}\n"
            f"vào CSDL:\n{DB_PATH}\n\n"
            "Các dòng đã tồn tại sẵn (trùng XML1_ID + ID_CP + Mã chuyên đề) sẽ tự động "
            "được bỏ qua, không lưu lại lần 2. Tiếp tục?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self.check_dup_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()

        self.worker = worker
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished_save.connect(self.on_save_done)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def on_save_done(self, so_da_luu, so_trung_bo_qua, tong_so):
        self.check_dup_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        QMessageBox.information(
            self, "Hoàn tất",
            f"Đã lưu mới: {so_da_luu} dòng.\n"
            f"Bỏ qua (đã có sẵn / trùng): {so_trung_bo_qua} dòng.\n"
            f"Tổng số dòng đưa vào: {tong_so} dòng."
        )

    def on_failed(self, msg):
        self.check_dup_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        QMessageBox.critical(self, "Lỗi", msg)


# ============================================================
# TRANG 5: QUY TẮC GIÁM ĐỊNH THEO CHUYÊN ĐỀ
# ============================================================

class RuleAuditPage(QWidget):
    def __init__(self):
        super().__init__()
        self.rule_file_columns = []
        self.check_file_columns = []
        self.violations_df = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        layout.addWidget(page_title("Quy tắc giám định theo chuyên đề"))
        note = QLabel(
            "Định nghĩa quy tắc: với mỗi Mã chuyên đề + Mã chi phí (MA_CP), khai báo danh "
            "sách các Mã bệnh (MA_BENH, theo ICD) được phép chỉ định. Khi kiểm tra hồ sơ, "
            "nếu MA_CP được dùng nhưng cả MA_BENH và MA_BENH_KHAC (nhiều mã cách nhau bởi "
            "dấu ';') đều KHÔNG thuộc danh sách cho phép, hồ sơ đó bị coi là chỉ định SAI "
            "quy định và được xuất ra kèm mã/lý do từ chối."
        )
        note.setWordWrap(True)
        note.setObjectName("noteLabel")
        layout.addWidget(note)

        # --- Nhóm 1: Nạp danh mục quy tắc ---
        rule_group = QGroupBox("1. Danh mục quy tắc (Mã chuyên đề + MA_CP → các MA_BENH cho phép)")
        rule_form = QFormLayout()

        self.rule_path_edit = QLineEdit()
        browse_rule_btn = QPushButton("Chọn file Excel danh mục...")
        browse_rule_btn.clicked.connect(self.browse_rule_file)
        row1 = QHBoxLayout()
        row1.addWidget(self.rule_path_edit)
        row1.addWidget(browse_rule_btn)
        rule_form.addRow("File danh mục quy tắc:", row1)

        self.read_rule_cols_btn = QPushButton("Đọc cột của file")
        self.read_rule_cols_btn.clicked.connect(self.load_rule_columns)
        rule_form.addRow("", self.read_rule_cols_btn)

        self.rule_col_ma_cd_combo = QComboBox()
        rule_form.addRow("Cột Mã chuyên đề:", self.rule_col_ma_cd_combo)
        self.rule_col_ten_cd_combo = QComboBox()
        rule_form.addRow("Cột Tên chuyên đề (nếu có):", self.rule_col_ten_cd_combo)
        self.rule_col_ma_cp_combo = QComboBox()
        rule_form.addRow("Cột MA_CP:", self.rule_col_ma_cp_combo)
        self.rule_col_ten_cp_combo = QComboBox()
        rule_form.addRow("Cột Tên chi phí (nếu có):", self.rule_col_ten_cp_combo)
        self.rule_col_ma_benh_combo = QComboBox()
        rule_form.addRow("Cột Mã bệnh (ICD) cho phép:", self.rule_col_ma_benh_combo)
        self.rule_col_ten_benh_combo = QComboBox()
        rule_form.addRow("Cột Tên bệnh (nếu có):", self.rule_col_ten_benh_combo)
        self.rule_col_nhom_benh_combo = QComboBox()
        rule_form.addRow("Cột Nhóm bệnh (nếu có, vd BONG):", self.rule_col_nhom_benh_combo)

        rule_group.setLayout(rule_form)
        layout.addWidget(rule_group)

        self.load_rules_btn = QPushButton("Nạp / Cập nhật danh mục quy tắc vào CSDL")
        self.load_rules_btn.clicked.connect(self.run_load_rules)
        layout.addWidget(self.load_rules_btn)

        rules_db_info = QLabel(
            f"📁 Quy tắc được lưu vào bảng QUY_TAC_BENH trong CSDL DÙNG CHUNG của toàn "
            f"phần mềm (gắn cố định kèm phần mềm): {DB_PATH}\n"
            "Nạp lại file danh mục nhiều lần sẽ tự động CẬP NHẬT quy tắc trùng "
            "(Mã chuyên đề + MA_CP + MA_BENH) thay vì tạo bản ghi trùng."
        )
        rules_db_info.setWordWrap(True)
        rules_db_info.setObjectName("noteLabel")
        layout.addWidget(rules_db_info)

        # --- Nhóm 1b: Import nhanh quy tắc bằng cách dán danh sách mã bệnh ---
        quick_group = QGroupBox(
            "1b. Import nhanh quy tắc (dán danh sách mã bệnh, không cần file Excel)"
        )
        quick_layout = QVBoxLayout()
        quick_desc = QLabel(
            "Dùng khi bạn có sẵn 1 bảng mã bệnh dạng như ví dụ: dán trực tiếp các dòng "
            "đã copy từ Word/Excel vào ô bên dưới (mỗi dòng 1 mã bệnh, các cột cách nhau "
            "bởi Tab hoặc nhiều khoảng trắng, có thể có số thứ tự ở đầu dòng - phần mềm "
            "tự bỏ qua). Ví dụ:\n"
            "1.\tL55\tBỏng nắng\tBONG\n"
            "2.\tL55.0\tBỏng nắng độ một\tBONG"
        )
        quick_desc.setWordWrap(True)
        quick_layout.addWidget(quick_desc)

        quick_form = QFormLayout()
        self.quick_ma_cd_edit = QLineEdit()
        quick_form.addRow("Mã chuyên đề:", self.quick_ma_cd_edit)
        self.quick_ten_cd_edit = QLineEdit()
        quick_form.addRow("Tên chuyên đề (nếu có):", self.quick_ten_cd_edit)
        self.quick_ma_cp_edit = QLineEdit()
        quick_form.addRow("MA_CP:", self.quick_ma_cp_edit)
        self.quick_ten_cp_edit = QLineEdit()
        quick_form.addRow("Tên chi phí (nếu có):", self.quick_ten_cp_edit)
        self.quick_nhom_benh_edit = QLineEdit()
        quick_form.addRow(
            "Nhóm bệnh áp dụng cho mọi dòng dán\n(để trống nếu mỗi dòng đã có sẵn cột nhóm bệnh):",
            self.quick_nhom_benh_edit
        )
        quick_layout.addLayout(quick_form)

        self.quick_paste_edit = QTextEdit()
        self.quick_paste_edit.setPlaceholderText(
            "Dán danh sách mã bệnh vào đây, mỗi dòng 1 mã...\n"
            "1.\tL55\tBỏng nắng\tBONG\n2.\tL55.0\tBỏng nắng độ một\tBONG"
        )
        self.quick_paste_edit.setMaximumHeight(140)
        quick_layout.addWidget(self.quick_paste_edit)

        self.quick_add_btn = QPushButton("Import danh sách đã dán vào CSDL quy tắc")
        self.quick_add_btn.clicked.connect(self.run_quick_add)
        quick_layout.addWidget(self.quick_add_btn)

        quick_group.setLayout(quick_layout)
        layout.addWidget(quick_group)

        # --- Nhóm 2: Kiểm tra hồ sơ theo quy tắc ---
        check_group = QGroupBox("2. Kiểm tra hồ sơ theo quy tắc")
        check_form = QFormLayout()

        self.check_source_combo = QComboBox()
        self.check_source_combo.addItems(["File Excel", "CSDL SQLite ngoài (ví dụ XML1...)"])
        self.check_source_combo.currentIndexChanged.connect(self.on_check_source_changed)
        check_form.addRow("Nguồn dữ liệu hồ sơ:", self.check_source_combo)

        # -- Nguồn: File Excel --
        self.check_excel_widget = QWidget()
        excel_layout = QVBoxLayout(self.check_excel_widget)
        excel_layout.setContentsMargins(0, 0, 0, 0)
        self.check_path_edit = QLineEdit()
        browse_check_btn = QPushButton("Chọn file Excel hồ sơ...")
        browse_check_btn.clicked.connect(self.browse_check_file)
        row2 = QHBoxLayout()
        row2.addWidget(self.check_path_edit)
        row2.addWidget(browse_check_btn)
        excel_layout.addLayout(row2)
        check_form.addRow("File Excel hồ sơ:", self.check_excel_widget)

        # -- Nguồn: CSDL SQLite ngoài --
        self.check_db_widget = QWidget()
        db_src_layout = QVBoxLayout(self.check_db_widget)
        db_src_layout.setContentsMargins(0, 0, 0, 0)
        self.check_db_path_edit = QLineEdit()
        browse_check_db_btn = QPushButton("Chọn file CSDL SQLite ngoài...")
        browse_check_db_btn.clicked.connect(self.browse_check_db)
        row_db = QHBoxLayout()
        row_db.addWidget(self.check_db_path_edit)
        row_db.addWidget(browse_check_db_btn)
        db_src_layout.addLayout(row_db)
        self.check_db_table_combo = QComboBox()
        db_src_layout.addWidget(self.check_db_table_combo)
        check_form.addRow("CSDL SQLite ngoài:", self.check_db_widget)
        self.check_db_widget.setVisible(False)

        self.read_check_cols_btn = QPushButton("Đọc cột dữ liệu")
        self.read_check_cols_btn.clicked.connect(self.load_check_columns)
        check_form.addRow("", self.read_check_cols_btn)

        self.check_col_ma_cp_combo = QComboBox()
        check_form.addRow("Cột MA_CP:", self.check_col_ma_cp_combo)
        self.check_col_ma_benh_combo = QComboBox()
        check_form.addRow("Cột MA_BENH:", self.check_col_ma_benh_combo)
        self.check_col_ma_benh_khac_combo = QComboBox()
        check_form.addRow("Cột MA_BENH_KHAC:", self.check_col_ma_benh_khac_combo)
        self.check_col_ma_chuyen_de_combo = QComboBox()
        check_form.addRow("Cột Mã chuyên đề (nếu có):", self.check_col_ma_chuyen_de_combo)

        self.prefix_match_check = QCheckBox(
            "Khớp theo tiền tố mã bệnh (ví dụ mã gốc L55 trong quy tắc sẽ khớp cả L55.9 trong hồ sơ)"
        )
        check_form.addRow("", self.prefix_match_check)

        self.ma_ly_do_tc_edit = QLineEdit("CHOT_3")
        check_form.addRow("Mã lý do từ chối (MA_LY_DO_TC):", self.ma_ly_do_tc_edit)

        self.ly_do_tc_edit = QLineEdit(DEFAULT_LY_DO_TC_TEMPLATE)
        check_form.addRow("Nội dung lý do (LY_DO_TC):", self.ly_do_tc_edit)
        ly_do_hint = QLabel(
            "Có thể dùng các chỗ giữ chỗ: {MA_CP} {MA_BENH} {MA_BENH_KHAC} {NHOM_BENH} "
            "{DANH_SACH_MA_BENH}"
        )
        ly_do_hint.setWordWrap(True)
        ly_do_hint.setObjectName("noteLabel")
        check_form.addRow("", ly_do_hint)

        check_group.setLayout(check_form)
        layout.addWidget(check_group)

        btn_row = QHBoxLayout()
        self.check_btn = QPushButton("Kiểm tra hồ sơ")
        self.check_btn.setObjectName("primaryBtn")
        self.check_btn.clicked.connect(self.run_check)
        self.export_btn = QPushButton("Xuất dữ liệu SAI ra Excel")
        self.export_btn.clicked.connect(self.export_violations)
        self.export_btn.setEnabled(False)
        btn_row.addWidget(self.check_btn)
        btn_row.addWidget(self.export_btn)
        layout.addLayout(btn_row)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        layout.addWidget(QLabel("Chi tiết các bước đang xử lý:"))
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box, stretch=1)

    def append_log(self, text):
        self.log_box.append(text)

    # --- Nhóm 1: danh mục quy tắc ---

    def browse_rule_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel danh mục quy tắc", "",
            "Excel/CSV (*.xlsx *.xls *.csv);;Tất cả file (*.*)"
        )
        if not path:
            return
        self.rule_path_edit.setText(path)
        self.load_rule_columns()

    def load_rule_columns(self):
        path = self.rule_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file danh mục quy tắc trước.")
            return
        try:
            df = read_table_any(path)
            self.rule_file_columns = list(df.columns)

            def fill_combo(combo, guesses, allow_none=False):
                combo.clear()
                if allow_none:
                    combo.addItem(NONE_OPTION)
                combo.addItems(self.rule_file_columns)
                for g in guesses:
                    if g in self.rule_file_columns:
                        combo.setCurrentText(g)
                        return g
                return None

            fill_combo(self.rule_col_ma_cd_combo, ["MA_CHUYEN_DE", "MA_CD"])
            fill_combo(self.rule_col_ten_cd_combo, ["TEN_CHUYEN_DE"], allow_none=True)
            fill_combo(self.rule_col_ma_cp_combo, ["MA_CP"])
            fill_combo(self.rule_col_ten_cp_combo, ["TEN_CP"], allow_none=True)
            fill_combo(self.rule_col_ma_benh_combo, ["MA_BENH"])
            fill_combo(self.rule_col_ten_benh_combo, ["TEN_BENH"], allow_none=True)
            fill_combo(self.rule_col_nhom_benh_combo, ["NHOM_BENH"], allow_none=True)

            self.append_log(f"Đã đọc {len(self.rule_file_columns)} cột, {len(df)} dòng từ file danh mục.")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không đọc được file:\n{e}")

    def run_load_rules(self):
        path = self.rule_path_edit.text().strip()
        if not path or not os.path.isfile(path):
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file danh mục quy tắc hợp lệ.")
            return
        col_ma_cd = self.rule_col_ma_cd_combo.currentText().strip()
        col_ma_cp = self.rule_col_ma_cp_combo.currentText().strip()
        col_ma_benh = self.rule_col_ma_benh_combo.currentText().strip()
        if not col_ma_cd or not col_ma_cp or not col_ma_benh:
            QMessageBox.warning(
                self, "Thiếu thông tin",
                "Vui lòng chọn đủ cột Mã chuyên đề, MA_CP và Mã bệnh cho phép."
            )
            return

        self.load_rules_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()

        self.rule_worker = RuleLoadWorker(
            path, col_ma_cd,
            self.rule_col_ten_cd_combo.currentText().strip(),
            col_ma_cp,
            self.rule_col_ten_cp_combo.currentText().strip(),
            col_ma_benh,
            self.rule_col_ten_benh_combo.currentText().strip(),
            self.rule_col_nhom_benh_combo.currentText().strip(),
        )
        self.rule_worker.log.connect(self.append_log)
        self.rule_worker.progress.connect(self.progress.setValue)
        self.rule_worker.finished_ok.connect(self.on_rules_loaded)
        self.rule_worker.failed.connect(self.on_rule_failed)
        self.rule_worker.start()

    def on_rules_loaded(self, so_dong, tong_so):
        self.load_rules_btn.setEnabled(True)
        QMessageBox.information(
            self, "Hoàn tất",
            f"Đã nạp {so_dong} dòng quy tắc từ file.\n"
            f"Tổng số quy tắc hiện có trong CSDL: {tong_so}."
        )

    def on_rule_failed(self, msg):
        self.load_rules_btn.setEnabled(True)
        QMessageBox.critical(self, "Lỗi", msg)

    # --- Nhóm 1b: import nhanh (dán danh sách) ---

    def run_quick_add(self):
        ma_cd = self.quick_ma_cd_edit.text().strip()
        ma_cp = self.quick_ma_cp_edit.text().strip()
        if not ma_cd or not ma_cp:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng nhập Mã chuyên đề và MA_CP.")
            return

        text = self.quick_paste_edit.toPlainText()
        parsed = parse_pasted_benh_lines(text)
        if not parsed:
            QMessageBox.warning(
                self, "Không đọc được dữ liệu",
                "Không phân tích được dòng mã bệnh nào từ văn bản đã dán. Mỗi dòng cần "
                "có ít nhất 1 mã bệnh, các cột cách nhau bởi Tab hoặc nhiều khoảng trắng."
            )
            return

        preview = "\n".join(
            f"- {p['MA_BENH']}  {p['TEN_BENH']}  "
            f"[{self.quick_nhom_benh_edit.text().strip() or p['NHOM_BENH'] or '(không nhóm)'}]"
            for p in parsed[:10]
        )
        more = f"\n... và {len(parsed) - 10} dòng khác" if len(parsed) > 10 else ""
        reply = QMessageBox.question(
            self, "Xác nhận import",
            f"Sẽ thêm {len(parsed)} mã bệnh cho MA_CHUYEN_DE={ma_cd}, MA_CP={ma_cp}:\n\n"
            f"{preview}{more}\n\nTiếp tục?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self.quick_add_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()

        self.quick_worker = QuickAddRuleWorker(
            ma_cd, self.quick_ten_cd_edit.text().strip(),
            ma_cp, self.quick_ten_cp_edit.text().strip(),
            self.quick_nhom_benh_edit.text().strip(),
            parsed
        )
        self.quick_worker.log.connect(self.append_log)
        self.quick_worker.progress.connect(self.progress.setValue)
        self.quick_worker.finished_ok.connect(self.on_quick_add_done)
        self.quick_worker.failed.connect(self.on_quick_add_failed)
        self.quick_worker.start()

    def on_quick_add_done(self, so_dong, tong_so):
        self.quick_add_btn.setEnabled(True)
        QMessageBox.information(
            self, "Hoàn tất",
            f"Đã thêm {so_dong} dòng quy tắc.\nTổng số quy tắc hiện có trong CSDL: {tong_so}."
        )

    def on_quick_add_failed(self, msg):
        self.quick_add_btn.setEnabled(True)
        QMessageBox.critical(self, "Lỗi", msg)

    # --- Nhóm 2: kiểm tra hồ sơ ---

    def on_check_source_changed(self, idx):
        is_sqlite = (idx == 1)
        self.check_excel_widget.setVisible(not is_sqlite)
        self.check_db_widget.setVisible(is_sqlite)

    def browse_check_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel hồ sơ cần kiểm tra", "",
            "Excel/CSV (*.xlsx *.xls *.csv);;Tất cả file (*.*)"
        )
        if not path:
            return
        self.check_path_edit.setText(path)
        self.load_check_columns()

    def browse_check_db(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file CSDL SQLite ngoài", "",
            "SQLite Database (*.sqlite *.sqlite3 *.db);;Tất cả file (*.*)"
        )
        if not path:
            return
        self.check_db_path_edit.setText(path)
        try:
            conn = sqlite3.connect(path)
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
            tables = [r[0] for r in cur.fetchall()]
            conn.close()
            self.check_db_table_combo.clear()
            self.check_db_table_combo.addItems(tables)
            self.append_log(f"Đã tìm thấy {len(tables)} bảng trong CSDL.")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể đọc danh sách bảng:\n{e}")

    def load_check_columns(self):
        is_sqlite = (self.check_source_combo.currentIndex() == 1)
        try:
            if is_sqlite:
                db_path = self.check_db_path_edit.text().strip()
                table = self.check_db_table_combo.currentText().strip()
                if not db_path or not table:
                    QMessageBox.warning(
                        self, "Thiếu thông tin",
                        "Vui lòng chọn file CSDL SQLite và bảng dữ liệu trước."
                    )
                    return
                conn = sqlite3.connect(db_path)
                cur = conn.execute(f'PRAGMA table_info("{table}")')
                columns = [row[1] for row in cur.fetchall()]
                cur2 = conn.execute(f'SELECT COUNT(*) FROM "{table}"')
                row_count = cur2.fetchone()[0]
                conn.close()
                self.check_file_columns = columns
                self.append_log(f"Đã đọc {len(columns)} cột, {row_count} dòng từ bảng '{table}'.")
            else:
                path = self.check_path_edit.text().strip()
                if not path:
                    QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file hồ sơ trước.")
                    return
                df = read_table_any(path)
                self.check_file_columns = list(df.columns)
                self.append_log(f"Đã đọc {len(self.check_file_columns)} cột, {len(df)} dòng từ file hồ sơ.")

            def fill_combo(combo, guesses, allow_none=False):
                combo.clear()
                if allow_none:
                    combo.addItem(NONE_OPTION)
                combo.addItems(self.check_file_columns)
                for g in guesses:
                    if g in self.check_file_columns:
                        combo.setCurrentText(g)
                        return g
                return None

            fill_combo(self.check_col_ma_cp_combo, ["MA_CP"])
            fill_combo(self.check_col_ma_benh_combo, ["MA_BENH"])
            fill_combo(self.check_col_ma_benh_khac_combo, ["MA_BENH_KHAC"])
            fill_combo(self.check_col_ma_chuyen_de_combo, ["MA_CHUYEN_DE"], allow_none=True)
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không đọc được cột dữ liệu:\n{e}")

    def run_check(self):
        is_sqlite = (self.check_source_combo.currentIndex() == 1)
        input_path, src_db_path, src_table_name = None, None, None
        if is_sqlite:
            src_db_path = self.check_db_path_edit.text().strip()
            src_table_name = self.check_db_table_combo.currentText().strip()
            if not src_db_path or not os.path.isfile(src_db_path):
                QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file CSDL SQLite hợp lệ.")
                return
            if not src_table_name:
                QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn bảng dữ liệu.")
                return
        else:
            input_path = self.check_path_edit.text().strip()
            if not input_path or not os.path.isfile(input_path):
                QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn file hồ sơ hợp lệ.")
                return

        col_ma_cp = self.check_col_ma_cp_combo.currentText().strip()
        if not col_ma_cp:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng chọn cột MA_CP.")
            return
        col_ma_benh = self.check_col_ma_benh_combo.currentText().strip()
        col_ma_benh_khac = self.check_col_ma_benh_khac_combo.currentText().strip()
        if (not col_ma_benh or col_ma_benh == NONE_OPTION) and \
           (not col_ma_benh_khac or col_ma_benh_khac == NONE_OPTION):
            QMessageBox.warning(
                self, "Thiếu thông tin",
                "Vui lòng chọn ít nhất 1 trong 2 cột MA_BENH hoặc MA_BENH_KHAC."
            )
            return
        ma_ly_do_tc = self.ma_ly_do_tc_edit.text().strip() or "CHOT_3"
        ly_do_tc_template = self.ly_do_tc_edit.text().strip() or DEFAULT_LY_DO_TC_TEMPLATE

        self.check_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()

        self.check_worker = RuleCheckWorker(
            "sqlite" if is_sqlite else "excel", input_path, src_db_path, src_table_name,
            col_ma_cp, col_ma_benh, col_ma_benh_khac,
            self.check_col_ma_chuyen_de_combo.currentText().strip(),
            self.prefix_match_check.isChecked(),
            ma_ly_do_tc, ly_do_tc_template,
        )
        self.check_worker.log.connect(self.append_log)
        self.check_worker.progress.connect(self.progress.setValue)
        self.check_worker.finished_ok.connect(self.on_check_done)
        self.check_worker.failed.connect(self.on_check_failed)
        self.check_worker.start()

    def on_check_done(self, violations_df, total, so_co_quy_tac, so_vi_pham):
        self.violations_df = violations_df
        self.check_btn.setEnabled(True)
        self.export_btn.setEnabled(so_vi_pham > 0)
        if so_vi_pham == 0:
            QMessageBox.information(
                self, "Hoàn tất",
                f"Đã kiểm tra {total} dòng ({so_co_quy_tac} dòng có quy tắc áp dụng). "
                "Không phát hiện dòng nào chỉ định sai quy định."
            )
        else:
            QMessageBox.warning(
                self, "Phát hiện chỉ định sai",
                f"Đã kiểm tra {total} dòng ({so_co_quy_tac} dòng có quy tắc áp dụng).\n"
                f"Phát hiện {so_vi_pham} dòng chỉ định SAI quy định. Bấm 'Xuất dữ liệu SAI "
                "ra Excel' để xem chi tiết."
            )

    def on_check_failed(self, msg):
        self.check_btn.setEnabled(True)
        QMessageBox.critical(self, "Lỗi", msg)

    def export_violations(self):
        if self.violations_df is None or self.violations_df.empty:
            QMessageBox.warning(self, "Không có dữ liệu", "Chưa có dòng nào sai để xuất.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Lưu dữ liệu chỉ định sai", "ho_so_chi_dinh_sai.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            self.violations_df.to_excel(path, index=False)
            QMessageBox.information(self, "Thành công", f"Đã lưu file:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể lưu file:\n{e}")


# ============================================================
# GIAO DIỆN CHÍNH - MENU HIỆN ĐẠI (SIDEBAR)
# ============================================================

def page_title(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setObjectName("pageTitle")
    return lbl


APP_STYLE = """
QWidget {
    font-family: "Segoe UI", "Roboto", sans-serif;
    font-size: 13px;
    color: #1f2937;
}
QMainWindow, #contentArea {
    background-color: #eef6fb;
}
#sidebar {
    background-color: #0066a3;
}
#sidebar QListWidget {
    background-color: #0066a3;
    border: none;
    outline: 0;
    padding-top: 8px;
}
#sidebar QListWidget::item {
    color: #d6ecf9;
    padding: 14px 20px;
    border-left: 4px solid transparent;
}
#sidebar QListWidget::item:selected {
    background-color: #008acd;
    color: #ffffff;
    border-left: 4px solid #ffffff;
}
#sidebar QListWidget::item:hover:!selected {
    background-color: #00588c;
}
#appTitle {
    color: #ffffff;
    font-size: 16px;
    font-weight: 600;
    padding: 18px 20px 6px 20px;
}
#appSubtitle {
    color: #bfe1f5;
    font-size: 11px;
    padding: 0 20px 16px 20px;
    border-bottom: 1px solid #00588c;
}
#pageTitle {
    font-size: 19px;
    font-weight: 600;
    color: #0066a3;
    padding-bottom: 4px;
}
#noteLabel {
    color: #044e77;
    background-color: #e3f3fb;
    border: 1px solid #99d6f0;
    border-radius: 6px;
    padding: 8px 10px;
}
QGroupBox {
    background-color: #ffffff;
    border: 1px solid #cfe8f6;
    border-radius: 8px;
    margin-top: 10px;
    padding: 12px;
    font-weight: 600;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
    color: #0066a3;
}
QLineEdit, QComboBox {
    background-color: #ffffff;
    border: 1px solid #b9dcef;
    border-radius: 6px;
    padding: 5px 8px;
    min-height: 22px;
}
QLineEdit:focus, QComboBox:focus {
    border: 1px solid #008acd;
}
QPushButton {
    background-color: #e3f3fb;
    border: none;
    border-radius: 6px;
    padding: 7px 14px;
    color: #0066a3;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #cbe9f7;
}
QPushButton:disabled {
    background-color: #f3f4f6;
    color: #9ca3af;
}
QPushButton#primaryBtn {
    background-color: #008acd;
    color: #ffffff;
    font-weight: 600;
}
QPushButton#primaryBtn:hover {
    background-color: #0066a3;
}
QPushButton#primaryBtn:disabled {
    background-color: #9ed4ee;
    color: #f0f9ff;
}
QTableWidget {
    background-color: #ffffff;
    border: 1px solid #cfe8f6;
    border-radius: 6px;
    gridline-color: #eef6fb;
}
QHeaderView::section {
    background-color: #e3f3fb;
    color: #0066a3;
    padding: 6px;
    border: none;
    border-bottom: 1px solid #cfe8f6;
    font-weight: 600;
}
QProgressBar {
    background-color: #dceefa;
    border-radius: 6px;
    text-align: center;
    height: 16px;
}
QProgressBar::chunk {
    background-color: #008acd;
    border-radius: 6px;
}
QTextEdit {
    background-color: #04324d;
    color: #d6ecf9;
    border-radius: 6px;
    padding: 6px;
    font-family: Consolas, monospace;
}
"""


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Phần mềm nghiệp vụ BHYT - Giám định XML1")
        self.resize(1250, 820)

        central = QWidget()
        central.setObjectName("contentArea")
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # --- Sidebar ---
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(260)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)

        title = QLabel("BHYT - GIÁM ĐỊNH")
        title.setObjectName("appTitle")
        subtitle = QLabel("Công cụ xử lý dữ liệu XML1")
        subtitle.setObjectName("appSubtitle")
        sidebar_layout.addWidget(title)
        sidebar_layout.addWidget(subtitle)

        self.menu_list = QListWidget()
        self.menu_list.setSpacing(2)
        menu_items = [
            "🔎  Tra cứu dữ liệu XML1",
            "🗂  Tách Excel theo MA_CSKCB",
            "📊  Tổng hợp trừ chi phí theo chuyên đề",
            "💾  Lưu hồ sơ đã trừ & Kiểm tra trùng",
            "⚖️  Quy tắc giám định theo chuyên đề",
        ]
        for text in menu_items:
            item = QListWidgetItem(text)
            item.setSizeHint(item.sizeHint())
            self.menu_list.addItem(item)
        self.menu_list.currentRowChanged.connect(self.change_page)
        sidebar_layout.addWidget(self.menu_list, stretch=1)

        root_layout.addWidget(sidebar)

        # --- Content stack ---
        self.stack = QStackedWidget()
        self.stack.addWidget(LookupPage())
        self.stack.addWidget(SplitPage())
        self.stack.addWidget(ExcludeSummaryPage())
        self.stack.addWidget(SaveDeductedPage())
        self.stack.addWidget(RuleAuditPage())
        root_layout.addWidget(self.stack, stretch=1)

        self.setCentralWidget(central)
        self.menu_list.setCurrentRow(0)

    def change_page(self, index):
        if index >= 0:
            self.stack.setCurrentIndex(index)


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
