"""Quản lý danh mục loại KCB và cơ sở KCB trong CSDL dùng chung."""
import sqlite3
from contextlib import closing

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QLabel, QLineEdit,
    QPushButton, QGroupBox, QTableWidget, QTableWidgetItem, QAbstractItemView,
    QHeaderView, QMessageBox, QSplitter,
)


CATALOGS = {
    "loai_kcb": ("DM_LOAI_KCB", "MA_LOAI_KCB", "TEN_LOAI_KCB", "Danh mục loại khám chữa bệnh"),
    "cskcb": ("DM_CSKCB", "MA_CSKCB", "TEN_CSKCB", "Danh mục cơ sở khám chữa bệnh"),
}


class CatalogStore:
    def __init__(self, db_path, kind):
        self.db_path = db_path
        self.table, self.code_column, self.name_column, self.title = CATALOGS[kind]
        with closing(sqlite3.connect(db_path)) as conn, conn:
            conn.execute(
                f'CREATE TABLE IF NOT EXISTS "{self.table}" ('
                f'"{self.code_column}" TEXT PRIMARY KEY NOT NULL, '
                f'"{self.name_column}" TEXT NOT NULL, GHI_CHU TEXT NOT NULL DEFAULT \'\')'
            )

    def rows(self):
        with closing(sqlite3.connect(self.db_path)) as conn:
            return conn.execute(
                f'SELECT "{self.code_column}", "{self.name_column}", GHI_CHU '
                f'FROM "{self.table}" ORDER BY "{self.code_column}"'
            ).fetchall()

    def save(self, code, name, note, editing=False):
        if not code.strip() or not name.strip():
            raise ValueError("Vui lòng nhập đầy đủ mã và tên.")
        with closing(sqlite3.connect(self.db_path)) as conn, conn:
            if editing:
                cursor = conn.execute(
                    f'UPDATE "{self.table}" SET "{self.name_column}"=?, GHI_CHU=? '
                    f'WHERE "{self.code_column}"=?', (name.strip(), note.strip(), code.strip())
                )
                if cursor.rowcount != 1:
                    raise ValueError("Mã đã bị xóa. Hãy tải lại danh sách.")
            else:
                conn.execute(
                    f'INSERT INTO "{self.table}" ("{self.code_column}", "{self.name_column}", GHI_CHU) '
                    'VALUES (?, ?, ?)', (code.strip(), name.strip(), note.strip())
                )

    def delete(self, code):
        with closing(sqlite3.connect(self.db_path)) as conn, conn:
            conn.execute(f'DELETE FROM "{self.table}" WHERE "{self.code_column}"=?', (code,))


class CatalogPage(QWidget):
    def __init__(self, db_path, kind):
        super().__init__()
        self.store = CatalogStore(db_path, kind)
        self.editing_code = None
        self.rows = []
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        title = QLabel(self.store.title)
        title.setObjectName("pageTitle")
        layout.addWidget(title)
        note = QLabel("Chọn một dòng để sửa tên hoặc ghi chú. Bấm Thêm mới để nhập mã khác. Mã được giữ nguyên khi sửa và giữ số 0 ở đầu.")
        note.setWordWrap(True)
        note.setObjectName("noteLabel")
        layout.addWidget(note)
        columns = QSplitter(Qt.Orientation.Horizontal)
        columns.setChildrenCollapsible(False)
        form_group = QGroupBox("1. Thêm / sửa danh mục")
        form_layout = QVBoxLayout(form_group)
        form = QFormLayout()
        self.code_edit = QLineEdit()
        self.name_edit = QLineEdit()
        self.note_edit = QLineEdit()
        form.addRow(self.store.code_column + ":", self.code_edit)
        form.addRow("Tên:", self.name_edit)
        form.addRow("Ghi chú:", self.note_edit)
        form_layout.addLayout(form)
        buttons = QHBoxLayout()
        self.new_btn = QPushButton("Thêm mới")
        self.new_btn.clicked.connect(self.new_record)
        self.save_btn = QPushButton("Lưu danh mục")
        self.save_btn.setObjectName("primaryBtn")
        self.save_btn.clicked.connect(self.save_record)
        buttons.addWidget(self.new_btn)
        buttons.addWidget(self.save_btn)
        form_layout.addLayout(buttons)
        form_layout.addStretch()
        columns.addWidget(form_group)
        list_group = QGroupBox("2. Danh sách danh mục")
        list_layout = QVBoxLayout(list_group)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Tìm theo mã, tên hoặc ghi chú...")
        self.search_edit.textChanged.connect(self.filter_rows)
        list_layout.addWidget(self.search_edit)
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels([self.store.code_column, "Tên", "Ghi chú"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.itemSelectionChanged.connect(self.select_record)
        list_layout.addWidget(self.table)
        actions = QHBoxLayout()
        self.refresh_btn = QPushButton("Tải lại danh sách")
        self.refresh_btn.clicked.connect(self.refresh)
        self.delete_btn = QPushButton("Xóa dòng đã chọn")
        self.delete_btn.clicked.connect(self.delete_record)
        actions.addWidget(self.refresh_btn)
        actions.addWidget(self.delete_btn)
        list_layout.addLayout(actions)
        self.status = QLabel()
        list_layout.addWidget(self.status)
        columns.addWidget(list_group)
        columns.setSizes([440, 760])
        columns.setStretchFactor(0, 1)
        columns.setStretchFactor(1, 2)
        layout.addWidget(columns, 1)
        self.refresh()

    def new_record(self):
        self.table.clearSelection()
        self.editing_code = None
        self.code_edit.setReadOnly(False)
        self.code_edit.clear()
        self.name_edit.clear()
        self.note_edit.clear()
        self.code_edit.setFocus()

    def refresh(self):
        try:
            self.rows = self.store.rows()
            self.new_record()
            self.filter_rows()
        except Exception as exc:
            QMessageBox.critical(self, "Lỗi đọc danh mục", str(exc))

    def filter_rows(self):
        query = self.search_edit.text().strip().casefold()
        self.visible_rows = [row for row in self.rows if any(query in str(v or "").casefold() for v in row)]
        self.table.blockSignals(True)
        self.table.clearSelection()
        self.table.setRowCount(len(self.visible_rows))
        for i, row in enumerate(self.visible_rows):
            for j, value in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(value or "")))
        self.table.blockSignals(False)
        self.status.setText(f"Hiển thị {len(self.visible_rows)} / {len(self.rows)} danh mục")

    def select_record(self):
        if not self.table.selectedItems():
            return
        index = self.table.currentRow()
        if 0 <= index < len(self.visible_rows):
            code, name, note = self.visible_rows[index]
            self.editing_code = code
            self.code_edit.setText(code)
            self.code_edit.setReadOnly(True)
            self.name_edit.setText(name or "")
            self.note_edit.setText(note or "")

    def save_record(self):
        try:
            self.store.save(self.editing_code if self.editing_code is not None else self.code_edit.text(),
                            self.name_edit.text(), self.note_edit.text(), self.editing_code is not None)
            self.refresh()
            self.status.setText("Đã lưu danh mục thành công.")
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Mã đã tồn tại", "Mã này đã có. Chọn dòng trong danh sách để sửa.")
        except ValueError as exc:
            QMessageBox.warning(self, "Kiểm tra thông tin", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "Không lưu được danh mục", str(exc))

    def delete_record(self):
        if not self.table.selectedItems():
            QMessageBox.warning(self, "Chưa chọn dòng", "Vui lòng chọn một dòng cần xóa.")
            return
        code = self.visible_rows[self.table.currentRow()][0]
        reply = QMessageBox.question(self, "Xác nhận xóa", f"Xóa mã '{code}' khỏi danh mục?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            self.store.delete(code)
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "Không xóa được danh mục", str(exc))
