"""Ribbon và cửa sổ chính; nhận các lớp trang từ ứng dụng, không chứa nghiệp vụ.

Thêm chức năng: khai báo RibbonCommand trong RIBBON_TABS và truyền lớp trang
ở cùng vị trí page_index trong danh sách page_factories của RibbonMainWindow.
"""
from dataclasses import dataclass
from collections.abc import Callable, Sequence

from PyQt6.QtCore import QSize, Qt, pyqtSignal
from PyQt6.QtWidgets import (
    QButtonGroup, QFrame, QHBoxLayout, QLabel, QMainWindow, QScrollArea,
    QStackedWidget, QStyle, QTabWidget, QToolButton, QVBoxLayout, QWidget,
)


@dataclass(frozen=True)
class RibbonCommand:
    page_index: int
    text: str
    tooltip: str
    icon: QStyle.StandardPixmap


RIBBON_TABS = (
    ("Danh mục", (
        ("Quản lý CSDL", (
            RibbonCommand(8, "Danh mục\nloại KCB", "Quản lý MA_LOAI_KCB: thêm, sửa, xóa và tìm kiếm", QStyle.StandardPixmap.SP_FileDialogListView),
            RibbonCommand(9, "Danh mục\ncơ sở KCB", "Quản lý MA_CSKCB: thêm, sửa, xóa và tìm kiếm", QStyle.StandardPixmap.SP_DirHomeIcon),
        )),
    )),
    ("Nghiệp vụ BHYT", (
        ("Tra cứu", (
            RibbonCommand(0, "Tra cứu\nXML1", "Tra cứu dữ liệu XML1 từ CSDL SQLite", QStyle.StandardPixmap.SP_FileDialogContentsView),
        )),
        ("Chi phí & hồ sơ", (
            RibbonCommand(2, "Tổng hợp\ntrừ chi phí", "Tổng hợp trừ chi phí theo cơ sở khám chữa bệnh và chuyên đề", QStyle.StandardPixmap.SP_FileDialogDetailedView),
            RibbonCommand(3, "Lưu hồ sơ\nKiểm tra trùng", "Lưu hồ sơ đã trừ và kiểm tra dữ liệu trùng", QStyle.StandardPixmap.SP_DialogSaveButton),
        )),
    )),
    ("Chuyên đề", (
        ("Quản lý chuyên đề", (
            RibbonCommand(4, "Định nghĩa\nchuyên đề SQL", "Thêm, sửa, xóa định nghĩa chuyên đề SQL", QStyle.StandardPixmap.SP_FileDialogInfoView),
            RibbonCommand(5, "Chạy chuyên đề\ntheo kỳ", "Chạy chuyên đề theo kỳ giám định", QStyle.StandardPixmap.SP_MediaPlay),
        )),
    )),
    ("Công cụ Excel", (
        ("Tách dữ liệu", (
            RibbonCommand(1, "Tách Excel\ntheo MA_CSKCB", "Tách file Excel theo MA_CSKCB hoặc cột được chọn", QStyle.StandardPixmap.SP_DirOpenIcon),
        )),
        ("Ghép dữ liệu", (
            RibbonCommand(6, "Ghép thư mục\nthành nhiều sheet", "Ghép các file Excel trong thư mục thành một file nhiều sheet", QStyle.StandardPixmap.SP_DirIcon),
            RibbonCommand(7, "Ghép nhiều sheet\nthành một sheet", "Ghép các sheet trong một file Excel thành một sheet", QStyle.StandardPixmap.SP_FileIcon),
        )),
    )),
)

RIBBON_STYLE = """
QTabWidget#ribbon::pane { border: 0; border-bottom: 1px solid #c3c6d6; background: white; }
QTabWidget#ribbon > QTabBar::tab { background: #e9edff; padding: 8px 22px; border: 0; }
QTabWidget#ribbon > QTabBar::tab:selected { background: white; color: #003d9b; border-top: 3px solid #003d9b; }
QTabWidget#ribbon > QTabBar::tab:hover { color: #003d9b; }
QWidget#ribbonPage, QFrame#ribbonGroup { background: white; }
QFrame#ribbonGroup { border-right: 1px solid #d8e2ff; }
QLabel#ribbonGroupTitle { color: #434654; font-size: 11px; }
QToolButton#ribbonCommand { background: white; border: 1px solid transparent; border-radius: 5px; padding: 5px; font-size: 12px; }
QToolButton#ribbonCommand:hover { background: #f1f3ff; border-color: #b2c5ff; }
QToolButton#ribbonCommand:checked { background: #dae2ff; border-color: #003d9b; color: #003d9b; }
QLabel#ribbonBrand { background: #003d9b; color: white; font-size: 15px; font-weight: 600; padding: 8px 16px; }
"""


class RibbonBar(QTabWidget):
    pageRequested = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ribbon")
        self.setStyleSheet(RIBBON_STYLE)
        self.setFixedHeight(166)
        self.buttons = {}
        self.page_tabs = {}
        self.button_group = QButtonGroup(self)
        self.button_group.setExclusive(True)
        self.button_group.idClicked.connect(self.pageRequested.emit)
        for tab_index, (tab_title, groups) in enumerate(RIBBON_TABS):
            page = QWidget()
            page.setObjectName("ribbonPage")
            row = QHBoxLayout(page)
            row.setContentsMargins(8, 4, 8, 4)
            row.setSpacing(8)
            for group_title, commands in groups:
                group = QFrame()
                group.setObjectName("ribbonGroup")
                column = QVBoxLayout(group)
                column.setContentsMargins(4, 0, 12, 0)
                column.setSpacing(2)
                buttons = QHBoxLayout()
                for command in commands:
                    button = QToolButton()
                    button.setObjectName("ribbonCommand")
                    button.setText(command.text)
                    button.setToolTip(command.tooltip)
                    button.setAccessibleName(command.tooltip)
                    button.setIcon(self.style().standardIcon(command.icon))
                    button.setIconSize(QSize(32, 32))
                    button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
                    button.setCursor(Qt.CursorShape.PointingHandCursor)
                    button.setCheckable(True)
                    button.setFixedSize(144, 88)
                    self.button_group.addButton(button, command.page_index)
                    self.buttons[command.page_index] = button
                    self.page_tabs[command.page_index] = tab_index
                    buttons.addWidget(button)
                column.addLayout(buttons)
                caption = QLabel(group_title)
                caption.setObjectName("ribbonGroupTitle")
                caption.setAlignment(Qt.AlignmentFlag.AlignCenter)
                column.addWidget(caption)
                row.addWidget(group)
            row.addStretch()
            scroll = QScrollArea()
            scroll.setFrameShape(QFrame.Shape.NoFrame)
            scroll.setWidgetResizable(True)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            scroll.setWidget(page)
            self.addTab(scroll, tab_title)

    def set_active_page(self, index: int):
        if index in self.buttons:
            self.buttons[index].setChecked(True)
            self.setCurrentIndex(self.page_tabs[index])


class RibbonMainWindow(QMainWindow):
    def __init__(self, page_factories: Sequence[Callable[[], QWidget]]):
        super().__init__()
        if len(page_factories) != len([c for _, groups in RIBBON_TABS for _, commands in groups for c in commands]):
            raise ValueError("Danh sách trang phải khớp với các chức năng ribbon.")
        self.setWindowTitle("Phần mềm nghiệp vụ BHYT - Giám định XML1")
        self.resize(1280, 940)
        central = QWidget()
        central.setObjectName("contentArea")
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        brand = QLabel("BHYT · GIÁM ĐỊNH XML1")
        brand.setObjectName("ribbonBrand")
        brand.setStyleSheet(RIBBON_STYLE)
        layout.addWidget(brand)
        self.ribbon = RibbonBar()
        layout.addWidget(self.ribbon)
        self.stack = QStackedWidget()
        for factory in page_factories:
            self.stack.addWidget(factory())
        layout.addWidget(self.stack, 1)
        self.setCentralWidget(central)
        self.ribbon.pageRequested.connect(self._set_active_index)
        self.stack.currentChanged.connect(self.ribbon.set_active_page)
        self._set_active_index(0)

    def _set_active_index(self, index: int):
        if 0 <= index < self.stack.count():
            self.stack.setCurrentIndex(index)
            self.ribbon.set_active_page(index)
