import sys
import json
import logging
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Set

import pandas as pd

from PySide6.QtCore import Qt, QPointF, Signal, QMimeData
from PySide6.QtGui import QAction, QColor, QBrush, QPen, QDrag, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QGroupBox,
    QHeaderView,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QDoubleSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
    QGraphicsItem,
    QGraphicsRectItem,
    QGraphicsScene,
    QGraphicsSimpleTextItem,
    QGraphicsView,
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
LOG_DIR = APP_DIR / "logs"
STYLE_PATH = APP_DIR / "styles" / "app_style.qss"
DB_PATH = DATA_DIR / "parts.db"
TEMPLATE_PATH = DATA_DIR / "parts_upload_template.xlsx"
RAW_EXPORT_DEFAULT = DATA_DIR / "parts_raw_data.xlsx"
CANVAS_SAVE_PATH = DATA_DIR / "breaker_canvas_layout.json"

DATA_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)


# ------------------------------------------------------------
# Logging
# ------------------------------------------------------------
def setup_logging() -> logging.Logger:
    logger = logging.getLogger("electrical_capacity_app")
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(funcName)s | %(message)s"
    )

    file_handler = logging.FileHandler(LOG_DIR / "app.log", encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    logger.debug("Logging initialized")
    return logger


logger = setup_logging()


# ------------------------------------------------------------
# Data layer
# ------------------------------------------------------------
@dataclass
class PartRecord:
    part_no: str
    part_name: str
    category: str
    voltage_v: float
    current_a: float
    power_w: float
    phase: str
    power_factor: float
    recommended_breaker_a: float
    note: str = ""


class DatabaseManager:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        logger.debug("DatabaseManager init: %s", db_path)
        self._initialize()

    def _connect(self):
        return sqlite3.connect(self.db_path)

    def _initialize(self):
        logger.info("Initializing database")
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS parts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    part_no TEXT NOT NULL UNIQUE,
                    part_name TEXT NOT NULL,
                    category TEXT,
                    voltage_v REAL NOT NULL,
                    current_a REAL NOT NULL,
                    power_w REAL NOT NULL,
                    phase TEXT NOT NULL,
                    power_factor REAL NOT NULL,
                    recommended_breaker_a REAL NOT NULL,
                    note TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            conn.commit()
        logger.info("Database ready")

    def upsert_part(self, part: PartRecord):
        logger.info("Upserting part: %s / %s", part.part_no, part.part_name)
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO parts (
                    part_no, part_name, category, voltage_v, current_a, power_w,
                    phase, power_factor, recommended_breaker_a, note
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(part_no) DO UPDATE SET
                    part_name=excluded.part_name,
                    category=excluded.category,
                    voltage_v=excluded.voltage_v,
                    current_a=excluded.current_a,
                    power_w=excluded.power_w,
                    phase=excluded.phase,
                    power_factor=excluded.power_factor,
                    recommended_breaker_a=excluded.recommended_breaker_a,
                    note=excluded.note
                """,
                (
                    part.part_no,
                    part.part_name,
                    part.category,
                    part.voltage_v,
                    part.current_a,
                    part.power_w,
                    part.phase,
                    part.power_factor,
                    part.recommended_breaker_a,
                    part.note,
                ),
            )
            conn.commit()

    def get_all_parts(self) -> pd.DataFrame:
        with self._connect() as conn:
            return pd.read_sql_query(
                "SELECT part_no, part_name, category, voltage_v, current_a, power_w, phase, power_factor, recommended_breaker_a, note, created_at FROM parts ORDER BY part_no",
                conn,
            )

    def get_part_by_no(self, part_no: str) -> Optional[pd.Series]:
        with self._connect() as conn:
            df = pd.read_sql_query(
                "SELECT * FROM parts WHERE part_no = ?",
                conn,
                params=(part_no,),
            )
        if df.empty:
            return None
        return df.iloc[0]

    def get_part_nos(self) -> List[str]:
        with self._connect() as conn:
            rows = conn.execute("SELECT part_no FROM parts ORDER BY part_no").fetchall()
        return [r[0] for r in rows]

    def bulk_upsert_from_dataframe(self, df: pd.DataFrame) -> int:
        required_columns = [
            "part_no",
            "part_name",
            "category",
            "voltage_v",
            "current_a",
            "power_w",
            "phase",
            "power_factor",
            "recommended_breaker_a",
            "note",
        ]
        missing = [c for c in required_columns if c not in df.columns]
        if missing:
            raise ValueError(f"엑셀 양식 컬럼이 누락되었습니다: {missing}")

        count = 0
        for _, row in df.iterrows():
            part = PartRecord(
                part_no=str(row["part_no"]).strip(),
                part_name=str(row["part_name"]).strip(),
                category=str(row.get("category", "")).strip(),
                voltage_v=float(row["voltage_v"]),
                current_a=float(row["current_a"]),
                power_w=float(row["power_w"]),
                phase=str(row["phase"]).strip(),
                power_factor=float(row["power_factor"]),
                recommended_breaker_a=float(row["recommended_breaker_a"]),
                note=str(row.get("note", "")).strip(),
            )
            self.upsert_part(part)
            count += 1
        return count


class ExcelManager:
    TEMPLATE_COLUMNS = {
        "part_no": ["MTR-001"],
        "part_name": ["Main Motor"],
        "category": ["Motor"],
        "voltage_v": [220],
        "current_a": [5.2],
        "power_w": [1144],
        "phase": ["1P"],
        "power_factor": [0.95],
        "recommended_breaker_a": [15],
        "note": ["Sample row. Delete after checking format."],
    }

    @staticmethod
    def create_template(path: Path):
        df = pd.DataFrame(ExcelManager.TEMPLATE_COLUMNS)
        df.to_excel(path, index=False)

    @staticmethod
    def load_excel(path: Path) -> pd.DataFrame:
        return pd.read_excel(path)

    @staticmethod
    def export_dataframe(df: pd.DataFrame, path: Path):
        df.to_excel(path, index=False)


class LoadCalculator:
    BREAKER_STANDARDS = [5, 10, 15, 20, 30, 40, 50, 60, 75, 100, 125, 150, 175, 200, 225, 250, 300, 400]

    @staticmethod
    def calculate_total(part_row: pd.Series, quantity: int, safety_factor: float = 1.25) -> dict:
        total_current = float(part_row["current_a"]) * quantity
        total_power = float(part_row["power_w"]) * quantity
        safety_current = total_current * safety_factor
        suggested_breaker = LoadCalculator.select_breaker(safety_current)
        return {
            "part_no": part_row["part_no"],
            "part_name": part_row["part_name"],
            "quantity": quantity,
            "voltage_v": float(part_row["voltage_v"]),
            "phase": part_row["phase"],
            "unit_current_a": float(part_row["current_a"]),
            "unit_power_w": float(part_row["power_w"]),
            "total_current_a": round(total_current, 2),
            "total_power_w": round(total_power, 2),
            "safety_factor": round(safety_factor, 2),
            "safety_current_a": round(safety_current, 2),
            "suggested_breaker_a": suggested_breaker,
        }

    @staticmethod
    def select_breaker(required_current: float) -> int:
        if required_current < 0:
            required_current = 0
        for size in LoadCalculator.BREAKER_STANDARDS:
            if size >= required_current:
                return size
        return LoadCalculator.BREAKER_STANDARDS[-1]


# ------------------------------------------------------------
# Parts / Calculator tabs
# ------------------------------------------------------------
class PartsTab(QWidget):
    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db = db
        self._build_ui()
        self.refresh_table()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        form_group = QGroupBox("파트 등록")
        form_layout = QGridLayout()

        self.part_no_edit = QLineEdit()
        self.part_name_edit = QLineEdit()
        self.category_edit = QLineEdit()
        self.voltage_spin = QDoubleSpinBox()
        self.voltage_spin.setRange(0, 100000)
        self.voltage_spin.setDecimals(2)
        self.voltage_spin.setValue(220)

        self.current_spin = QDoubleSpinBox()
        self.current_spin.setRange(0, 100000)
        self.current_spin.setDecimals(2)

        self.power_spin = QDoubleSpinBox()
        self.power_spin.setRange(0, 10000000)
        self.power_spin.setDecimals(2)

        self.phase_combo = QComboBox()
        self.phase_combo.addItems(["1P", "3P", "DC"])

        self.pf_spin = QDoubleSpinBox()
        self.pf_spin.setRange(0, 1)
        self.pf_spin.setSingleStep(0.01)
        self.pf_spin.setValue(0.95)

        self.breaker_spin = QDoubleSpinBox()
        self.breaker_spin.setRange(0, 10000)
        self.breaker_spin.setDecimals(0)

        self.note_edit = QLineEdit()

        widgets = [
            ("Part No", self.part_no_edit),
            ("Part Name", self.part_name_edit),
            ("Category", self.category_edit),
            ("Voltage (V)", self.voltage_spin),
            ("Current (A)", self.current_spin),
            ("Power (W)", self.power_spin),
            ("Phase", self.phase_combo),
            ("Power Factor", self.pf_spin),
            ("Breaker (A)", self.breaker_spin),
            ("Note", self.note_edit),
        ]

        for i, (label, widget) in enumerate(widgets):
            row, col = divmod(i, 2)
            form_layout.addWidget(QLabel(label), row * 2, col)
            form_layout.addWidget(widget, row * 2 + 1, col)

        form_group.setLayout(form_layout)

        button_row = QHBoxLayout()
        self.save_btn = QPushButton("파트 저장")
        self.template_btn = QPushButton("엑셀 양식 다운로드")
        self.import_btn = QPushButton("엑셀 일괄 등록")
        self.export_btn = QPushButton("등록 파트 Raw Data 다운로드")
        self.refresh_btn = QPushButton("목록 새로고침")
        for btn in [self.save_btn, self.template_btn, self.import_btn, self.export_btn, self.refresh_btn]:
            button_row.addWidget(btn)
        button_row.addStretch()

        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels(
            [
                "Part No", "Name", "Category", "Voltage(V)", "Current(A)",
                "Power(W)", "Phase", "PF", "Breaker(A)", "Note",
            ]
        )
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)

        layout.addWidget(form_group)
        layout.addLayout(button_row)
        layout.addWidget(self.table)

        self.save_btn.clicked.connect(self.save_part)
        self.template_btn.clicked.connect(self.download_template)
        self.import_btn.clicked.connect(self.import_excel)
        self.export_btn.clicked.connect(self.export_raw_data)
        self.refresh_btn.clicked.connect(self.refresh_table)

    def save_part(self):
        try:
            part = PartRecord(
                part_no=self.part_no_edit.text().strip(),
                part_name=self.part_name_edit.text().strip(),
                category=self.category_edit.text().strip(),
                voltage_v=self.voltage_spin.value(),
                current_a=self.current_spin.value(),
                power_w=self.power_spin.value(),
                phase=self.phase_combo.currentText(),
                power_factor=self.pf_spin.value(),
                recommended_breaker_a=self.breaker_spin.value(),
                note=self.note_edit.text().strip(),
            )
            if not part.part_no or not part.part_name:
                raise ValueError("Part No와 Part Name은 필수입니다.")
            self.db.upsert_part(part)
            self.refresh_table()
            QMessageBox.information(self, "완료", "파트가 저장되었습니다.")
        except Exception as e:
            logger.exception("Failed to save part")
            QMessageBox.critical(self, "오류", str(e))

    def refresh_table(self):
        df = self.db.get_all_parts()
        self.table.setRowCount(len(df))
        columns = [
            "part_no", "part_name", "category", "voltage_v", "current_a",
            "power_w", "phase", "power_factor", "recommended_breaker_a", "note",
        ]
        for row_idx, (_, row) in enumerate(df.iterrows()):
            for col_idx, col in enumerate(columns):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(row[col])))

    def download_template(self):
        try:
            save_path, _ = QFileDialog.getSaveFileName(
                self, "엑셀 양식 저장", str(TEMPLATE_PATH), "Excel Files (*.xlsx)"
            )
            if not save_path:
                return
            ExcelManager.create_template(Path(save_path))
            QMessageBox.information(self, "완료", f"양식이 저장되었습니다.\n{save_path}")
        except Exception as e:
            logger.exception("Failed to create template")
            QMessageBox.critical(self, "오류", str(e))

    def import_excel(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "엑셀 파일 선택", str(APP_DIR), "Excel Files (*.xlsx *.xls)"
            )
            if not file_path:
                return
            df = ExcelManager.load_excel(Path(file_path))
            count = self.db.bulk_upsert_from_dataframe(df)
            self.refresh_table()
            QMessageBox.information(self, "완료", f"{count}개 파트가 등록/업데이트되었습니다.")
        except Exception as e:
            logger.exception("Failed to import excel")
            QMessageBox.critical(self, "오류", str(e))

    def export_raw_data(self):
        try:
            save_path, _ = QFileDialog.getSaveFileName(
                self, "등록 파트 Raw Data 저장", str(RAW_EXPORT_DEFAULT), "Excel Files (*.xlsx)"
            )
            if not save_path:
                return
            df = self.db.get_all_parts()
            ExcelManager.export_dataframe(df, Path(save_path))
            QMessageBox.information(self, "완료", f"Raw Data가 저장되었습니다.\n{save_path}")
        except Exception as e:
            logger.exception("Failed to export raw data")
            QMessageBox.critical(self, "오류", str(e))


class CalcTab(QWidget):
    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db = db
        self._build_ui()
        self.reload_part_numbers()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        input_group = QGroupBox("부하 계산 / 차단기 선정")
        form = QFormLayout()

        self.part_combo = QComboBox()
        self.qty_spin = QSpinBox()
        self.qty_spin.setRange(1, 100000)
        self.qty_spin.setValue(1)
        self.safety_factor_spin = QDoubleSpinBox()
        self.safety_factor_spin.setRange(1.0, 5.0)
        self.safety_factor_spin.setSingleStep(0.05)
        self.safety_factor_spin.setValue(1.25)
        self.calc_btn = QPushButton("계산 실행")
        self.reload_btn = QPushButton("파트 목록 다시 불러오기")

        form.addRow("등록 파트", self.part_combo)
        form.addRow("수량", self.qty_spin)
        form.addRow("안전율", self.safety_factor_spin)
        form.addRow(self.calc_btn, self.reload_btn)
        input_group.setLayout(form)

        result_group = QGroupBox("계산 결과")
        result_layout = QGridLayout()
        self.result_labels = {}
        result_fields = [
            "part_no", "part_name", "quantity", "voltage_v", "phase",
            "unit_current_a", "unit_power_w", "total_current_a", "total_power_w",
            "safety_factor", "safety_current_a", "suggested_breaker_a",
        ]
        for i, key in enumerate(result_fields):
            title = QLabel(key)
            value = QLabel("-")
            self.result_labels[key] = value
            r, c = divmod(i, 2)
            result_layout.addWidget(title, r, c * 2)
            result_layout.addWidget(value, r, c * 2 + 1)
        result_group.setLayout(result_layout)

        layout.addWidget(input_group)
        layout.addWidget(result_group)
        layout.addStretch()

        self.calc_btn.clicked.connect(self.run_calculation)
        self.reload_btn.clicked.connect(self.reload_part_numbers)

    def reload_part_numbers(self):
        self.part_combo.clear()
        self.part_combo.addItems(self.db.get_part_nos())

    def run_calculation(self):
        try:
            part_no = self.part_combo.currentText().strip()
            if not part_no:
                raise ValueError("먼저 등록된 파트를 선택해 주세요.")
            part_row = self.db.get_part_by_no(part_no)
            if part_row is None:
                raise ValueError("선택한 파트를 찾을 수 없습니다.")
            result = LoadCalculator.calculate_total(
                part_row, self.qty_spin.value(), self.safety_factor_spin.value()
            )
            for key, value in result.items():
                self.result_labels[key].setText(str(value))
        except Exception as e:
            logger.exception("Failed to calculate")
            QMessageBox.critical(self, "오류", str(e))


# ------------------------------------------------------------
# Canvas drag sources
# ------------------------------------------------------------
class DraggablePartListWidget(QListWidget):
    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db = db
        self.setDragEnabled(True)
        self.setAlternatingRowColors(True)
        self.refresh_parts()

    def refresh_parts(self):
        self.clear()
        df = self.db.get_all_parts()
        for _, row in df.iterrows():
            text = f"{row['part_no']} | {row['part_name']} | {row['current_a']}A | {row['power_w']}W"
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, row['part_no'])
            self.addItem(item)

    def startDrag(self, supportedActions):
        item = self.currentItem()
        if item is None:
            return
        part_no = item.data(Qt.UserRole)
        if not part_no:
            return
        drag = QDrag(self)
        mime = QMimeData()
        mime.setText(str(part_no))
        mime.setData("application/x-item-type", b"part")
        drag.setMimeData(mime)
        drag.exec(Qt.CopyAction)


class BreakerTemplateListWidget(QListWidget):
    def __init__(self):
        super().__init__()
        self.setDragEnabled(True)
        self.setAlternatingRowColors(True)
        item = QListWidgetItem("하위 차단기 템플릿")
        item.setData(Qt.UserRole, "breaker_template")
        self.addItem(item)

    def startDrag(self, supportedActions):
        item = self.currentItem()
        if item is None:
            return
        drag = QDrag(self)
        mime = QMimeData()
        mime.setText("breaker_template")
        mime.setData("application/x-item-type", b"breaker")
        drag.setMimeData(mime)
        drag.exec(Qt.CopyAction)


# ------------------------------------------------------------
# Canvas items
# ------------------------------------------------------------
class BreakerSettingsDialog(QDialog):
    def __init__(self, breaker_name: str, safety_factor: float, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"차단기 설정 - {breaker_name}")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.safety_spin = QDoubleSpinBox()
        self.safety_spin.setRange(1.0, 5.0)
        self.safety_spin.setSingleStep(0.05)
        self.safety_spin.setValue(safety_factor)
        form.addRow("안전율", self.safety_spin)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class LoadPartItem(QGraphicsRectItem):
    WIDTH = 200
    HEIGHT = 68

    def __init__(self, scene_ref: 'BreakerCanvasScene', parent_breaker: 'BreakerItem', part_row: pd.Series, quantity: int = 1):
        super().__init__(0, 0, self.WIDTH, self.HEIGHT)
        self.scene_ref = scene_ref
        self.parent_breaker = parent_breaker
        self.part_no = str(part_row['part_no'])
        self.part_name = str(part_row['part_name'])
        self.unit_current = float(part_row['current_a'])
        self.unit_power = float(part_row['power_w'])
        self.quantity = quantity

        self.setBrush(QBrush(QColor("#fff8ef")))
        self.setPen(QPen(QColor("#d9c2a6"), 1.2))
        self.setFlag(QGraphicsItem.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)
        self.setAcceptHoverEvents(True)
        self.setZValue(3)

        self.label = QGraphicsSimpleTextItem(self)
        self.label.setBrush(QBrush(QColor("#4a3f35")))
        self.update_display()

    def update_display(self):
        total_current = self.unit_current * self.quantity
        total_power = self.unit_power * self.quantity
        self.label.setText(
            f"{self.part_no}\n{self.part_name}\nQty {self.quantity} | {total_current:.2f}A | {total_power:.0f}W"
        )
        self.label.setPos(10, 8)

    def itemChange(self, change, value):
        if change == QGraphicsItem.ItemPositionHasChanged:
            self.scene_ref.notify_layout_changed()
            if self.parent_breaker:
                self.parent_breaker.refresh_recursive()
        return super().itemChange(change, value)

    def mouseDoubleClickEvent(self, event):
        menu = QMessageBox()
        menu.setWindowTitle("부하 파트 편집")
        menu.setText(f"{self.part_no} / {self.part_name}")
        qty_btn = menu.addButton("수량 변경", QMessageBox.AcceptRole)
        del_btn = menu.addButton("삭제", QMessageBox.DestructiveRole)
        menu.addButton("취소", QMessageBox.RejectRole)
        menu.exec()

        clicked = menu.clickedButton()
        if clicked == qty_btn:
            qty, ok = QInputDialog.getInt(
                None,
                "수량 변경",
                f"{self.part_no} 수량",
                self.quantity,
                1,
                100000,
                1,
            )
            if ok:
                self.quantity = qty
                self.update_display()
                if self.parent_breaker:
                    self.parent_breaker.refresh_recursive()
                self.scene_ref.notify_layout_changed()
        elif clicked == del_btn:
            self.scene_ref.delete_item(self)
        super().mouseDoubleClickEvent(event)

    def to_dict(self) -> dict:
        pos = self.scenePos()
        return {
            "part_no": self.part_no,
            "quantity": self.quantity,
            "x": round(pos.x(), 2),
            "y": round(pos.y(), 2),
        }


class BreakerItem(QGraphicsRectItem):
    WIDTH = 280
    HEIGHT = 120

    def __init__(self, scene_ref: 'BreakerCanvasScene', name: str, pos_x: float, pos_y: float,
                 safety_factor: float = 1.25, parent_breaker: Optional['BreakerItem'] = None,
                 is_top_level: bool = False):
        super().__init__(0, 0, self.WIDTH, self.HEIGHT)
        self.scene_ref = scene_ref
        self.name = name
        self.safety_factor = safety_factor
        self.parent_breaker = parent_breaker
        self.child_breakers: List['BreakerItem'] = []
        self.load_items: List[LoadPartItem] = []
        self.is_top_level = is_top_level

        self.setPos(pos_x, pos_y)
        self.setBrush(QBrush(QColor("#eef4ff")))
        self.setPen(QPen(QColor("#9cb4de"), 1.6))
        self.setFlag(QGraphicsItem.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)
        self.setAcceptDrops(False)
        self.setZValue(2)

        self.title_text = QGraphicsSimpleTextItem(self)
        self.title_text.setBrush(QBrush(QColor("#334155")))
        self.summary_text = QGraphicsSimpleTextItem(self)
        self.summary_text.setBrush(QBrush(QColor("#475569")))
        self.hint_text = QGraphicsSimpleTextItem("더블클릭: 안전율 설정", self)
        self.hint_text.setBrush(QBrush(QColor("#64748b")))
        self.title_text.setPos(12, 10)
        self.summary_text.setPos(12, 38)
        self.hint_text.setPos(12, 90)
        self.update_summary()

    def itemChange(self, change, value):
        if change == QGraphicsItem.ItemPositionHasChanged:
            self.scene_ref.notify_layout_changed()
        return super().itemChange(change, value)

    def mouseDoubleClickEvent(self, event):
        dialog = BreakerSettingsDialog(self.name, self.safety_factor)
        if dialog.exec() == QDialog.Accepted:
            self.safety_factor = dialog.safety_spin.value()
            self.refresh_recursive()
            self.scene_ref.notify_layout_changed()
        super().mouseDoubleClickEvent(event)

    def is_ancestor_of(self, other: 'BreakerItem') -> bool:
        cur = other.parent_breaker
        while cur is not None:
            if cur is self:
                return True
            cur = cur.parent_breaker
        return False

    def can_accept_breaker(self, child: 'BreakerItem') -> bool:
        if child is self:
            return False
        if child.is_ancestor_of(self):
            return False
        return True

    def set_parent_breaker(self, new_parent: Optional['BreakerItem']):
        old_parent = self.parent_breaker
        if old_parent is new_parent:
            return
        if old_parent is not None and self in old_parent.child_breakers:
            old_parent.child_breakers.remove(self)
            old_parent.refresh_recursive()
        self.parent_breaker = new_parent
        if new_parent is not None and self not in new_parent.child_breakers:
            new_parent.child_breakers.append(self)
        self.refresh_recursive()

    def add_part(self, part_row: pd.Series, quantity: int = 1, drop_pos: Optional[QPointF] = None):
        load_item = LoadPartItem(self.scene_ref, self, part_row, quantity)
        self.scene_ref.addItem(load_item)
        self.load_items.append(load_item)
        if drop_pos is None:
            drop_pos = QPointF(self.scenePos().x() + 40, self.scenePos().y() + self.HEIGHT + 60)
        load_item.setPos(drop_pos)
        self.refresh_recursive()
        self.scene_ref.notify_layout_changed()
        return load_item

    def add_child_breaker(self, child_breaker: 'BreakerItem') -> bool:
        if not self.can_accept_breaker(child_breaker):
            QMessageBox.warning(None, "연결 불가", "자기 자신 또는 상위 차단기를 하위에 연결할 수 없습니다.")
            return False
        child_breaker.set_parent_breaker(self)
        self.refresh_recursive()
        self.scene_ref.notify_layout_changed()
        return True

    def remove_load_item(self, load_item: LoadPartItem):
        if load_item in self.load_items:
            self.load_items.remove(load_item)
            self.refresh_recursive()

    def get_total_current(self, visited: Optional[Set[int]] = None) -> float:
        if visited is None:
            visited = set()
        obj_id = id(self)
        if obj_id in visited:
            logger.warning("Cycle detected in get_total_current: %s", self.name)
            return 0.0
        visited.add(obj_id)
        total = sum(item.unit_current * item.quantity for item in self.load_items)
        for child in self.child_breakers:
            total += child.get_total_current(visited)
        return total

    def get_total_power(self, visited: Optional[Set[int]] = None) -> float:
        if visited is None:
            visited = set()
        obj_id = id(self)
        if obj_id in visited:
            logger.warning("Cycle detected in get_total_power: %s", self.name)
            return 0.0
        visited.add(obj_id)
        total = sum(item.unit_power * item.quantity for item in self.load_items)
        for child in self.child_breakers:
            total += child.get_total_power(visited)
        return total

    def get_total_load_count(self, visited: Optional[Set[int]] = None) -> int:
        if visited is None:
            visited = set()
        obj_id = id(self)
        if obj_id in visited:
            logger.warning("Cycle detected in get_total_load_count: %s", self.name)
            return 0
        visited.add(obj_id)
        total = len(self.load_items)
        for child in self.child_breakers:
            total += child.get_total_load_count(visited)
        return total

    def suggested_breaker(self) -> int:
        return LoadCalculator.select_breaker(self.get_total_current() * self.safety_factor)

    def update_summary(self):
        self.title_text.setText(self.name)
        self.summary_text.setText(
            f"안전율 {self.safety_factor:.2f} | 부하 {self.get_total_load_count()}개\n"
            f"합계 {self.get_total_current():.2f}A / {self.get_total_power():.0f}W | 추천 {self.suggested_breaker()}A"
        )

    def refresh_recursive(self):
        self.update_summary()
        if self.parent_breaker is not None:
            self.parent_breaker.refresh_recursive()

    def contains_scene_pos(self, scene_pos: QPointF) -> bool:
        return self.sceneBoundingRect().contains(scene_pos)

    def to_dict(self) -> dict:
        pos = self.scenePos()
        return {
            "name": self.name,
            "x": round(pos.x(), 2),
            "y": round(pos.y(), 2),
            "safety_factor": self.safety_factor,
            "is_top_level": self.is_top_level,
            "loads": [item.to_dict() for item in self.load_items],
            "children": [child.to_dict() for child in self.child_breakers],
        }


# ------------------------------------------------------------
# Canvas scene / view
# ------------------------------------------------------------
class BreakerCanvasScene(QGraphicsScene):
    layoutChanged = Signal()

    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db = db
        self.top_breaker: Optional[BreakerItem] = None
        self.all_breakers: List[BreakerItem] = []
        self.connection_lines = []
        self.breaker_seq = 1
        self.setSceneRect(0, 0, 2600, 1800)
        self.setBackgroundBrush(QBrush(QColor("#f8fafc")))

    def new_breaker_name(self) -> str:
        while True:
            name = f"MCCB-{self.breaker_seq:02d}"
            self.breaker_seq += 1
            if not any(b.name == name for b in self.all_breakers):
                return name

    def get_top_breakers(self) -> List[BreakerItem]:
        return [b for b in self.all_breakers if b.parent_breaker is None]

    def add_top_breaker(self, name: Optional[str] = None, x: float = 120, y: float = 80) -> Optional[BreakerItem]:
        if self.top_breaker is not None:
            QMessageBox.information(None, "안내", "최상위 차단기는 1개만 생성할 수 있습니다.")
            return None
        breaker = BreakerItem(self, name or self.new_breaker_name(), x, y, is_top_level=True)
        self.addItem(breaker)
        self.all_breakers.append(breaker)
        self.top_breaker = breaker
        self.notify_layout_changed()
        return breaker

    def create_child_breaker(self, parent_breaker: BreakerItem, pos: Optional[QPointF] = None, name: Optional[str] = None) -> BreakerItem:
        if pos is None:
            pos = QPointF(parent_breaker.scenePos().x() + 380, parent_breaker.scenePos().y() + 220)
        breaker = BreakerItem(self, name or self.new_breaker_name(), pos.x(), pos.y(), parent_breaker=parent_breaker)
        self.addItem(breaker)
        self.all_breakers.append(breaker)
        parent_breaker.child_breakers.append(breaker)
        parent_breaker.refresh_recursive()
        self.notify_layout_changed()
        return breaker

    def find_breaker_at(self, scene_pos: QPointF) -> Optional[BreakerItem]:
        items = self.items(scene_pos)
        for item in items:
            if isinstance(item, BreakerItem):
                return item
            if isinstance(item, QGraphicsSimpleTextItem) and isinstance(item.parentItem(), BreakerItem):
                return item.parentItem()
        return None

    def update_connection_lines(self):
        for line in self.connection_lines:
            self.removeItem(line)
        self.connection_lines.clear()

        pen = QPen(QColor("#b8c4d6"), 1.2)
        for breaker in self.all_breakers:
            start_x = breaker.scenePos().x() + breaker.WIDTH / 2
            start_y = breaker.scenePos().y() + breaker.HEIGHT

            for load_item in breaker.load_items:
                end_x = load_item.scenePos().x() + load_item.WIDTH / 2
                end_y = load_item.scenePos().y()
                line = self.addLine(start_x, start_y, end_x, end_y, pen)
                line.setZValue(-10)
                self.connection_lines.append(line)

            for child in breaker.child_breakers:
                end_x = child.scenePos().x() + child.WIDTH / 2
                end_y = child.scenePos().y()
                line = self.addLine(start_x, start_y, end_x, end_y, pen)
                line.setZValue(-10)
                self.connection_lines.append(line)

    def notify_layout_changed(self):
        for breaker in self.all_breakers:
            breaker.update_summary()
        self.update_connection_lines()
        self.layoutChanged.emit()

    def delete_item(self, item):
        if isinstance(item, LoadPartItem):
            if item.parent_breaker:
                item.parent_breaker.remove_load_item(item)
            self.removeItem(item)
            self.notify_layout_changed()
            return

        if isinstance(item, BreakerItem):
            if item.is_top_level:
                QMessageBox.information(None, "안내", "최상위 차단기는 삭제할 수 없습니다.")
                return
            for child in list(item.child_breakers):
                child.set_parent_breaker(item.parent_breaker)
            for load in list(item.load_items):
                item.remove_load_item(load)
                self.removeItem(load)
            if item.parent_breaker and item in item.parent_breaker.child_breakers:
                item.parent_breaker.child_breakers.remove(item)
                item.parent_breaker.refresh_recursive()
            if item in self.all_breakers:
                self.all_breakers.remove(item)
            self.removeItem(item)
            self.notify_layout_changed()

    def delete_selected_items(self):
        selected = list(self.selectedItems())
        if not selected:
            QMessageBox.information(None, "안내", "삭제할 항목을 먼저 선택하세요.")
            return
        for item in selected:
            if isinstance(item, (LoadPartItem, BreakerItem)):
                self.delete_item(item)

    def handle_drop(self, mime, scene_pos: QPointF):
        item_type = bytes(mime.data("application/x-item-type")).decode(errors="ignore")
        target_breaker = self.find_breaker_at(scene_pos)

        if item_type == "part":
            part_no = mime.text()
            if not target_breaker:
                QMessageBox.warning(None, "배치 실패", "차단기 카드 위에 드롭해야 합니다.")
                return
            part_row = self.db.get_part_by_no(part_no)
            if part_row is None:
                QMessageBox.warning(None, "배치 실패", "해당 파트를 찾을 수 없습니다.")
                return
            target_breaker.add_part(part_row, 1, scene_pos)
            return

        if item_type == "breaker":
            if not target_breaker:
                QMessageBox.warning(None, "배치 실패", "하위 차단기는 기존 차단기 위에 드롭해야 합니다.")
                return
            self.create_child_breaker(target_breaker, QPointF(scene_pos.x(), scene_pos.y() + 160))

    def to_dict(self) -> dict:
        if self.top_breaker is None:
            return {"top_breaker": None}
        return {"top_breaker": self.top_breaker.to_dict()}

    def clear_canvas(self):
        self.clear()
        self.connection_lines.clear()
        self.all_breakers.clear()
        self.top_breaker = None

    def save_to_file(self, path: Path):
        path.write_text(json.dumps(self.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")

    def _restore_breaker_tree(self, data: dict, parent: Optional[BreakerItem] = None):
        breaker = BreakerItem(
            self,
            data.get("name", self.new_breaker_name()),
            float(data.get("x", 120)),
            float(data.get("y", 80)),
            float(data.get("safety_factor", 1.25)),
            parent_breaker=parent,
            is_top_level=bool(data.get("is_top_level", parent is None)),
        )
        self.addItem(breaker)
        self.all_breakers.append(breaker)
        if parent is None:
            self.top_breaker = breaker
        else:
            parent.child_breakers.append(breaker)

        for load in data.get("loads", []):
            part_row = self.db.get_part_by_no(load.get("part_no", ""))
            if part_row is not None:
                breaker.add_part(
                    part_row,
                    int(load.get("quantity", 1)),
                    QPointF(float(load.get("x", breaker.scenePos().x() + 40)), float(load.get("y", breaker.scenePos().y() + 180))),
                )

        for child_data in data.get("children", []):
            self._restore_breaker_tree(child_data, breaker)

        breaker.refresh_recursive()
        return breaker

    def load_from_file(self, path: Path):
        self.clear_canvas()
        if not path.exists():
            self.add_top_breaker("MAIN-MCCB", 120, 80)
            return
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            top_data = data.get("top_breaker")
            if not top_data:
                self.add_top_breaker("MAIN-MCCB", 120, 80)
            else:
                self._restore_breaker_tree(top_data, None)
            self.notify_layout_changed()
        except Exception:
            logger.exception("Failed to load canvas layout")
            self.clear_canvas()
            self.add_top_breaker("MAIN-MCCB", 120, 80)


class BreakerCanvasView(QGraphicsView):
    def __init__(self, scene: BreakerCanvasScene):
        super().__init__(scene)
        self.setRenderHint(QPainter.Antialiasing)
        self.setAcceptDrops(True)
        self.setDragMode(QGraphicsView.RubberBandDrag)
        self.setViewportUpdateMode(QGraphicsView.BoundingRectViewportUpdate)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        scene_pos = self.mapToScene(event.position().toPoint())
        self.scene().handle_drop(event.mimeData(), scene_pos)
        event.acceptProposedAction()


# ------------------------------------------------------------
# Canvas tab
# ------------------------------------------------------------
class CanvasTab(QWidget):
    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db = db
        self.scene = BreakerCanvasScene(self.db)
        self._build_ui()
        self.scene.load_from_file(CANVAS_SAVE_PATH)
        self.refresh_summary_table()

    def _apply_widget_shadow(self, widget: QWidget):
        effect = QGraphicsDropShadowEffect(widget)
        effect.setBlurRadius(20)
        effect.setOffset(0, 4)
        effect.setColor(QColor(0, 0, 0, 40))
        widget.setGraphicsEffect(effect)

    def _build_ui(self):
        root = QHBoxLayout(self)

        left_panel = QVBoxLayout()
        lib_group = QGroupBox("파트 라이브러리")
        lib_layout = QVBoxLayout()
        self.part_list = DraggablePartListWidget(self.db)
        self.part_list.setMinimumWidth(330)
        lib_layout.addWidget(QLabel("파트를 드래그해서 차단기 위에 놓으면 하위 부하로 추가됩니다."))
        lib_layout.addWidget(self.part_list)
        lib_group.setLayout(lib_layout)

        breaker_group = QGroupBox("차단기 템플릿")
        breaker_layout = QVBoxLayout()
        self.breaker_template_list = BreakerTemplateListWidget()
        breaker_layout.addWidget(QLabel("하위 차단기 템플릿을 드래그해서 기존 차단기 위에 놓으세요."))
        breaker_layout.addWidget(self.breaker_template_list)
        breaker_group.setLayout(breaker_layout)

        btn_row = QHBoxLayout()
        self.refresh_library_btn = QPushButton("라이브러리 새로고침")
        self.add_top_breaker_btn = QPushButton("최상위 차단기 생성")
        self.delete_selected_btn = QPushButton("선택 항목 삭제")
        self.save_canvas_btn = QPushButton("캔버스 저장")
        self.reset_canvas_btn = QPushButton("기본 배치 복원")
        for btn in [
            self.refresh_library_btn,
            self.add_top_breaker_btn,
            self.delete_selected_btn,
            self.save_canvas_btn,
            self.reset_canvas_btn,
        ]:
            btn_row.addWidget(btn)

        summary_group = QGroupBox("차단기별 계산 요약")
        summary_layout = QVBoxLayout()
        self.summary_table = QTableWidget()
        self.summary_table.setColumnCount(6)
        self.summary_table.setHorizontalHeaderLabels([
            "차단기", "안전율", "전체 부하 수", "합계 전류(A)", "합계 전력(W)", "추천 차단기(A)"
        ])
        self.summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        summary_layout.addWidget(self.summary_table)
        summary_group.setLayout(summary_layout)

        canvas_group = QGroupBox("단선 구성 캔버스")
        canvas_layout = QVBoxLayout()
        self.canvas_view = BreakerCanvasView(self.scene)
        canvas_layout.addWidget(self.canvas_view)
        canvas_group.setLayout(canvas_layout)

        for w in [lib_group, breaker_group, summary_group, canvas_group]:
            self._apply_widget_shadow(w)

        left_panel.addWidget(lib_group)
        left_panel.addWidget(breaker_group)
        left_panel.addLayout(btn_row)
        left_panel.addWidget(summary_group)

        root.addLayout(left_panel, 1)
        root.addWidget(canvas_group, 2)

        self.refresh_library_btn.clicked.connect(self.reload_library)
        self.add_top_breaker_btn.clicked.connect(self.add_top_breaker)
        self.delete_selected_btn.clicked.connect(self.scene.delete_selected_items)
        self.save_canvas_btn.clicked.connect(self.save_canvas)
        self.reset_canvas_btn.clicked.connect(self.reset_canvas)
        self.scene.layoutChanged.connect(self.refresh_summary_table)

    def reload_library(self):
        self.part_list.refresh_parts()

    def add_top_breaker(self):
        self.scene.add_top_breaker("MAIN-MCCB", 120, 80)

    def save_canvas(self):
        try:
            self.scene.save_to_file(CANVAS_SAVE_PATH)
            QMessageBox.information(self, "완료", f"캔버스 구성이 저장되었습니다.\n{CANVAS_SAVE_PATH}")
        except Exception as e:
            logger.exception("Failed to save canvas")
            QMessageBox.critical(self, "오류", str(e))

    def reset_canvas(self):
        self.scene.clear_canvas()
        self.scene.add_top_breaker("MAIN-MCCB", 120, 80)
        self.scene.notify_layout_changed()
        self.save_canvas()

    def _collect_breakers(self, root_breaker: Optional[BreakerItem]) -> List[BreakerItem]:
        if root_breaker is None:
            return []
        result = []

        def walk(b: BreakerItem):
            result.append(b)
            for child in b.child_breakers:
                walk(child)

        walk(root_breaker)
        return result

    def refresh_summary_table(self):
        breakers = self._collect_breakers(self.scene.top_breaker)
        self.summary_table.setRowCount(len(breakers))
        for row, breaker in enumerate(breakers):
            values = [
                breaker.name,
                f"{breaker.safety_factor:.2f}",
                str(breaker.get_total_load_count()),
                f"{breaker.get_total_current():.2f}",
                f"{breaker.get_total_power():.0f}",
                str(breaker.suggested_breaker()),
            ]
            for col, value in enumerate(values):
                self.summary_table.setItem(row, col, QTableWidgetItem(value))


# ------------------------------------------------------------
# Main window
# ------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Electrical Capacity & Breaker Helper")
        self.resize(1680, 980)
        self.db = DatabaseManager(DB_PATH)
        self._build_ui()

    def _build_ui(self):
        tabs = QTabWidget()
        self.parts_tab = PartsTab(self.db)
        self.calc_tab = CalcTab(self.db)
        self.canvas_tab = CanvasTab(self.db)
        tabs.addTab(self.parts_tab, "파트 관리")
        tabs.addTab(self.calc_tab, "부하 계산")
        tabs.addTab(self.canvas_tab, "차단기 캔버스")
        self.setCentralWidget(tabs)

        refresh_action = QAction("새로고침", self)
        refresh_action.triggered.connect(self.refresh_all)
        self.menuBar().addAction(refresh_action)

    def refresh_all(self):
        self.parts_tab.refresh_table()
        self.calc_tab.reload_part_numbers()
        self.canvas_tab.reload_library()
        self.canvas_tab.refresh_summary_table()

    def closeEvent(self, event):
        try:
            self.canvas_tab.scene.save_to_file(CANVAS_SAVE_PATH)
        except Exception:
            logger.exception("Failed to save canvas on close")
        super().closeEvent(event)


def load_stylesheet(app: QApplication):
    if STYLE_PATH.exists():
        app.setStyleSheet(STYLE_PATH.read_text(encoding="utf-8"))
    else:
        logger.warning("Stylesheet file not found: %s", STYLE_PATH)


def main():
    app = QApplication(sys.argv)
    load_stylesheet(app)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
