"""
미트웰 발주서 자동화
스마트스토어 + 카페24 → 미트웰 발주 엑셀 생성 + 카카오톡 전송
"""
import sys
import os
import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QProgressBar, QStatusBar, QGroupBox, QSplitter, QHeaderView,
    QMessageBox, QLineEdit, QFrame,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QColor, QPalette, QIcon

from processor import (
    process_smartstore, process_cafe24, build_output, save_output_excel
)


# ──────────────────────────────────────────
# 스타일 상수
# ──────────────────────────────────────────
STYLE = """
QMainWindow, QWidget {
    background-color: #1E1E2E;
    color: #CDD6F4;
    font-family: 'Malgun Gothic', sans-serif;
    font-size: 13px;
}
QGroupBox {
    border: 1px solid #45475A;
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 8px;
    font-weight: bold;
    color: #89B4FA;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 6px;
}
QPushButton {
    background-color: #89B4FA;
    color: #1E1E2E;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    font-weight: bold;
    font-size: 13px;
}
QPushButton:hover {
    background-color: #B4BEFE;
}
QPushButton:pressed {
    background-color: #7287FD;
}
QPushButton:disabled {
    background-color: #45475A;
    color: #6C7086;
}
QPushButton#btn_process {
    background-color: #A6E3A1;
    font-size: 15px;
    padding: 10px 24px;
}
QPushButton#btn_process:hover {
    background-color: #94E2D5;
}
QPushButton#btn_kakao {
    background-color: #FAE3B0;
    color: #1E1E2E;
    font-size: 14px;
}
QPushButton#btn_kakao:hover {
    background-color: #F9E2AF;
}
QPushButton#btn_save {
    background-color: #CBA6F7;
}
QPushButton#btn_save:hover {
    background-color: #F5C2E7;
}
QLineEdit {
    background-color: #313244;
    border: 1px solid #45475A;
    border-radius: 6px;
    padding: 6px 10px;
    color: #CDD6F4;
}
QLineEdit:focus {
    border-color: #89B4FA;
}
QTableWidget {
    background-color: #181825;
    border: 1px solid #45475A;
    border-radius: 6px;
    gridline-color: #313244;
    color: #CDD6F4;
}
QTableWidget::item:selected {
    background-color: #89B4FA;
    color: #1E1E2E;
}
QHeaderView::section {
    background-color: #313244;
    color: #89B4FA;
    border: none;
    border-bottom: 1px solid #45475A;
    padding: 6px;
    font-weight: bold;
}
QProgressBar {
    background-color: #313244;
    border: none;
    border-radius: 4px;
    height: 8px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #89B4FA;
    border-radius: 4px;
}
QStatusBar {
    background-color: #11111B;
    color: #6C7086;
    border-top: 1px solid #313244;
}
QLabel#lbl_file {
    background-color: #313244;
    border: 1px dashed #45475A;
    border-radius: 6px;
    padding: 8px;
    color: #6C7086;
}
QLabel#lbl_file_set {
    background-color: #1E3A5F;
    border: 1px solid #89B4FA;
    border-radius: 6px;
    padding: 8px;
    color: #89B4FA;
}
"""


# ──────────────────────────────────────────
# 처리 워커 스레드
# ──────────────────────────────────────────
class ProcessWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(object)  # DataFrame or None
    error    = pyqtSignal(str)

    def __init__(self, ss_paths, c24_path, ss_password="1111"):
        super().__init__()
        self.ss_paths    = ss_paths
        self.c24_path    = c24_path
        self.ss_password = ss_password

    def run(self):
        try:
            import pandas as pd

            all_ss_rows   = []
            all_ss_saeum  = {}

            total = len(self.ss_paths) + (1 if self.c24_path else 0)
            step  = 0

            for path in self.ss_paths:
                self.progress.emit(
                    int(step / total * 80),
                    f"스마트스토어 처리 중: {os.path.basename(path)}"
                )
                rows, saeum = process_smartstore(path, self.ss_password)
                all_ss_rows.append(rows)
                for k, v in saeum.items():
                    all_ss_saeum[k] = all_ss_saeum.get(k, 0) + v
                step += 1

            ss_combined = pd.concat(all_ss_rows, ignore_index=True) if all_ss_rows else pd.DataFrame()

            c24_rows  = pd.DataFrame()
            c24_saeum = {}
            if self.c24_path:
                self.progress.emit(int(step / total * 80), "카페24 처리 중...")
                c24_rows, c24_saeum = process_cafe24(self.c24_path)
                step += 1

            self.progress.emit(85, "발주서 생성 중...")
            result = build_output(ss_combined, c24_rows, all_ss_saeum, c24_saeum)

            self.progress.emit(100, "완료!")
            self.finished.emit(result)

        except Exception as e:
            import traceback
            self.error.emit(f"{e}\n\n{traceback.format_exc()}")


# ──────────────────────────────────────────
# 메인 윈도우
# ──────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ss_paths   = []
        self.c24_path   = ""
        self.result_df  = None
        self.out_path   = ""
        self.worker     = None

        self.setWindowTitle("미트웰 발주서 자동화")
        self.setMinimumSize(1100, 750)
        self.setStyleSheet(STYLE)

        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(16, 16, 16, 8)
        root.setSpacing(12)

        # ── 제목
        title = QLabel("🥩 미트웰 발주서 자동화")
        title.setFont(QFont("Malgun Gothic", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #CBA6F7; margin-bottom: 4px;")
        root.addWidget(title)

        # ── 파일 선택 그룹
        file_group = QGroupBox("파일 선택")
        fg_layout  = QVBoxLayout(file_group)
        fg_layout.setSpacing(8)

        # 스마트스토어
        ss_row = QHBoxLayout()
        ss_row.addWidget(QLabel("스마트스토어 XLSX:"))
        self.lbl_ss = QLabel("파일을 선택하세요 (여러 파일 가능)")
        self.lbl_ss.setObjectName("lbl_file")
        self.lbl_ss.setWordWrap(True)
        ss_row.addWidget(self.lbl_ss, 1)
        btn_ss = QPushButton("파일 선택")
        btn_ss.setFixedWidth(100)
        btn_ss.clicked.connect(self._pick_ss)
        ss_row.addWidget(btn_ss)
        fg_layout.addLayout(ss_row)

        # 스마트스토어 비밀번호
        pwd_row = QHBoxLayout()
        pwd_row.addWidget(QLabel("스마트스토어 암호:"))
        self.txt_pwd = QLineEdit("1111")
        self.txt_pwd.setEchoMode(QLineEdit.EchoMode.Password)
        self.txt_pwd.setFixedWidth(120)
        pwd_row.addWidget(self.txt_pwd)
        pwd_row.addStretch()
        fg_layout.addLayout(pwd_row)

        # 카페24
        c24_row = QHBoxLayout()
        c24_row.addWidget(QLabel("카페24 CSV:          "))
        self.lbl_c24 = QLabel("파일을 선택하세요")
        self.lbl_c24.setObjectName("lbl_file")
        c24_row.addWidget(self.lbl_c24, 1)
        btn_c24 = QPushButton("파일 선택")
        btn_c24.setFixedWidth(100)
        btn_c24.clicked.connect(self._pick_c24)
        c24_row.addWidget(btn_c24)
        fg_layout.addLayout(c24_row)

        root.addWidget(file_group)

        # ── 실행 버튼 + 진행바
        action_row = QHBoxLayout()
        self.btn_process = QPushButton("▶  발주서 변환")
        self.btn_process.setObjectName("btn_process")
        self.btn_process.setFixedHeight(44)
        self.btn_process.clicked.connect(self._run_process)
        action_row.addWidget(self.btn_process)
        root.addLayout(action_row)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        root.addWidget(self.progress)

        # ── 미리보기 테이블
        preview_group = QGroupBox("미리보기")
        pg_layout     = QVBoxLayout(preview_group)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        pg_layout.addWidget(self.table)

        self.lbl_count = QLabel("0행")
        self.lbl_count.setStyleSheet("color: #6C7086; font-size: 12px;")
        pg_layout.addWidget(self.lbl_count)

        root.addWidget(preview_group, 1)

        # ── 하단 버튼
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        self.btn_save = QPushButton("💾  엑셀 저장")
        self.btn_save.setObjectName("btn_save")
        self.btn_save.setEnabled(False)
        self.btn_save.setFixedHeight(40)
        self.btn_save.clicked.connect(self._save_excel)
        btn_row.addWidget(self.btn_save)

        self.btn_kakao = QPushButton("💬  카카오톡 전송")
        self.btn_kakao.setObjectName("btn_kakao")
        self.btn_kakao.setEnabled(False)
        self.btn_kakao.setFixedHeight(40)
        self.btn_kakao.clicked.connect(self._send_kakao)
        btn_row.addWidget(self.btn_kakao)

        root.addLayout(btn_row)

        # ── 상태바
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("준비됨")

    # ── 파일 선택
    def _pick_ss(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "스마트스토어 파일 선택", "",
            "Excel 파일 (*.xlsx *.xls)"
        )
        if paths:
            self.ss_paths = paths
            names = "\n".join(os.path.basename(p) for p in paths)
            self.lbl_ss.setText(names)
            self.lbl_ss.setObjectName("lbl_file_set")
            self.lbl_ss.setStyleSheet(
                "background:#1E3A5F; border:1px solid #89B4FA; "
                "border-radius:6px; padding:8px; color:#89B4FA;"
            )
            self.status.showMessage(f"스마트스토어: {len(paths)}개 파일 선택됨")

    def _pick_c24(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "카페24 CSV 선택", "", "CSV 파일 (*.csv)"
        )
        if path:
            self.c24_path = path
            self.lbl_c24.setText(os.path.basename(path))
            self.lbl_c24.setObjectName("lbl_file_set")
            self.lbl_c24.setStyleSheet(
                "background:#1E3A5F; border:1px solid #89B4FA; "
                "border-radius:6px; padding:8px; color:#89B4FA;"
            )
            self.status.showMessage(f"카페24: {os.path.basename(path)} 선택됨")

    # ── 변환 실행
    def _run_process(self):
        if not self.ss_paths and not self.c24_path:
            QMessageBox.warning(self, "파일 없음", "파일을 먼저 선택하세요.")
            return

        self.btn_process.setEnabled(False)
        self.btn_save.setEnabled(False)
        self.btn_kakao.setEnabled(False)
        self.progress.setValue(0)
        self.table.setRowCount(0)
        self.status.showMessage("처리 중...")

        self.worker = ProcessWorker(
            self.ss_paths, self.c24_path,
            ss_password=self.txt_pwd.text()
        )
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_progress(self, val, msg):
        self.progress.setValue(val)
        self.status.showMessage(msg)

    def _on_finished(self, df):
        self.result_df = df
        self._fill_table(df)
        self.btn_process.setEnabled(True)
        self.btn_save.setEnabled(True)
        self.btn_kakao.setEnabled(True)
        self.progress.setValue(100)
        self.status.showMessage(f"완료! 총 {len(df)}행 생성됨")

    def _on_error(self, msg):
        self.btn_process.setEnabled(True)
        self.progress.setValue(0)
        self.status.showMessage("오류 발생!")
        QMessageBox.critical(self, "처리 오류", msg)

    # ── 테이블 채우기
    def _fill_table(self, df):
        cols = list(df.columns)
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        self.table.setRowCount(len(df))

        for ri, row in enumerate(df.itertuples(index=False)):
            for ci, val in enumerate(row):
                if val is None or (hasattr(val, '__class__') and val.__class__.__name__ == 'float'):
                    import math
                    try:
                        if math.isnan(float(val)):
                            val = ""
                    except (TypeError, ValueError):
                        pass
                item = QTableWidgetItem(str(val) if val is not None else "")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(ri, ci, item)

        self.lbl_count.setText(f"총 {len(df)}행 / {df['수하인명'].nunique()}명")
        self.table.resizeColumnsToContents()

    # ── 엑셀 저장
    def _save_excel(self):
        if self.result_df is None:
            return

        today = datetime.datetime.now().strftime("%m%d")
        default_name = f"미트웰발주_{today}.xlsx"

        path, _ = QFileDialog.getSaveFileName(
            self, "엑셀 저장", default_name,
            "Excel 파일 (*.xlsx)"
        )
        if not path:
            return

        try:
            save_output_excel(self.result_df, path)
            self.out_path = path
            self.status.showMessage(f"저장 완료: {path}")
            QMessageBox.information(self, "저장 완료", f"파일이 저장되었습니다:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "저장 오류", str(e))

    # ── 카카오톡 전송
    def _send_kakao(self):
        if not self.out_path or not os.path.exists(self.out_path):
            # 먼저 저장
            reply = QMessageBox.question(
                self, "파일 저장 필요",
                "카카오톡 전송 전에 파일을 먼저 저장해야 합니다.\n저장하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._save_excel()
            if not self.out_path:
                return

        try:
            from kakao import send_file_simple
            ok = send_file_simple(
                contact_name="미트엘",
                file_path=self.out_path,
                status_callback=lambda m: self.status.showMessage(m)
            )
            if ok:
                QMessageBox.information(self, "전송 완료", "카카오톡으로 파일을 전송했습니다!")
            else:
                QMessageBox.warning(self, "전송 실패", "카카오톡 전송에 실패했습니다.\n카카오톡이 실행 중인지 확인하세요.")
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))


# ──────────────────────────────────────────
# 실행
# ──────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
