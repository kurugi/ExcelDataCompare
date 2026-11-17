"""Excel Compare Pro — PySide6 / Qt Designer 기반 비교 분석 화면.

설치: pip install PySide6 openpyxl
실행: python main.py
화면 편집: Qt Designer에서 ui/mainwindow.ui 열기
"""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QLabel, QLineEdit, QMessageBox,
    QPushButton, QSpinBox, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
    QLayout, QSizePolicy, QWidget, QDialog, QDialogButtonBox, QVBoxLayout, QHBoxLayout,
    QScrollArea,
)
from PySide6.QtUiTools import QUiLoader

BASE = Path(__file__).resolve().parent
COLUMNS = ['상태', '식별 키', '비교 열', 'A (기준 파일)', 'B (대상 파일)', '상세 내용']
STATUS_FILTER = {
    '전체 상태': None, '값 불일치': '값 불일치', 'A에만 존재': 'A에만 존재',
    'B에만 존재': 'B에만 존재', '일치': '일치',
    '중복 키': '중복 키', '키 누락': '키 누락',
}
DISPLAY_LIMIT = 5000
HEADER_SCAN_LIMIT = 100


def detect_header_row(rows):
    """Return the first row with at least three nonblank cells."""
    for row_no, row in enumerate(rows, 1):
        if sum(value is not None and bool(to_text(value).strip()) for value in row) >= 3:
            return row_no
    return None


def to_text(value) -> str:
    if value is None:
        return ''
    if isinstance(value, (date, datetime)):
        return value.isoformat(sep=' ') if isinstance(value, datetime) else value.isoformat()
    return str(value)


def clean_header(value, idx: int) -> str:
    return (to_text(value).strip() or f'열{idx + 1}')


class ExcelCompareApp:
    def __init__(self):
        loader = QUiLoader()
        self.window = loader.load(str(BASE / 'ui' / 'mainwindow.ui'))
        if self.window is None:
            raise RuntimeError('ui/mainwindow.ui를 열 수 없습니다. UI 파일 위치를 확인하세요.')
        style_path = BASE / 'theme.qss'
        if style_path.is_file():
            self.window.setStyleSheet(style_path.read_text(encoding='utf-8'))
        self.file_paths = {'A': None, 'B': None}
        self.headers = {'A': [], 'B': []}
        self.compare_pairs = []
        self.results: list[tuple[str, str, str, str, str, str]] = []
        self.metrics = {'Total': 0, 'Same': 0, 'Different': 0, 'OnlyA': 0, 'OnlyB': 0}

        self.table = self.w('resultTable', QTableWidget)
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.setSortingEnabled(False)
        self.table.setColumnWidth(0, 110)
        self.table.setColumnWidth(1, 145)
        self.table.setColumnWidth(2, 145)
        self.table.setColumnWidth(3, 170)
        self.table.setColumnWidth(4, 170)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(34)

        for side in 'AB':
            self.w(f'btnFile{side}', QPushButton).clicked.connect(lambda checked=False, s=side: self.choose_file(s))
            self.w(f'comboSheet{side}', QComboBox).currentIndexChanged.connect(lambda _=0, s=side: self.update_headers(s, auto_detect=True))
            self.w(f'spinHeader{side}', QSpinBox).valueChanged.connect(lambda _=0, s=side: self.update_headers(s))
            self.w(f'spinHeader{side}', QSpinBox).setToolTip(
                '파일·시트 선택 시 상단 100행에서 값이 있는 셀 3개 이상인 첫 행을 자동 인식합니다. 직접 수정할 수 있습니다.'
            )

        self.w('btnCompare', QPushButton).clicked.connect(self.run_compare)
        self.w('btnReset', QPushButton).clicked.connect(self.reset)
        self.w('btnExport', QPushButton).clicked.connect(self.export_excel)
        self.w('btnNavExport', QPushButton).clicked.connect(self.export_excel)
        self.w('btnNavHelp', QPushButton).clicked.connect(self.show_help)
        self.w('btnNavCompare', QPushButton).clicked.connect(lambda: self.w('editPathA', QLineEdit).setFocus())
        self.w('comboFilter', QComboBox).currentTextChanged.connect(self.refresh_result_table)
        self.w('editSearch', QLineEdit).textChanged.connect(self.refresh_result_table)

        # 이전 비교 결과를 잘못 저장하지 않도록 설정이 바뀌면 결과를 초기화합니다.
        for side in 'AB':
            for combo in self.key_combos(side):
                combo.currentIndexChanged.connect(
                    lambda _=0, s=side: self.on_key_changed(s)
                )
        for checkbox_name in ('checkTrim', 'checkCase', 'checkBlank'):
            self.w(checkbox_name, QCheckBox).toggled.connect(self.invalidate_results)
        edit = self.w('editCompareFields', QLineEdit)
        edit.setReadOnly(True)
        edit.setPlaceholderText('비교 열 선택 · 미선택 시 식별 키의 존재 여부만 비교')
        button = self.w('btnCompareFields', QPushButton)
        button.clicked.connect(self.choose_compare_fields)
        self.configure_layout()

    def refresh_compare_fields(self):
        self.compare_pairs = [(a, b) for a, b in self.compare_pairs
                              if a in self.headers['A'] and b in self.headers['B']]
        text = ' / '.join(f'{a} → {b}' for a, b in self.compare_pairs)
        self.w('editCompareFields', QLineEdit).setText(text)
        self.w('editCompareFields', QLineEdit).setToolTip(text)

    def choose_compare_fields(self):
        if not all(self.headers.values()):
            QMessageBox.information(self.window, '파일 선택', '기준파일과 대상파일을 먼저 선택하세요.')
            return
        dialog = QDialog(self.window)
        dialog.setWindowTitle('비교 열 선택')
        dialog.resize(680, 420)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel('같은 행의 기준파일 열과 대상파일 열을 서로 비교합니다.'))
        table = QTableWidget(0, 2, dialog)
        table.setHorizontalHeaderLabels(['A · 기준파일 열', 'B · 대상파일 열'])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(table)

        def add_pair(a=None, b=None):
            row = table.rowCount()
            table.insertRow(row)
            for column, side, selected in ((0, 'A', a), (1, 'B', b)):
                combo = QComboBox(table)
                combo.addItem('열을 선택하세요', None)
                for name in self.headers[side]:
                    combo.addItem(name, name)
                if selected in self.headers[side]:
                    combo.setCurrentIndex(combo.findData(selected))
                table.setCellWidget(row, column, combo)

        actions = QHBoxLayout()
        add = QPushButton('비교 쌍 추가', dialog)
        remove = QPushButton('선택한 쌍 삭제', dialog)
        add.clicked.connect(lambda: add_pair())
        remove.clicked.connect(lambda: table.removeRow(table.currentRow()) if table.currentRow() >= 0 else None)
        actions.addWidget(add)
        actions.addWidget(remove)
        layout.addLayout(actions)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText('적용')
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText('취소')
        layout.addWidget(buttons)

        def apply():
            pairs = []
            for row in range(table.rowCount()):
                pair = tuple(table.cellWidget(row, col).currentData() for col in (0, 1))
                if None in pair:
                    QMessageBox.warning(dialog, '비교 열 선택', f'{row + 1}번째 쌍의 양쪽 열을 모두 선택하세요.')
                    return
                if pair not in pairs:
                    pairs.append(pair)
            if pairs != self.compare_pairs:
                self.invalidate_results()
                self.compare_pairs = pairs
                self.refresh_compare_fields()
            dialog.accept()

        buttons.accepted.connect(apply)
        buttons.rejected.connect(dialog.reject)
        for a, b in self.compare_pairs:
            add_pair(a, b)
        if not self.compare_pairs:
            add_pair()
        dialog.exec()

    def configure_layout(self):
        """Reserve each section's minimum height and scroll on smaller screens."""
        main_area = self.w('mainArea', QWidget)
        grid = main_area.findChild(QLayout, 'keyConditionsGrid')
        for column in range(1, 5):
            grid.setColumnStretch(column, 1)
        for side in 'AB':
            first = self.w(f'comboKey{side}', QComboBox)
            first.setToolTip('조건 1 (필수)')
            for index in range(2, 5):
                combo = self.w(f'comboKey{side}{index}', QComboBox)
                combo.clear()
                combo.addItem(f'조건 {index}: 선택 안 함', None)
                combo.setToolTip(f'조건 {index} (선택) · 같은 번호의 A/B 열을 AND로 연결합니다.')
        self.w('ruleNote', QLabel).setText('최대 4개 조건 AND · 같은 번호의 A/B 키가 모두 같은 행 비교')
        content_layout = main_area.layout()
        content_layout.setContentsMargins(20, 16, 20, 14)
        content_layout.setSpacing(10)
        for name in ('cardFiles', 'cardRules', 'cardResults'):
            layout = self.w(name, QWidget).layout()
            layout.setContentsMargins(16, 12, 16, 12)
            layout.setSpacing(8)
        for side in 'AB':
            layout = self.w(f'file{side}', QWidget).layout()
            layout.setContentsMargins(12, 10, 12, 10)
            layout.setSpacing(6)

        # Polish before measuring: stylesheet padding and fonts affect size hints.
        self.window.ensurePolished()
        for widget in main_area.findChildren(QWidget):
            widget.ensurePolished()
            if isinstance(widget, (QLabel, QPushButton, QLineEdit, QComboBox, QSpinBox, QCheckBox)):
                widget.setMinimumHeight(max(widget.minimumHeight(), widget.sizeHint().height()))

        for side in 'AB':
            badge = self.w(f'fileBadge{side}', QLabel)
            badge.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            badge.setMinimumSize(badge.sizeHint())
            button = self.w(f'btnFile{side}', QPushButton)
            button.setMinimumWidth(max(button.minimumWidth(), button.sizeHint().width()))
            sheet = self.w(f'comboSheet{side}', QComboBox)
            sheet.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
            sheet.setMinimumContentsLength(16)
            # Reserve the full placeholder width, including padding and the arrow.
            placeholder_width = sheet.fontMetrics().horizontalAdvance('파일을 먼저 선택하세요')
            sheet.setMinimumWidth(max(sheet.sizeHint().width(), placeholder_width + 56))
            sheet.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            sheet_row = main_area.findChild(QLayout, f'sheetRow{side}')
            sheet_row.setStretch(1, 1)
            self.w(f'spinHeader{side}', QSpinBox).setFixedWidth(70)
            self.w(f'file{side}', QWidget).layout().setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)

        version = self.w('versionLabel', QLabel)
        version.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        header = main_area.findChild(QLayout, 'headerLayout')
        header.setStretch(0, 1)
        header.setAlignment(version, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        for name in ('cardFiles', 'cardRules'):
            card = self.w(name, QWidget)
            card.layout().setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)
            card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Minimum)
        for key in ('Total', 'Same', 'Different', 'OnlyA', 'OnlyB'):
            card = self.w(f'metric{key}', QWidget)
            card.layout().setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)
            card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        self.w('resultTable', QTableWidget).setMinimumHeight(100)
        # Only the result panel consumes spare height; the summary row stays compact.
        for index in range(content_layout.count()):
            content_layout.setStretch(index, 0)
        content_layout.setStretch(content_layout.indexOf(self.w('cardResults', QWidget)), 1)
        content_layout.setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)

        # A fixed viewport must not squeeze controls below their required heights.
        # The scroll area appears only when the screen cannot fit the full form.
        root = self.window.findChild(QLayout, 'rootLayout')
        scroll = QScrollArea(self.window)
        scroll.setObjectName('contentScroll')
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        old_item = root.replaceWidget(main_area, scroll)
        scroll.setWidget(main_area)
        del old_item
        self.window.setWindowFlag(Qt.WindowType.WindowMaximizeButtonHint, False)
        self.window.layout().activate()
        available = self.window.screen().availableGeometry()
        # Qt screen geometry uses logical pixels, so Windows display scaling is respected.
        width = min(1440, available.width() - 20)
        height = min(900, available.height() - 50)
        self.window.setFixedSize(width, height)
        self.window.move(
            available.x() + (available.width() - width) // 2,
            available.y() + max(0, (available.height() - height - 30) // 2),
        )

    def key_combos(self, side):
        return [self.w(f'comboKey{side}{suffix}', QComboBox) for suffix in ('', '2', '3', '4')]

    def on_key_changed(self, side):
        self.refresh_key_options(side)
        self.invalidate_results()

    def refresh_key_options(self, side):
        """Disable columns already selected by another condition in this file."""
        combos = self.key_combos(side)
        selected = [
            combo.currentText()
            if combo.currentIndex() >= 0 and (index == 0 or combo.currentData() is not None)
            else None
            for index, combo in enumerate(combos)
        ]
        for index, combo in enumerate(combos):
            used = {value for other, value in enumerate(selected) if other != index and value is not None}
            for row in range(combo.count()):
                item = combo.model().item(row)
                is_placeholder = index > 0 and combo.itemData(row) is None
                item.setEnabled(is_placeholder or combo.itemText(row) not in used)

    def selected_keys(self, side):
        combos = self.key_combos(side)
        keys = [combos[0].currentText()] if combos[0].currentIndex() >= 0 else []
        keys.extend(combo.currentText() for combo in combos[1:] if combo.currentData() is not None)
        if not keys:
            raise ValueError(f'{side} 파일에서 첫 번째 식별 키를 선택하세요.')
        if len(keys) != len(set(keys)):
            raise ValueError(f'{side} 파일의 식별 키에 같은 열을 중복 선택할 수 없습니다.')
        return keys

    def w(self, name, klass):
        result = self.window.findChild(klass, name)
        if result is None:
            raise RuntimeError(f'UI 위젯을 찾을 수 없습니다: {name}')
        return result

    def status(self, message: str):
        self.w('statusLabel', QLabel).setText('●  ' + message)

    def invalidate_results(self, *_):
        """비교 조건이 바뀌면 오래된 결과와 통계를 제거합니다."""
        if not self.results and all(value == 0 for value in self.metrics.values()):
            return
        self.results.clear()
        self.metrics = dict.fromkeys(self.metrics, 0)
        for key in self.metrics:
            self.w(f'metricNum{key}', QLabel).setText('0')
        self.refresh_result_table()
        self.w('resultDescription', QLabel).setText('설정이 변경되었습니다. 비교를 다시 실행해 주세요.')
        self.status('비교 조건 변경됨 · 비교를 다시 실행해 주세요.')

    def choose_file(self, side: str):
        path, _ = QFileDialog.getOpenFileName(
            self.window, f'{side} 엑셀 파일 선택', '',
            'Excel 파일 (*.xlsx *.xlsm);;모든 파일 (*.*)'
        )
        if not path:
            return
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            names = list(wb.sheetnames)
            wb.close()
        except Exception as exc:
            QMessageBox.warning(self.window, '파일을 열 수 없음', f'{side} 파일을 열 수 없습니다.\n\n{exc}')
            return
        self.invalidate_results()
        self.file_paths[side] = path
        self.w(f'editPath{side}', QLineEdit).setText(path)
        combo = self.w(f'comboSheet{side}', QComboBox)
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(names)
        combo.blockSignals(False)
        self.update_headers(side, auto_detect=True)

    def update_headers(self, side: str, auto_detect=False):
        self.invalidate_results()
        path = self.file_paths[side]
        if not path:
            return
        sheet_name = self.w(f'comboSheet{side}', QComboBox).currentText()
        header_row = self.w(f'spinHeader{side}', QSpinBox).value()
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            detected = None
            try:
                ws = wb[sheet_name]
                if auto_detect:
                    detected = detect_header_row(ws.iter_rows(max_row=HEADER_SCAN_LIMIT, values_only=True))
                    header_row = detected or 1
                    spin = self.w(f'spinHeader{side}', QSpinBox)
                    spin.blockSignals(True)
                    spin.setValue(header_row)
                    spin.blockSignals(False)
                raw = next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True), ())
            finally:
                wb.close()
            names, counts = [], Counter()
            for i, x in enumerate(raw):
                name = clean_header(x, i)
                counts[name] += 1
                names.append(name if counts[name] == 1 else f'{name} ({counts[name]})')
            self.headers[side] = names
            key_combo = self.w(f'comboKey{side}', QComboBox)
            old = key_combo.currentText()
            key_combo.blockSignals(True)
            key_combo.clear()
            key_combo.addItems(names)
            if old in names:
                key_combo.setCurrentText(old)
            elif names:
                # ID, 코드, 상품코드 등 자주 쓰는 식별 키를 우선으로 선택
                for i, name in enumerate(names):
                    if name.strip().lower() in {'id', '상품코드', '제품코드', '관리번호', '바코드', '사번'}:
                        key_combo.setCurrentIndex(i)
                        break
            key_combo.blockSignals(False)
            for index, combo in enumerate(self.key_combos(side)[1:], 2):
                old = combo.currentData()
                combo.blockSignals(True)
                combo.clear()
                combo.addItem(f'조건 {index}: 선택 안 함', None)
                for name in names:
                    combo.addItem(name, name)
                if old in names:
                    combo.setCurrentIndex(combo.findData(old))
                combo.blockSignals(False)
            self.refresh_key_options(side)
            self.refresh_compare_fields()
            if auto_detect:
                if detected is not None:
                    self.status(f'{side} 파일 헤더 {header_row}행 자동 인식 · 비교 기준을 확인해 주세요.')
                else:
                    self.status(f'{side} 파일 상단 {HEADER_SCAN_LIMIT}행에서 헤더를 찾지 못했습니다. 헤더 행을 직접 지정해 주세요. (현재 1행)')
        except Exception as exc:
            self.headers[side] = []
            self.refresh_compare_fields()
            for combo in self.key_combos(side):
                combo.clear()
            self.status(f'{side} 헤더를 읽을 수 없습니다: {exc}')

    def get_rows(self, side: str):
        from openpyxl import load_workbook
        path = self.file_paths[side]
        sheet_name = self.w(f'comboSheet{side}', QComboBox).currentText()
        header_row = self.w(f'spinHeader{side}', QSpinBox).value()
        key_indices = [self.headers[side].index(name) for name in self.selected_keys(side)]
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb[sheet_name]
            headers = self.headers[side]
            records = defaultdict(list)
            missing_keys = []
            for row_no, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1):
                values = list(row[:len(headers)]) + [None] * max(0, len(headers) - len(row))
                if not any(v is not None and to_text(v).strip() for v in values):
                    continue
                raw_parts = [to_text(values[index]) for index in key_indices]
                key = tuple(self.normalize(part) for part in raw_parts)
                raw_key = raw_parts[0] if len(raw_parts) == 1 else repr(tuple(raw_parts))
                if any(not part for part in key):
                    missing_keys.append((row_no, values))
                    continue
                records[key].append((raw_key, dict(zip(headers, values))))
            return records, missing_keys
        finally:
            wb.close()

    def normalize(self, value) -> str:
        val = to_text(value)
        if self.w('checkTrim', QCheckBox).isChecked():
            val = val.strip()
        if self.w('checkBlank', QCheckBox).isChecked() and not val.strip():
            val = ''
        if not self.w('checkCase', QCheckBox).isChecked():
            val = val.casefold()
        return val

    def run_compare(self):
        if not self.file_paths['A'] or not self.file_paths['B']:
            QMessageBox.information(self.window, '파일 선택', '기준 파일 A와 대상 파일 B를 모두 선택하세요.')
            return
        self.status('데이터를 읽고 비교하는 중입니다...')
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            self.selected_keys('A')
            self.selected_keys('B')
            slots_a = [True] + [c.currentData() is not None for c in self.key_combos('A')[1:]]
            slots_b = [True] + [c.currentData() is not None for c in self.key_combos('B')[1:]]
            if slots_a != slots_b:
                raise ValueError('기준 파일과 대상 파일에서 같은 번호의 식별 키 조건을 선택하세요.')
            fields = list(self.compare_pairs)
            if any(a not in self.headers['A'] or b not in self.headers['B'] for a, b in fields):
                raise ValueError('비교할 열을 다시 선택하세요.')
            a, a_missing = self.get_rows('A')
            b, b_missing = self.get_rows('B')
            keys = sorted(set(a) | set(b))
            results = []
            stats = dict(Total=len(keys), Same=0, Different=0, OnlyA=0, OnlyB=0)
            for k in keys:
                aa, bb = a.get(k, []), b.get(k, [])
                shown_key = (aa or bb)[0][0]
                if len(aa) > 1 or len(bb) > 1:
                    results.append(('중복 키', shown_key, '—', str(len(aa)), str(len(bb)), '키 중복: A / B 해당 행 수'))
                    continue
                if not aa:
                    stats['OnlyB'] += 1
                    results.append(('B에만 존재', shown_key, '—', '', '', '대상 파일(B)에만 존재하는 식별 키'))
                elif not bb:
                    stats['OnlyA'] += 1
                    results.append(('A에만 존재', shown_key, '—', '', '', '기준 파일(A)에만 존재하는 식별 키'))
                else:
                    a_record, b_record = aa[0][1], bb[0][1]
                    mismatches = []
                    for field_a, field_b in fields:
                        field = field_a if field_a == field_b else f'{field_a} → {field_b}'
                        av = to_text(a_record.get(field_a))
                        bv = to_text(b_record.get(field_b))
                        if self.normalize(av) != self.normalize(bv):
                            mismatches.append(('값 불일치', shown_key, field, av, bv, '같은 식별 키의 값이 다릅니다'))
                    if mismatches:
                        stats['Different'] += 1
                        results.extend(mismatches)
                    else:
                        stats['Same'] += 1
                        detail = '지정한 모든 비교 열이 일치합니다' if fields else '같은 식별 키가 두 파일에 모두 존재합니다 (값 비교 안 함)'
                        results.append(('일치', shown_key, '—', '', '', detail))
            for side, missing in [('A', a_missing), ('B', b_missing)]:
                for row_no, _ in missing:
                    results.append(('키 누락', f'{side}행 {row_no}', '식별 키', '', '', f'{side} 파일 {row_no}행의 식별 키가 비어 있습니다'))
            self.results, self.metrics = results, stats
            for key, count in stats.items():
                self.w(f'metricNum{key}', QLabel).setText(f'{count:,}')
            self.refresh_result_table()
            self.status(f'비교 완료 · 고유 키 {len(keys):,}개 · 결과 {len(results):,}건')
            self.w('resultDescription', QLabel).setText(
                (f'비교 열 {len(fields)}개' if fields else '식별 키 기준 존재 여부만 비교')
                + ' · 중복 키와 키 누락은 별도 표시됩니다.'
            )
        except Exception as exc:
            self.status('비교 중 오류가 발생했습니다.')
            QMessageBox.warning(self.window, '비교 실패', str(exc))
        finally:
            QApplication.restoreOverrideCursor()

    def refresh_result_table(self, *_):
        active = STATUS_FILTER.get(self.w('comboFilter', QComboBox).currentText())
        q = self.w('editSearch', QLineEdit).text().strip().casefold()
        filtered = [r for r in self.results if (active is None or r[0] == active) and (not q or q in ' '.join(r).casefold())]
        shown = filtered[:DISPLAY_LIMIT]
        self.table.setSortingEnabled(False)
        self.table.setUpdatesEnabled(False)
        try:
            self.table.setRowCount(len(shown))
            for i, row in enumerate(shown):
                for j, value in enumerate(row):
                    item = QTableWidgetItem(value)
                    if j == 0:
                        color = {'일치': '#16896E', '값 불일치': '#CB841E',
                                 'A에만 존재': '#C35462', 'B에만 존재': '#5A62C7',
                                 '중복 키': '#A268B5', '키 누락': '#CC4B4B'}.get(value)
                        if color:
                            item.setForeground(QColor(color))
                    self.table.setItem(i, j, item)
        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)
        self.w('labelResultCount', QLabel).setText(
            f'표시 결과: {len(shown):,}건 / 필터 결과 {len(filtered):,}건'
            + (f'  (표에서는 최대 {DISPLAY_LIMIT:,}건, 저장은 전체)' if len(filtered) > DISPLAY_LIMIT else '')
        )

    def export_excel(self):
        if not self.results:
            QMessageBox.information(self.window, '저장할 결과 없음', '비교를 먼저 실행하세요.')
            return
        file, _ = QFileDialog.getSaveFileName(self.window, '비교 결과 Excel 저장', 'ExcelCompare_결과.xlsx', 'Excel 파일 (*.xlsx)')
        if not file:
            return
        if not file.lower().endswith('.xlsx'):
            file += '.xlsx'
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment
            from openpyxl.utils import get_column_letter
            wb = Workbook()
            ws = wb.active
            ws.title = '비교결과'
            ws.append(COLUMNS)
            for row in self.results:
                ws.append(list(row))
                # 셀 내용이 '='로 시작해도 Excel 수식으로 실행되지 않도록 문자열로 저장
                for cell in ws[ws.max_row]:
                    if isinstance(cell.value, str):
                        cell.data_type = 's'
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = f'A1:F{ws.max_row}'
            for cell in ws[1]:
                cell.fill = PatternFill('solid', fgColor='193A70')
                cell.font = Font(color='FFFFFF', bold=True)
                cell.alignment = Alignment(vertical='center')
            ws.row_dimensions[1].height = 28
            for col, width in enumerate([17, 23, 22, 33, 33, 52], 1):
                ws.column_dimensions[get_column_letter(col)].width = width
            info = wb.create_sheet('요약')
            info.append(['항목', '건수'])
            labels = {'Total':'전체 고유 키', 'Same':'일치', 'Different':'값 불일치', 'OnlyA':'A에만 존재', 'OnlyB':'B에만 존재'}
            for key in ['Total','Same','Different','OnlyA','OnlyB']:
                info.append([labels[key], self.metrics[key]])
            info.column_dimensions['A'].width = 24
            info.column_dimensions['B'].width = 16
            wb.save(file)
        except Exception as exc:
            QMessageBox.warning(self.window, '저장 실패', str(exc))
            return

        # Open only after saving succeeds, without a blocking success dialog.
        self.status('결과 저장 완료 · Excel 파일 열기: ' + file)
        try:
            opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(file).resolve())))
        except Exception:
            opened = False
        if not opened:
            self.status('결과 저장 완료 · 파일 자동 열기 실패: ' + file)
            QMessageBox.warning(
                self.window, '파일 열기 실패',
                f'파일은 저장했지만 자동으로 열지 못했습니다.\n'
                f'Excel 설치 및 .xlsx 기본 앱 연결을 확인해 주세요.\n\n{file}',
            )

    def reset(self):
        for side in 'AB':
            self.file_paths[side], self.headers[side] = None, []
            self.w(f'editPath{side}', QLineEdit).clear()
            cb = self.w(f'comboSheet{side}', QComboBox)
            cb.blockSignals(True)
            cb.clear(); cb.addItem('파일을 먼저 선택하세요')
            cb.blockSignals(False)
            for index, combo in enumerate(self.key_combos(side), 1):
                combo.clear()
                if index > 1:
                    combo.addItem(f'조건 {index}: 선택 안 함', None)
            spin = self.w(f'spinHeader{side}', QSpinBox)
            spin.blockSignals(True); spin.setValue(1); spin.blockSignals(False)
        self.results.clear()
        self.compare_pairs.clear()
        self.refresh_compare_fields()
        self.w('editSearch', QLineEdit).clear()
        self.w('comboFilter', QComboBox).setCurrentIndex(0)
        self.metrics = dict.fromkeys(self.metrics, 0)
        for key in self.metrics:
            self.w(f'metricNum{key}', QLabel).setText('0')
        self.refresh_result_table()
        self.w('resultDescription', QLabel).setText('비교를 실행하면 결과가 아래 표에 표시됩니다.')
        self.status('초기화 완료. 파일을 선택해 주세요.')

    def show_help(self):
        QMessageBox.information(
            self.window, '사용 방법',
            '1. 기준 파일(A)과 대상 파일(B)을 선택합니다.\n'
            '2. 각각 시트, 헤더 행, 식별 키를 최대 4개 지정합니다.\n'
            '   헤더는 상단 100행 중 값이 있는 셀 3개 이상인 첫 행으로 자동 설정됩니다. 필요하면 직접 수정하세요.\n'
            '   같은 번호의 A/B 조건이 모두 같은 행을 비교합니다(AND).\n'
            '   첫 번째 키는 필수이며, 선택한 키 중 하나라도 비면 키 누락입니다.\n'
            '3. [비교 열 선택]에서 A/B 열을 쌍으로 지정합니다. 헤더가 달라도 비교할 수 있습니다.\n'
            '   비교 쌍을 모두 삭제하면 식별 키의 존재 여부만 확인합니다.\n'
            '4. [비교 실행]을 클릭합니다.\n'
            '5. 필터와 검색으로 결과를 확인하고 [Excel 저장]을 누릅니다. 저장 후 파일이 자동으로 열립니다.\n\n'
            '지원 형식: .xlsx, .xlsm\n'
            '주의: 수식 결과값을 비교합니다. 중복 키는 별도로 표시하고 자동 대조하지 않습니다.'
        )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName('Excel Compare Pro')
    instance = ExcelCompareApp()
    instance.window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
