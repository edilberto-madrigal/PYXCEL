from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTableView,
    QHeaderView,
    QLineEdit,
    QLabel,
    QComboBox,
    QPushButton,
    QSpacerItem,
    QSizePolicy,
)
from PySide6.QtCore import Qt, QModelIndex, Signal, QItemSelectionModel
from PySide6.QtGui import QFont, QKeyEvent

from ..models.spreadsheet import SpreadsheetModel
from .theme import theme_manager, ThemeColors, ThemeFonts, get_app_stylesheet


class SpreadsheetView(QTableView):
    cell_selected = Signal(int, int)
    cell_edited = Signal(int, int, str)
    selection_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_col = 0
        self.update_style()

    def update_style(self):
        self.setAlternatingRowColors(True)
        self.setShowGrid(True)
        self.setSelectionBehavior(QTableView.SelectionBehavior.SelectItems)
        self.setSelectionMode(QTableView.SelectionMode.ExtendedSelection)
        self.setCornerButtonEnabled(True)
        self.setWordWrap(False)
        self.setEditTriggers(
            QTableView.EditTrigger.DoubleClicked
            | QTableView.EditTrigger.EditKeyPressed
            | QTableView.EditTrigger.AnyKeyPressed
        )

        # Use JetBrains Mono or fallback
        font = QFont("JetBrains Mono")
        font.setStyleHint(QFont.StyleHint.Monospace)
        font.setPointSize(10)
        self.setFont(font)

        self.horizontalHeader().setStretchLastSection(False)
        self.verticalHeader().setStretchLastSection(False)

        self.horizontalHeader().setDefaultSectionSize(90)
        self.verticalHeader().setDefaultSectionSize(26)

        # Theme stylesheet
        C = theme_manager.colors
        self.setStyleSheet(f"""
            QTableView {{
                background-color: {C.BASE};
                color: {C.TEXT};
                gridline-color: {C.GRID};
                alternate-background-color: {C.ROW_ALT};
                border: none;
            }}
            QTableView::item {{
                background-color: {C.ROW_DEFAULT};
                color: {C.TEXT};
                padding: 2px 8px;
                border: none;
            }}
            QTableView::item:selected {{
                background-color: {C.SELECTION};
                color: {C.SELECTION_TEXT};
            }}
            QTableView::item:hover {{
                background-color: {C.HOVER};
            }}
            QHeaderView::section {{
                background-color: {C.HEADER_BG};
                color: {C.SUBTEXT0};
                padding: 8px 12px;
                border: none;
                border-bottom: 2px solid {C.BLUE};
                font-weight: 600;
                font-size: {ThemeFonts.SIZES["small"]}px;
            }}
            QHeaderView::section:hover {{
                background-color: {C.SURFACE0};
                color: {C.TEXT};
            }}
            QTableCornerButton::section {{
                background-color: {C.HEADER_BG};
                border: none;
            }}
        """)

        self.clicked.connect(self._on_cell_clicked)
        self.doubleClicked.connect(self._on_cell_double_clicked)

    def _col_num_to_letter(self, col: int) -> str:
        result = ""
        while col >= 0:
            result = chr(65 + (col % 26)) + result
            col = col // 26 - 1
        return result

    def set_column_count(self, count: int):
        pass

    def _on_cell_clicked(self, index: QModelIndex):
        self._current_row = index.row()
        self._current_col = index.column()
        self.cell_selected.emit(self._current_row, self._current_col)

    def _on_cell_double_clicked(self, index: QModelIndex):
        self._current_row = index.row()
        self._current_col = index.column()
        self.edit(index)

    def get_current_position(self) -> tuple:
        return (self._current_row, self._current_col)

    def set_current_position(self, row: int, col: int):
        self._current_row = row
        self._current_col = col
        index = self.model().index(row, col)
        if index.isValid():
            self.setCurrentIndex(index)
            self.selectionModel().setCurrentIndex(
                index, QItemSelectionModel.SelectionFlag.Select
            )

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key.Key_Return or event.key() == Qt.Key.Key_Enter:
            if self._current_row < self.model().rowCount() - 1:
                self.set_current_position(self._current_row + 1, self._current_col)
            event.accept()
        elif event.key() == Qt.Key.Key_Tab:
            if self._current_col < self.model().columnCount() - 1:
                self.set_current_position(self._current_row, self._current_col + 1)
            event.accept()
        elif event.key() == Qt.Key.Key_Delete or event.key() == Qt.Key.Key_Backspace:
            self._clear_current_cell()
            event.accept()
        else:
            super().keyPressEvent(event)

    def _clear_current_cell(self):
        if self._current_row >= 0 and self._current_col >= 0:
            index = self.model().index(self._current_row, self._current_col)
            if index.isValid():
                self.model().setData(index, "", Qt.ItemDataRole.EditRole)


class CellEditor(QWidget):
    editing_finished = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self.update_style()

    def update_style(self):
        C = theme_manager.colors
        self.position_label.setStyleSheet(f"""
            QLabel {{
                background-color: {C.SURFACE0};
                color: {C.SUBTEXT0};
                padding: 8px 10px;
                border: none;
                border-right: 1px solid {C.SURFACE1};
                font-family: '{ThemeFonts.SPREADSHEET}', monospace;
                font-size: {ThemeFonts.SIZES["small"]}px;
                font-weight: 600;
            }}
        """)
        self.formula_bar.setStyleSheet(f"""
            QLineEdit {{
                background-color: {C.BASE};
                color: {C.TEXT};
                border: none;
                padding: 8px 12px;
                font-family: '{ThemeFonts.SPREADSHEET}', monospace;
                font-size: 13px;
            }}
            QLineEdit:focus {{
                border: 2px solid {C.BLUE};
            }}
        """)

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.position_label = QLabel("A1")
        self.position_label.setFixedWidth(65)

        self.formula_bar = QLineEdit()
        self.formula_bar.setFont(QFont(ThemeFonts.SPREADSHEET, 10))

        layout.addWidget(self.position_label)
        layout.addWidget(self.formula_bar)

        self.formula_bar.returnPressed.connect(self._on_editing_finished)
        self.formula_bar.editingFinished.connect(self._on_editing_finished)

    def _on_editing_finished(self):
        self.editing_finished.emit(self.formula_bar.text())

    def set_position(self, row: int, col: int):
        col_letter = ""
        while col >= 0:
            col_letter = chr(65 + (col % 26)) + col_letter
            col = col // 26 - 1
        self.position_label.setText(f"{col_letter}{row + 1}")

    def set_text(self, text: str):
        self.formula_bar.setText(text)

    def get_text(self) -> str:
        return self.formula_bar.text()

    def clear(self):
        self.formula_bar.clear()
        self.position_label.setText("A1")


class SpreadsheetWidget(QWidget):
    cell_changed = Signal(int, int, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self.update_style()

    def update_style(self):
        self.cell_editor.update_style()
        self.table_view.update_style()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.cell_editor = CellEditor()
        self.table_view = SpreadsheetView()

        layout.addWidget(self.cell_editor)
        layout.addWidget(self.table_view)

        self.table_view.cell_selected.connect(self._on_cell_selected)
        self.cell_editor.editing_finished.connect(self._on_editing_finished)

    def set_model(self, model: SpreadsheetModel):
        self.table_view.setModel(model)

        self.table_view.set_column_count(model.columnCount())

        header = self.table_view.horizontalHeader()
        for i in range(min(model.columnCount(), 100)):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)

        self._connect_model_signals()

    def _connect_model_signals(self):
        model = self.table_view.model()
        if model:
            model.cell_changed.connect(self._on_model_cell_changed)

    def _on_cell_selected(self, row: int, col: int):
        self.cell_editor.set_position(row, col)

        model = self.table_view.model()
        if model:
            formula = model.get_cell_formula(row, col)
            if formula:
                self.cell_editor.set_text(f"={formula}")
            else:
                value = model.get_cell_value(row, col)
                self.cell_editor.set_text(str(value) if value is not None else "")

    def _on_editing_finished(self, text: str):
        row, col = self.table_view.get_current_position()

        model = self.table_view.model()
        if model:
            model.setData(model.index(row, col), text, Qt.ItemDataRole.EditRole)
            self.cell_changed.emit(row, col, text)

    def _on_model_cell_changed(self, row: int, col: int):
        if row == self.table_view._current_row and col == self.table_view._current_col:
            model = self.table_view.model()
            if model:
                formula = model.get_cell_formula(row, col)
                if formula:
                    self.cell_editor.set_text(f"={formula}")
                else:
                    value = model.get_cell_value(row, col)
                    self.cell_editor.set_text(str(value) if value is not None else "")

    def get_view(self) -> SpreadsheetView:
        return self.table_view

    def get_editor(self) -> CellEditor:
        return self.cell_editor
