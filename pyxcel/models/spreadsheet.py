from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex, Signal
from PySide6.QtGui import QColor, QFont
from typing import Any, Dict, Optional, List, Tuple
import re


class CellData:
    def __init__(self, value: Any = None, formula: str = None):
        self.value = value
        self.formula = formula
        self.format = CellFormat()
        self.is_calculated = False

    def __repr__(self):
        return f"CellData(value={self.value}, formula={self.formula})"


class CellFormat:
    # Default dark theme colors
    DEFAULT_BG = QColor("#11111B")
    DEFAULT_TEXT = QColor("#CDD6F4")

    def __init__(self):
        self.bold = False
        self.italic = False
        self.underline = False
        self.font_size = 10
        self.font_color = self.DEFAULT_TEXT
        self.bg_color = self.DEFAULT_BG
        self.alignment = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        self.number_format = "General"
        self.border_top = False
        self.border_bottom = False
        self.border_left = False
        self.border_right = False
        self.border_color = QColor("#45475A")
        self.font_family = "JetBrains Mono"


class SpreadsheetModel(QAbstractTableModel):
    cell_changed = Signal(int, int)
    formula_evaluated = Signal(int, int, object)

    MAX_ROWS = 10000
    MAX_COLS = 1000

    def __init__(self, rows: int = 100, cols: int = 26):
        super().__init__()
        self._rows = min(rows, self.MAX_ROWS)
        self._cols = min(cols, self.MAX_COLS)
        self._data: Dict[Tuple[int, int], CellData] = {}
        self._formula_engine = None
        self._undo_stack: List[Dict] = []
        self._redo_stack: List[Dict] = []

    def set_formula_engine(self, engine):
        self._formula_engine = engine

    def rowCount(self, parent=QModelIndex()):
        return self._rows

    def columnCount(self, parent=QModelIndex()):
        return self._cols

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return (
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEditable
        )

    def _col_num_to_letter(self, col: int) -> str:
        result = ""
        while col >= 0:
            result = chr(65 + (col % 26)) + result
            col = col // 26 - 1
        return result

    def headerData(
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.DisplayRole,
    ):
        if role == Qt.DisplayRole or role == 0:
            if orientation == Qt.Horizontal or orientation == 1:
                return self._col_num_to_letter(section)
            elif orientation == Qt.Vertical or orientation == 2:
                return str(section + 1)
        return None

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        row, col = index.row(), index.column()
        cell = self._data.get((row, col))

        if cell is None:
            if role == Qt.ItemDataRole.DisplayRole:
                return ""
            elif role == Qt.ItemDataRole.BackgroundRole:
                return CellFormat.DEFAULT_BG
            elif role == Qt.ItemDataRole.ForegroundRole:
                return CellFormat.DEFAULT_TEXT
            return None

        if role == Qt.ItemDataRole.DisplayRole:
            if cell.formula and not cell.is_calculated:
                return f"={cell.formula}"
            return str(cell.value) if cell.value is not None else ""

        elif role == Qt.ItemDataRole.ForegroundRole:
            return cell.format.font_color

        elif role == Qt.ItemDataRole.EditRole:
            if cell.formula:
                return f"={cell.formula}" if cell.formula else str(cell.value)
            return str(cell.value) if cell.value is not None else ""

        elif role == Qt.ItemDataRole.BackgroundRole:
            return cell.format.bg_color

        elif role == Qt.ItemDataRole.FontRole:
            font = QFont(cell.format.font_family)
            font.setPointSize(cell.format.font_size)
            font.setBold(cell.format.bold)
            font.setItalic(cell.format.italic)
            font.setUnderline(cell.format.underline)
            return font

        elif role == Qt.ItemDataRole.TextAlignmentRole:
            return cell.format.alignment

        return None

    def setData(
        self, index: QModelIndex, value: Any, role: int = Qt.ItemDataRole.EditRole
    ):
        if not index.isValid():
            return False

        row, col = index.row(), index.column()

        if role == Qt.ItemDataRole.EditRole:
            old_cell = self._data.get((row, col))
            old_value = old_cell.value if old_cell else None
            old_formula = old_cell.formula if old_cell else None

            self._undo_stack.append(
                {
                    "type": "cell_change",
                    "row": row,
                    "col": col,
                    "old_value": old_value,
                    "old_formula": old_formula,
                    "new_value": value,
                    "new_formula": None,
                }
            )
            self._redo_stack.clear()

            if isinstance(value, str) and value.startswith("="):
                formula = value[1:]
                cell = self._data.get((row, col), CellData())
                cell.formula = formula
                cell.value = None
                cell.is_calculated = False
                self._data[(row, col)] = cell
                self._evaluate_formula(row, col, formula)
            else:
                cell = self._data.get((row, col), CellData())
                cell.value = value
                cell.formula = None
                cell.is_calculated = True
                self._data[(row, col)] = cell

            self.dataChanged.emit(index, index, [role])
            self.cell_changed.emit(row, col)
            return True

        return False

    def _evaluate_formula(self, row: int, col: int, formula: str):
        if self._formula_engine is None:
            return

        try:
            result = self._formula_engine.evaluate(formula, self, row, col)
            cell = self._data.get((row, col))
            if cell:
                cell.value = result
                cell.is_calculated = True
                index = self.index(row, col)
                self.dataChanged.emit(index, index, [Qt.ItemDataRole.DisplayRole])
                self.formula_evaluated.emit(row, col, result)
        except Exception as e:
            print(f"Error evaluating formula: {e}")

    def evaluate_all_formulas(self):
        for (row, col), cell in self._data.items():
            if cell.formula:
                self._evaluate_formula(row, col, cell.formula)

    def get_cell_value(self, row: int, col: int) -> Any:
        cell = self._data.get((row, col))
        return cell.value if cell else None

    def get_cell_formula(self, row: int, col: int) -> Optional[str]:
        cell = self._data.get((row, col))
        return cell.formula if cell else None

    def get_cell_format(self, row: int, col: int) -> CellFormat:
        cell = self._data.get((row, col))
        if cell is None:
            cell = CellData()
            self._data[(row, col)] = cell
        return cell.format

    def set_cell_format(self, row: int, col: int, fmt: CellFormat):
        cell = self._data.get((row, col), CellData())
        cell.format = fmt
        self._data[(row, col)] = cell
        index = self.index(row, col)
        self.dataChanged.emit(
            index,
            index,
            [
                Qt.ItemDataRole.BackgroundRole,
                Qt.ItemDataRole.ForegroundRole,
                Qt.ItemDataRole.FontRole,
                Qt.ItemDataRole.TextAlignmentRole,
            ],
        )

    def insert_rows(self, row: int, count: int = 1):
        self.beginInsertRows(QModelIndex(), row, row + count - 1)

        new_data = {}
        for (r, c), cell in self._data.items():
            if r >= row:
                new_data[(r + count, c)] = cell
            else:
                new_data[(r, c)] = cell

        self._data = new_data
        self.endInsertRows()

    def insert_columns(self, col: int, count: int = 1):
        self.beginInsertColumns(QModelIndex(), col, col + count - 1)

        new_data = {}
        for (r, c), cell in self._data.items():
            if c >= col:
                new_data[(r, c + count)] = cell
            else:
                new_data[(r, c)] = cell

        self._data = new_data
        self.endInsertColumns()

    def remove_rows(self, row: int, count: int = 1):
        self.beginRemoveRows(QModelIndex(), row, row + count - 1)

        new_data = {}
        for (r, c), cell in self._data.items():
            if r < row:
                new_data[(r, c)] = cell
            elif r >= row + count:
                new_data[(r - count, c)] = cell

        self._data = new_data
        self.endRemoveRows()

    def remove_columns(self, col: int, count: int = 1):
        self.beginRemoveColumns(QModelIndex(), col, col + count - 1)

        new_data = {}
        for (r, c), cell in self._data.items():
            if c < col:
                new_data[(r, c)] = cell
            elif c >= col + count:
                new_data[(r, c - count)] = cell

        self._data = new_data
        self.endRemoveColumns()

    def undo(self):
        if not self._undo_stack:
            return

        action = self._undo_stack.pop()

        if action["type"] == "cell_change":
            row, col = action["row"], action["col"]

            self._redo_stack.append(
                {
                    "type": "cell_change",
                    "row": row,
                    "col": col,
                    "old_value": action["new_value"],
                    "old_formula": action["new_formula"],
                    "new_value": action["old_value"],
                    "new_formula": action["old_formula"],
                }
            )

            cell = self._data.get((row, col), CellData())
            cell.value = action["old_value"]
            cell.formula = action["old_formula"]
            cell.is_calculated = True
            self._data[(row, col)] = cell

            index = self.index(row, col)
            self.dataChanged.emit(
                index, index, [Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole]
            )

    def redo(self):
        if not self._redo_stack:
            return

        action = self._redo_stack.pop()

        if action["type"] == "cell_change":
            row, col = action["row"], action["col"]

            self._undo_stack.append(
                {
                    "type": "cell_change",
                    "row": row,
                    "col": col,
                    "old_value": action["new_value"],
                    "old_formula": action["new_formula"],
                    "new_value": action["old_value"],
                    "new_formula": action["old_formula"],
                }
            )

            cell = self._data.get((row, col), CellData())
            cell.value = action["old_value"]
            cell.formula = action["old_formula"]
            cell.is_calculated = True
            self._data[(row, col)] = cell

            index = self.index(row, col)
            self.dataChanged.emit(
                index, index, [Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole]
            )

    def can_undo(self) -> bool:
        return len(self._undo_stack) > 0

    def can_redo(self) -> bool:
        return len(self._redo_stack) > 0

    def get_undo_stack(self) -> List[Dict]:
        return self._undo_stack

    def clear(self):
        self.beginResetModel()
        self._data.clear()
        self._undo_stack.clear()
        self._redo_stack.clear()
        self.endResetModel()

    def get_all_data(self) -> Dict[Tuple[int, int], CellData]:
        return self._data.copy()

    def set_all_data(self, data: Dict[Tuple[int, int], CellData]):
        self.beginResetModel()
        self._data = data
        self.endResetModel()
        self.evaluate_all_formulas()

    def find_and_replace(
        self, find_text: str, replace_text: str, case_sensitive: bool = False
    ):
        changes = []
        for (row, col), cell in self._data.items():
            if cell.value and isinstance(cell.value, str):
                if case_sensitive:
                    if find_text in cell.value:
                        cell.value = cell.value.replace(find_text, replace_text)
                        changes.append((row, col))
                else:
                    if find_text.lower() in cell.value.lower():
                        cell.value = cell.value.replace(find_text, replace_text, 1)
                        changes.append((row, col))
        return changes

    def sort_by_column(self, col: int, ascending: bool = True):
        items = []
        for row in range(self._rows):
            cell = self._data.get((row, col))
            value = cell.value if cell else None
            items.append((value, row))

        items.sort(key=lambda x: (x[0] is None, x[0]), reverse=not ascending)

        new_data = {}
        for new_row, (value, old_row) in enumerate(items):
            for c in range(self._cols):
                if (old_row, c) in self._data:
                    new_data[(new_row, c)] = self._data[(old_row, c)]

        self._data = new_data
        self.beginResetModel()
        self.endResetModel()

    def sort_by_multiple_columns(self, sort_levels: list):
        if not sort_levels:
            return

        items = []
        for row in range(self._rows):
            values = []
            for level in sort_levels:
                col = level["column"]
                cell = self._data.get((row, col))
                value = cell.value if cell else None
                values.append(value)
            items.append((values, row))

        def sort_key(item):
            values = item[0]
            result = []
            for i, level in enumerate(sort_levels):
                val = values[i] if i < len(values) else None
                result.append((val is None, val))
            return tuple(result)

        items.sort(key=sort_key, reverse=False)

        new_data = {}
        for new_row, (values, old_row) in enumerate(items):
            for c in range(self._cols):
                if (old_row, c) in self._data:
                    new_data[(new_row, c)] = self._data[(old_row, c)]

        self._data = new_data
        self.beginResetModel()
        self.endResetModel()
