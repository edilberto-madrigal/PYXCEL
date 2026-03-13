import pytest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from PyQt6.QtCore import Qt
except ImportError:
    Qt = None

from pyxcel.models.spreadsheet import CellData, CellFormat, SpreadsheetModel

Horizontal = 1
Vertical = 2
DisplayRole = 0
EditRole = 2


class TestCellData:
    def test_cell_data_creation(self):
        cell = CellData()
        assert cell.value is None
        assert cell.formula is None
        assert cell.is_calculated is False

    def test_cell_data_with_value(self):
        cell = CellData(value=42)
        assert cell.value == 42

    def test_cell_data_with_formula(self):
        cell = CellData(value=10, formula="=5*2")
        assert cell.value == 10
        assert cell.formula == "=5*2"

    def test_cell_data_repr(self):
        cell = CellData(value=10, formula="=5*2")
        assert "CellData" in repr(cell)
        assert "10" in repr(cell)


class TestCellFormat:
    def test_cell_format_defaults(self):
        fmt = CellFormat()
        assert fmt.bold is False
        assert fmt.italic is False
        assert fmt.underline is False
        assert fmt.font_size == 10
        assert fmt.number_format == "General"

    def test_cell_format_custom_values(self):
        fmt = CellFormat()
        fmt.bold = True
        fmt.italic = True
        fmt.font_size = 14
        assert fmt.bold is True
        assert fmt.italic is True
        assert fmt.font_size == 14


class TestSpreadsheetModel:
    def test_model_creation(self):
        model = SpreadsheetModel(10, 5)
        assert model.rowCount() == 10
        assert model.columnCount() == 5

    def test_model_flags(self):
        model = SpreadsheetModel(10, 5)
        flags = model.flags(model.index(0, 0))
        assert flags != 0

    def test_set_and_get_cell_value(self):
        model = SpreadsheetModel(10, 5)
        model.setData(model.index(0, 0), "Test Value")
        value = model.data(model.index(0, 0))
        assert value == "Test Value"

    def test_set_cell_formula(self):
        model = SpreadsheetModel(10, 5)
        model.setData(model.index(0, 0), "=A1+B1", role=EditRole)
        value = model.data(model.index(0, 0), role=EditRole)
        assert value == "=A1+B1"

    def test_header_data_horizontal(self):
        model = SpreadsheetModel(10, 5)
        header = model.headerData(0, Horizontal, DisplayRole)
        assert header == "A"
        header = model.headerData(1, Horizontal, DisplayRole)
        assert header == "B"

    def test_header_data_vertical(self):
        model = SpreadsheetModel(10, 5)
        header = model.headerData(0, Vertical, DisplayRole)
        assert header == "1"
        header = model.headerData(4, Vertical, DisplayRole)
        assert header == "5"

    def test_get_cell_format(self):
        model = SpreadsheetModel(10, 5)
        fmt = model.get_cell_format(0, 0)
        assert isinstance(fmt, CellFormat)

    def test_set_cell_format(self):
        model = SpreadsheetModel(10, 5)
        fmt = model.get_cell_format(0, 0)
        fmt.bold = True
        fmt.font_size = 14
        model.set_cell_format(0, 0, fmt)

        new_fmt = model.get_cell_format(0, 0)
        assert new_fmt.bold is True
        assert new_fmt.font_size == 14

    def test_sort_by_column(self):
        model = SpreadsheetModel(5, 2)
        model.setData(model.index(0, 0), 3)
        model.setData(model.index(1, 0), 1)
        model.setData(model.index(2, 0), 2)

        model.sort_by_column(0, ascending=True)

        values = [
            model.data(model.index(0, 0)),
            model.data(model.index(1, 0)),
            model.data(model.index(2, 0)),
        ]
        assert values == ["1", "2", "3"]

    def test_undo_redo(self):
        model = SpreadsheetModel(10, 5)

        model.setData(model.index(0, 0), "Value 1")
        model.setData(model.index(1, 0), "Value 2")

        undo_stack = model.get_undo_stack()
        assert len(undo_stack) >= 2

    def test_max_rows_cols(self):
        model = SpreadsheetModel(20000, 2000)
        assert model.rowCount() == model.MAX_ROWS
        assert model.columnCount() == model.MAX_COLS
