import pytest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pyxcel.models.workbook import Workbook
from pyxcel.models.spreadsheet import SpreadsheetModel


class TestWorkbook:
    def test_workbook_creation(self):
        wb = Workbook()
        assert wb is not None
        assert wb.sheet_count() == 0

    def test_add_sheet(self):
        wb = Workbook()
        sheet = wb.add_sheet("Sheet1")
        assert sheet is not None
        assert wb.sheet_count() == 1

    def test_add_default_sheet(self):
        wb = Workbook()
        sheet = wb.add_sheet()
        assert sheet is not None
        assert "Hoja" in wb.get_sheet_name(0)

    def test_get_sheet(self):
        wb = Workbook()
        sheet = wb.add_sheet("TestSheet")
        retrieved = wb.get_sheet(0)
        assert retrieved is sheet

    def test_get_sheet_by_name(self):
        wb = Workbook()
        wb.add_sheet("TestSheet")
        sheet = wb.get_sheet_by_name("TestSheet")
        assert sheet is not None

    def test_get_sheet_name(self):
        wb = Workbook()
        wb.add_sheet("MySheet")
        assert wb.get_sheet_name(0) == "MySheet"

    def test_rename_sheet(self):
        wb = Workbook()
        wb.add_sheet("OldName")
        wb.rename_sheet(0, "NewName")
        assert wb.get_sheet_name(0) == "NewName"

    def test_insert_sheet(self):
        wb = Workbook()
        wb.add_sheet("Sheet1")
        wb.insert_sheet(0, "Sheet0")
        assert wb.sheet_count() == 2
        assert wb.get_sheet_name(0) == "Sheet0"

    def test_remove_sheet(self):
        wb = Workbook()
        wb.add_sheet("Sheet1")
        wb.add_sheet("Sheet2")
        wb.remove_sheet(0)
        assert wb.sheet_count() == 1
        assert wb.get_sheet_name(0) == "Sheet2"

    def test_cannot_remove_last_sheet(self):
        wb = Workbook()
        wb.add_sheet("Sheet1")
        wb.remove_sheet(0)
        assert wb.sheet_count() == 1

    def test_set_active_sheet(self):
        wb = Workbook()
        wb.add_sheet("Sheet1")
        wb.add_sheet("Sheet2")
        wb.set_active_sheet(1)
        assert wb.get_active_sheet() == 1

    def test_file_path(self):
        wb = Workbook()
        assert wb.get_file_path() is None
        wb.set_file_path("/path/to/file.xlsx")
        assert wb.get_file_path() == "/path/to/file.xlsx"

    def test_is_modified(self):
        wb = Workbook()
        assert wb.is_modified() is False
        wb.set_modified(True)
        assert wb.is_modified() is True

    def test_current_sheet(self):
        wb = Workbook()
        sheet = wb.add_sheet("TestSheet")
        current = wb.current_sheet()
        assert current is sheet

    def test_clear_workbook(self):
        wb = Workbook()
        wb.add_sheet("Sheet1")
        wb.add_sheet("Sheet2")
        wb.clear()
        assert wb.sheet_count() == 0
