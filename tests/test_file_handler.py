import pytest
import sys
import os
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pyxcel.models.workbook import Workbook
from pyxcel.utils.file_handler import FileManager


class TestFileManager:
    def test_file_manager_creation(self):
        fm = FileManager()
        assert fm is not None
        assert fm.get_current_file_path() is None

    def test_save_and_load_csv(self):
        fm = FileManager()
        wb = Workbook()
        wb.add_sheet("Sheet1")
        sheet = wb.get_sheet(0)

        sheet.setData(sheet.index(0, 0), "Name")
        sheet.setData(sheet.index(1, 0), "John")
        sheet.setData(sheet.index(2, 0), "Jane")

        sheet.setData(sheet.index(0, 1), "Age")
        sheet.setData(sheet.index(1, 1), "30")
        sheet.setData(sheet.index(2, 1), "25")

        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            temp_path = f.name

        try:
            success = fm.save_file(temp_path, wb)
            assert success is True

            wb2 = Workbook()
            success = fm.open_file(temp_path, wb2)
            assert success is True

            sheet2 = wb2.get_sheet(0)
            assert sheet2.data(sheet2.index(0, 0)) == "Name"
            assert sheet2.data(sheet2.index(1, 0)) == "John"
            assert sheet2.data(sheet2.index(2, 0)) == "Jane"

        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    def test_save_and_load_xlsx(self):
        fm = FileManager()
        wb = Workbook()
        wb.add_sheet("Sheet1")
        sheet = wb.get_sheet(0)

        sheet.setData(sheet.index(0, 0), "Product")
        sheet.setData(sheet.index(1, 0), "Apple")
        sheet.setData(sheet.index(2, 0), "Banana")

        sheet.setData(sheet.index(0, 1), "Price")
        sheet.setData(sheet.index(1, 1), "1.50")
        sheet.setData(sheet.index(2, 1), "0.75")

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xlsx", delete=False) as f:
            temp_path = f.name

        try:
            success = fm.save_file(temp_path, wb)
            assert success is True

            wb2 = Workbook()
            success = fm.open_file(temp_path, wb2)
            assert success is True

            sheet2 = wb2.get_sheet(0)
            assert sheet2.data(sheet2.index(0, 0)) == "Product"
            assert sheet2.data(sheet2.index(1, 0)) == "Apple"
            assert sheet2.data(sheet2.index(2, 0)) == "Banana"

        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    def test_save_xlsx_with_multiple_sheets(self):
        fm = FileManager()
        wb = Workbook()

        wb.add_sheet("Sheet1")
        sheet1 = wb.get_sheet(0)
        sheet1.setData(sheet1.index(0, 0), "Sheet1 Data")

        wb.add_sheet("Sheet2")
        sheet2 = wb.get_sheet(1)
        sheet2.setData(sheet2.index(0, 0), "Sheet2 Data")

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xlsx", delete=False) as f:
            temp_path = f.name

        try:
            success = fm.save_file(temp_path, wb)
            assert success is True

            wb2 = Workbook()
            success = fm.open_file(temp_path, wb2)
            assert success is True

            assert wb2.sheet_count() >= 2

        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    def test_unsupported_format(self):
        fm = FileManager()
        wb = Workbook()

        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            temp_path = f.name

        try:
            success = fm.save_file(temp_path, wb)
            assert success is False

        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    def test_get_current_file_path(self):
        fm = FileManager()
        assert fm.get_current_file_path() is None
