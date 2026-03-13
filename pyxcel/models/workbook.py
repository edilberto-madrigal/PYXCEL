from PySide6.QtCore import QObject, Signal
from typing import List, Optional, Dict, Any
from .spreadsheet import SpreadsheetModel


class Workbook:
    def __init__(self):
        self._sheets: List[SpreadsheetModel] = []
        self._sheet_names: List[str] = []
        self._active_sheet: int = 0
        self._file_path: Optional[str] = None
        self._is_modified: bool = False
        self._macros: Dict[str, Any] = {}

    def set_macros(self, macros: Dict[str, Any]):
        self._macros = macros
        self._is_modified = True

    def get_macros(self) -> Dict[str, Any]:
        return self._macros

    def add_sheet(
        self, name: str = None, rows: int = 100, cols: int = 26
    ) -> SpreadsheetModel:
        if name is None:
            name = f"Hoja{len(self._sheets) + 1}"

        sheet = SpreadsheetModel(rows, cols)
        self._sheets.append(sheet)
        self._sheet_names.append(name)
        self._is_modified = True
        return sheet

    def insert_sheet(
        self, index: int, name: str = None, rows: int = 100, cols: int = 26
    ) -> SpreadsheetModel:
        if name is None:
            name = f"Hoja{index + 1}"

        sheet = SpreadsheetModel(rows, cols)
        self._sheets.insert(index, sheet)
        self._sheet_names.insert(index, name)
        self._is_modified = True
        if self._active_sheet >= index:
            self._active_sheet += 1
        return sheet

    def remove_sheet(self, index: int):
        if 0 <= index < len(self._sheets):
            if len(self._sheets) == 1:
                return
            self._sheets.pop(index)
            self._sheet_names.pop(index)
            if self._active_sheet >= len(self._sheets):
                self._active_sheet = len(self._sheets) - 1
            self._is_modified = True

    def get_sheet(self, index: int) -> Optional[SpreadsheetModel]:
        if 0 <= index < len(self._sheets):
            return self._sheets[index]
        return None

    def get_sheet_by_name(self, name: str) -> Optional[SpreadsheetModel]:
        if name in self._sheet_names:
            return self._sheets[self._sheet_names.index(name)]
        return None

    def get_sheet_index(self, name: str) -> int:
        try:
            return self._sheet_names.index(name)
        except ValueError:
            return -1

    def get_sheet_name(self, index: int) -> str:
        if 0 <= index < len(self._sheet_names):
            return self._sheet_names[index]
        return ""

    def set_sheet_name(self, index: int, name: str):
        if 0 <= index < len(self._sheet_names):
            self._sheet_names[index] = name
            self._is_modified = True

    def rename_sheet(self, index: int, name: str):
        self.set_sheet_name(index, name)

    def sheet_count(self) -> int:
        return len(self._sheets)

    def active_sheet(self) -> int:
        return self._active_sheet

    def set_active_sheet(self, index: int):
        if 0 <= index < len(self._sheets):
            self._active_sheet = index

    def get_active_sheet(self) -> int:
        return self._active_sheet

    def current_sheet(self) -> Optional[SpreadsheetModel]:
        if 0 <= self._active_sheet < len(self._sheets):
            return self._sheets[self._active_sheet]
        return None

    def file_path(self) -> Optional[str]:
        return self._file_path

    def get_file_path(self) -> Optional[str]:
        return self._file_path

    def set_file_path(self, path: str):
        self._file_path = path

    def is_modified(self) -> bool:
        return self._is_modified

    def set_modified(self, modified: bool = True):
        self._is_modified = modified

    def get_all_sheet_names(self) -> List[str]:
        return self._sheet_names.copy()

    def clear(self):
        self._sheets.clear()
        self._sheet_names.clear()
        self._active_sheet = 0
        self._file_path = None
        self._is_modified = False

    def get_workbook_data(self) -> Dict[str, Any]:
        return {
            "sheets": [
                {"name": name, "data": sheet.get_all_data()}
                for name, sheet in zip(self._sheet_names, self._sheets)
            ],
            "active_sheet": self._active_sheet,
        }

    def load_workbook_data(self, data: Dict[str, Any]):
        self.clear()
        for sheet_data in data.get("sheets", []):
            name = sheet_data.get("name", "Sheet")
            sheet_data_dict = sheet_data.get("data", {})
            rows = max([r for r, c in sheet_data_dict.keys()], default=0) + 1
            cols = max([c for r, c in sheet_data_dict.keys()], default=0) + 1
            sheet = self.add_sheet(name, max(rows, 100), max(cols, 26))
            sheet.set_all_data(sheet_data_dict)

        active = data.get("active_sheet", 0)
        if 0 <= active < len(self._sheets):
            self._active_sheet = active
