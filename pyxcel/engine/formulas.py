import re
from typing import Any, Optional, Dict
from formulas import Parser, ExcelModel
from formulas import functions as excel_functions
import math


class FormulaEngine:
    def __init__(self):
        self._parser = None
        self._model = None
        self._spreadsheet = None
        self._current_row = 0
        self._current_col = 0

    def set_spreadsheet(self, spreadsheet):
        self._spreadsheet = spreadsheet

    def _col_to_letter(self, col: int) -> str:
        result = ""
        while col >= 0:
            result = chr(65 + (col % 26)) + result
            col = col // 26 - 1
        return result

    def _letter_to_col(self, letter: str) -> int:
        col = 0
        for char in letter.upper():
            col = col * 26 + (ord(char) - ord("A") + 1)
        return col - 1

    def _cell_ref_to_indices(self, ref: str) -> tuple:
        match = re.match(r"([A-Z]+)(\d+)", ref.upper())
        if match:
            col = self._letter_to_col(match.group(1))
            row = int(match.group(2)) - 1
            return (row, col)
        return (0, 0)

    def _get_cell_address(self, row: int, col: int) -> str:
        return f"{self._col_to_letter(col)}{row + 1}"

    def evaluate(self, formula: str, spreadsheet, row: int = 0, col: int = 0) -> Any:
        self._spreadsheet = spreadsheet
        self._current_row = row
        self._current_col = col

        try:
            parser = Parser()
            ast = parser.parse(formula)

            custom_env = self._create_environment()

            result = ast.evaluate(env=custom_env)
            return result
        except Exception as e:
            return f"#{str(e)}"

    def _create_environment(self) -> Dict:
        env = {
            "TRUE": True,
            "FALSE": False,
            "PI": math.pi,
            "E": math.e,
        }

        env.update(excel_functions)

        env["THIS_CELL"] = lambda: self._get_cell_value(
            self._current_row, self._current_col
        )

        return env

    def _get_cell_value(self, row: int, col: int) -> Any:
        if self._spreadsheet:
            return self._spreadsheet.get_cell_value(row, col)
        return None

    def evaluate_range(
        self, start_row: int, start_col: int, end_row: int, end_col: int
    ) -> list:
        values = []
        for row in range(start_row, end_row + 1):
            row_values = []
            for col in range(start_col, end_col + 1):
                row_values.append(self._get_cell_value(row, col))
            values.append(row_values)
        return values

    def load_workbook(self, file_path: str):
        try:
            self._model = ExcelModel()
            self._model.loads(file_path).finish()
            return True
        except Exception as e:
            print(f"Error loading workbook: {e}")
            return False

    def evaluate_cell(self, cell_ref: str) -> Any:
        if self._model:
            try:
                return self._model[cell_ref].value
            except:
                return None
        return None

    def get_dependencies(self, cell_ref: str) -> list:
        if self._model:
            try:
                return self._model[cell_ref].inputs
            except:
                return []
        return []

    def recalculate_all(self):
        if self._model:
            self._model.calculate()
