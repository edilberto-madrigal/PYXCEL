import pytest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from unittest.mock import Mock
from pyxcel.engine.formulas import FormulaEngine


class MockSpreadsheet:
    def __init__(self):
        self.data = {}

    def get_cell_value(self, row, col):
        return self.data.get((row, col), 0)


class TestFormulaEngine:
    def test_engine_creation(self):
        engine = FormulaEngine()
        assert engine is not None

    def test_col_to_letter(self):
        engine = FormulaEngine()
        assert engine._col_to_letter(0) == "A"
        assert engine._col_to_letter(1) == "B"
        assert engine._col_to_letter(25) == "Z"
        assert engine._col_to_letter(26) == "AA"

    def test_letter_to_col(self):
        engine = FormulaEngine()
        assert engine._letter_to_col("A") == 0
        assert engine._letter_to_col("B") == 1
        assert engine._letter_to_col("Z") == 25
        assert engine._letter_to_col("AA") == 26

    def test_cell_ref_to_indices(self):
        engine = FormulaEngine()
        assert engine._cell_ref_to_indices("A1") == (0, 0)
        assert engine._cell_ref_to_indices("B2") == (1, 1)
        assert engine._cell_ref_to_indices("C10") == (9, 2)

    def test_get_cell_address(self):
        engine = FormulaEngine()
        assert engine._get_cell_address(0, 0) == "A1"
        assert engine._get_cell_address(1, 1) == "B2"
        assert engine._get_cell_address(9, 2) == "C10"

    def test_range_to_cells(self):
        engine = FormulaEngine()
        cells = engine._range_to_cells("A1:B2")
        assert len(cells) == 4
        assert (0, 0) in cells
        assert (0, 1) in cells
        assert (1, 0) in cells
        assert (1, 1) in cells

    def test_range_to_cells_single(self):
        engine = FormulaEngine()
        cells = engine._range_to_cells("A1")
        assert len(cells) == 1
        assert cells[0] == (0, 0)

    def test_execute_function_sum(self):
        engine = FormulaEngine()
        engine.set_spreadsheet(MockSpreadsheet())

        result = engine._execute_function("SUM", [[1, 2, 3, 4, 5]])
        assert result == 15

    def test_execute_function_average(self):
        engine = FormulaEngine()
        engine.set_spreadsheet(MockSpreadsheet())

        result = engine._execute_function("AVERAGE", [[10, 20, 30]])
        assert result == 20

    def test_execute_function_max_min(self):
        engine = FormulaEngine()
        engine.set_spreadsheet(MockSpreadsheet())

        result = engine._execute_function("MAX", [[1, 5, 3]])
        assert result == 5

        result = engine._execute_function("MIN", [[1, 5, 3]])
        assert result == 1

    def test_get_range_values(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        spreadsheet.data[(0, 0)] = 10
        spreadsheet.data[(1, 0)] = 20
        spreadsheet.data[(2, 0)] = 30
        engine.set_spreadsheet(spreadsheet)

        values = engine._get_range_values("A1:A3")
        assert values == [[10], [20], [30]]

    def test_match_criteria(self):
        engine = FormulaEngine()

        assert engine._match_criteria(5, 5) is True
        assert engine._match_criteria("test", "test") is True
        assert engine._match_criteria(5, 10) is False
        assert engine._match_criteria("hello", "*ell*") is True

    def test_evaluate_basic_arithmetic(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        spreadsheet.data[(0, 0)] = 10
        spreadsheet.data[(0, 1)] = 5
        engine.set_spreadsheet(spreadsheet)

        result = engine.evaluate("=A1+B1", spreadsheet, 0, 0)
        assert result == 15

    def test_evaluate_functions(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        spreadsheet.data[(0, 0)] = 10
        spreadsheet.data[(1, 0)] = 20
        spreadsheet.data[(2, 0)] = 30
        engine.set_spreadsheet(spreadsheet)

        result = engine.evaluate("=SUM(A1:A3)", spreadsheet, 0, 0)
        assert result == 60

    def test_vlookup_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        table = [["A", 100], ["B", 200], ["C", 300]]
        result = engine._vlookup("B", table, 2)
        assert result == 200

    def test_index_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        array = [[1, 2], [3, 4], [5, 6]]
        result = engine._index(array, 2, 1)
        assert result == 3

    def test_match_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        array = [["A"], ["B"], ["C"]]
        result = engine._match("B", array, 0)
        assert result == 2

    def test_xlookup_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        lookup_array = ["A", "B", "C"]
        return_array = [100, 200, 300]
        result = engine._xlookup("B", lookup_array, return_array)
        assert result == 200

    def test_filter_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        array = [[1, "A"], [2, "B"], [3, "A"], [4, "C"]]
        result = engine._filter(array, lambda x: x[1] == "A", "No results")
        assert len(result) == 2

    def test_unique_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        array = [[1], [2], [1], [3], [2]]
        result = engine._unique(array)
        assert set(result) == {1, 2, 3}

    def test_sort_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        array = [[3], [1], [2]]
        result = engine._sort(array)
        assert result == [[1], [2], [3]]

    def test_sumif_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        range_vals = [10, 20, 30, 40]
        result = engine._sumif(range_vals, ">20")
        assert result == 70

    def test_countif_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        range_vals = ["A", "B", "A", "C", "A"]
        result = engine._countif(range_vals, "A")
        assert result == 3

    def test_averageif_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        range_vals = [10, 20, 30, 40]
        result = engine._averageif(range_vals, ">20")
        assert result == 35

    def test_if_function(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        result = engine._execute_function("IF", [True, "Yes", "No"])
        assert result == "Yes"

        result = engine._execute_function("IF", [False, "Yes", "No"])
        assert result == "No"

    def test_text_functions(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        result = engine._execute_function("UPPER", ["hello"])
        assert result == "HELLO"

        result = engine._execute_function("LOWER", ["HELLO"])
        assert result == "hello"

        result = engine._execute_function("LEN", ["hello"])
        assert result == 5

        result = engine._execute_function("LEFT", ["hello", 2])
        assert result == "he"

        result = engine._execute_function("RIGHT", ["hello", 2])
        assert result == "lo"

    def test_nested_functions(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        spreadsheet.data[(0, 0)] = 10
        spreadsheet.data[(1, 0)] = 20
        spreadsheet.data[(2, 0)] = 30
        engine.set_spreadsheet(spreadsheet)

        result = engine.evaluate("=SUM(A1:A2)+10", spreadsheet, 0, 0)
        assert result == 40

    def test_divide_by_zero(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        spreadsheet.data[(0, 0)] = 10
        spreadsheet.data[(0, 1)] = 0
        engine.set_spreadsheet(spreadsheet)

        result = engine.evaluate("=A1/B1", spreadsheet, 0, 0)
        assert result == "#DIV/0!"

    def test_concatenate(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        result = engine._execute_function("CONCATENATE", ["Hello", " ", "World"])
        assert result == "Hello World"

    def test_count_functions(self):
        engine = FormulaEngine()
        spreadsheet = MockSpreadsheet()
        engine.set_spreadsheet(spreadsheet)

        result = engine._execute_function("COUNT", [[1, 2, "a", ""]])
        assert result == 4

        result = engine._execute_function("COUNTA", [[1, 2, "b", ""]])
        assert result == 3

        result = engine._execute_function("COUNTBLANK", [[1, "", 3, ""]])
        assert result == 2
