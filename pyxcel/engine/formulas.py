import re
import math
from typing import Any, Optional, Dict, List
import operator


class FormulaEngine:
    def __init__(self):
        self._spreadsheet = None
        self._current_row = 0
        self._current_col = 0
        self._operators = {
            "+": operator.add,
            "-": operator.sub,
            "*": operator.mul,
            "/": operator.truediv,
            "^": operator.pow,
            "=": operator.eq,
            "<>": lambda a, b: a != b,
            "<": operator.lt,
            ">": operator.gt,
            "<=": operator.le,
            ">=": operator.ge,
            "&": lambda a, b: str(a) + str(b),
        }

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

    def _cell_ref_to_indices(self, cell_ref: str) -> tuple:
        match = re.match(r"^([A-Z]+)(\d+)$", cell_ref.upper())
        if match:
            col = self._letter_to_col(match.group(1))
            row = int(match.group(2)) - 1
            return (row, col)
        return (0, 0)

    def _get_cell_address(self, row: int, col: int) -> str:
        return f"{self._col_to_letter(col)}{row + 1}"

    def _range_to_cells(self, range_ref: str) -> List[tuple]:
        cells = []
        if ":" in range_ref:
            start, end = range_ref.split(":")
            start_row, start_col = self._cell_ref_to_indices(start)
            end_row, end_col = self._cell_ref_to_indices(end)

            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    cells.append((row, col))
        else:
            cells.append(self._cell_ref_to_indices(range_ref))
        return cells

    def _get_cell_value(self, row: int, col: int) -> Any:
        if self._spreadsheet:
            try:
                return self._spreadsheet.get_cell_value(row, col)
            except:
                return 0
        return 0

    def _get_range_values(self, range_ref: str) -> List[List[Any]]:
        cells = self._range_to_cells(range_ref)
        if not cells:
            return []

        min_row = min(c[0] for c in cells)
        max_row = max(c[0] for c in cells)
        min_col = min(c[1] for c in cells)
        max_col = max(c[1] for c in cells)

        result = []
        for row in range(min_row, max_row + 1):
            row_data = []
            for col in range(min_col, max_col + 1):
                row_data.append(self._get_cell_value(row, col))
            result.append(row_data)

        return result

    def _to_number(self, value: Any) -> float:
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            try:
                return float(value.replace(",", "."))
            except:
                return 0
        return 0

    def _is_number(self, value: Any) -> bool:
        try:
            self._to_number(value)
            return True
        except:
            return False

    def evaluate(self, formula: str, spreadsheet, row: int = 0, col: int = 0) -> Any:
        self._spreadsheet = spreadsheet
        self._current_row = row
        self._current_col = col

        formula = formula.strip()
        if not formula.startswith("="):
            return formula

        expression = formula[1:]

        try:
            return self._evaluate_expression(expression)
        except Exception as e:
            return f"#ERROR: {str(e)}"

    def _evaluate_expression(self, expr: str) -> Any:
        expr = expr.strip()

        for op in ["<=", ">=", "<>", "="]:
            if op in expr:
                parts = expr.split(op)
                if len(parts) == 2:
                    left = self._evaluate_expression(parts[0])
                    right = self._evaluate_expression(parts[1])
                    return self._operators[op](left, right)

        for op in ["+", "-"]:
            if expr.count(op) == 1 and expr[0] not in [op]:
                parts = expr.split(op)
                if len(parts) == 2:
                    left = self._evaluate_expression(parts[0])
                    right = self._evaluate_expression(parts[1])
                    return self._operators[op](left, right)

        for op in ["*", "/"]:
            if op in expr:
                parts = expr.split(op)
                if len(parts) == 2:
                    left = self._evaluate_expression(parts[0])
                    right = self._evaluate_expression(parts[1])
                    if op == "/" and right == 0:
                        return "#DIV/0!"
                    return self._operators[op](left, right)

        for op in ["^", "&"]:
            if op in expr:
                parts = expr.split(op)
                if len(parts) == 2:
                    left = self._evaluate_expression(parts[0])
                    right = self._evaluate_expression(parts[1])
                    return self._operators[op](left, right)

        if "(" in expr:
            return self._evaluate_function(expr)

        return self._get_value(expr)

    def _get_value(self, token: str) -> Any:
        token = token.strip()

        if token.upper() == "TRUE":
            return True
        if token.upper() == "FALSE":
            return False
        if token.upper() == "PI":
            return math.pi
        if token.upper() == "E":
            return math.e

        try:
            if "." in token:
                return float(token.replace(",", "."))
            return int(token)
        except:
            pass

        if re.match(r"^[A-Z]+\d+$", token.upper()):
            row, col = self._cell_ref_to_indices(token)
            return self._get_cell_value(row, col)

        if ":" in token:
            return self._get_range_values(token)

        return token

    def _evaluate_function(self, expr: str) -> Any:
        expr = expr.strip()

        if not expr.endswith(")"):
            return expr

        paren_idx = expr.rfind("(")
        if paren_idx == -1:
            return self._get_value(expr)

        func_name = expr[:paren_idx].strip().upper()
        args_str = expr[paren_idx + 1 : -1]

        args = self._parse_arguments(args_str)

        return self._execute_function(func_name, args)

    def _parse_arguments(self, args_str: str) -> List[Any]:
        args = []
        current = ""
        depth = 0

        for char in args_str:
            if char == "(":
                depth += 1
                current += char
            elif char == ")":
                depth -= 1
                current += char
            elif char == "," and depth == 0:
                args.append(current.strip())
                current = ""
            else:
                current += char

        if current.strip():
            args.append(current.strip())

        return [self._evaluate_expression(a) for a in args]

    def _execute_function(self, func_name: str, args: List[Any]) -> Any:
        env = {
            "SUM": lambda a: sum(self._to_number(x) for x in self._flatten(a)),
            "AVERAGE": lambda a: sum(self._to_number(x) for x in self._flatten(a))
            / len(self._flatten(a))
            if self._flatten(a)
            else 0,
            "MAX": lambda a: max(self._to_number(x) for x in self._flatten(a))
            if self._flatten(a)
            else 0,
            "MIN": lambda a: min(self._to_number(x) for x in self._flatten(a))
            if self._flatten(a)
            else 0,
            "COUNT": lambda a: len([x for x in self._flatten(a) if self._is_number(x)]),
            "COUNTA": lambda a: len(
                [x for x in self._flatten(a) if x is not None and x != ""]
            ),
            "COUNTBLANK": lambda a: len(
                [x for x in self._flatten(a) if x is None or x == ""]
            ),
            "IF": lambda cond, true_val, false_val=None: true_val
            if cond
            else (false_val if false_val is not None else true_val),
            "AND": lambda *a: all(a),
            "OR": lambda *a: any(a),
            "NOT": lambda a: not a,
            "ABS": lambda a: abs(self._to_number(a)),
            "ROUND": lambda a, d=0: round(
                self._to_number(a), int(d) if self._is_number(d) else 0
            ),
            "ROUNDUP": lambda a, d=0: math.ceil(self._to_number(a) * (10 ** int(d)))
            / (10 ** int(d))
            if self._is_number(d)
            else math.ceil(self._to_number(a)),
            "ROUNDDOWN": lambda a, d=0: math.floor(self._to_number(a) * (10 ** int(d)))
            / (10 ** int(d))
            if self._is_number(d)
            else math.floor(self._to_number(a)),
            "INT": lambda a: int(self._to_number(a)),
            "MOD": lambda a, b: self._to_number(a) % self._to_number(b),
            "POWER": lambda a, b: self._to_number(a) ** self._to_number(b),
            "SQRT": lambda a: math.sqrt(self._to_number(a)),
            "LEN": lambda a: len(str(a)),
            "LEFT": lambda a, n=1: str(a)[: int(n) if self._is_number(n) else 1],
            "RIGHT": lambda a, n=1: str(a)[-int(n) :]
            if self._is_number(n)
            else str(a)[-1:],
            "MID": lambda a, start, length: str(a)[
                int(start) - 1 : int(start) - 1 + int(length)
            ],
            "UPPER": lambda a: str(a).upper(),
            "LOWER": lambda a: str(a).lower(),
            "TRIM": lambda a: str(a).strip(),
            "CONCATENATE": lambda *a: "".join(str(x) for x in a),
            "TEXT": lambda a, fmt: str(a),
            "VALUE": lambda a: self._to_number(a),
            "DATE": lambda y, m, d: f"{int(y)}-{int(m):02d}-{int(d):02d}",
            "TODAY": lambda: self._get_today(),
            "NOW": lambda: self._get_now(),
            "YEAR": lambda d: self._get_year(d),
            "MONTH": lambda d: self._get_month(d),
            "DAY": lambda d: self._get_day(d),
            "VLOOKUP": lambda *a: self._vlookup(*a),
            "HLOOKUP": lambda *a: self._hlookup(*a),
            "INDEX": lambda *a: self._index(*a),
            "MATCH": lambda *a: self._match(*a),
            "XLOOKUP": lambda *a: self._xlookup(*a),
            "FILTER": lambda *a: self._filter(*a),
            "UNIQUE": lambda *a: self._unique(*a),
            "SORT": lambda *a: self._sort(*a),
            "SUMIF": lambda *a: self._sumif(*a),
            "COUNTIF": lambda *a: self._countif(*a),
            "AVERAGEIF": lambda *a: self._averageif(*a),
            "IFERROR": lambda val, err_val: val
            if not str(val).startswith("#")
            else err_val,
            "ISERROR": lambda a: str(a).startswith("#"),
            "ISBLANK": lambda a: a is None or a == "",
            "ISNUMBER": lambda a: self._is_number(a),
            "ISTEXT": lambda a: isinstance(a, str) and not self._is_number(a),
            "TRUE": True,
            "FALSE": False,
        }

        if func_name in env:
            return env[func_name](*args)

        return f"#{func_name}?"

    def _flatten(self, lst: Any) -> List[Any]:
        if isinstance(lst, list):
            result = []
            for item in lst:
                result.extend(self._flatten(item))
            return result
        return [lst]

    def _get_today(self) -> str:
        from datetime import datetime

        return datetime.now().strftime("%Y-%m-%d")

    def _get_now(self) -> str:
        from datetime import datetime

        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _get_year(self, date_str: str) -> int:
        try:
            return int(date_str.split("-")[0])
        except:
            return 0

    def _get_month(self, date_str: str) -> int:
        try:
            return int(date_str.split("-")[1])
        except:
            return 0

    def _get_day(self, date_str: str) -> int:
        try:
            return int(date_str.split("-")[2])
        except:
            return 0

    def _vlookup(self, lookup_value, table_array, col_index, range_lookup=True) -> Any:
        if not table_array or not isinstance(table_array, list):
            return "#N/A"

        for row in table_array:
            if not isinstance(row, list) or len(row) < col_index:
                continue
            if row[0] == lookup_value:
                return row[col_index - 1]

        return "#N/A"

    def _hlookup(self, lookup_value, table_array, row_index, range_lookup=True) -> Any:
        if not table_array or not isinstance(table_array, list):
            return "#N/A"

        for col in range(len(table_array[0])):
            if table_array[0][col] == lookup_value:
                if row_index <= len(table_array):
                    return table_array[row_index - 1][col]

        return "#N/A"

    def _index(self, array, row_num, col_num=None) -> Any:
        if not array:
            return "#REF!"

        flat = self._flatten(array)

        if col_num is None:
            row_num = int(row_num) if self._is_number(row_num) else 1
            idx = row_num - 1
            if 0 <= idx < len(flat):
                return flat[idx]
            return "#REF!"

        row_num = int(row_num) if self._is_number(row_num) else 1
        col_num = int(col_num) if self._is_number(col_num) else 1

        if isinstance(array, list) and len(array) >= row_num:
            if (
                isinstance(array[row_num - 1], list)
                and len(array[row_num - 1]) >= col_num
            ):
                return array[row_num - 1][col_num - 1]

        return "#REF!"

    def _match(self, lookup_value, lookup_array, match_type=1) -> Any:
        if not lookup_array:
            return "#N/A"

        flat = self._flatten(lookup_array)

        if match_type == 0:
            for i, val in enumerate(flat):
                if val == lookup_value:
                    return i + 1
        elif match_type == 1:
            for i, val in enumerate(flat):
                if val == lookup_value:
                    return i + 1
                elif val > lookup_value:
                    return i if i > 0 else 1
        elif match_type == -1:
            for i, val in enumerate(flat):
                if val == lookup_value:
                    return i + 1
                elif val < lookup_value:
                    return i if i > 0 else 1

        return "#N/A"

    def _xlookup(
        self,
        lookup_value,
        lookup_array,
        return_array,
        if_not_found=None,
        match_mode=0,
        search_mode=1,
    ) -> Any:
        if not lookup_array or not return_array:
            return if_not_found if if_not_found else "#N/A"

        flat_lookup = self._flatten(lookup_array)
        flat_return = self._flatten(return_array)

        for i, val in enumerate(flat_lookup):
            if val == lookup_value:
                if i < len(flat_return):
                    return flat_return[i]

        return if_not_found if if_not_found else "#N/A"

    def _filter(self, array, include, if_empty=None):
        if not array:
            return if_empty if if_empty else []

        if callable(include):
            result = [row for row in array if include(row)]
        else:
            result = [row for row in array if any(include == val for val in row)]

        if not result:
            return if_empty if if_empty else []

        return result

    def _unique(self, array, by_col=False, exactly_once=False):
        if not array:
            return []

        flat = self._flatten(array)
        seen = {}

        for val in flat:
            seen[val] = seen.get(val, 0) + 1

        if exactly_once:
            return [k for k, v in seen.items() if v == 1]

        return list(seen.keys())

    def _sort(self, array, sort_index=0, ascending=True):
        if not array:
            return []

        if not isinstance(array[0], list):
            array = [[x] for x in array]
            sort_index = 0

        return sorted(
            array,
            key=lambda x: x[sort_index] if len(x) > sort_index else "",
            reverse=not ascending,
        )

    def _sumif(self, range_vals, criteria, sum_range=None):
        if sum_range is None:
            sum_range = range_vals

        range_flat = self._flatten(range_vals)
        sum_flat = self._flatten(sum_range)

        total = 0
        for i, val in enumerate(range_flat):
            if self._match_criteria(val, criteria) and i < len(sum_flat):
                total += self._to_number(sum_flat[i])

        return total

    def _countif(self, range_vals, criteria):
        range_flat = self._flatten(range_vals)

        count = 0
        for val in range_flat:
            if self._match_criteria(val, criteria):
                count += 1

        return count

    def _averageif(self, range_vals, criteria, average_range=None):
        if average_range is None:
            average_range = range_vals

        range_flat = self._flatten(range_vals)
        avg_flat = self._flatten(average_range)

        total = 0
        count = 0

        for i, val in enumerate(range_flat):
            if self._match_criteria(val, criteria) and i < len(avg_flat):
                total += self._to_number(avg_flat[i])
                count += 1

        return total / count if count > 0 else 0

    def _match_criteria(self, value, criteria) -> bool:
        if isinstance(criteria, str):
            if criteria.startswith("*") and criteria.endswith("*"):
                return criteria[1:-1] in str(value)
            elif criteria.startswith("*"):
                return str(value).endswith(criteria[1:])
            elif criteria.endswith("*"):
                return str(value).startswith(criteria[:-1])
            elif "<" in criteria:
                op, num = criteria.split("<")
                return self._to_number(value) < self._to_number(num)
            elif ">" in criteria:
                op, num = criteria.split(">")
                return self._to_number(value) > self._to_number(num)
            elif "=" in criteria:
                return str(value) == criteria[1:]
            else:
                return str(value) == criteria
        return value == criteria
