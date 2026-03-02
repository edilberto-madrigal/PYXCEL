from typing import Any, Callable, Dict, List, Optional
import inspect
import sys


class Macro:
    def __init__(self, name: str, func: Callable, description: str = ""):
        self.name = name
        self.func = func
        self.description = description
        self._is_running = False

    def run(self, *args, **kwargs) -> Any:
        self._is_running = True
        try:
            result = self.func(*args, **kwargs)
            return result
        finally:
            self._is_running = False

    def is_running(self) -> bool:
        return self._is_running

    def get_signature(self) -> inspect.Signature:
        return inspect.signature(self.func)

    def get_docstring(self) -> str:
        return self.func.__doc__ or ""


class MacroManager:
    def __init__(self):
        self._macros: Dict[str, Macro] = {}
        self._macro_groups: Dict[str, List[str]] = {}
        self._workbook = None
        self._current_sheet = None

    def set_workbook(self, workbook):
        self._workbook = workbook

    def set_current_sheet(self, sheet):
        self._current_sheet = sheet

    def register(self, name: str = None, description: str = "", group: str = "General"):
        def decorator(func: Callable) -> Callable:
            macro_name = name or func.__name__

            doc = description
            if not doc and func.__doc__:
                doc = func.__doc__.strip().split("\n")[0]

            macro = Macro(macro_name, func, doc)
            self._macros[macro_name] = macro

            if group not in self._macro_groups:
                self._macro_groups[group] = []
            self._macro_groups[group].append(macro_name)

            return func

        return decorator

    def unregister(self, name: str):
        if name in self._macros:
            macro = self._macros.pop(name)
            for group, macros in self._macro_groups.items():
                if name in macros:
                    macros.remove(name)

    def get_macro(self, name: str) -> Optional[Macro]:
        return self._macros.get(name)

    def run_macro(self, name: str, *args, **kwargs) -> Any:
        macro = self.get_macro(name)
        if macro:
            return macro.run(*args, **kwargs)
        raise ValueError(f"Macro '{name}' no encontrada")

    def list_macros(self, group: str = None) -> List[str]:
        if group:
            return self._macro_groups.get(group, []).copy()
        return list(self._macros.keys())

    def list_groups(self) -> List[str]:
        return list(self._macro_groups.keys())

    def get_macro_info(self, name: str) -> Dict:
        macro = self.get_macro(name)
        if macro:
            return {
                "name": macro.name,
                "description": macro.description,
                "signature": str(macro.get_signature()),
                "docstring": macro.get_docstring(),
            }
        return {}

    def clear(self):
        self._macros.clear()
        self._macro_groups.clear()


macro_manager = MacroManager()


def macro(name: str = None, description: str = "", group: str = "General"):
    return macro_manager.register(name, description, group)


class MacroRunner:
    def __init__(self, macro_manager: MacroManager):
        self._manager = macro_manager

    def run(self, macro_name: str, *args, **kwargs):
        return self._manager.run_macro(macro_name, *args, **kwargs)

    def run_all(self, group: str = None):
        macro_names = self._manager.list_macros(group)
        results = {}
        for name in macro_names:
            try:
                results[name] = self._manager.run_macro(name)
            except Exception as e:
                results[name] = f"Error: {str(e)}"
        return results


_default_macros = """
# Macros predeterminadas para PYXCEL
# Estas macros están disponibles automáticamente

@macro(name="sum_column", description="Suma todos los valores de una columna", group="Matemáticas")
def sum_column(sheet, col_letter: str):
    '''Suma todos los valores numéricos en una columna'''
    col_idx = ord(col_letter.upper()) - ord('A')
    total = 0
    for row in range(sheet.rowCount()):
        value = sheet.get_cell_value(row, col_idx)
        if isinstance(value, (int, float)):
            total += value
    return total

@macro(name="average_column", description="Calcula el promedio de una columna", group="Estadística")
def average_column(sheet, col_letter: str):
    '''Calcula el promedio de valores numéricos en una columna'''
    col_idx = ord(col_letter.upper()) - ord('A')
    values = []
    for row in range(sheet.rowCount()):
        value = sheet.get_cell_value(row, col_idx)
        if isinstance(value, (int, float)):
            values.append(value)
    return sum(values) / len(values) if values else 0

@macro(name="clear_sheet", description="Limpia todos los datos de la hoja", group="Manipulación")
def clear_sheet(sheet):
    '''Elimina todos los datos de la hoja actual'''
    sheet.clear()
    return "Hoja limpiada"

@macro(name="count_filled", description="Cuenta celdas con datos", group="Estadística")
def count_filled(sheet):
    '''Cuenta el número de celdas que contienen datos'''
    count = 0
    for (row, col), cell in sheet.get_all_data().items():
        if cell.value is not None and cell.value != "":
            count += 1
    return count

@macro(name="find_max", description="Encuentra el valor máximo en una columna", group="Estadística")
def find_max(sheet, col_letter: str):
    '''Encuentra el valor máximo en una columna'''
    col_idx = ord(col_letter.upper()) - ord('A')
    max_val = None
    for row in range(sheet.rowCount()):
        value = sheet.get_cell_value(row, col_idx)
        if isinstance(value, (int, float)):
            if max_val is None or value > max_val:
                max_val = value
    return max_val

@macro(name="find_min", description="Encuentra el valor mínimo en una columna", group="Estadística")
def find_min(sheet, col_letter: str):
    '''Encuentra el valor mínimo en una columna'''
    col_idx = ord(col_letter.upper()) - ord('A')
    min_val = None
    for row in range(sheet.rowCount()):
        value = sheet.get_cell_value(row, col_idx)
        if isinstance(value, (int, float)):
            if min_val is None or value < min_val:
                min_val = value
    return min_val

@macro(name="insert_date", description="Inserta la fecha actual en una celda", group="Manipulación")
def insert_date(sheet, row: int, col: int):
    '''Inserta la fecha actual en la celda especificada'''
    from datetime import datetime
    sheet.setData(sheet.index(row, col), datetime.now().strftime("%Y-%m-%d"))
    return "Fecha insertada"
"""

exec(_default_macros)
