from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTabWidget,
    QMessageBox,
    QInputDialog,
    QFileDialog,
    QLabel,
    QSplitter,
    QDockWidget,
    QTextEdit,
)
from PySide6.QtCore import Qt, Signal, QTimer
from PySide6.QtGui import QKeySequence, QAction, QClipboard

import sys
import os

from .models.workbook import Workbook
from .models.spreadsheet import SpreadsheetModel
from .ui.table import SpreadsheetWidget
from .ui.toolbar import ToolbarManager, MenuManager
from .ui.dialogs import (
    FindReplaceDialog,
    FormatCellsDialog,
    InsertChartDialog,
    SortDialog,
    FilterDialog,
)
from .utils.file_handler import FileManager
from .utils.chart_builder import ChartBuilder
from .engine.formulas import FormulaEngine
from .macros.macro_system import macro_manager, MacroRunner


class PYXCEL(QMainWindow):
    def __init__(self):
        super().__init__()

        self.workbook = Workbook()
        self.formula_engine = FormulaEngine()
        self.file_manager = FileManager()
        self.chart_builder = ChartBuilder()
        self.macro_runner = MacroRunner(macro_manager)

        self._setup_ui()
        self._connect_signals()
        self._create_default_sheet()

        self.setWindowTitle("PYXCEL - Hoja de Cálculo")
        self.resize(1200, 800)

    def _setup_ui(self):
        self.toolbar_manager = ToolbarManager(self)
        self.menu_manager = MenuManager(self)

        self.setMenuBar(self.menu_manager.create_menus())
        self.addToolBar(self.toolbar_manager.create_toolbar())

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        self.sheet_tabs = QTabWidget()
        self.sheet_tabs.setTabPosition(QTabWidget.TabPosition.South)
        self.sheet_tabs.setTabsClosable(True)
        self.sheet_tabs.tabCloseRequested.connect(self._close_tab)

        main_layout.addWidget(self.sheet_tabs)

        self.statusBar().showMessage("Listo")

        self._create_sheet_tab("Hoja1")

        self.sheet_tabs.currentChanged.connect(self._on_sheet_changed)

    def _create_sheet_tab(self, name: str = None):
        if name is None:
            name = f"Hoja{self.workbook.sheet_count() + 1}"

        sheet = self.workbook.add_sheet(name, 1000, 100)

        spreadsheet_widget = SpreadsheetWidget()
        spreadsheet_widget.set_model(sheet)

        sheet.set_formula_engine(self.formula_engine)

        index = self.sheet_tabs.addTab(spreadsheet_widget, name)
        self.sheet_tabs.setCurrentIndex(index)

        return spreadsheet_widget

    def _create_default_sheet(self):
        if self.workbook.sheet_count() == 0:
            self._create_sheet_tab()

    def _connect_signals(self):
        self.file_manager.file_loaded.connect(self._on_file_loaded)
        self.file_manager.file_saved.connect(self._on_file_saved)
        self.file_manager.error_occurred.connect(self._on_file_error)

    def _on_file_loaded(self, file_path: str):
        self.statusBar().showMessage(f"Archivo abierto: {file_path}")

        self.sheet_tabs.clear()

        for i in range(self.workbook.sheet_count()):
            sheet = self.workbook.get_sheet(i)
            sheet_name = self.workbook.get_sheet_name(i)

            spreadsheet_widget = SpreadsheetWidget()
            spreadsheet_widget.set_model(sheet)
            sheet.set_formula_engine(self.formula_engine)

            self.sheet_tabs.addTab(spreadsheet_widget, sheet_name)

        self.setWindowTitle(f"PYXCEL - {os.path.basename(file_path)}")
        self.statusBar().showMessage("Archivo cargado correctamente")

    def _on_file_saved(self, file_path: str):
        self.statusBar().showMessage(f"Archivo guardado: {file_path}")
        self.setWindowTitle(f"PYXCEL - {os.path.basename(file_path)}")

    def _on_file_error(self, error: str):
        QMessageBox.critical(self, "Error", error)

    def _on_sheet_changed(self, index: int):
        if index >= 0:
            self.workbook.set_active_sheet(index)

    def _close_tab(self, index: int):
        if self.sheet_tabs.count() <= 1:
            QMessageBox.warning(
                self, "Advertencia", "No se puede cerrar la última hoja"
            )
            return

        reply = QMessageBox.question(
            self,
            "Cerrar Hoja",
            "¿Está seguro de cerrar esta hoja?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.sheet_tabs.removeTab(index)
            self.workbook.remove_sheet(index)

    def get_current_sheet_widget(self) -> SpreadsheetWidget:
        return self.sheet_tabs.currentWidget()

    def get_current_sheet(self) -> SpreadsheetModel:
        widget = self.get_current_sheet_widget()
        if widget:
            return widget.get_view().model()
        return None

    def new_file(self):
        if self.workbook.is_modified():
            reply = QMessageBox.question(
                self,
                "Guardar cambios",
                "¿Desea guardar los cambios antes de crear un nuevo archivo?",
                QMessageBox.StandardButton.Save
                | QMessageBox.StandardButton.Discard
                | QMessageBox.StandardButton.Cancel,
            )

            if reply == QMessageBox.StandardButton.Save:
                self.save_file()
            elif reply == QMessageBox.StandardButton.Cancel:
                return

        self.workbook.clear()
        self.sheet_tabs.clear()
        self._create_default_sheet()
        self.setWindowTitle("PYXCEL - Sin título")
        self.statusBar().showMessage("Nuevo archivo creado")

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Abrir archivo",
            "",
            "Archivos Excel (*.xlsx *.xls);;Archivos CSV (*.csv);;Todos los archivos (*)",
        )

        if file_path:
            self.file_manager.open_file(file_path, self.workbook)

    def save_file(self):
        if self.workbook.file_path():
            self.file_manager.save_file(self.workbook.file_path(), self.workbook)
        else:
            self.save_file_as()

    def save_file_as(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Guardar archivo", "", "Archivos Excel (*.xlsx);;Archivos CSV (*.csv)"
        )

        if file_path:
            self.file_manager.save_file(file_path, self.workbook)

    def export_to_pdf(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar a PDF", "", "PDF Files (*.pdf)"
        )

        if file_path:
            QMessageBox.information(
                self, "Exportar PDF", "Función de exportación a PDF en desarrollo"
            )

    def undo(self):
        sheet = self.get_current_sheet()
        if sheet:
            sheet.undo()

    def redo(self):
        sheet = self.get_current_sheet()
        if sheet:
            sheet.redo()

    def cut(self):
        widget = self.get_current_sheet_widget()
        if widget:
            clipboard = QApplication.clipboard()
            view = widget.get_view()
            row, col = view.get_current_position()
            sheet = self.get_current_sheet()

            if sheet:
                value = sheet.get_cell_value(row, col)
                if value:
                    clipboard.setText(str(value))

    def copy(self):
        widget = self.get_current_sheet_widget()
        if widget:
            clipboard = QApplication.clipboard()
            view = widget.get_view()
            row, col = view.get_current_position()
            sheet = self.get_current_sheet()

            if sheet:
                value = sheet.get_cell_value(row, col)
                if value:
                    clipboard.setText(str(value))

    def paste(self):
        widget = self.get_current_sheet_widget()
        if widget:
            clipboard = QApplication.clipboard()
            text = clipboard.text()

            if text:
                view = widget.get_view()
                row, col = view.get_current_position()
                sheet = self.get_current_sheet()

                if sheet:
                    index = sheet.index(row, col)
                    sheet.setData(index, text, Qt.ItemDataRole.EditRole)

    def show_find_dialog(self):
        dialog = FindReplaceDialog(self)
        dialog.find_next.connect(self._find_next)
        dialog.replace.connect(self._replace)
        dialog.replace_all.connect(self._replace_all)
        dialog.show()

    def show_replace_dialog(self):
        dialog = FindReplaceDialog(self)
        dialog.show()

    def _find_next(self, text: str, match_case: bool):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        view = self.get_current_sheet_widget().get_view()
        start_row, start_col = view.get_current_position()

        for row in range(start_row, sheet.rowCount()):
            for col in range(start_col if row == start_row else 0, sheet.columnCount()):
                value = sheet.get_cell_value(row, col)
                if value:
                    value_str = str(value)
                    if match_case:
                        if text in value_str:
                            view.set_current_position(row, col)
                            return
                    else:
                        if text.lower() in value_str.lower():
                            view.set_current_position(row, col)
                            return

        QMessageBox.information(self, "Buscar", "No se encontró más coincidencias")

    def _replace(self, find: str, replace: str, match_case: bool):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        view = self.get_current_sheet_widget().get_view()
        row, col = view.get_current_position()

        index = sheet.index(row, col)
        value = sheet.get_cell_value(row, col)

        if value:
            value_str = str(value)
            new_value = (
                value_str.replace(find, replace)
                if match_case
                else value_str.replace(find, replace)
            )
            sheet.setData(index, new_value, Qt.ItemDataRole.EditRole)

    def _replace_all(self, find: str, replace: str, match_case: bool):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        changes = sheet.find_and_replace(find, replace, match_case)

        QMessageBox.information(
            self, "Reemplazar", f"Se reemplazaron {len(changes)} ocurrencias"
        )

    def insert_row(self):
        sheet = self.get_current_sheet()
        if sheet:
            widget = self.get_current_sheet_widget()
            view = widget.get_view()
            row, col = view.get_current_position()
            sheet.insert_rows(row, 1)

    def insert_column(self):
        sheet = self.get_current_sheet()
        if sheet:
            widget = self.get_current_sheet_widget()
            view = widget.get_view()
            row, col = view.get_current_position()
            sheet.insert_columns(col, 1)

    def insert_chart(self):
        dialog = InsertChartDialog(self)
        if dialog.exec():
            QMessageBox.information(
                self, "Gráfico", "Función de gráficos en desarrollo"
            )

    def format_cells(self):
        dialog = FormatCellsDialog(self)
        if dialog.exec():
            pass

    def toggle_autofilter(self):
        QMessageBox.information(
            self, "Autofiltro", "Función de autofiltro en desarrollo"
        )

    def sort_data(self):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        columns = [
            f"Columna {chr(65 + i)}" for i in range(min(sheet.columnCount(), 26))
        ]

        dialog = SortDialog(self)
        dialog.set_columns(columns)

        if dialog.exec():
            col_index = dialog.get_column_index()
            ascending = dialog.get_ascending()
            sheet.sort_by_column(col_index, ascending)

    def filter_data(self):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        columns = [
            f"Columna {chr(65 + i)}" for i in range(min(sheet.columnCount(), 26))
        ]

        dialog = FilterDialog(self)
        dialog.set_columns(columns)

        if dialog.exec():
            QMessageBox.information(self, "Filtrar", "Función de filtro en desarrollo")

    def show_about(self):
        QMessageBox.about(
            self,
            "Acerca de PYXCEL",
            "PYXCEL v0.1.0\n\n"
            "Una aplicación de hoja de cálculo\n"
            "similar a Excel, escrita en Python.\n\n"
            "Librerías utilizadas:\n"
            "- PySide6 (Qt)\n"
            "- pandas, numpy\n"
            "- matplotlib, seaborn\n"
            "- openpyxl\n"
            "- formulas",
        )

    def format_font(self, font_family: str):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        widget = self.get_current_sheet_widget()
        view = widget.get_view()
        row, col = view.get_current_position()

        fmt = sheet.get_cell_format(row, col)
        fmt.font_family = font_family
        sheet.set_cell_format(row, col, fmt)

    def format_font_size(self, size: int):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        widget = self.get_current_sheet_widget()
        view = widget.get_view()
        row, col = view.get_current_position()

        fmt = sheet.get_cell_format(row, col)
        fmt.font_size = size
        sheet.set_cell_format(row, col, fmt)

    def format_bold(self):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        widget = self.get_current_sheet_widget()
        view = widget.get_view()
        row, col = view.get_current_position()

        fmt = sheet.get_cell_format(row, col)
        fmt.bold = not fmt.bold
        sheet.set_cell_format(row, col, fmt)

    def format_italic(self):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        widget = self.get_current_sheet_widget()
        view = widget.get_view()
        row, col = view.get_current_position()

        fmt = sheet.get_cell_format(row, col)
        fmt.italic = not fmt.italic
        sheet.set_cell_format(row, col, fmt)

    def format_underline(self):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        widget = self.get_current_sheet_widget()
        view = widget.get_view()
        row, col = view.get_current_position()

        fmt = sheet.get_cell_format(row, col)
        fmt.underline = not fmt.underline
        sheet.set_cell_format(row, col, fmt)

    def format_background_color(self, color):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        widget = self.get_current_sheet_widget()
        view = widget.get_view()
        row, col = view.get_current_position()

        fmt = sheet.get_cell_format(row, col)
        fmt.bg_color = color
        sheet.set_cell_format(row, col, fmt)

    def format_alignment(self, alignment):
        sheet = self.get_current_sheet()
        if not sheet:
            return

        widget = self.get_current_sheet_widget()
        view = widget.get_view()
        row, col = view.get_current_position()

        fmt = sheet.get_cell_format(row, col)
        fmt.alignment = alignment
        sheet.set_cell_format(row, col, fmt)

    def closeEvent(self, event):
        if self.workbook.is_modified():
            reply = QMessageBox.question(
                self,
                "Guardar cambios",
                "¿Desea guardar los cambios antes de salir?",
                QMessageBox.StandardButton.Save
                | QMessageBox.StandardButton.Discard
                | QMessageBox.StandardButton.Cancel,
            )

            if reply == QMessageBox.StandardButton.Save:
                self.save_file()
                event.accept()
            elif reply == QMessageBox.StandardButton.Discard:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = PYXCEL()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
