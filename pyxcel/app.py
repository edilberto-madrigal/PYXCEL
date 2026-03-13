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
    QListWidget,
    QListWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QLineEdit,
    QDialog,
    QScrollArea,
    QPushButton,
    QSplashScreen,
)
from PySide6.QtCore import Qt, Signal, QTimer
from PySide6.QtGui import QKeySequence, QAction, QClipboard, QFont, QTextCursor, QColor

from .ui.theme import (
    theme_manager,
    ThemeColors,
    ThemeFonts,
    get_terminal_style,
    get_app_stylesheet,
    get_palette,
)

import sys
import os
import io
from contextlib import redirect_stdout

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


class TerminalWidget(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)

        # Modern terminal styling
        self.setFont(QFont(ThemeFonts.MONO, 11))
        self.update_style()

    def update_style(self):
        self.setStyleSheet(get_terminal_style())
        
        # Persistent globals for the shell
        self._globals = {
            "__builtins__": __builtins__,
            "print": print, # Ensure print is available
        }

        # Python namespace for variables
        import builtins

        self._namespace = {"__builtins__": builtins}

        self.append(
            f"<span style='color: {ThemeColors.OVERLAY0};'>+-------------------------------------------------------------+</span>"
        )
        self.append(
            f"<span style='color: {ThemeColors.BLUE};'>|</span> "
            f"<span style='color: {ThemeColors.MAUVE}; font-weight: bold;'>PYXCEL Python Terminal</span> "
            f"<span style='color: {ThemeColors.BLUE};'>|</span>"
        )
        self.append(
            f"<span style='color: {ThemeColors.OVERLAY0};'>|</span> <span style='color: {ThemeColors.SUBTEXT0};'>Use !pip install &lt;package&gt; to install libraries</span> "
            f"<span style='color: {ThemeColors.BLUE};'>|</span>"
        )
        self.append(
            f"<span style='color: {ThemeColors.OVERLAY0};'>+-------------------------------------------------------------+</span>"
        )
        self.append("")
        self._prompt()

    def _prompt(self):
        self.append(
            f"<span style='color: {ThemeColors.TEAL};'>&gt;&gt;&gt;</span> "
            f"<span style='color: {ThemeColors.GREEN};'></span>"
        )
        self.moveCursor(QTextCursor.MoveOperation.End)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Return:
            self._execute_command()
            return
        elif event.key() == Qt.Key.Key_Up:
            self._history_navigate(-1)
            return
        elif event.key() == Qt.Key.Key_Down:
            self._history_navigate(1)
            return

        super().keyPressEvent(event)

    def _run_python(self, command):
        import io
        import sys
        
        output_buffer = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = output_buffer
        
        try:
            # Try to compile and run
            try:
                # First try eval (for expressions like '1+1')
                code = compile(command, '<string>', 'eval')
                result = eval(code, self._globals)
            except SyntaxError:
                # If syntax error, it might be a statement (like 'x = 1' or 'import os')
                code = compile(command, '<string>', 'exec')
                result = None
                exec(code, self._globals)
            
            # Show captured print() outputs
            output = output_buffer.getvalue()
            if output:
                self.append(f"<span style='color: {ThemeColors.TEXT};'>{output.rstrip()}</span>")
            
            # Show the result if it was an expression (like the result of 1+1)
            if result is not None:
                self.append(f"<span style='color: {ThemeColors.MAUVE};'>{repr(result)}</span>")
                
        except Exception as e:
            self.append(f"<span style='color: {ThemeColors.RED};'>✗ Error: {e}</span>")
        finally:
            sys.stdout = old_stdout

    def _execute_command(self):
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.StartOfLine)
        cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
        command = cursor.selectedText().strip()

        # Remove prompt symbols
        # Handle different prompt symbols
        if command.startswith("❯ "):
            command = command[2:]
        elif command.startswith(">>> "):
            command = command[4:]
        elif command.startswith("> "):
            command = command[2:]
        elif "❯" in command: # Fallback for just '❯' without space, or if it's part of a longer prompt
            command = command.split("❯")[-1].strip()
        elif command.startswith(">"): # Fallback for just '>' without space
            command = command[1:]

        command = command.strip()

        if not command:
            self.append("")
            self._prompt()
            return

        if command.lower() in ["clear", "cls", "clear()"]:
            self.clear()
            self._prompt()
            return

        self.append("")
        self._history.append(command)
        self._history_index = len(self._history)

        # Handle pip commands specially
        if command.startswith("!pip ") or command.startswith("pip "):
            self._run_pip_command(command)
        else:
            self._run_python(command)

        self._prompt()

    def _run_pip_command(self, command):
        """Run pip command to install packages"""
        import subprocess
        import sys

        # Remove !pip or pip prefix
        if command.startswith("!pip"):
            command = command[5:].strip()
        else:
            command = command[4:].strip()

        self.append(
            f"<span style='color: {ThemeColors.YELLOW};'>[Running pip {command}]</span>"
        )

        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip"] + command.split(),
                capture_output=True,
                text=True,
                timeout=120,
            )

            if result.stdout:
                self.append(
                    f"<span style='color: {ThemeColors.TEXT};'>{result.stdout}</span>"
                )
            if result.stderr:
                self.append(
                    f"<span style='color: {ThemeColors.RED};'>{result.stderr}</span>"
                )

            if result.returncode == 0:
                self.append(
                    f"<span style='color: {ThemeColors.GREEN};'>[Package installed successfully]</span>"
                )
            else:
                self.append(
                    f"<span style='color: {ThemeColors.RED};'>[Error: {result.returncode}]</span>"
                )

        except subprocess.TimeoutExpired:
            self.append(
                f"<span style='color: {ThemeColors.RED};'>[Error: Installation timed out]</span>"
            )
        except Exception as e:
            self.append(
                f"<span style='color: {ThemeColors.RED};'>[Error: {str(e)}]</span>"
            )

    def _history_navigate(self, direction):
        if not self._history:
            return

        self._history_index += direction

        if self._history_index < 0:
            self._history_index = 0
        elif self._history_index >= len(self._history):
            self._history_index = len(self._history) - 1
            return

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.StartOfLine)
        cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
        cursor.removeSelectedText()
        cursor.insertText(f"&gt;&gt;&gt; {self._history[self._history_index]}")


class PYXCEL(QMainWindow):
    def __init__(self):
        super().__init__()

        self.workbook = Workbook()
        self.formula_engine = FormulaEngine()
        self.file_manager = FileManager()
        self.chart_builder = ChartBuilder()
        self.macro_runner = MacroRunner(macro_manager)

        self._setup_ui()
        self._set_theme()
        self._connect_signals()
        self._create_default_sheet()

        self.setWindowTitle("PYXCEL")
        self.resize(900, 600)

    def toggle_theme(self):
        """Alternar entre tema oscuro (Dark) y claro (Light)"""
        current_mode = theme_manager.get_mode()
        new_mode = "light" if current_mode == "dark" else "dark"
        theme_manager.set_mode(new_mode)
        
        # Apply theme globally to the QApplication
        app = QApplication.instance()
        if app:
            app.setPalette(get_palette())
            app.setStyleSheet(get_app_stylesheet())
        
        # Refresh specific widgets and managers
        self._refresh_all_styles()

    def _refresh_all_styles(self):
        """Actualizar los estilos de todos los componentes visuales"""
        # Global Window Style
        self.setPalette(get_palette())
        self.setStyleSheet(get_app_stylesheet())
        
        # Toolbar & Menus
        self.toolbar_manager.update_style()
        self.menu_manager.update_style()
        
        # Right Panel Components
        self.terminal.update_style()
        if hasattr(self, 'spreadsheet') and self.spreadsheet:
            self.spreadsheet.update_style()
            
        # Left Panel (Sheets and Charts list)
        self._refresh_left_panel_style()

    def _refresh_left_panel_style(self):
        """Actualizar el estilo del panel izquierdo"""
        C = theme_manager.colors
        
        self.left_tabs.setStyleSheet(f"""
            QTabWidget::pane {{ background-color: {C.CRUST}; border: none; border-left: 1px solid {C.SURFACE0}; }}
            QTabBar::tab {{ background-color: {C.MANTLE}; color: {C.SUBTEXT0}; padding: 12px 16px; border: none; border-bottom: 1px solid {C.SURFACE0}; font-weight: 500; }}
            QTabBar::tab:selected {{ background-color: {C.CRUST}; color: {C.BLUE}; border-left: 3px solid {C.BLUE}; }}
            QTabBar::tab:hover {{ background-color: {C.SURFACE0}; }}
        """)
        
        self.shell_list.setStyleSheet(f"""
            QListWidget {{ background-color: {C.MANTLE}; color: {C.SUBTEXT0}; border: none; padding: 8px; }}
            QListWidget::item {{ padding: 6px 10px; border-radius: 4px; color: {C.SUBTEXT1}; }}
            QListWidget::item:selected {{ background-color: {C.BLUE}; color: {C.CRUST}; }}
            QListWidget::item:hover {{ background-color: {C.SURFACE0}; }}
        """)
        
        self.browser_tree.setStyleSheet(f"""
            QTreeWidget {{ background-color: {C.MANTLE}; color: {C.TEXT}; border: none; padding: 8px; }}
            QTreeWidget::item {{ padding: 6px 10px; border-radius: 4px; }}
            QTreeWidget::item:selected {{ background-color: {C.BLUE}; color: {C.CRUST}; }}
            QTreeWidget::item:hover {{ background-color: {C.SURFACE0}; }}
            QHeaderView::section {{ background-color: {C.MANTLE}; color: {C.SUBTEXT0}; border: none; font-weight: bold; padding: 8px; }}
        """)
        
        if hasattr(self, 'left_panel') and self.left_panel:
            self.left_panel.setStyleSheet(f"background-color: {C.CRUST}; border-right: 1px solid {C.SURFACE0};")

    def _set_theme(self):
        self._refresh_all_styles()

    def _setup_ui(self):
        self.toolbar_manager = ToolbarManager(self)
        self.menu_manager = MenuManager(self)

        self.setMenuBar(self.menu_manager.create_menus())
        self.addToolBar(self.toolbar_manager.create_toolbar())

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self._create_left_panel()

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.addWidget(self.left_panel)

        self.right_panel = QWidget()
        right_layout = QVBoxLayout(self.right_panel)
        right_layout.setContentsMargins(1, 1, 1, 1)
        right_layout.setSpacing(0)

        self.output_tabs = QTabWidget()
        self.output_tabs.setTabPosition(QTabWidget.TabPosition.South)

        self.terminal = TerminalWidget()

        self.spreadsheet_tab = QWidget()
        spreadsheet_layout = QVBoxLayout(self.spreadsheet_tab)
        spreadsheet_layout.setContentsMargins(0, 0, 0, 0)

        self.sheet_tabs = QTabWidget()
        self.sheet_tabs.setTabPosition(QTabWidget.TabPosition.South)
        self.sheet_tabs.setTabsClosable(True)
        self.sheet_tabs.tabCloseRequested.connect(self._close_tab)

        spreadsheet_layout.addWidget(self.sheet_tabs)

        self.output_tabs.addTab(self.terminal, "Terminal")
        self.output_tabs.addTab(self.spreadsheet_tab, "Editor")

        right_layout.addWidget(self.output_tabs)

        self.splitter.addWidget(self.right_panel)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 4)

        main_layout.addWidget(self.splitter)

        self.statusBar().showMessage("Ready")

        self._create_sheet_tab("Hoja1")

        self.sheet_tabs.currentChanged.connect(self._on_sheet_changed)

    def _create_left_panel(self):
        self.left_panel = QWidget()
        left_layout = QVBoxLayout(self.left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)

        self.left_tabs = QTabWidget()
        self.left_tabs.setTabPosition(QTabWidget.TabPosition.West)

        self.shell_list = QListWidget()
        self.shell_list.addItems(
            [
                ">>> Restarts: 1",
                "KeyboardInterrupt",
                "--",
            ]
        )

        self.browser_tree = QTreeWidget()
        self.browser_tree.setHeaderLabel("Explorador")
        self.browser_tree.setIndentation(16)

        root_item = QTreeWidgetItem(["[B] PYXCEL"])
        root_item.setForeground(0, QColor(ThemeColors.BLUE))
        child1 = QTreeWidgetItem(["[S] Hoja1"])
        child1.setForeground(0, QColor(ThemeColors.TEXT))
        root_item.addChild(child1)
        self.browser_tree.addTopLevelItem(root_item)
        self.browser_tree.expandAll()

        self.left_tabs.addTab(self.shell_list, "Shell")
        self.left_tabs.addTab(self.browser_tree, "Archivos")

        left_layout.addWidget(self.left_tabs)

        self.left_panel.setFixedWidth(170)

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
        sheet = self.get_current_sheet()
        if not sheet:
            QMessageBox.warning(self, "Gráfico", "No hay hoja de cálculo activa")
            return

        dialog = InsertChartDialog(self)
        if dialog.exec():
            chart_type = dialog.get_chart_type()
            data_range = dialog.get_data_range()
            labels_range = dialog.get_labels_range()

            if not data_range:
                QMessageBox.warning(
                    self, "Gráfico", "Por favor especifique un rango de datos"
                )
                return

            data = self._parse_range_to_data(sheet, data_range, labels_range)
            if not data:
                QMessageBox.warning(
                    self,
                    "Gráfico",
                    "No se pudieron obtener datos del rango especificado",
                )
                return

            chart_builder = ChartBuilder()
            canvas = None

            if "Barras" in chart_type and "apiladas" not in chart_type:
                canvas = chart_builder.create_bar_chart(data, "Gráfico de Barras")
            elif "Barras apiladas" in chart_type:
                canvas = chart_builder.create_stacked_bar_chart(data, "Barras Apiladas")
            elif "Líneas" in chart_type and "apiladas" not in chart_type:
                canvas = chart_builder.create_line_chart(data, "Gráfico de Líneas")
            elif "Líneas apiladas" in chart_type:
                canvas = chart_builder.create_stacked_line_chart(
                    data, "Líneas Apiladas"
                )
            elif "Circular" in chart_type or "Pie" in chart_type:
                canvas = chart_builder.create_pie_chart(data, "Gráfico Circular")
            elif "Dispersión" in chart_type or "Scatter" in chart_type:
                canvas = chart_builder.create_scatter_chart(
                    data, "Gráfico de Dispersión"
                )
            elif "Área" in chart_type and "apilada" not in chart_type:
                canvas = chart_builder.create_area_chart(data, "Gráfico de Área")
            elif "Área apilada" in chart_type:
                canvas = chart_builder.create_stacked_area_chart(data, "Área Apilada")

            if canvas:
                self._show_chart_window(canvas, chart_type)

    def _parse_range_to_data(
        self, sheet, data_range: str, labels_range: str = None
    ) -> dict:
        import re

        match = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", data_range.upper())
        if not match:
            return {}

        start_col = self._letter_to_col(match.group(1))
        start_row = int(match.group(2)) - 1
        end_col = self._letter_to_col(match.group(3))
        end_row = int(match.group(4)) - 1

        labels = []
        if labels_range:
            label_match = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", labels_range.upper())
            if label_match:
                l_start_col = self._letter_to_col(label_match.group(1))
                l_start_row = int(label_match.group(2)) - 1
                for col in range(l_start_col, l_start_col + (end_col - start_col + 1)):
                    val = sheet.get_cell_value(l_start_row, col)
                    labels.append(str(val) if val is not None else f"Item{col}")
        else:
            for col in range(start_col, end_col + 1):
                labels.append(f"Col{chr(65 + col)}")

        data = {}
        for col in range(start_col, end_col + 1):
            key = (
                labels[col - start_col]
                if col - start_col < len(labels)
                else f"Col{col}"
            )
            values = []
            for row in range(start_row, end_row + 1):
                val = sheet.get_cell_value(row, col)
                if val is not None:
                    try:
                        values.append(float(val))
                    except (ValueError, TypeError):
                        pass
            if values:
                data[key] = sum(values) / len(values) if len(values) > 1 else values[0]

        return data

    def _letter_to_col(self, letter: str) -> int:
        col = 0
        for char in letter.upper():
            col = col * 26 + (ord(char) - ord("A") + 1)
        return col - 1

    def _show_chart_window(self, canvas, chart_type: str):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{chart_type} - PYXCEL")
        dialog.setMinimumSize(700, 500)

        layout = QVBoxLayout(dialog)

        scroll = QScrollArea()
        scroll.setWidget(canvas)
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"""
            QScrollArea {{
                background-color: {ThemeColors.BASE};
                border: none;
            }}
        """)

        layout.addWidget(scroll)

        close_btn = QPushButton("Cerrar")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)

        dialog.exec()

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
            sort_levels = dialog.get_sort_levels()
            if sort_levels:
                sheet.sort_by_multiple_columns(sort_levels)

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
    from PySide6.QtGui import QPixmap, QPainter, QFont
    from PySide6.QtCore import Qt, QTimer

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Create splash screen
    splash_pixmap = QPixmap(500, 350)
    splash_pixmap.fill(QColor(ThemeColors.CRUST))

    # Draw splash content
    painter = QPainter(splash_pixmap)
    painter.setPen(QColor(ThemeColors.BLUE))
    painter.setFont(QFont("Segoe UI", 28, QFont.Weight.Bold))
    painter.drawText(
        splash_pixmap.rect().adjusted(0, 60, 0, -180),
        Qt.AlignmentFlag.AlignHCenter,
        "PYXCEL",
    )

    painter.setPen(QColor(ThemeColors.SUBTEXT0))
    painter.setFont(QFont("Segoe UI", 12))
    painter.drawText(
        splash_pixmap.rect().adjusted(0, 120, 0, -120),
        Qt.AlignmentFlag.AlignHCenter,
        "Hoja de Calculo Python",
    )

    # Loading bar background
    painter.setBrush(QColor(ThemeColors.SURFACE0))
    painter.setPen(QColor(ThemeColors.SURFACE1))
    bar_rect = splash_pixmap.rect().adjusted(50, 250, -50, -80)
    painter.drawRect(bar_rect)

    # Loading bar progress
    painter.setBrush(QColor(ThemeColors.BLUE))
    painter.setPen(Qt.PenStyle.NoPen)
    progress_rect = bar_rect.adjusted(2, 2, -2, -2)
    progress_width = int(progress_rect.width() * 0.25)
    painter.drawRect(
        progress_rect.adjusted(0, 0, -progress_rect.width() + progress_width, 0)
    )

    painter.setPen(QColor(ThemeColors.SUBTEXT1))
    painter.setFont(QFont("Segoe UI", 10))
    painter.drawText(
        splash_pixmap.rect().adjusted(0, 290, 0, -30),
        Qt.AlignmentFlag.AlignHCenter,
        "Cargando componentes...",
    )
    painter.end()

    splash = QSplashScreen(splash_pixmap, Qt.WindowType.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    # Apply global theme
    app.setPalette(get_palette())
    app.setStyleSheet(get_app_stylesheet())

    # Set application-level attributes
    app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps)

    # Simulate loading time
    import time

    for i in range(1, 5):
        time.sleep(0.5)
        # Update progress bar
        splash_pixmap.fill(QColor(ThemeColors.CRUST))
        painter = QPainter(splash_pixmap)
        painter.setPen(QColor(ThemeColors.BLUE))
        painter.setFont(QFont("Segoe UI", 28, QFont.Weight.Bold))
        painter.drawText(
            splash_pixmap.rect().adjusted(0, 60, 0, -180),
            Qt.AlignmentFlag.AlignHCenter,
            "PYXCEL",
        )

        painter.setPen(QColor(ThemeColors.SUBTEXT0))
        painter.setFont(QFont("Segoe UI", 12))
        painter.drawText(
            splash_pixmap.rect().adjusted(0, 120, 0, -120),
            Qt.AlignmentFlag.AlignHCenter,
            "Hoja de Calculo Python",
        )

        painter.setBrush(QColor(ThemeColors.SURFACE0))
        painter.setPen(QColor(ThemeColors.SURFACE1))
        bar_rect = splash_pixmap.rect().adjusted(50, 250, -50, -80)
        painter.drawRect(bar_rect)

        painter.setBrush(QColor(ThemeColors.BLUE))
        painter.setPen(Qt.PenStyle.NoPen)
        progress_rect = bar_rect.adjusted(2, 2, -2, -2)
        progress_width = int(progress_rect.width() * (i / 4))
        painter.drawRect(
            progress_rect.adjusted(0, 0, -progress_rect.width() + progress_width, 0)
        )

        painter.setPen(QColor(ThemeColors.SUBTEXT1))
        painter.setFont(QFont("Segoe UI", 10))
        messages = [
            "Cargando componentes...",
            "Inicializando motor...",
            "Preparando interfaz...",
            "Cargando hojas de calculo...",
        ]
        painter.drawText(
            splash_pixmap.rect().adjusted(0, 290, 0, -30),
            Qt.AlignmentFlag.AlignHCenter,
            messages[i - 1],
        )
        painter.end()

        splash.setPixmap(splash_pixmap)
        app.processEvents()

    window = PYXCEL()
    window.show()
    splash.finish(window)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
