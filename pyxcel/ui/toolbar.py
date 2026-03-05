from PySide6.QtWidgets import (
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QToolBar,
    QStatusBar,
    QMenuBar,
    QMenu,
    QDialog,
    QLabel,
    QPushButton,
    QComboBox,
    QFileDialog,
    QMessageBox,
    QInputDialog,
    QColorDialog,
    QFontComboBox,
    QSpinBox,
    QApplication,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QAction, QKeySequence, QIcon, QColor, QFont

from .table import SpreadsheetWidget


class ToolbarManager:
    def __init__(self, main_window):
        self.main_window = main_window
        self._toolbar = None
        self._font_combo = None
        self._font_size_spin = None

    def create_toolbar(self) -> QToolBar:
        self._toolbar = QToolBar()
        self._toolbar.setMovable(False)

        open_btn = QPushButton("Abrir")
        open_btn.setFixedSize(50, 25)
        open_btn.clicked.connect(self.main_window.open_file)

        save_btn = QPushButton("Guardar")
        save_btn.setFixedSize(60, 25)
        save_btn.clicked.connect(self.main_window.save_file)

        self._toolbar.addWidget(open_btn)
        self._toolbar.addWidget(save_btn)
        self._toolbar.addSeparator()

        self._font_combo = QFontComboBox()
        self._font_combo.setFixedWidth(120)
        self._font_combo.currentFontChanged.connect(self._on_font_changed)

        self._font_size_spin = QSpinBox()
        self._font_size_spin.setRange(6, 72)
        self._font_size_spin.setValue(10)
        self._font_size_spin.setFixedWidth(40)
        self._font_size_spin.valueChanged.connect(self._on_font_size_changed)

        bold_btn = QPushButton("N")
        bold_btn.setFixedSize(25, 22)
        bold_btn.setStyleSheet("font-weight: bold;")
        bold_btn.clicked.connect(self._on_bold_clicked)

        italic_btn = QPushButton("I")
        italic_btn.setFixedSize(25, 22)
        italic_btn.setStyleSheet("font-style: italic;")
        italic_btn.clicked.connect(self._on_italic_clicked)

        underline_btn = QPushButton("U")
        underline_btn.setFixedSize(25, 22)
        underline_btn.setStyleSheet("text-decoration: underline;")
        underline_btn.clicked.connect(self._on_underline_clicked)

        self._toolbar.addWidget(self._font_combo)
        self._toolbar.addWidget(self._font_size_spin)
        self._toolbar.addSeparator()
        self._toolbar.addWidget(bold_btn)
        self._toolbar.addWidget(italic_btn)
        self._toolbar.addWidget(underline_btn)

        return self._toolbar

    def _on_font_changed(self, font: QFont):
        self.main_window.format_font(font.family())

    def _on_font_size_changed(self, size: int):
        self.main_window.format_font_size(size)

    def _on_bold_clicked(self):
        self.main_window.format_bold()

    def _on_italic_clicked(self):
        self.main_window.format_italic()

    def _on_underline_clicked(self):
        self.main_window.format_underline()

    def update_selection(self, row: int, col: int):
        pass


class MenuManager:
    def __init__(self, main_window):
        self.main_window = main_window
        self._menu_bar = None

    def create_menus(self) -> QMenuBar:
        self._menu_bar = QMenuBar()

        self._create_file_menu()
        self._create_edit_menu()
        self._create_shell_menu()
        self._create_debug_menu()
        self._create_options_menu()
        self._create_window_menu()
        self._create_help_menu()

        return self._menu_bar

    def _create_file_menu(self):
        file_menu = self._menu_bar.addMenu("Archivo")

        nuevo_action = QAction("Nuevo", self.main_window)
        nuevo_action.setShortcut(QKeySequence.StandardKey.New)
        nuevo_action.triggered.connect(self.main_window.new_file)
        file_menu.addAction(nuevo_action)

        abrir_action = QAction("Abrir...", self.main_window)
        abrir_action.setShortcut(QKeySequence.StandardKey.Open)
        abrir_action.triggered.connect(self.main_window.open_file)
        file_menu.addAction(abrir_action)

        file_menu.addSeparator()

        guardar_action = QAction("Guardar", self.main_window)
        guardar_action.setShortcut(QKeySequence.StandardKey.Save)
        guardar_action.triggered.connect(self.main_window.save_file)
        file_menu.addAction(guardar_action)

        guardar_como_action = QAction("Guardar como...", self.main_window)
        guardar_como_action.setShortcut(QKeySequence.StandardKey.SaveAs)
        guardar_como_action.triggered.connect(self.main_window.save_file_as)
        file_menu.addAction(guardar_como_action)

        file_menu.addSeparator()

        exportar_pdf_action = QAction("Exportar a PDF...", self.main_window)
        exportar_pdf_action.triggered.connect(self.main_window.export_to_pdf)
        file_menu.addAction(exportar_pdf_action)

        file_menu.addSeparator()

        salir_action = QAction("Salir", self.main_window)
        salir_action.setShortcut(QKeySequence.StandardKey.Quit)
        salir_action.triggered.connect(self.main_window.close)
        file_menu.addAction(salir_action)

    def _create_edit_menu(self):
        edit_menu = self._menu_bar.addMenu("Editar")

        deshacer_action = QAction("Deshacer", self.main_window)
        deshacer_action.setShortcut(QKeySequence.StandardKey.Undo)
        deshacer_action.triggered.connect(self.main_window.undo)
        edit_menu.addAction(deshacer_action)

        rehacer_action = QAction("Rehacer", self.main_window)
        rehacer_action.setShortcut(QKeySequence.StandardKey.Redo)
        rehacer_action.triggered.connect(self.main_window.redo)
        edit_menu.addAction(rehacer_action)

        edit_menu.addSeparator()

        cortar_action = QAction("Cortar", self.main_window)
        cortar_action.setShortcut(QKeySequence.StandardKey.Cut)
        cortar_action.triggered.connect(self.main_window.cut)
        edit_menu.addAction(cortar_action)

        copiar_action = QAction("Copiar", self.main_window)
        copiar_action.setShortcut(QKeySequence.StandardKey.Copy)
        copiar_action.triggered.connect(self.main_window.copy)
        edit_menu.addAction(copiar_action)

        pegar_action = QAction("Pegar", self.main_window)
        pegar_action.setShortcut(QKeySequence.StandardKey.Paste)
        pegar_action.triggered.connect(self.main_window.paste)
        edit_menu.addAction(pegar_action)

        edit_menu.addSeparator()

        buscar_action = QAction("Buscar...", self.main_window)
        buscar_action.setShortcut(QKeySequence.StandardKey.Find)
        buscar_action.triggered.connect(self.main_window.show_find_dialog)
        edit_menu.addAction(buscar_action)

        reemplazar_action = QAction("Reemplazar...", self.main_window)
        reemplazar_action.setShortcut(QKeySequence("Ctrl+H"))
        reemplazar_action.triggered.connect(self.main_window.show_replace_dialog)
        edit_menu.addAction(reemplazar_action)

        edit_menu.addSeparator()

        select_all_action = QAction("Seleccionar todo", self.main_window)
        select_all_action.setShortcut(QKeySequence.StandardKey.SelectAll)
        edit_menu.addAction(select_all_action)

    def _create_shell_menu(self):
        shell_menu = self._menu_bar.addMenu("Shell")

        restart_action = QAction("Reiniciar Shell", self.main_window)
        restart_action.triggered.connect(self._restart_shell)
        shell_menu.addAction(restart_action)

    def _create_debug_menu(self):
        debug_menu = self._menu_bar.addMenu("Depurar")

        go_action = QAction("Ir", self.main_window)
        go_action.setShortcut(QKeySequence("F5"))
        debug_menu.addAction(go_action)

        step_action = QAction("Paso", self.main_window)
        step_action.setShortcut(QKeySequence("F6"))
        debug_menu.addAction(step_action)

        over_action = QAction("Sobre", self.main_window)
        over_action.setShortcut(QKeySequence("F7"))
        debug_menu.addAction(over_action)

        out_action = QAction("Fuera", self.main_window)
        out_action.setShortcut(QKeySequence("F8"))
        debug_menu.addAction(out_action)

        debug_menu.addSeparator()

        quit_debug_action = QAction("Salir", self.main_window)
        quit_debug_action.setShortcut(QKeySequence("F9"))
        debug_menu.addAction(quit_debug_action)

    def _create_options_menu(self):
        options_menu = self._menu_bar.addMenu("Opciones")

        configure_action = QAction("Configurar IDLE...", self.main_window)
        options_menu.addAction(configure_action)

        options_menu.addSeparator()

        show_action = QAction("Mostrar contexto de código", self.main_window)
        show_action.setCheckable(True)
        show_action.setChecked(True)
        options_menu.addAction(show_action)

    def _create_window_menu(self):
        window_menu = self._menu_bar.addMenu("Ventana")

        zoom_action = QAction("Zoom/Altura", self.main_window)
        zoom_action.triggered.connect(self._zoom_window)
        window_menu.addAction(zoom_action)

    def _create_help_menu(self):
        help_menu = self._menu_bar.addMenu("Ayuda")

        about_action = QAction("Acerca de PYXCEL", self.main_window)
        about_action.triggered.connect(self.main_window.show_about)
        help_menu.addAction(about_action)

    def _restart_shell(self):
        QMessageBox.information(self.main_window, "Shell", "Shell reiniciado")

    def _zoom_window(self):
        QMessageBox.information(self.main_window, "Ventana", "Función de zoom")
