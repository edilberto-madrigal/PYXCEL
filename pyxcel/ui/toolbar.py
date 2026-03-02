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
        self._toolbar = QToolBar("Herramientas")
        self._toolbar.setMovable(False)

        self._font_combo = QFontComboBox()
        self._font_combo.setFixedWidth(150)
        self._font_combo.currentFontChanged.connect(self._on_font_changed)

        self._font_size_spin = QSpinBox()
        self._font_size_spin.setRange(6, 72)
        self._font_size_spin.setValue(10)
        self._font_size_spin.setFixedWidth(50)
        self._font_size_spin.valueChanged.connect(self._on_font_size_changed)

        bold_btn = QPushButton("N")
        bold_btn.setFixedSize(30, 25)
        bold_btn.setStyleSheet("font-weight: bold;")
        bold_btn.clicked.connect(self._on_bold_clicked)

        italic_btn = QPushButton("I")
        italic_btn.setFixedSize(30, 25)
        italic_btn.setStyleSheet("font-style: italic;")
        italic_btn.clicked.connect(self._on_italic_clicked)

        underline_btn = QPushButton("U")
        underline_btn.setFixedSize(30, 25)
        underline_btn.setStyleSheet("text-decoration: underline;")
        underline_btn.clicked.connect(self._on_underline_clicked)

        bg_color_btn = QPushButton(" Color ")
        bg_color_btn.clicked.connect(self._on_bg_color_clicked)

        align_left_btn = QPushButton("Izq")
        align_left_btn.setFixedSize(40, 25)
        align_left_btn.clicked.connect(
            lambda: self._on_align_clicked(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
        )

        align_center_btn = QPushButton("Cen")
        align_center_btn.setFixedSize(40, 25)
        align_center_btn.clicked.connect(
            lambda: self._on_align_clicked(
                Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
            )
        )

        align_right_btn = QPushButton("Der")
        align_right_btn.setFixedSize(40, 25)
        align_right_btn.clicked.connect(
            lambda: self._on_align_clicked(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
        )

        self._toolbar.addWidget(QLabel("Fuente:"))
        self._toolbar.addWidget(self._font_combo)
        self._toolbar.addWidget(QLabel(" Tamaño:"))
        self._toolbar.addWidget(self._font_size_spin)
        self._toolbar.addSeparator()
        self._toolbar.addWidget(bold_btn)
        self._toolbar.addWidget(italic_btn)
        self._toolbar.addWidget(underline_btn)
        self._toolbar.addSeparator()
        self._toolbar.addWidget(bg_color_btn)
        self._toolbar.addSeparator()
        self._toolbar.addWidget(align_left_btn)
        self._toolbar.addWidget(align_center_btn)
        self._toolbar.addWidget(align_right_btn)

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

    def _on_bg_color_clicked(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.main_window.format_background_color(color)

    def _on_align_clicked(self, alignment: Qt.AlignmentFlag):
        self.main_window.format_alignment(alignment)

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
        self._create_insert_menu()
        self._create_format_menu()
        self._create_data_menu()
        self._create_help_menu()

        return self._menu_bar

    def _create_file_menu(self):
        file_menu = self._menu_bar.addMenu("&Archivo")

        nuevo_action = QAction("&Nuevo", self.main_window)
        nuevo_action.setShortcut(QKeySequence.StandardKey.New)
        nuevo_action.triggered.connect(self.main_window.new_file)
        file_menu.addAction(nuevo_action)

        abrir_action = QAction("&Abrir...", self.main_window)
        abrir_action.setShortcut(QKeySequence.StandardKey.Open)
        abrir_action.triggered.connect(self.main_window.open_file)
        file_menu.addAction(abrir_action)

        file_menu.addSeparator()

        guardar_action = QAction("&Guardar", self.main_window)
        guardar_action.setShortcut(QKeySequence.StandardKey.Save)
        guardar_action.triggered.connect(self.main_window.save_file)
        file_menu.addAction(guardar_action)

        guardar_como_action = QAction("Guardar &como...", self.main_window)
        guardar_como_action.setShortcut(QKeySequence.StandardKey.SaveAs)
        guardar_como_action.triggered.connect(self.main_window.save_file_as)
        file_menu.addAction(guardar_como_action)

        file_menu.addSeparator()

        exportar_pdf_action = QAction("Exportar a &PDF...", self.main_window)
        exportar_pdf_action.triggered.connect(self.main_window.export_to_pdf)
        file_menu.addAction(exportar_pdf_action)

        file_menu.addSeparator()

        salir_action = QAction("&Salir", self.main_window)
        salir_action.setShortcut(QKeySequence.StandardKey.Quit)
        salir_action.triggered.connect(self.main_window.close)
        file_menu.addAction(salir_action)

    def _create_edit_menu(self):
        edit_menu = self._menu_bar.addMenu("&Editar")

        deshacer_action = QAction("&Deshacer", self.main_window)
        deshacer_action.setShortcut(QKeySequence.StandardKey.Undo)
        deshacer_action.triggered.connect(self.main_window.undo)
        edit_menu.addAction(deshacer_action)

        rehacer_action = QAction("&Rehacer", self.main_window)
        rehacer_action.setShortcut(QKeySequence.StandardKey.Redo)
        rehacer_action.triggered.connect(self.main_window.redo)
        edit_menu.addAction(rehacer_action)

        edit_menu.addSeparator()

        cortar_action = QAction("Cor&tar", self.main_window)
        cortar_action.setShortcut(QKeySequence.StandardKey.Cut)
        cortar_action.triggered.connect(self.main_window.cut)
        edit_menu.addAction(cortar_action)

        copiar_action = QAction("&Copiar", self.main_window)
        copiar_action.setShortcut(QKeySequence.StandardKey.Copy)
        copiar_action.triggered.connect(self.main_window.copy)
        edit_menu.addAction(copiar_action)

        pegar_action = QAction("&Pegar", self.main_window)
        pegar_action.setShortcut(QKeySequence.StandardKey.Paste)
        pegar_action.triggered.connect(self.main_window.paste)
        edit_menu.addAction(pegar_action)

        edit_menu.addSeparator()

        buscar_action = QAction("&Buscar...", self.main_window)
        buscar_action.setShortcut(QKeySequence.StandardKey.Find)
        buscar_action.triggered.connect(self.main_window.show_find_dialog)
        edit_menu.addAction(buscar_action)

        reemplazar_action = QAction("Re&emplazar...", self.main_window)
        reemplazar_action.setShortcut(QKeySequence("Ctrl+H"))
        reemplazar_action.triggered.connect(self.main_window.show_replace_dialog)
        edit_menu.addAction(reemplazar_action)

    def _create_insert_menu(self):
        insert_menu = self._menu_bar.addMenu("&Insertar")

        fila_action = QAction("&Fila", self.main_window)
        fila_action.triggered.connect(self.main_window.insert_row)
        insert_menu.addAction(fila_action)

        columna_action = QAction("&Columna", self.main_window)
        columna_action.triggered.connect(self.main_window.insert_column)
        insert_menu.addAction(columna_action)

        insert_menu.addSeparator()

        grafico_action = QAction("&Gráfico...", self.main_window)
        grafico_action.triggered.connect(self.main_window.insert_chart)
        insert_menu.addAction(grafico_action)

    def _create_format_menu(self):
        format_menu = self._menu_bar.addMenu("F&ormato")

        celdas_action = QAction("Celdas...", self.main_window)
        celdas_action.triggered.connect(self.main_window.format_cells)
        format_menu.addAction(celdas_action)

        format_menu.addSeparator()

        autofiltro_action = QAction("&Autofiltro", self.main_window)
        autofiltro_action.triggered.connect(self.main_window.toggle_autofilter)
        format_menu.addAction(autofiltro_action)

    def _create_data_menu(self):
        data_menu = self._menu_bar.addMenu("&Datos")

        ordenar_action = QAction("&Ordenar", self.main_window)
        ordenar_action.triggered.connect(self.main_window.sort_data)
        data_menu.addAction(ordenar_action)

        filtro_action = QAction("F&iltro", self.main_window)
        filtro_action.triggered.connect(self.main_window.filter_data)
        data_menu.addAction(filtro_action)

    def _create_help_menu(self):
        help_menu = self._menu_bar.addMenu("A&yuda")

        acerca_action = QAction("&Acerca de PYXCEL", self.main_window)
        acerca_action.triggered.connect(self.main_window.show_about)
        help_menu.addAction(acerca_action)
