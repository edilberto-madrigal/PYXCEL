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
    QToolButton,
    QStyle,
)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QAction, QKeySequence, QIcon, QColor, QFont

from .table import SpreadsheetWidget
from .theme import theme_manager, ThemeColors, get_accent_button_style, get_icon_button_style


class ToolbarManager:
    def __init__(self, main_window):
        self.main_window = main_window
        self._toolbar = None
        self._font_combo = None
        self._font_size_spin = None
        self._buttons = {}

    def _create_icon_button(self, icon_name: str, tooltip: str, slot) -> QToolButton:
        """Create a modern icon button with tooltip"""
        btn = QToolButton()

        # Use more common icons to avoid AttributeError
        standard_icons = {
            "open": QStyle.StandardPixmap.SP_DialogOpenButton,
            "save": QStyle.StandardPixmap.SP_DialogSaveButton,
            "bold": QStyle.StandardPixmap.SP_FileDialogInfoView,
            "italic": QStyle.StandardPixmap.SP_FileDialogContentsView,
            "underline": QStyle.StandardPixmap.SP_ArrowRight,
            "new": QStyle.StandardPixmap.SP_FileIcon,
            "print": QStyle.StandardPixmap.SP_FileIcon,
            "cut": QStyle.StandardPixmap.SP_MediaPlay,
            "copy": QStyle.StandardPixmap.SP_FileDialogDetailedView,
            "paste": QStyle.StandardPixmap.SP_FileDialogListView,
            "undo": QStyle.StandardPixmap.SP_ArrowLeft,
            "redo": QStyle.StandardPixmap.SP_ArrowRight,
        }

        # Safe attribute access
        icon_enum = QStyle.StandardPixmap.SP_DesktopIcon
        if icon_name in standard_icons:
            icon_enum = standard_icons[icon_name]
        
        try:
            btn.setIcon(self.main_window.style().standardIcon(icon_enum))
        except (AttributeError, Exception):
            # Fallback to a very safe icon
            try:
                btn.setIcon(self.main_window.style().standardIcon(QStyle.StandardPixmap.SP_DesktopIcon))
            except:
                pass

        btn.setToolTip(tooltip)
        btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        btn.setFixedSize(32, 32)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.clicked.connect(slot)

        btn.setStyleSheet(f"""
            QToolButton {{
                background-color: transparent;
                border: none;
                border-radius: 6px;
                padding: 4px;
            }}
            QToolButton:hover {{
                background-color: {theme_manager.colors.SURFACE0};
            }}
            QToolButton:pressed {{
                background-color: {theme_manager.colors.SURFACE1};
            }}
        """)

        return btn

    def create_toolbar(self) -> QToolBar:
        self._toolbar = QToolBar()
        self._toolbar.setMovable(False)
        self._toolbar.setIconSize(QSize(20, 20))
        self._toolbar.setStyleSheet(f"""
            QToolBar {{
                background-color: {theme_manager.colors.MANTLE};
                border: none;
                padding: 4px 8px;
                spacing: 2px;
            }}
        """)

        # File operations group
        new_btn = self._create_icon_button(
            "new", "Nuevo archivo (Ctrl+N)", self.main_window.new_file
        )
        open_btn = self._create_icon_button(
            "open", "Abrir archivo (Ctrl+O)", self.main_window.open_file
        )
        save_btn = self._create_icon_button(
            "save", "Guardar (Ctrl+S)", self.main_window.save_file
        )

        for btn in [new_btn, open_btn, save_btn]:
            self._toolbar.addWidget(btn)

        self._add_separator()

        # Font family selector
        self._font_combo = QFontComboBox()
        self._font_combo.setFixedWidth(150)
        self._font_combo.setStyleSheet(f"""
            QFontComboBox {{
                background-color: {ThemeColors.SURFACE0};
                color: {ThemeColors.TEXT};
                border: 1px solid {ThemeColors.SURFACE1};
                border-radius: 6px;
                padding: 4px 8px;
                min-width: 120px;
            }}
            QFontComboBox::drop-down {{
                border: none;
                width: 20px;
            }}
        """)
        self._font_combo.currentFontChanged.connect(self._on_font_changed)
        self._toolbar.addWidget(self._font_combo)

        # Font size
        self._font_size_spin = QSpinBox()
        self._font_size_spin.setRange(6, 72)
        self._font_size_spin.setValue(11)
        self._font_size_spin.setFixedWidth(55)
        self._font_size_spin.valueChanged.connect(self._on_font_size_changed)
        self._toolbar.addWidget(self._font_size_spin)
        
        self._add_separator()

        # Format buttons
        bold_btn = self._create_format_button(
            "N", "Negrita (Ctrl+B)", self._on_bold_clicked
        )
        italic_btn = self._create_format_button(
            "I", "Cursiva (Ctrl+I)", self._on_italic_clicked
        )
        underline_btn = self._create_format_button(
            "U", "Subrayado (Ctrl+U)", self._on_underline_clicked
        )

        bold_btn.setStyleSheet(bold_btn.styleSheet() + "font-weight: bold;")
        italic_btn.setStyleSheet(italic_btn.styleSheet() + "font-style: italic;")
        underline_btn.setStyleSheet(
            underline_btn.styleSheet() + "text-decoration: underline;"
        )

        self._format_buttons = [bold_btn, italic_btn, underline_btn]

        for btn in self._format_buttons:
            self._toolbar.addWidget(btn)

        self._add_separator()

        # Add alignment buttons (Hidden for now)
        align_left = self._create_icon_button("cut", "Alinear a la izquierda", lambda: None)
        align_center = self._create_icon_button("copy", "Centrar", lambda: None)
        align_right = self._create_icon_button("paste", "Alinear a la derecha", lambda: None)

        for btn in [align_left, align_center, align_right]:
            btn.setVisible(False)
            self._toolbar.addWidget(btn)

        self.update_style()
        return self._toolbar

    def update_style(self):
        """Actualizar el estilo visual de la barra de herramientas"""
        C = theme_manager.colors
        
        # Main Toolbar
        self._toolbar.setStyleSheet(f"""
            QToolBar {{
                background-color: {C.MANTLE};
                border: none;
                padding: 4px 8px;
                spacing: 2px;
            }}
        """)
        
        # Font ComboBox
        if self._font_combo:
            self._font_combo.setStyleSheet(f"""
                QFontComboBox {{
                    background-color: {C.SURFACE0};
                    color: {C.TEXT};
                    border: 1px solid {C.SURFACE1};
                    border-radius: 6px;
                    padding: 4px 8px;
                    min-width: 120px;
                }}
            """)
            
        # Font Size SpinBox
        if self._font_size_spin:
            self._font_size_spin.setStyleSheet(f"""
                QSpinBox {{
                    background-color: {C.SURFACE0};
                    color: {C.TEXT};
                    border: 1px solid {C.SURFACE1};
                    border-radius: 6px;
                    padding: 4px 6px;
                }}
            """)

        # Format Buttons (Negrita, Cursiva, etc)
        if hasattr(self, '_format_buttons'):
            for btn in self._format_buttons:
                btn.setStyleSheet(f"""
                    QToolButton {{
                        background-color: {C.SURFACE0};
                        color: {C.TEXT};
                        border: 1px solid {C.SURFACE1};
                        border-radius: 4px;
                        font-size: 12px;
                        font-weight: 600;
                    }}
                    QToolButton:hover {{
                        background-color: {C.SURFACE1};
                        border-color: {C.BLUE};
                    }}
                    QToolButton:pressed {{
                        background-color: {C.BLUE};
                        color: {C.CRUST};
                    }}
                """)

    def _create_format_button(self, text: str, tooltip: str, slot) -> QToolButton:
        """Create a text-based format button"""
        btn = QToolButton()
        btn.setText(text)
        btn.setToolTip(tooltip)
        btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        btn.setFixedSize(28, 28)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.clicked.connect(slot)
        btn.setStyleSheet(f"""
            QToolButton {{
                background-color: {theme_manager.colors.SURFACE0};
                color: {theme_manager.colors.TEXT};
                border: 1px solid {theme_manager.colors.SURFACE1};
                border-radius: 4px;
                font-size: 12px;
                font-weight: 600;
            }}
            QToolButton:hover {{
                background-color: {theme_manager.colors.SURFACE1};
                border-color: {theme_manager.colors.BLUE};
            }}
            QToolButton:pressed {{
                background-color: {theme_manager.colors.BLUE};
                color: {theme_manager.colors.CRUST};
            }}
        """)
        return btn

    def _add_separator(self):
        sep = QWidget()
        sep.setFixedWidth(1)
        sep.setStyleSheet(f"background-color: {theme_manager.colors.SURFACE0}; margin: 4px 8px;")
        self._toolbar.addWidget(sep)

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

        self.update_style()
        return self._menu_bar

    def update_style(self):
        """Actualizar el estilo visual de la barra de menú"""
        C = theme_manager.colors
        self._menu_bar.setStyleSheet(f"""
            QMenuBar {{
                background-color: {C.MANTLE};
                color: {C.TEXT};
                border-bottom: 1px solid {C.SURFACE0};
                padding: 2px;
            }}
            QMenuBar::item:selected {{
                background-color: {C.SURFACE0};
            }}
            QMenu {{
                background-color: {C.MANTLE};
                color: {C.TEXT};
                border: 1px solid {C.SURFACE0};
            }}
            QMenu::item:selected {{
                background-color: {C.BLUE};
                color: {C.CRUST};
            }}
        """)

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

        theme_action = QAction("Cambiar Tema (Oscuro/Claro)", self.main_window)
        theme_action.triggered.connect(self.main_window.toggle_theme)
        options_menu.addAction(theme_action)

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
