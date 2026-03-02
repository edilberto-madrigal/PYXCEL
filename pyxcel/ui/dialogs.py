from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QCheckBox,
    QGroupBox,
    QFormLayout,
    QSpinBox,
    QComboBox,
    QDialogButtonBox,
    QTabWidget,
    QWidget,
    QColorDialog,
    QFontComboBox,
)
from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QColor, QFont


class FindReplaceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("Buscar y Reemplazar")
        self.setModal(True)
        self.setMinimumWidth(400)

        layout = QVBoxLayout(self)

        tab_widget = QTabWidget()

        find_tab = self._create_find_tab()
        replace_tab = self._create_replace_tab()

        tab_widget.addTab(find_tab, "Buscar")
        tab_widget.addTab(replace_tab, "Reemplazar")

        layout.addWidget(tab_widget)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.FindNext
            | QDialogButtonBox.StandardButton.Replace
            | QDialogButtonBox.StandardButton.ReplaceAll
            | QDialogButtonBox.StandardButton.Close
        )

        buttons.button(QDialogButtonBox.StandardButton.FindNext).clicked.connect(
            self._on_find_next
        )
        buttons.button(QDialogButtonBox.StandardButton.Replace).clicked.connect(
            self._on_replace
        )
        buttons.button(QDialogButtonBox.StandardButton.ReplaceAll).clicked.connect(
            self._on_replace_all
        )
        buttons.button(QDialogButtonBox.StandardButton.Close).clicked.connect(
            self.close
        )

        layout.addWidget(buttons)

        self.find_button = buttons.button(QDialogButtonBox.StandardButton.FindNext)
        self.replace_button = buttons.button(QDialogButtonBox.StandardButton.Replace)
        self.replace_all_button = buttons.button(
            QDialogButtonBox.StandardButton.ReplaceAll
        )

    def _create_find_tab(self):
        widget = QWidget()
        layout = QFormLayout(widget)

        self.find_text = QLineEdit()
        layout.addRow("Buscar:", self.find_text)

        self.match_case = QCheckBox("Distinguir mayúsculas")
        layout.addRow("", self.match_case)

        self.regex_check = QCheckBox("Expresión regular")
        layout.addRow("", self.regex_check)

        return widget

    def _create_replace_tab(self):
        widget = QWidget()
        layout = QFormLayout(widget)

        self.replace_find_text = QLineEdit()
        layout.addRow("Buscar:", self.replace_find_text)

        self.replace_text = QLineEdit()
        layout.addRow("Reemplazar con:", self.replace_text)

        self.replace_match_case = QCheckBox("Distinguir mayúsculas")
        layout.addRow("", self.replace_match_case)

        return widget

    def _on_find_next(self):
        self.accept()
        self.find_next.emit(self.get_find_text(), self.match_case_checked())

    def _on_replace(self):
        self.accept()
        self.replace.emit(
            self.get_find_text(),
            self.get_replace_text(),
            self.replace_match_case_checked(),
        )

    def _on_replace_all(self):
        self.accept()
        self.replace_all.emit(
            self.get_find_text(),
            self.get_replace_text(),
            self.replace_match_case_checked(),
        )

    def get_find_text(self) -> str:
        if self.find_text.text():
            return self.find_text.text()
        return self.replace_find_text.text()

    def get_replace_text(self) -> str:
        return self.replace_text.text()

    def match_case_checked(self) -> bool:
        return self.match_case.isChecked()

    def replace_match_case_checked(self) -> bool:
        return self.replace_match_case.isChecked()

    find_next = Signal(str, bool)
    replace = Signal(str, str, bool)
    replace_all = Signal(str, str, bool)


class FormatCellsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("Formato de Celdas")
        self.setModal(True)
        self.setMinimumWidth(400)

        layout = QVBoxLayout(self)

        tab_widget = QTabWidget()

        tab_widget.addTab(self._create_number_tab(), "Número")
        tab_widget.addTab(self._create_alignment_tab(), "Alineación")
        tab_widget.addTab(self._create_font_tab(), "Fuente")
        tab_widget.addTab(self._create_border_tab(), "Bordes")
        tab_widget.addTab(self._create_fill_tab(), "Relleno")

        layout.addWidget(tab_widget)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(buttons)

    def _create_number_tab(self):
        widget = QWidget()
        layout = QFormLayout(widget)

        self.number_format_combo = QComboBox()
        self.number_format_combo.addItems(
            [
                "General",
                "Número",
                "Moneda",
                "Contabilidad",
                "Fecha",
                "Hora",
                "Porcentaje",
                "Fracción",
                "Científica",
                "Texto",
            ]
        )
        layout.addRow("Categoría:", self.number_format_combo)

        self.decimal_places = QSpinBox()
        self.decimal_places.setRange(0, 10)
        self.decimal_places.setValue(2)
        layout.addRow("Decimales:", self.decimal_places)

        return widget

    def _create_alignment_tab(self):
        widget = QWidget()
        layout = QFormLayout(widget)

        self.horizontal_align = QComboBox()
        self.horizontal_align.addItems(
            ["General", "Izquierda", "Centro", "Derecha", "Relleno"]
        )
        layout.addRow("Horizontal:", self.horizontal_align)

        self.vertical_align = QComboBox()
        self.vertical_align.addItems(["Superior", "Centro", "Inferior"])
        layout.addRow("Vertical:", self.vertical_align)

        self.wrap_text = QCheckBox("Ajuste de texto")
        layout.addRow("", self.wrap_text)

        return widget

    def _create_font_tab(self):
        widget = QWidget()
        layout = QFormLayout(widget)

        self.font_combo = QFontComboBox()
        layout.addRow("Fuente:", self.font_combo)

        self.font_size = QSpinBox()
        self.font_size.setRange(6, 72)
        self.font_size.setValue(10)
        layout.addRow("Tamaño:", self.font_size)

        self.bold_check = QCheckBox("Negrita")
        layout.addRow("", self.bold_check)

        self.italic_check = QCheckBox("Cursiva")
        layout.addRow("", self.italic_check)

        self.underline_check = QCheckBox("Subrayado")
        layout.addRow("", self.underline_check)

        self.font_color_btn = QPushButton("Elegir color...")
        layout.addRow("Color:", self.font_color_btn)

        return widget

    def _create_border_tab(self):
        widget = QWidget()
        layout = QFormLayout(widget)

        self.border_top = QCheckBox("Superior")
        layout.addRow("", self.border_top)

        self.border_bottom = QCheckBox("Inferior")
        layout.addRow("", self.border_bottom)

        self.border_left = QCheckBox("Izquierdo")
        layout.addRow("", self.border_left)

        self.border_right = QCheckBox("Derecho")
        layout.addRow("", self.border_right)

        return widget

    def _create_fill_tab(self):
        widget = QWidget()
        layout = QFormLayout(widget)

        self.fill_color_btn = QPushButton("Elegir color...")
        layout.addRow("Color de relleno:", self.fill_color_btn)

        return widget


class InsertChartDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("Insertar Gráfico")
        self.setModal(True)
        self.setMinimumWidth(500)
        self.setMinimumHeight(400)

        layout = QVBoxLayout(self)

        chart_types_group = QGroupBox("Tipo de gráfico")
        chart_layout = QHBoxLayout(chart_types_group)

        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems(
            [
                "Barras",
                "Barras apiladas",
                "Líneas",
                "Líneas apiladas",
                "Circular",
                "Dispersión",
                "Área",
                "Área apilada",
            ]
        )
        chart_layout.addWidget(self.chart_type_combo)

        layout.addWidget(chart_types_group)

        data_group = QGroupBox("Rango de datos")
        data_layout = QFormLayout(data_group)

        self.data_range = QLineEdit()
        self.data_range.setPlaceholderText("Ej: A1:B10")
        data_layout.addRow("Rango:", self.data_range)

        self.labels_range = QLineEdit()
        self.labels_range.setPlaceholderText("Ej: A1:A10 (opcional)")
        data_layout.addRow("Etiquetas:", self.labels_range)

        layout.addWidget(data_group)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(buttons)

    def get_chart_type(self) -> str:
        return self.chart_type_combo.currentText()

    def get_data_range(self) -> str:
        return self.data_range.text()

    def get_labels_range(self) -> str:
        return self.labels_range.text()


class SortDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("Ordenar")
        self.setModal(True)

        layout = QVBoxLayout(self)

        form_layout = QFormLayout()

        self.column_combo = QComboBox()
        form_layout.addRow("Columna:", self.column_combo)

        self.order_combo = QComboBox()
        self.order_combo.addItems(["Ascendente", "Descendente"])
        form_layout.addRow("Orden:", self.order_combo)

        layout.addLayout(form_layout)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(buttons)

    def set_columns(self, columns: list):
        self.column_combo.clear()
        for col in columns:
            self.column_combo.addItem(col)

    def get_column_index(self) -> int:
        return self.column_combo.currentIndex()

    def get_ascending(self) -> bool:
        return self.order_combo.currentIndex() == 0


class FilterDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("Filtrar")
        self.setModal(True)

        layout = QVBoxLayout(self)

        form_layout = QFormLayout()

        self.column_combo = QComboBox()
        form_layout.addRow("Columna:", self.column_combo)

        self.condition_combo = QComboBox()
        self.condition_combo.addItems(
            [
                "Es igual a",
                "Es diferente de",
                "Es mayor que",
                "Es menor que",
                "Contiene",
                "No contiene",
            ]
        )
        form_layout.addRow("Condición:", self.condition_combo)

        self.value_edit = QLineEdit()
        form_layout.addRow("Valor:", self.value_edit)

        layout.addLayout(form_layout)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(buttons)

    def set_columns(self, columns: list):
        self.column_combo.clear()
        for col in columns:
            self.column_combo.addItem(col)

    def get_filter_config(self) -> dict:
        return {
            "column": self.column_combo.currentIndex(),
            "condition": self.condition_combo.currentIndex(),
            "value": self.value_edit.text(),
        }
