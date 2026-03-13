"""
PYXCEL Theme System
Supports Dark and Light modes with dynamic switching
"""

from PySide6.QtGui import QColor, QFont, QPalette


class DarkColors:
    # Base colors (Mocha)
    CRUST = "#1E1E2E"
    MANTLE = "#181825"
    BASE = "#11111B"
    SURFACE0 = "#313244"
    SURFACE1 = "#45475A"
    SURFACE2 = "#585B70"
    BLUE = "#89B4FA"
    SAPPHIRE = "#74C7EC"
    GREEN = "#A6E3A1"
    RED = "#F38BA8"
    MAUVE = "#CBA6F7"
    LAVENDER = "#B4BEFE"
    TEXT = "#CDD6F4"
    SUBTEXT0 = "#A6ADC8"
    SUBTEXT1 = "#BAC2DE"
    OVERLAY0 = "#6C7086"
    OVERLAY1 = "#7F849C"
    OVERLAY2 = "#9399B2"
    YELLOW = "#F9E2AF"
    PEACH = "#FAB387"
    MAROON = "#EBA0AC"
    TEAL = "#94E2D5"
    SKY = "#89DCEB"
    GRID = "#45475A"
    HEADER_BG = "#181825"
    SELECTION = "#89B4FA"
    SELECTION_TEXT = "#1E1E2E"
    ROW_ALT = "#1E1E2E"
    ROW_DEFAULT = "#11111B"
    HOVER = "#313244"


class LightColors:
    # Base colors (Latte)
    CRUST = "#DCE0E8"
    MANTLE = "#E6E9EF"
    BASE = "#EFF1F5"
    SURFACE0 = "#CCD0DA"
    SURFACE1 = "#BCC0CC"
    SURFACE2 = "#ACB0BE"
    BLUE = "#1E66F5"
    SAPPHIRE = "#209FB5"
    GREEN = "#40A02B"
    RED = "#D20F39"
    MAUVE = "#8839EF"
    LAVENDER = "#7287FD"
    TEXT = "#4C4F69"
    SUBTEXT0 = "#5C5F77"
    SUBTEXT1 = "#6C6F85"
    OVERLAY0 = "#9CA0B0"
    OVERLAY1 = "#8C8FA1"
    OVERLAY2 = "#7C7F93"
    YELLOW = "#DF8E1D"
    PEACH = "#FE640B"
    MAROON = "#E64553"
    TEAL = "#179299"
    SKY = "#04A5E5"
    GRID = "#BCC0CC"
    HEADER_BG = "#E6E9EF"
    SELECTION = "#1E66F5"
    SELECTION_TEXT = "#EFF1F5"
    ROW_ALT = "#DCE0E8"
    ROW_DEFAULT = "#EFF1F5"
    HOVER = "#CCD0DA"


class ThemeManager:
    def __init__(self):
        self._mode = "dark"
        self.colors = DarkColors

    def set_mode(self, mode: str):
        self._mode = mode
        if mode == "light":
            self.colors = LightColors
        else:
            self.colors = DarkColors

    def get_mode(self) -> str:
        return self._mode


# Theme manager instance
theme_manager = ThemeManager()


class ThemeFonts:
    SPREADSHEET = "JetBrains Mono"
    MONO = "JetBrains Mono"
    UI = "Segoe UI"
    SIZES = {
        "tiny": 8, "small": 10, "normal": 11,
        "medium": 12, "large": 14, "title": 16, "header": 18,
    }


def get_app_stylesheet() -> str:
    C = theme_manager.colors
    return f"""
    QMainWindow {{ background-color: {C.CRUST}; color: {C.TEXT}; font-family: '{ThemeFonts.UI}'; font-size: {ThemeFonts.SIZES["normal"]}px; }}
    QWidget {{ background-color: {C.CRUST}; color: {C.TEXT}; }}
    QLabel {{ color: {C.TEXT}; background: transparent; }}
    QMenuBar {{ background-color: {C.MANTLE}; color: {C.TEXT}; border-bottom: 1px solid {C.SURFACE0}; padding: 2px; }}
    QMenuBar::item {{ background: transparent; padding: 6px 12px; border-radius: 4px; }}
    QMenuBar::item:selected {{ background-color: {C.SURFACE0}; }}
    QMenu {{ background-color: {C.MANTLE}; color: {C.TEXT}; border: 1px solid {C.SURFACE0}; border-radius: 8px; padding: 4px; }}
    QMenu::item:selected {{ background-color: {C.BLUE}; color: {C.CRUST}; }}
    QToolBar {{ background-color: {C.MANTLE}; border: none; padding: 4px 8px; spacing: 4px; }}
    QPushButton {{ background-color: {C.SURFACE0}; color: {C.TEXT}; border: 1px solid {C.SURFACE1}; border-radius: 6px; padding: 6px 14px; }}
    QPushButton:hover {{ background-color: {C.SURFACE1}; border-color: {C.BLUE}; }}
    QLineEdit, QTextEdit {{ background-color: {C.MANTLE}; color: {C.TEXT}; border: 1px solid {C.SURFACE0}; border-radius: 6px; padding: 8px 12px; }}
    QTabWidget::pane {{ background-color: {C.CRUST}; border: 1px solid {C.SURFACE0}; border-radius: 4px; }}
    QTabBar::tab {{ background-color: {C.MANTLE}; color: {C.SUBTEXT0}; padding: 10px 20px; border: 1px solid {C.SURFACE0}; border-bottom: none; border-top-left-radius: 4px; border-top-right-radius: 4px; margin-right: 2px; }}
    QTabBar::tab:selected {{ background-color: {C.CRUST}; color: {C.BLUE}; border-bottom: 2px solid {C.BLUE}; font-weight: bold; }}
    QTabBar::tab:hover {{ background-color: {C.SURFACE0}; }}
    QTableView {{ background-color: {C.BASE}; color: {C.TEXT}; gridline-color: {C.GRID}; }}
    QHeaderView::section {{ background-color: {C.HEADER_BG}; color: {C.SUBTEXT0}; padding: 8px 12px; border: 1px solid {C.SURFACE0}; border-bottom: 2px solid {C.BLUE}; }}
    QStatusBar {{ background-color: {C.MANTLE}; color: {C.SUBTEXT0}; border-top: 1px solid {C.SURFACE0}; padding: 4px 12px; }}
    QSplitter::handle {{ background-color: {C.SURFACE0}; }}
    QSplitter::handle:horizontal {{ width: 2px; }}
    QSplitter::handle:vertical {{ height: 2px; }}
    """


def get_palette() -> QPalette:
    C = theme_manager.colors
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(C.CRUST))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(C.TEXT))
    palette.setColor(QPalette.ColorRole.Base, QColor(C.MANTLE))
    palette.setColor(QPalette.ColorRole.Text, QColor(C.TEXT))
    palette.setColor(QPalette.ColorRole.Button, QColor(C.SURFACE0))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(C.TEXT))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(C.BLUE))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(C.CRUST))
    return palette


def get_terminal_style() -> str:
    C = theme_manager.colors
    return f"""
    QTextEdit {{
        background-color: {C.BASE};
        color: {C.GREEN if theme_manager.get_mode() == "dark" else "#1A5E1A"};
        border: none;
        font-family: '{ThemeFonts.MONO}', monospace;
        font-size: {ThemeFonts.SIZES["normal"]}px;
        line-height: 1.4;
    }}
    """

class ThemeColorsProxy:
    def __getattr__(self, name):
        return getattr(theme_manager.colors, name)

# Global proxy for backward compatibility and dynamic lookup
ThemeColors = ThemeColorsProxy()


def get_accent_button_style() -> str:
    C = theme_manager.colors
    return f"""
    QPushButton {{ background-color: {C.BLUE}; color: {C.CRUST if hasattr(C, 'CRUST') else '#FFFFFF'}; border: none; border-radius: 6px; padding: 6px 12px; font-weight: 600; }}
    QPushButton:hover {{ background-color: {C.SAPPHIRE}; }}
    """


def get_icon_button_style() -> str:
    C = theme_manager.colors
    return f"""
    QPushButton {{ background-color: transparent; color: {C.TEXT}; border: none; border-radius: 6px; padding: 6px 8px; min-width: 28px; min-height: 28px; }}
    QPushButton:hover {{ background-color: {C.SURFACE0}; }}
    """

def get_cell_editor_style() -> str:
    C = theme_manager.colors
    return f"""
    QLabel {{ background-color: {C.SURFACE0}; color: {C.SUBTEXT0}; padding: 6px 10px; font-family: '{ThemeFonts.SPREADSHEET}'; }}
    QLineEdit {{ background-color: {C.BASE}; color: {C.TEXT}; border: 2px solid {C.BLUE}; padding: 6px 10px; font-family: '{ThemeFonts.SPREADSHEET}'; }}
    """
