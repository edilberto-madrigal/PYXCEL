from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from typing import List, Tuple, Optional


class ChartBuilder:
    # Catppuccin color palette for charts
    COLORS = [
        "#89B4FA",  # Blue
        "#A6E3A1",  # Green
        "#F9E2AF",  # Yellow
        "#FAB387",  # Peach
        "#CBA6F7",  # Mauve
        "#94E2D5",  # Teal
        "#F38BA8",  # Red
        "#74C7EC",  # Sapphire
    ]

    def __init__(self):
        self.figure = None
        self.canvas = None
        self._setup_style()

    def _setup_style(self):
        sns.set_style("darkgrid")
        plt.rcParams["font.size"] = 10
        plt.rcParams["text.color"] = "#CDD6F4"
        plt.rcParams["axes.labelcolor"] = "#CDD6F4"
        plt.rcParams["xtick.color"] = "#CDD6F4"
        plt.rcParams["ytick.color"] = "#CDD6F4"
        plt.rcParams["axes.edgecolor"] = "#45475A"
        plt.rcParams["axes.facecolor"] = "#1E1E2E"
        plt.rcParams["figure.facecolor"] = "#1E1E2E"

    def create_canvas(self, width: float = 6, height: float = 4) -> FigureCanvasQTAgg:
        self.figure = Figure(figsize=(width, height), dpi=100)
        self.figure.patch.set_facecolor("#1E1E2E")
        self.canvas = FigureCanvasQTAgg(self.figure)
        return self.canvas

    def create_bar_chart(
        self,
        data: dict,
        title: str = "Gráfico de Barras",
        x_label: str = "",
        y_label: str = "",
    ):
        if not self.figure:
            self.create_canvas()

        ax = self.figure.add_subplot(111)

        labels = list(data.keys())
        values = list(data.values())

        colors = self.COLORS[: len(labels)]
        bars = ax.bar(labels, values, color=colors)

        ax.set_title(title, color="#CDD6F4")
        ax.set_xlabel(x_label, color="#A6ADC8") if x_label else None
        ax.set_ylabel(y_label, color="#A6ADC8") if y_label else None
        ax.tick_params(colors="#CDD6F4")

        for bar in bars:
            height = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2.0,
                height,
                f"{height:.0f}",
                ha="center",
                va="bottom",
                color="#CDD6F4",
            )

        self.figure.tight_layout()
        return self.canvas

    def create_line_chart(
        self,
        data: dict,
        title: str = "Gráfico de Líneas",
        x_label: str = "",
        y_label: str = "",
    ):
        if not self.figure:
            self.create_canvas()

        ax = self.figure.add_subplot(111)

        labels = list(data.keys())
        values = list(data.values())

        ax.plot(
            labels,
            values,
            marker="o",
            linewidth=2,
            markersize=6,
            color=sns.color_palette("husl", 1)[0],
        )

        ax.set_title(title)
        if x_label:
            ax.set_xlabel(x_label)
        if y_label:
            ax.set_ylabel(y_label)

        self.figure.tight_layout()
        return self.canvas

    def create_pie_chart(self, data: dict, title: str = "Gráfico Circular"):
        if not self.figure:
            self.create_canvas()

        ax = self.figure.add_subplot(111)

        labels = list(data.keys())
        values = list(data.values())

        colors = sns.color_palette("husl", len(labels))

        wedges, texts, autotexts = ax.pie(
            values, labels=labels, colors=colors, autopct="%1.1f%%", startangle=90
        )

        for text in texts:
            text.set_fontsize(9)
        for autotext in autotexts:
            autotext.set_color("white")
            autotext.set_fontsize(9)
            autotext.set_weight("bold")

        ax.set_title(title)

        self.figure.tight_layout()
        return self.canvas

    def create_scatter_chart(
        self,
        x_data: List,
        y_data: List,
        title: str = "Gráfico de Dispersión",
        x_label: str = "",
        y_label: str = "",
    ):
        if not self.figure:
            self.create_canvas()

        ax = self.figure.add_subplot(111)

        ax.scatter(
            x_data,
            y_data,
            c=sns.color_palette("husl", 1)[0],
            s=100,
            alpha=0.6,
            edgecolors="black",
            linewidth=0.5,
        )

        if len(x_data) > 1:
            z = np.polyfit(x_data, y_data, 1)
            p = np.poly1d(z)
            ax.plot(x_data, p(x_data), "r--", alpha=0.8, linewidth=2)

        ax.set_title(title)
        if x_label:
            ax.set_xlabel(x_label)
        if y_label:
            ax.set_ylabel(y_label)

        self.figure.tight_layout()
        return self.canvas

    def create_area_chart(
        self,
        data: dict,
        title: str = "Gráfico de Área",
        x_label: str = "",
        y_label: str = "",
    ):
        if not self.figure:
            self.create_canvas()

        ax = self.figure.add_subplot(111)

        labels = list(data.keys())
        values = list(data.values())

        ax.fill_between(
            labels, values, alpha=0.4, color=sns.color_palette("husl", 1)[0]
        )
        ax.plot(labels, values, color=sns.color_palette("husl", 1)[0], linewidth=2)

        ax.set_title(title)
        if x_label:
            ax.set_xlabel(x_label)
        if y_label:
            ax.set_ylabel(y_label)

        self.figure.tight_layout()
        return self.canvas

    def create_stacked_bar_chart(
        self,
        data: dict,
        title: str = "Barras Apiladas",
        x_label: str = "",
        y_label: str = "",
    ):
        if not self.figure:
            self.create_canvas()

        ax = self.figure.add_subplot(111)

        categories = list(data.keys())
        datasets = {}

        for category in categories:
            for key, value in data[category].items():
                if key not in datasets:
                    datasets[key] = [0] * len(categories)
                datasets[key][categories.index(category)] = value

        bottom = np.zeros(len(categories))
        colors = sns.color_palette("husl", len(datasets))

        for i, (key, values) in enumerate(datasets.items()):
            ax.bar(categories, values, bottom=bottom, label=key, color=colors[i])
            bottom += values

        ax.set_title(title)
        if x_label:
            ax.set_xlabel(x_label)
        if y_label:
            ax.set_ylabel(y_label)
        ax.legend()

        self.figure.tight_layout()
        return self.canvas

    def create_stacked_line_chart(
        self,
        data: dict,
        title: str = "Líneas Apiladas",
        x_label: str = "",
        y_label: str = "",
    ):
        if not self.figure:
            self.create_canvas()

        ax = self.figure.add_subplot(111)

        categories = list(data.keys())

        for i, (key, value) in enumerate(data.items()):
            ax.plot(
                categories,
                list(value.values()) if isinstance(value, dict) else [value],
                marker="o",
                linewidth=2,
                label=key,
                color=self.COLORS[i % len(self.COLORS)],
            )

        ax.set_title(title, color="#CDD6F4")
        if x_label:
            ax.set_xlabel(x_label, color="#A6ADC8")
        if y_label:
            ax.set_ylabel(y_label, color="#A6ADC8")
        ax.tick_params(colors="#CDD6F4")
        ax.legend(labelcolor="#CDD6F4")

        self.figure.tight_layout()
        return self.canvas

    def create_stacked_area_chart(
        self,
        data: dict,
        title: str = "Área Apilada",
        x_label: str = "",
        y_label: str = "",
    ):
        if not self.figure:
            self.create_canvas()

        ax = self.figure.add_subplot(111)

        categories = list(data.keys())

        y_data = []
        for key, value in data.items():
            y_data.append(list(value.values()) if isinstance(value, dict) else [value])

        ax.stackplot(
            categories,
            *y_data,
            labels=list(data.keys()),
            colors=self.COLORS[: len(data)],
        )

        ax.set_title(title, color="#CDD6F4")
        if x_label:
            ax.set_xlabel(x_label, color="#A6ADC8")
        if y_label:
            ax.set_ylabel(y_label, color="#A6ADC8")
        ax.tick_params(colors="#CDD6F4")
        ax.legend(labelcolor="#CDD6F4")

        self.figure.tight_layout()
        return self.canvas

    def clear(self):
        if self.figure:
            self.figure.clear()
            self.figure = None
            self.canvas = None

    def save_chart(self, file_path: str, format: str = "png"):
        if self.figure:
            self.figure.savefig(file_path, format=format, dpi=300, bbox_inches="tight")
            return True
        return False
