from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from typing import List, Tuple, Optional


class ChartBuilder:
    def __init__(self):
        self.figure = None
        self.canvas = None
        self._setup_style()

    def _setup_style(self):
        sns.set_style("whitegrid")
        plt.rcParams["font.size"] = 10

    def create_canvas(self, width: float = 6, height: float = 4) -> FigureCanvasQTAgg:
        self.figure = Figure(figsize=(width, height), dpi=100)
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

        bars = ax.bar(labels, values, color=sns.color_palette("husl", len(labels)))

        ax.set_title(title)
        if x_label:
            ax.set_xlabel(x_label)
        if y_label:
            ax.set_ylabel(y_label)

        for bar in bars:
            height = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2.0,
                height,
                f"{height:.0f}",
                ha="center",
                va="bottom",
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
