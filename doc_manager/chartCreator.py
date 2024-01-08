from io import BytesIO

import matplotlib.pyplot as plt


class ChartCreator:
    def __init__(self, percent_open, percent_close):
        self.percent_open = percent_open
        self.percent_close = percent_close
        self.sizes = self.size_formatter()
        self.colors = ["green", "red"]

    def size_formatter(self):
        open_str = self.percent_open[:-1]
        close_str = self.percent_close[:-1]
        open_size_float = float(open_str)
        close_size_float = float(close_str)
        return [open_size_float, close_size_float]

    def generate_chart(self):
        # Create a 2D pie chart with only 2 colors and percentages
        fig, ax = plt.subplots(figsize=(3, 3))
        ax.pie(
            self.sizes,
            colors=self.colors,
            autopct="%1.1f%%",
            startangle=90,
            textprops={"fontsize": 20,"fontweight": "bold", "color": "white"},
        )
        fig.patch.set_facecolor("none")
        ax.set_facecolor("none")
        plt.tight_layout()
        buffer = BytesIO()
        plt.savefig(
            buffer,
            format="png",
            bbox_inches="tight",
            pad_inches=0.0,
            transparent=True,
        )
        plt.close()
        return buffer
