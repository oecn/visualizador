"""Visualizador historico de flujo (Debe/Haber) acumulando todos los Excel de `a/`.

Cada archivo de la carpeta representa un mes; esta app (PySide6 + Matplotlib)
los combina para construir una linea de tiempo completa de ingresos y egresos.
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Tuple

import pandas as pd
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT
from matplotlib.figure import Figure
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = BASE_DIR / "a"
DATE_COLUMN = "FECHAMOVI"
VIEW_MODES = {
    "Diario": None,
    "Mensual": "M",
    "Trimestral": "Q",
    "Anual": "A",
}


def parse_number(value) -> float:
    """Normaliza numeros con separadores de miles y decimales."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return 0.0
    normalized = text.replace(".", "").replace(",", ".")
    try:
        return float(normalized)
    except ValueError as exc:
        raise ValueError(f"No se pudo convertir el valor numerico: {value!r}") from exc


def human_size(num_bytes: int) -> str:
    """Da formato compacto a tamanos de archivo."""
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(num_bytes)
    for unit in units:
        if size < 1024.0:
            return f"{size:0.1f} {unit}"
        size /= 1024.0
    return f"{size:0.1f} PB"


def list_extracts(path: Path) -> List[Path]:
    """Devuelve los Excel (.xls/.xlsx) ordenados por fecha descendente."""
    return sorted(path.glob("*.xls*"), key=lambda file: file.stat().st_mtime, reverse=True)


def prepare_dataframe(excel_path: Path) -> pd.DataFrame:
    """Convierte un extracto individual en un DataFrame normalizado."""
    df = pd.read_excel(excel_path)
    if DATE_COLUMN not in df:
        raise KeyError(
            f"El archivo {excel_path.name} no contiene la columna {DATE_COLUMN!r}"
        )

    df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], errors="coerce")
    df["DEBE_NUM"] = df["DEBE"].apply(parse_number)
    df["HABER_NUM"] = df["HABER"].apply(parse_number)

    df = df.dropna(subset=[DATE_COLUMN])
    df["FECHA_SOLO_DIA"] = df[DATE_COLUMN].dt.date
    return df[["FECHA_SOLO_DIA", "DEBE_NUM", "HABER_NUM"]]


def build_daily_summary(directory: Path) -> Tuple[pd.DataFrame, List[Path]]:
    """Carga todos los Excel de la carpeta y los consolida por dia."""
    files = list_extracts(directory)
    if not files:
        raise FileNotFoundError(f"No se encontraron archivos .xls/.xlsx dentro de {directory}")

    frames = []
    for file in files:
        frames.append(prepare_dataframe(file))

    combined = pd.concat(frames, ignore_index=True).sort_values("FECHA_SOLO_DIA")
    daily = (
        combined.groupby("FECHA_SOLO_DIA")[["DEBE_NUM", "HABER_NUM"]]
        .sum()
        .rename(columns={"DEBE_NUM": "Debe", "HABER_NUM": "Haber"})
    )
    daily["Flujo Neto"] = daily["Haber"] - daily["Debe"]
    daily["Flujo Acumulado"] = daily["Flujo Neto"].cumsum()
    daily.index = pd.to_datetime(daily.index)
    return daily, files


def format_currency(value: float) -> str:
    """Formatea montos con separador latino."""
    formatted = f"{value:,.2f}"
    return formatted.replace(",", "_").replace(".", ",").replace("_", ".")


def format_compact(value: float) -> str:
    """Formatea montos en estilo compacto (1.2M, 450k) sin prefijos."""
    if value == 0:
        return "0"
    prefix = "-" if value < 0 else ""
    abs_value = abs(value)
    if abs_value >= 1_000_000_000:
        return f"{prefix}{abs_value/1_000_000_000:.1f}B"
    if abs_value >= 1_000_000:
        return f"{prefix}{abs_value/1_000_000:.1f}M"
    if abs_value >= 1_000:
        return f"{prefix}{abs_value/1_000:.1f}k"
    return f"{prefix}{abs_value:.0f}"


class FlowWindow(QMainWindow):
    """Ventana principal de la aplicacion."""

    def __init__(self, data_dir: Path):
        super().__init__()
        self.setWindowTitle("Flujo historico (Debe/Haber)")
        self.resize(1180, 720)

        self.data_dir = data_dir
        self.daily: pd.DataFrame | None = None
        self.files: List[Path] = []
        self.view_mode = "Mensual"
        self.dark_mode = True

        central = QWidget()
        layout = QVBoxLayout(central)
        self.setCentralWidget(central)

        header_layout = QHBoxLayout()
        self.dir_label = QLabel(f"Carpeta: {self.data_dir}")
        self.dir_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        header_layout.addWidget(self.dir_label, stretch=1)

        self.view_selector = QComboBox()
        self.view_selector.addItems(list(VIEW_MODES.keys()))
        self.view_selector.setCurrentText(self.view_mode)
        self.view_selector.currentTextChanged.connect(self.change_view_mode)
        header_layout.addWidget(QLabel("Vista:"))
        header_layout.addWidget(self.view_selector)

        self.dark_button = QPushButton("Modo claro" if self.dark_mode else "Modo oscuro")
        self.dark_button.clicked.connect(self.toggle_dark_mode)
        header_layout.addWidget(self.dark_button)

        self.refresh_button = QPushButton("Actualizar")
        self.refresh_button.clicked.connect(self.refresh_data)
        header_layout.addWidget(self.refresh_button)

        self.change_dir_button = QPushButton("Cambiar carpeta")
        self.change_dir_button.clicked.connect(self.change_directory)
        header_layout.addWidget(self.change_dir_button)

        layout.addLayout(header_layout)

        self.summary_label = QLabel("Carga inicial pendiente.")
        self.summary_label.setWordWrap(True)
        self.summary_label.setTextFormat(Qt.RichText)
        layout.addWidget(self.summary_label)
        self.files_list = QListWidget()
        self.files_list.setAlternatingRowColors(True)
        self.files_list.setMinimumHeight(130)
        layout.addWidget(self.files_list)

        self.canvas = FigureCanvas(Figure(figsize=(10, 5)))
        layout.addWidget(self.canvas, stretch=1)
        self.toolbar = NavigationToolbar2QT(self.canvas, self)
        layout.addWidget(self.toolbar)

        self.apply_palette()
        self.update_summary_style()
        self.refresh_data()

    # --- UI helpers -----------------------------------------------------
    def change_directory(self) -> None:
        """Permite elegir otra carpeta que contenga los extractos."""
        new_dir = QFileDialog.getExistingDirectory(
            self,
            "Selecciona la carpeta con los extractos",
            str(self.data_dir),
        )
        if new_dir:
            self.data_dir = Path(new_dir).resolve()
            self.dir_label.setText(f"Carpeta: {self.data_dir}")
            self.refresh_data()

    def change_view_mode(self, mode: str) -> None:
        """Actualiza el modo de agregacion y redibuja."""
        if mode not in VIEW_MODES:
            return
        self.view_mode = mode
        if self.daily is not None:
            self.update_summary()
            self.plot_history()

    def toggle_dark_mode(self) -> None:
        """Alterna entre modo claro y oscuro."""
        self.dark_mode = not self.dark_mode
        self.apply_palette()
        self.update_summary_style()
        self.dark_button.setText("Modo claro" if self.dark_mode else "Modo oscuro")
        if self.daily is not None:
            self.plot_history()

    def apply_palette(self) -> None:
        """Configura la paleta global segun el modo."""
        app = QApplication.instance()
        if app is None:
            return

        if not self.dark_mode:
            app.setPalette(app.style().standardPalette())
            self.toolbar.setStyleSheet("")
            return

        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(30, 30, 30))
        palette.setColor(QPalette.WindowText, Qt.white)
        palette.setColor(QPalette.Base, QColor(25, 25, 25))
        palette.setColor(QPalette.AlternateBase, QColor(35, 35, 35))
        palette.setColor(QPalette.ToolTipBase, Qt.white)
        palette.setColor(QPalette.ToolTipText, Qt.white)
        palette.setColor(QPalette.Text, Qt.white)
        palette.setColor(QPalette.Button, QColor(45, 45, 45))
        palette.setColor(QPalette.ButtonText, Qt.white)
        palette.setColor(QPalette.Highlight, QColor(0, 120, 215))
        palette.setColor(QPalette.HighlightedText, Qt.black)
        app.setPalette(palette)
        self.toolbar.setStyleSheet(
            "QToolBar { background-color: #1b1f25; border: 1px solid #2f3540; } "
            "QToolButton { color: #f4f6fb; }"
        )

    def update_summary_style(self) -> None:
        """Ajusta el estilo del resumen para el modo actual."""
        if self.dark_mode:
            style = (
                "font-size:15px;font-weight:500;color:#f5f5f5;"
                "background:#1e1f26;border:1px solid #3a414a;padding:10px;"
                "border-radius:8px;"
            )
        else:
            style = (
                "font-size:15px;font-weight:500;color:#222;"
                "background:#fdfdfd;border:1px solid #e4e7eb;padding:10px;"
                "border-radius:8px;"
            )
        self.summary_label.setStyleSheet(style)
        if hasattr(self, "toolbar"):
            if not self.dark_mode:
                self.toolbar.setStyleSheet("")
            else:
                self.toolbar.setStyleSheet(
                    "QToolBar { background-color: #1b1f25; border: 1px solid #2f3540; } "
                    "QToolButton { color: #f4f6fb; }"
                )

    def refresh_data(self) -> None:
        """Recarga los Excel y actualiza la UI."""
        try:
            daily, files = build_daily_summary(self.data_dir)
        except Exception as exc:  # pragma: no cover - manejo centralizado
            self.show_error(str(exc))
            return

        if daily.empty:
            self.show_error("No hay registros con fechas validas para graficar.")
            return

        self.daily = daily
        self.files = files
        self.update_summary()
        self.update_files_list()
        self.plot_history()

    def show_error(self, message: str) -> None:
        QMessageBox.critical(self, "Problema al cargar datos", message)

    def current_data(self) -> pd.DataFrame:
        assert self.daily is not None
        if self.view_mode == "Diario":
            return self.daily
        rule = VIEW_MODES[self.view_mode]
        aggregated = self.daily.resample(rule).agg({"Debe": "sum", "Haber": "sum", "Flujo Neto": "sum"})
        aggregated["Flujo Acumulado"] = aggregated["Flujo Neto"].cumsum()
        aggregated = aggregated.dropna(how="all")
        return aggregated

    def update_summary(self) -> None:
        assert self.daily is not None
        data = self.current_data()
        total_debe = data["Debe"].sum()
        total_haber = data["Haber"].sum()
        neto = data["Flujo Neto"].sum()
        acumulado = data["Flujo Acumulado"].iloc[-1]
        period_start = data.index[0].date()
        period_end = data.index[-1].date()
        summary_html = (
            "<div>"
            f"<span style='color:#5bc0de;font-weight:600;'>Vista:</span> {self.view_mode} "
            f"· <span style='color:#795548;font-weight:600;'>Archivos:</span> {len(self.files)} meses "
            f"· <span style='color:#607d8b;font-weight:600;'>Periodo:</span> {period_start} → {period_end}"
            "<br>"
            f"<span style='color:#d9534f;font-weight:600;'>Debe:</span> {format_currency(total_debe)} "
            f"· <span style='color:#5cb85c;font-weight:600;'>Haber:</span> {format_currency(total_haber)} "
            f"· <span style='color:#0275d8;font-weight:600;'>Flujo neto:</span> {format_currency(neto)} "
            f"· <span style='color:#f0ad4e;font-weight:600;'>Acumulado:</span> {format_currency(acumulado)}"
            "</div>"
        )
        self.summary_label.setText(summary_html)
        self.update_summary_style()

    def update_files_list(self) -> None:
        self.files_list.clear()
        for file in self.files:
            stamp = datetime.fromtimestamp(file.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            size = human_size(file.stat().st_size)
            item = QListWidgetItem(f"{file.name} | {stamp} | {size}")
            self.files_list.addItem(item)

    def plot_history(self) -> None:
        if self.daily is None:
            return
        daily = self.current_data()

        fig = self.canvas.figure
        fig.clear()
        colors = self.chart_colors()
        fig.patch.set_facecolor(colors["figure_bg"])
        ax = fig.add_subplot(111, facecolor=colors["axes_bg"])

        idx = range(len(daily))
        if self.view_mode == "Trimestral":
            labels = [f"T{d.quarter} {d.year}" for d in daily.index]
        else:
            date_format = "%Y-%m-%d"
            if self.view_mode == "Mensual":
                date_format = "%Y-%m"
            elif self.view_mode == "Anual":
                date_format = "%Y"
            labels = [d.strftime(date_format) for d in daily.index]
        bar_width = 0.4

        ax.bar(
            [i - bar_width / 2 for i in idx],
            daily["Debe"],
            width=bar_width,
            label="Debe",
            color=colors["bar_debe"],
        )
        ax.bar(
            [i + bar_width / 2 for i in idx],
            daily["Haber"],
            width=bar_width,
            label="Haber",
            color=colors["bar_haber"],
        )
        for i, value in enumerate(daily["Debe"]):
            if value:
                ax.text(
                    i - bar_width / 2,
                    value,
                    format_compact(value),
                    ha="center",
                    va="bottom",
                    fontsize=8,
                    color=colors["bar_debe"],
                )
        for i, value in enumerate(daily["Haber"]):
            if value:
                ax.text(
                    i + bar_width / 2,
                    value,
                    format_compact(value),
                    ha="center",
                    va="bottom",
                    fontsize=8,
                    color=colors["bar_haber"],
                )

        ax.set_xticks(list(idx))
        ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=8, color=colors["text"])
        ax.set_ylabel("Monto", color=colors["text"])
        ax.set_title("Flujo historico combinado", color=colors["text"], fontsize=14, fontweight="bold")
        ax.tick_params(colors=colors["text"])
        for spine in ("bottom", "left"):
            ax.spines[spine].set_color(colors["spine"])
        ax.grid(color=colors["grid"], linewidth=0.7, alpha=0.6)

        ax2 = ax.twinx()
        ax2.spines["right"].set_color(colors["spine"])
        net_series = daily["Flujo Neto"]
        ax2.plot(
            idx,
            net_series,
            color=colors["line_neto"],
            marker="o",
            linewidth=2,
            label="Flujo Neto",
        )
        ax2.plot(
            idx,
            daily["Flujo Acumulado"],
            color=colors["line_acum"],
            marker="s",
            linewidth=2,
            label="Flujo Acumulado",
        )
        window = min(5, max(2, len(daily) // 3)) if len(daily) > 1 else 1
        trend = net_series.rolling(window=window, min_periods=1).mean()
        ax2.plot(
            idx,
            trend,
            color=colors["trend_line"],
            linestyle="--",
            linewidth=1.5,
            label="Tendencia Flujo Neto",
        )
        positive_idx = [i for i, v in enumerate(net_series) if v > 0]
        negative_idx = [i for i, v in enumerate(net_series) if v < 0]
        if positive_idx:
            ax2.scatter(
                positive_idx,
                [net_series.iloc[i] for i in positive_idx],
                marker="^",
                color=colors["pos_marker"],
                s=60,
                label="Racha +",
                zorder=4,
            )
        if negative_idx:
            ax2.scatter(
                negative_idx,
                [net_series.iloc[i] for i in negative_idx],
                marker="v",
                color=colors["neg_marker"],
                s=60,
                label="Racha -",
                zorder=4,
            )

        max_idx = net_series.idxmax()
        min_idx = net_series.idxmin()
        if not net_series.empty:
            idx_map = {date: i for i, date in enumerate(daily.index)}
            max_pos = idx_map[max_idx]
            min_pos = idx_map[min_idx]
            ax2.annotate(
                f"Max {format_compact(net_series.loc[max_idx])}",
                xy=(max_pos, net_series.loc[max_idx]),
                xytext=(0, 12),
                textcoords="offset points",
                ha="center",
                color=colors["text"],
                fontsize=9,
                bbox=dict(facecolor=colors["label_bg"], edgecolor="none", alpha=0.85, boxstyle="round,pad=0.2"),
            )
            ax2.annotate(
                f"Min {format_compact(net_series.loc[min_idx])}",
                xy=(min_pos, net_series.loc[min_idx]),
                xytext=(0, -18),
                textcoords="offset points",
                ha="center",
                color=colors["text"],
                fontsize=9,
                bbox=dict(facecolor=colors["label_bg"], edgecolor="none", alpha=0.85, boxstyle="round,pad=0.2"),
            )

        ax2.set_ylabel("Flujo neto", color=colors["text"])
        ax2.tick_params(colors=colors["text"])

        handles, labels = ax.get_legend_handles_labels()
        handles2, labels2 = ax2.get_legend_handles_labels()
        legend = ax.legend(handles + handles2, labels + labels2, loc="upper left", fontsize=9)
        legend.get_frame().set_facecolor(colors["legend_bg"])
        legend.get_frame().set_edgecolor(colors["spine"])
        for text in legend.get_texts():
            text.set_color(colors["text"])

        fig.tight_layout()
        self.canvas.draw_idle()

    def chart_colors(self) -> dict[str, str]:
        """Devuelve la paleta del grafico segun el modo."""
        if self.dark_mode:
            return {
                "figure_bg": "#121418",
                "axes_bg": "#1b1f25",
                "text": "#f4f6fb",
                "grid": "#2f3540",
                "spine": "#49505f",
                "legend_bg": "#1f232b",
                "bar_debe": "#ff6b6b",
                "bar_haber": "#4bd292",
                "line_neto": "#4ca8ff",
                "line_acum": "#ffd166",
                "trend_line": "#b084f8",
                "pos_marker": "#2fd573",
                "neg_marker": "#ff5c5c",
                "label_bg": "#2a2f3a",
            }
        return {
            "figure_bg": "#f7f9fc",
            "axes_bg": "#ffffff",
            "text": "#212529",
            "grid": "#dfe4ec",
            "spine": "#b6bcc6",
            "legend_bg": "#ffffff",
            "bar_debe": "#d9534f",
            "bar_haber": "#5cb85c",
            "line_neto": "#0275d8",
            "line_acum": "#f0ad4e",
            "trend_line": "#7857ce",
            "pos_marker": "#1f9d55",
            "neg_marker": "#d6336c",
            "label_bg": "#f2f4f8",
        }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Visualiza los extractos historicos combinando todos los Excel de una carpeta.",
    )
    parser.add_argument(
        "--carpeta",
        help="Carpeta donde se encuentran los extractos (por defecto ./a).",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    data_dir = Path(args.carpeta).expanduser().resolve() if args.carpeta else DEFAULT_DATA_DIR
    if not data_dir.exists():
        print(f"La carpeta {data_dir} no existe. Crea la carpeta o usa --carpeta para indicar otra.")
        sys.exit(1)

    app = QApplication(sys.argv)
    window = FlowWindow(data_dir)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
