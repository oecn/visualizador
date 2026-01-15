"""Tab de evolucion mensual por proveedor."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter

from ..database import SalesRepository


class ProviderMonthlyTrendTab(QWidget):
    """Evolucion mensual por proveedor con grafico."""

    def __init__(self, repository: SalesRepository) -> None:
        super().__init__()
        self.repository = repository
        self.provider_combo = QComboBox()
        self.year_combo = QComboBox()
        self.metric_combo = QComboBox()
        self.chart_mode_combo = QComboBox()
        self.table = QTableWidget()
        self.chart = FigureCanvas(Figure(figsize=(5, 3)))
        self.summary_label = QLabel("Sin datos.")
        self._build_ui()
        self._connect_signals()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Proveedor:"))
        filter_layout.addWidget(self.provider_combo, 2)
        filter_layout.addWidget(QLabel("Ano:"))
        filter_layout.addWidget(self.year_combo, 1)
        self.metric_combo.addItems(["VENTAS", "UNIDADES"])
        filter_layout.addWidget(QLabel("Metrica:"))
        filter_layout.addWidget(self.metric_combo)
        self.chart_mode_combo.addItems(["Lineas", "Barras"])
        filter_layout.addWidget(QLabel("Grafico:"))
        filter_layout.addWidget(self.chart_mode_combo)
        layout.addLayout(filter_layout)

        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Mes", "Unidades", "Ventas (Gs.)"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        layout.addWidget(self.table, 1)

        layout.addWidget(self.chart, 2)
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

    def _connect_signals(self) -> None:
        self.provider_combo.currentIndexChanged.connect(lambda *_: self.load_monthly())
        self.year_combo.currentIndexChanged.connect(lambda *_: self.load_monthly())
        self.metric_combo.currentIndexChanged.connect(lambda *_: self.load_monthly())
        self.chart_mode_combo.currentIndexChanged.connect(lambda *_: self.load_monthly())

    def refresh(self) -> None:
        providers = self.repository.list_providers()
        current_provider = self.provider_combo.currentData()
        self.provider_combo.blockSignals(True)
        self.provider_combo.clear()
        self.provider_combo.addItem("Todos", None)
        for provider in providers:
            self.provider_combo.addItem(provider, provider)
        self.provider_combo.blockSignals(False)
        if current_provider:
            idx = self.provider_combo.findData(current_provider)
            if idx >= 0:
                self.provider_combo.setCurrentIndex(idx)

        years = self.repository.list_years()
        current_year = self.year_combo.currentData()
        self.year_combo.blockSignals(True)
        self.year_combo.clear()
        self.year_combo.addItem("Todos", None)
        for year in years:
            self.year_combo.addItem(str(year), year)
        self.year_combo.blockSignals(False)
        if current_year in years:
            idx = self.year_combo.findData(current_year)
            if idx >= 0:
                self.year_combo.setCurrentIndex(idx)

        self.load_monthly()

    def load_monthly(self) -> None:
        provider = self.provider_combo.currentData()
        year = self.year_combo.currentData()
        rows = self.repository.fetch_provider_monthly_totals(provider, year)
        if not rows:
            self._clear_view("No hay datos mensuales para el filtro actual.")
            return
        self.table.setRowCount(len(rows))
        months: list[str] = []
        amounts: list[float] = []
        quantities: list[float] = []
        for idx, row in enumerate(rows):
            month_key = str(row["month_key"])
            qty = float(row["total_quantity"] or 0.0)
            amt = float(row["total_amount"] or 0.0)
            months.append(month_key)
            amounts.append(amt)
            quantities.append(qty)
            self.table.setItem(idx, 0, QTableWidgetItem(month_key))
            self.table.setItem(idx, 1, QTableWidgetItem(self._format_number(qty)))
            self.table.setItem(idx, 2, QTableWidgetItem(self._format_currency(amt)))

        self._plot_chart(months, amounts, quantities)
        total_amt = sum(amounts)
        total_qty = sum(quantities)
        provider_label = provider or "Todos"
        year_label = str(year) if year else "Todos"
        summary = (
            f"Proveedor: {provider_label} | Ano: {year_label} | Meses: {len(months)} "
            f"| Ventas {self._format_currency(total_amt)} | Unidades {self._format_number(total_qty)}"
        )
        self.summary_label.setText(summary)

    def _plot_chart(self, months: list[str], amounts: list[float], quantities: list[float]) -> None:
        figure = self.chart.figure
        figure.clear()
        ax = figure.add_subplot(111)
        metric = self.metric_combo.currentText().strip().upper()
        chart_mode = self.chart_mode_combo.currentText().strip().upper()
        is_sales = metric == "VENTAS"
        data = amounts if is_sales else quantities
        label = "Ventas (Gs.)" if is_sales else "Unidades"
        color = "#1a9c47" if is_sales else "#0f3057"
        x_positions = list(range(len(months)))
        if chart_mode == "BARRAS":
            ax.bar(x_positions, data, color=color, alpha=0.8)
        else:
            ax.plot(x_positions, data, marker="o", color=color, linewidth=2)
        ax.set_xticks(x_positions, months, rotation=45, ha="right")
        ax.set_ylabel(label)
        ax.set_xlabel("Mes")
        ax.set_title("Evolucion mensual por proveedor")
        ax.yaxis.set_major_formatter(FuncFormatter(lambda val, _: self._format_short(val)))
        ax.grid(True, linestyle="--", linewidth=0.6, alpha=0.35)
        figure.tight_layout()
        self.chart.draw_idle()

    def _clear_view(self, message: str) -> None:
        self.table.setRowCount(0)
        fig = self.chart.figure
        fig.clear()
        fig.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        self.chart.draw_idle()
        self.summary_label.setText(message)

    def _format_currency(self, value: float) -> str:
        return f"Gs. {value:,.0f}".replace(",", ".")

    def _format_number(self, value: float) -> str:
        if value.is_integer():
            return f"{int(value):,}".replace(",", ".")
        return f"{value:,.2f}".replace(",", ".")

    def _format_short(self, value: float) -> str:
        abs_val = abs(value)
        suffix = ""
        divisor = 1.0
        if abs_val >= 1_000_000:
            suffix = "M"
            divisor = 1_000_000
        elif abs_val >= 1_000:
            suffix = "K"
            divisor = 1_000
        short = value / divisor
        if divisor == 1:
            text = f"{int(short)}"
        elif abs(short) >= 100:
            text = f"{short:,.0f}"
        elif abs(short) >= 10:
            text = f"{short:,.1f}"
        else:
            text = f"{short:,.2f}"
        text = text.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{text}{suffix}"
