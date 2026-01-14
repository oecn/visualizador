from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush, QPalette
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QHeaderView,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter

from ..database import SalesRepository


@dataclass
class BranchTotals:
    branch: str
    qty: float
    amt: float


class CompararAyAADialog(QDialog):
    """
    Dialogo flotante con dos pestañas: Unidades y Montos.

    - Unidades: muestra unidades vendidas por sucursal, total y Δ vs mes anterior.
    - Montos: muestra Gs. por sucursal, Gs. del mismo mes del año anterior y Δ Gs. por sucursal,
      además de total y Δ vs mes anterior.
    """

    BRANCHES = ["Casa Central", "CAACUPEMI", "ITAUGUA"]

    def __init__(
        self,
        parent: QWidget,
        repository: SalesRepository,
        product_code: str | None = None,
        product_desc: str | None = None,
        product_label: str | None = None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.product_code = product_code
        self.product_desc = product_desc
        self.product_label = product_label or "Todos los productos"
        self.setWindowTitle("Comparar AyAA - Crecimiento mensual por sucursal")
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowCloseButtonHint
        )
        self.resize(1100, 640)
        self.setWindowState(Qt.WindowMaximized)

        layout = QVBoxLayout(self)
        header = QHBoxLayout()
        header.addWidget(QLabel(f"Comparativo mensual por sucursal (unidades / montos) · {self.product_label}"), 1)
        self.growth_btn = QPushButton("Ver gráfica de crecimiento")
        self.growth_btn.clicked.connect(self._open_growth_chart)
        header.addWidget(self.growth_btn)
        layout.addLayout(header)

        self.tabs = QTabWidget()
        self.units_table = QTableWidget()
        self.amount_table = QTableWidget()
        self.total_table = QTableWidget()
        self._setup_total_table()
        self._setup_units_table()
        self._setup_amount_table()
        self._apply_selection_style()
        self.tabs.addTab(self.total_table, "Total Gs. YoY")
        self.tabs.addTab(self.units_table, "Unidades")
        self.tabs.addTab(self.amount_table, "Montos (Gs.)")
        layout.addWidget(self.tabs)

        self.summary_label = QLabel("Sin datos para comparar.")
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        # Dataset cache para graficar crecimiento.
        self._last_rows: list[tuple[str, dict[str, BranchTotals]]] = []
        self._load_data()

    def _setup_units_table(self) -> None:
        headers = ["Mes"] + [f"{b} (Unidades)" for b in self.BRANCHES] + ["Total unidades", "Δ vs mes anterior"]
        self.units_table.setColumnCount(len(headers))
        self.units_table.setHorizontalHeaderLabels(headers)
        self.units_table.setAlternatingRowColors(True)
        self.units_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.units_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.units_table.horizontalHeader().setStretchLastSection(True)
        self._tune_header(self.units_table)
        self._apply_branch_header_colors(self.units_table, mode="units")

    def _setup_amount_table(self) -> None:
        headers = (
            ["Mes"]
            + [item for b in self.BRANCHES for item in (f"{b} (Gs.)", f"{b} (Gs. año-1)", f"Δ {b} (Gs.)")]
            + ["Total Gs."]
        )
        self.amount_table.setColumnCount(len(headers))
        self.amount_table.setHorizontalHeaderLabels(headers)
        self.amount_table.setAlternatingRowColors(True)
        self.amount_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.amount_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.amount_table.horizontalHeader().setStretchLastSection(True)
        self._tune_header(self.amount_table, amount_mode=True)
        self._apply_branch_header_colors(self.amount_table, mode="amounts")

    def _load_data(self) -> None:
        months = self.repository.list_months()
        if not months:
            self.units_table.setRowCount(0)
            self.amount_table.setRowCount(0)
            self.summary_label.setText("No hay meses importados.")
            return
        months = list(months)  # mantener el ultimo mes primero
        rows: list[tuple[str, dict[str, BranchTotals]]] = []
        for row in months:
            month_key = str(row["month_key"])
            if self.product_code or self.product_desc:
                branch_rows = self.repository.fetch_monthly_product_branch_totals(month_key, self.product_code, self.product_desc)
            else:
                branch_rows = self.repository.fetch_monthly_branch_totals(month_key)
            branch_map: dict[str, BranchTotals] = {}
            for b_row in branch_rows:
                branch = b_row["branch"]
                branch_map[branch] = BranchTotals(
                    branch=branch,
                    qty=float(b_row["total_quantity"] or 0.0),
                    amt=float(b_row["total_amount"] or 0.0),
                )
            rows.append((month_key, branch_map))

        self._last_rows = rows
        self._fill_units_table(rows)
        self._fill_amount_table(rows)
        self._fill_total_table(rows)

        # Para el resumen, comparamos del mes más antiguo al más reciente.
        first = rows[-1][0]
        last = rows[0][0]
        growth = ""
        if len(rows) > 1:
            first_total = sum(b.amt for b in rows[-1][1].values())
            last_total = sum(b.amt for b in rows[0][1].values())
            diff = last_total - first_total
            pct = (diff / first_total * 100) if abs(first_total) > 1e-6 else 0.0
            growth = f"Crecimiento desde {first} a {last}: {self._format_currency(diff)} ({pct:+.1f}%)".replace(".", ",")
        self.summary_label.setText(
            f"Meses comparados: {len(rows)} · Sucursales: {len(self.BRANCHES)}. {growth}"
        )

    def _fill_units_table(self, rows: list[tuple[str, dict[str, BranchTotals]]]) -> None:
        self.units_table.setRowCount(len(rows))
        prev_total_qty: Optional[float] = None
        for idx, (month_key, branch_map) in enumerate(rows):
            self.units_table.setItem(idx, 0, QTableWidgetItem(month_key))
            total_qty = 0.0
            for b_idx, branch in enumerate(self.BRANCHES):
                data = branch_map.get(branch, BranchTotals(branch, 0.0, 0.0))
                total_qty += data.qty
                self.units_table.setItem(idx, 1 + b_idx, QTableWidgetItem(self._format_number(data.qty)))
            self.units_table.setItem(idx, 1 + len(self.BRANCHES), QTableWidgetItem(self._format_number(total_qty)))

            if prev_total_qty is None:
                delta_text = "N/A"
                color = None
            else:
                diff = total_qty - prev_total_qty
                pct = (diff / prev_total_qty * 100) if abs(prev_total_qty) > 1e-6 else 0.0
                delta_text = f"{self._format_number(diff)} ({pct:+.1f}%)".replace(".", ",")
                color = QColor("#1a9c47") if diff >= 0 else QColor("#d64541")
            delta_item = QTableWidgetItem(delta_text)
            if color:
                delta_item.setForeground(QBrush(color))
            self.units_table.setItem(idx, self.units_table.columnCount() - 1, delta_item)
            prev_total_qty = total_qty

    def _fill_amount_table(self, rows: list[tuple[str, dict[str, BranchTotals]]]) -> None:
        self.amount_table.setRowCount(len(rows))
        for idx, (month_key, branch_map) in enumerate(rows):
            self.amount_table.setItem(idx, 0, QTableWidgetItem(month_key))
            total_amt = 0.0
            prev_year_key = self._prev_year_key(month_key)
            prev_branch_rows = (
                {r["branch"]: r for r in self.repository.fetch_monthly_product_branch_totals(prev_year_key, self.product_code, self.product_desc)}
                if prev_year_key
                else {}
            )
            for b_idx, branch in enumerate(self.BRANCHES):
                data = branch_map.get(branch, BranchTotals(branch, 0.0, 0.0))
                total_amt += data.amt
                col_base = 1 + b_idx * 3
                self.amount_table.setItem(idx, col_base, QTableWidgetItem(self._format_currency(data.amt)))

                prev_data = prev_branch_rows.get(branch)
                prev_amt = float(prev_data["total_amount"] or 0.0) if prev_data else 0.0
                prev_item = QTableWidgetItem(self._format_currency(prev_amt))
                delta_item = self._delta_item(data.amt, prev_amt, is_currency=True)
                self.amount_table.setItem(idx, col_base + 1, prev_item)
                self.amount_table.setItem(idx, col_base + 2, delta_item)

            offset_after_yoy = 1 + len(self.BRANCHES) * 3
            self.amount_table.setItem(idx, offset_after_yoy, QTableWidgetItem(self._format_currency(total_amt)))

    def _setup_total_table(self) -> None:
        headers = ["Mes", "Total Gs.", "Total Gs. año-1", "Δ total YoY"]
        self.total_table.setColumnCount(len(headers))
        self.total_table.setHorizontalHeaderLabels(headers)
        self.total_table.setAlternatingRowColors(True)
        self.total_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.total_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.total_table.horizontalHeader().setStretchLastSection(True)
        self._tune_header(self.total_table, amount_mode=True)
        self._apply_branch_header_colors(self.total_table, mode="total")

    def _fill_total_table(self, rows: list[tuple[str, dict[str, BranchTotals]]]) -> None:
        self.total_table.setRowCount(len(rows))
        for idx, (month_key, branch_map) in enumerate(rows):
            self.total_table.setItem(idx, 0, QTableWidgetItem(month_key))
            total_amt = sum(b.amt for b in branch_map.values())
            self.total_table.setItem(idx, 1, QTableWidgetItem(self._format_currency(total_amt)))
            prev_key = self._prev_year_key(month_key)
            prev_total = 0.0
            if prev_key:
                if self.product_code or self.product_desc:
                    prev_rows = self.repository.fetch_monthly_product_branch_totals(prev_key, self.product_code, self.product_desc)
                else:
                    prev_rows = self.repository.fetch_monthly_branch_totals(prev_key)
                prev_total = sum(float(r["total_amount"] or 0.0) for r in prev_rows)
            self.total_table.setItem(idx, 2, QTableWidgetItem(self._format_currency(prev_total)))
            delta_item = self._delta_item(total_amt, prev_total, is_currency=True)
            self.total_table.setItem(idx, 3, delta_item)

    def _prev_year_key(self, month_key: str | None) -> Optional[str]:
        if not month_key or "-" not in month_key:
            return None
        try:
            year_str, month_str = month_key.split("-", 1)
            return f"{int(year_str) - 1:04d}-{month_str}"
        except Exception:
            return None

    def _tune_header(self, table: QTableWidget, amount_mode: bool = False) -> None:
        header = table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        if amount_mode:
            for idx in range(1, table.columnCount()):
                if idx == table.columnCount() - 1:
                    header.setSectionResizeMode(idx, QHeaderView.ResizeToContents)
                else:
                    header.setSectionResizeMode(idx, QHeaderView.Stretch)
        else:
            for idx in range(1, table.columnCount()):
                if idx >= table.columnCount() - 2:
                    header.setSectionResizeMode(idx, QHeaderView.ResizeToContents)
                else:
                    header.setSectionResizeMode(idx, QHeaderView.Stretch)

    def _format_currency(self, value: float) -> str:
        return f"Gs. {value:,.0f}".replace(",", ".")

    def _format_number(self, value: float) -> str:
        if value.is_integer():
            return f"{int(value):,}".replace(",", ".")
        return f"{value:,.2f}".replace(",", ".")

    def _delta_item(self, current: float, previous: float, *, is_currency: bool = False) -> QTableWidgetItem:
        diff = current - previous
        text = self._format_currency(diff) if is_currency else self._format_number(diff)
        pct = ""
        if abs(previous) > 1e-6:
            pct = f" ({(diff / previous) * 100:+.1f}%)".replace(".", ",")
        item = QTableWidgetItem(f"{text}{pct}")
        color = QColor("#1a9c47")
        if diff < 0:
            color = QColor("#d64541")
        item.setForeground(QBrush(color))
        return item

    def _apply_branch_header_colors(self, table: QTableWidget, mode: str) -> None:
        central_bg = QColor("#fff3cd")  # amarillo pastel
        central_fg = QColor("#7a5e00")
        caacu_bg = QColor("#f8d7da")  # rojo pastel
        caacu_fg = QColor("#7c1d1d")
        itau_bg = QColor("#dbeafe")  # azul pastel
        itau_fg = QColor("#0f3b78")
        mes_bg = QColor("#e0e0e0")
        mes_fg = QColor("#333333")
        total_bg = QColor("#e2f0d9")
        total_fg = QColor("#1c4d2f")

        def paint_cols(cols: list[int], bg: QColor, fg: QColor) -> None:
            for col in cols:
                item = table.horizontalHeaderItem(col)
                if item:
                    item.setBackground(QBrush(bg))
                    item.setForeground(QBrush(fg))

        # Mes
        paint_cols([0], mes_bg, mes_fg)

        if mode == "units":
            paint_cols([1], central_bg, central_fg)
            paint_cols([2], caacu_bg, caacu_fg)
            paint_cols([3], itau_bg, itau_fg)
            # Total unidades
            total_idx = 1 + len(self.BRANCHES)
            paint_cols([total_idx], total_bg, total_fg)
        elif mode == "amounts":
            paint_cols([1, 2, 3], central_bg, central_fg)
            paint_cols([4, 5, 6], caacu_bg, caacu_fg)
            paint_cols([7, 8, 9], itau_bg, itau_fg)
            total_idx = 1 + len(self.BRANCHES) * 3
            paint_cols([total_idx], total_bg, total_fg)
        elif mode == "total":
            # Mes, total actual, total año-1, delta total
            paint_cols([1, 2], total_bg, total_fg)

    def _open_growth_chart(self) -> None:
        if not self._last_rows:
            return
        # Construye series de totales mensuales
        month_keys: list[str] = []
        totals: list[float] = []
        for month_key, branch_map in self._last_rows:
            month_keys.append(month_key)
            totals.append(sum(b.amt for b in branch_map.values()))
        # Orden cronológico (ascendente) para la gráfica.
        ordered = sorted(zip(month_keys, totals), key=lambda item: item[0])
        month_keys = [m for m, _ in ordered]
        totals = [t for _, t in ordered]

        dialog = QDialog(self)
        dialog.setWindowTitle(f"Gráfico de crecimiento · {self.product_label}")
        dialog.resize(900, 520)
        vbox = QVBoxLayout(dialog)
        fig = Figure(figsize=(8, 4))
        canvas = FigureCanvas(fig)
        vbox.addWidget(canvas)

        ax = fig.add_subplot(111)
        x = list(range(len(month_keys)))
        ax.plot(x, totals, marker="o", linewidth=2, color="#1a73e8", label="Total mensual")

        # Rolling de tendencia trimestral aprox (ventana 3)
        window = 3
        if len(totals) >= 2:
            trend = []
            for i in range(len(totals)):
                start = max(0, i - window + 1)
                sub = totals[start : i + 1]
                trend.append(sum(sub) / len(sub))
            ax.plot(x, trend, linestyle="--", color="#6c757d", linewidth=1.5, label="Tendencia (rolling 3)")

        # Picos y valles
        if totals:
            max_idx = totals.index(max(totals))
            min_idx = totals.index(min(totals))
            ax.scatter([max_idx], [totals[max_idx]], color="#1a9c47", s=80, zorder=5, label="Pico")
            ax.scatter([min_idx], [totals[min_idx]], color="#d64541", s=80, zorder=5, label="Valle")

        ax.set_xticks(x, month_keys, rotation=45, ha="right")
        ax.set_ylabel("Total Gs.")
        ax.yaxis.set_major_formatter(FuncFormatter(lambda val, _: f"{self._format_millions(val)}M"))
        ax.set_title("Crecimiento mensual (picos y valles)")
        ax.grid(True, linestyle="--", alpha=0.4)
        ax.legend()
        fig.tight_layout()
        canvas.draw_idle()
        dialog.exec()

    def _apply_selection_style(self) -> None:
        highlight = QColor("#cfe8ff")  # azul claro
        text_color = QColor("#0b2a4a")

        def update_palette(widget: QWidget) -> None:
            palette = widget.palette()
            palette.setColor(QPalette.ColorRole.Highlight, highlight)
            palette.setColor(QPalette.ColorRole.HighlightedText, text_color)
            widget.setPalette(palette)

        for table in (self.units_table, self.amount_table, self.total_table):
            update_palette(table)

    def _format_millions(self, value: float) -> str:
        scaled = value / 1_000_000
        if abs(scaled) >= 100:
            text = f"{scaled:,.0f}"
        elif abs(scaled) >= 10:
            text = f"{scaled:,.1f}"
        else:
            text = f"{scaled:,.2f}"
        return text.replace(",", ".")
