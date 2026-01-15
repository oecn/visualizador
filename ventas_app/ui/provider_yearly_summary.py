"""Panel de resumen anual por proveedor."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
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

from ..database import SalesRepository

TREND_COLORS = {
    "up": QColor("#1a9c47"),
    "down": QColor("#d64541"),
    "flat": QColor("#6d7a88"),
}


class ProviderYearlySummaryTab(QWidget):
    """Resumen anual por proveedor con participacion y ranking."""

    def __init__(self, repository: SalesRepository) -> None:
        super().__init__()
        self.repository = repository
        self.year_combo = QComboBox()
        self.rank_metric_combo = QComboBox()
        self.summary_label = QLabel("Sin datos.")
        self.table = QTableWidget()
        self._build_ui()
        self._connect_signals()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Ano:"))
        filter_layout.addWidget(self.year_combo, 1)
        filter_layout.addWidget(QLabel("Ranking:"))
        self.rank_metric_combo.addItem("Unidades", "qty")
        self.rank_metric_combo.addItem("Ventas", "amt")
        filter_layout.addWidget(self.rank_metric_combo)
        layout.addLayout(filter_layout)

        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(
            [
                "Proveedor",
                "Unidades",
                "Ventas (Gs.)",
                "% participacion",
                "Var unidades vs ano-1",
                "Var ventas vs ano-1",
                "Rank crecimiento",
                "Rank volumen",
            ]
        )
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        for idx in range(1, self.table.columnCount()):
            header.setSectionResizeMode(idx, QHeaderView.ResizeToContents)
        layout.addWidget(self.table, 1)

        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

    def _connect_signals(self) -> None:
        self.year_combo.currentIndexChanged.connect(lambda *_: self.load_year())
        self.rank_metric_combo.currentIndexChanged.connect(lambda *_: self.load_year())

    def refresh(self) -> None:
        years = self.repository.list_years()
        current = self.year_combo.currentData()
        self.year_combo.blockSignals(True)
        self.year_combo.clear()
        for year in years:
            self.year_combo.addItem(str(year), year)
        self.year_combo.blockSignals(False)
        if years:
            if current in years:
                idx = self.year_combo.findData(current)
                if idx >= 0:
                    self.year_combo.setCurrentIndex(idx)
            else:
                self.year_combo.setCurrentIndex(0)
            self.load_year()
        else:
            self._clear_table("Importa datos para ver el resumen anual.")

    def load_year(self) -> None:
        year = self.year_combo.currentData()
        if not year:
            self._clear_table("Selecciona un ano para ver el resumen.")
            return
        cur_rows = self.repository.fetch_provider_yearly_totals(int(year))
        if not cur_rows:
            self._clear_table(f"No hay datos para {year}.")
            return
        prev_rows = self.repository.fetch_provider_yearly_totals(int(year) - 1)
        prev_map = {(row["provider"] or "Sin proveedor").strip() or "Sin proveedor": row for row in prev_rows}
        total_amt = sum(float(row["total_amount"] or 0.0) for row in cur_rows)
        total_qty = sum(float(row["total_quantity"] or 0.0) for row in cur_rows)

        items: list[dict[str, object]] = []
        for row in cur_rows:
            provider = (row["provider"] or "Sin proveedor").strip() or "Sin proveedor"
            qty = float(row["total_quantity"] or 0.0)
            amt = float(row["total_amount"] or 0.0)
            prev = prev_map.get(provider)
            prev_qty = float(prev["total_quantity"] or 0.0) if prev else 0.0
            prev_amt = float(prev["total_amount"] or 0.0) if prev else 0.0
            share = (amt / total_amt * 100) if total_amt > 0 else 0.0
            items.append(
                {
                    "provider": provider,
                    "qty": qty,
                    "amt": amt,
                    "prev_qty": prev_qty,
                    "prev_amt": prev_amt,
                    "share": share,
                    "growth_score": self._growth_score(amt, prev_amt),
                }
            )

        rank_metric = self.rank_metric_combo.currentData() or "qty"
        metric_label = "Rank unidades" if rank_metric == "qty" else "Rank ventas"
        self.table.setHorizontalHeaderItem(7, QTableWidgetItem(metric_label))
        volume_order = sorted(
            items,
            key=lambda item: (float(item[rank_metric]), str(item["provider"]).lower()),
            reverse=True,
        )
        for idx, item in enumerate(volume_order, start=1):
            item["rank_volume"] = idx

        growth_order = sorted(
            items,
            key=lambda item: (
                float(item["growth_score"]),
                float(item["amt"]),
                str(item["provider"]).lower(),
            ),
            reverse=True,
        )
        for idx, item in enumerate(growth_order, start=1):
            item["rank_growth"] = idx

        self.table.setRowCount(len(volume_order))
        for row_idx, item in enumerate(volume_order):
            self.table.setItem(row_idx, 0, QTableWidgetItem(str(item["provider"])))
            self.table.setItem(row_idx, 1, QTableWidgetItem(self._format_number(float(item["qty"]))))
            self.table.setItem(row_idx, 2, QTableWidgetItem(self._format_currency(float(item["amt"]))))
            share_text = f"{float(item['share']):.1f}%".replace(".", ",")
            self.table.setItem(row_idx, 3, QTableWidgetItem(share_text))
            self.table.setItem(
                row_idx,
                4,
                self._delta_item(float(item["qty"]), float(item["prev_qty"])),
            )
            self.table.setItem(
                row_idx,
                5,
                self._delta_item(float(item["amt"]), float(item["prev_amt"]), is_currency=True),
            )
            self.table.setItem(row_idx, 6, QTableWidgetItem(str(item.get("rank_growth", ""))))
            self.table.setItem(row_idx, 7, QTableWidgetItem(str(item.get("rank_volume", ""))))

        summary = (
            f"Ano {year}: proveedores {len(volume_order)} | Ventas {self._format_currency(total_amt)}"
            f" | Unidades {self._format_number(total_qty)}"
        )
        self.summary_label.setText(summary)

    def _clear_table(self, message: str) -> None:
        self.table.setRowCount(0)
        self.summary_label.setText(message)

    def _growth_score(self, current: float, previous: float) -> float:
        if previous > 0:
            return (current - previous) / previous
        if current > 0:
            return float("inf")
        return float("-inf")

    def _delta_item(self, current: float, previous: float, *, is_currency: bool = False) -> QTableWidgetItem:
        diff = current - previous
        text = self._format_currency(diff) if is_currency else self._format_number(diff)
        pct = ""
        if abs(previous) > 1e-6:
            pct = f" ({(diff / previous) * 100:+.1f}%)".replace(".", ",")
        item = QTableWidgetItem(f"{text}{pct}")
        color = TREND_COLORS["flat"]
        if diff > 0:
            color = TREND_COLORS["up"]
        elif diff < 0:
            color = TREND_COLORS["down"]
        item.setForeground(QBrush(color))
        return item

    def _format_currency(self, value: float) -> str:
        return f"Gs. {value:,.0f}".replace(",", ".")

    def _format_number(self, value: float) -> str:
        if value.is_integer():
            return f"{int(value):,}".replace(",", ".")
        return f"{value:,.2f}".replace(",", ".")
