"""Ventana principal de la app de ventas."""

from __future__ import annotations

import csv
import sqlite3
from pathlib import Path
from datetime import date
from typing import Optional

from PySide6.QtCore import Qt, QDate, QStringListModel
from PySide6.QtGui import QFont, QColor, QBrush, QPalette
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QCompleter,
    QDateEdit,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter

from ..database import SalesRepository
from ..excel_importer import ExcelImporter
from .comparar_ayaa import CompararAyAADialog
from .provider_yearly_summary import ProviderYearlySummaryTab
from .provider_monthly_trend import ProviderMonthlyTrendTab
from .user_panel import UserPanel


TREND_COLORS = {
    "up": QColor("#1a9c47"),
    "down": QColor("#d64541"),
    "flat": QColor("#6d7a88"),
}

TREND_SYMBOLS = {
    "up": "\u25B2",
    "down": "\u25BC",
    "flat": "\u25B6",
}


class MainWindow(QMainWindow):
    """Interfaz principal para importar y revisar las ventas."""

    MONTHLY_BRANCHES = ["Casa Central", "CAACUPEMI", "ITAUGUA"]

    def __init__(
        self,
        repository: SalesRepository,
        importer: ExcelImporter,
        default_folder: Optional[Path] = None,
        current_user: Optional[sqlite3.Row] = None,
    ) -> None:
        super().__init__()
        self.repository = repository
        self.importer = importer
        self.default_folder = default_folder or Path.cwd()
        self.current_user = current_user
        self.setWindowTitle("Seguimiento de ventas por sucursal")
        self.resize(1200, 700)

        self.period_list = QListWidget()
        self.branch_filter = QComboBox()
        self.table = QTableWidget()
        self.summary_label = QLabel("Sin datos")
        self.tabs = QTabWidget()
        self.month_filter = QComboBox()
        self.provider_filter = QComboBox()
        self.product_filter_combo = QComboBox()
        self.product_filter = QLineEdit()
        self.month_metric_combo = QComboBox()
        self.product_type_filter = QComboBox()
        self.product_completer = QCompleter()
        self.start_date_filter = QDateEditWithReset()
        self.end_date_filter = QDateEditWithReset()
        self.monthly_tree = QTreeWidget()
        self.month_summary_label = QLabel("Selecciona un mes para ver la venta mensual.")
        self.yoy_month_combo = QComboBox()
        self.yoy_branch_table = QTableWidget()
        self.yoy_gainers_table = QTableWidget()
        self.yoy_losers_table = QTableWidget()
        self.yoy_summary_label = QLabel("Sin datos anuales.")
        self.yoy_status_label = QLabel("Selecciona un mes para ver el comparativo.")
        self.yoy_quick_label = QLabel("")
        self.yoy_product_search = QLineEdit()
        self.yoy_product_completer = QCompleter()
        self.yoy_product_status = QLabel("Escribe o elige un producto para ver su YoY.")
        self.yoy_export_btn = QPushButton("Exportar CSV")
        self.yoy_copy_btn = QPushButton("Copiar fila sucursal")
        self.growth_provider_filter = QComboBox()
        self.growth_metric_combo = QComboBox()
        self.growth_product_filter = QLineEdit()
        self.growth_table = QTableWidget()
        self.growth_chart = FigureCanvas(Figure(figsize=(5, 3)))
        self.growth_summary_label = QLabel("Selecciona un proveedor para ver el crecimiento anual.")
        self.growth_status_badge = QLabel("")
        self._product_search_map: dict[str, tuple[str, str]] = {}
        self.provider_yearly_tab = ProviderYearlySummaryTab(self.repository)
        self.provider_monthly_tab = ProviderMonthlyTrendTab(self.repository)
        username = (current_user["username"] if current_user else "usuario") if current_user else "usuario"
        is_admin = bool(current_user["is_admin"]) if current_user else False
        self.user_panel = UserPanel(self.repository, username, is_admin)

        self._build_ui()
        self._connect_signals()
        self.refresh_periods()
        self._apply_selection_style()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        toolbar = QHBoxLayout()

        import_btn = QPushButton("Importar Excel")
        import_folder_btn = QPushButton("Importar carpeta")
        select_folder_btn = QPushButton("Seleccionar carpeta")
        refresh_btn = QPushButton("Refrescar")
        delete_period_btn = QPushButton("Eliminar Excel seleccionado")
        clear_btn = QPushButton("Borrar importaciones")
        if not (self.current_user and bool(self.current_user["is_admin"])):
            clear_btn.setVisible(False)

        toolbar.addWidget(import_btn)
        toolbar.addWidget(import_folder_btn)
        toolbar.addWidget(select_folder_btn)
        toolbar.addWidget(refresh_btn)
        toolbar.addWidget(delete_period_btn)
        toolbar.addWidget(clear_btn)
        toolbar.addStretch()
        main_layout.addLayout(toolbar)

        body_layout = QHBoxLayout()
        self.period_list.setSelectionMode(QAbstractItemView.SingleSelection)
        body_layout.addWidget(self.period_list, 1)

        detail_tab = QWidget()
        detail_layout = QVBoxLayout(detail_tab)
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Sucursal:"))
        self.branch_filter.addItem("Todas", None)
        filter_layout.addWidget(self.branch_filter, 1)
        filter_layout.addWidget(QLabel("Desde:"))
        filter_layout.addWidget(self.start_date_filter)
        filter_layout.addWidget(QLabel("Hasta:"))
        filter_layout.addWidget(self.end_date_filter)
        detail_layout.addLayout(filter_layout)

        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Periodo", "Sucursal", "Codigo", "Descripcion", "Cantidad", "Venta"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        detail_layout.addWidget(self.table, 1)
        detail_layout.addWidget(self.summary_label)

        monthly_tab = QWidget()
        monthly_layout = QVBoxLayout(monthly_tab)
        month_filter_layout = QHBoxLayout()
        month_filter_layout.addWidget(QLabel("Mes:"))
        month_filter_layout.addWidget(self.month_filter, 1)
        month_filter_layout.addWidget(QLabel("Proveedor:"))
        month_filter_layout.addWidget(self.provider_filter, 1)
        month_filter_layout.addWidget(QLabel("Metrica:"))
        self.month_metric_combo.addItem("UNIDADES", "qty")
        self.month_metric_combo.addItem("MONTO (GS.)", "amount")
        month_filter_layout.addWidget(self.month_metric_combo)
        monthly_layout.addLayout(month_filter_layout)

        product_filter_layout = QHBoxLayout()
        product_filter_layout.addWidget(QLabel("Filtro producto:"))
        self.product_filter_combo.addItem("EL CACIQUE", "el cacique")
        self.product_filter_combo.addItem("Todos", "")
        product_filter_layout.addWidget(self.product_filter_combo)
        self.product_filter.setPlaceholderText("Filtrar por codigo o descripcion")
        self.product_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.product_completer.setFilterMode(Qt.MatchContains)
        self.product_filter.setCompleter(self.product_completer)
        product_filter_layout.addWidget(self.product_filter, 1)
        product_filter_layout.addWidget(QLabel("Tipo:"))
        self.product_type_filter.addItem("Todos", "")
        self.product_type_filter.addItem("Fideo", "fideo")
        self.product_type_filter.addItem("Harina", "harina")
        self.product_type_filter.addItem("Granos", "el cacique")
        product_filter_layout.addWidget(self.product_type_filter)
        monthly_layout.addLayout(product_filter_layout)

        toggle_layout = QHBoxLayout()
        collapse_all_btn = QPushButton("Colapsar todo")
        expand_all_btn = QPushButton("Desplegar todo")
        toggle_layout.addWidget(collapse_all_btn)
        toggle_layout.addWidget(expand_all_btn)
        toggle_layout.addStretch()
        monthly_layout.addLayout(toggle_layout)

        headers = ["Mes / Producto"] + self.MONTHLY_BRANCHES + ["Total unidades"]
        self.monthly_tree.setColumnCount(len(headers))
        self.monthly_tree.setHeaderLabels(headers)
        self.monthly_tree.setAlternatingRowColors(True)
        self.monthly_tree.setRootIsDecorated(True)
        header = self.monthly_tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.resizeSection(0, 450)
        for idx in range(1, len(headers) - 1):
            header.setSectionResizeMode(idx, QHeaderView.ResizeToContents)
            header.resizeSection(idx, 135)
        header.setSectionResizeMode(len(headers) - 1, QHeaderView.ResizeToContents)
        header.resizeSection(len(headers) - 1, 150)
        monthly_layout.addWidget(self.monthly_tree, 1)
        monthly_layout.addWidget(self.month_summary_label)

        yoy_tab = QWidget()
        yoy_layout = QVBoxLayout(yoy_tab)
        yoy_filter = QHBoxLayout()
        yoy_filter.addWidget(QLabel("Mes:"))
        yoy_filter.addWidget(self.yoy_month_combo, 1)
        compare_ayaa_btn = QPushButton("Comparar AyAA")
        compare_ayaa_btn.clicked.connect(self.open_compare_ayaa)
        yoy_filter.addWidget(compare_ayaa_btn)
        yoy_filter.addWidget(self.yoy_export_btn)
        yoy_filter.addWidget(self.yoy_copy_btn)
        yoy_filter.addStretch()
        yoy_layout.addLayout(yoy_filter)

        product_filter_layout = QHBoxLayout()
        product_filter_layout.addWidget(QLabel("Producto:"))
        self.yoy_product_search.setPlaceholderText("Codigo o descripcion")
        self.yoy_product_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.yoy_product_completer.setFilterMode(Qt.MatchContains)
        self.yoy_product_search.setCompleter(self.yoy_product_completer)
        product_filter_layout.addWidget(self.yoy_product_search, 1)
        yoy_layout.addLayout(product_filter_layout)
        self.yoy_product_status.setWordWrap(True)
        yoy_layout.addWidget(self.yoy_product_status)
        self.yoy_status_label.setWordWrap(True)
        self.yoy_quick_label.setWordWrap(True)
        yoy_layout.addWidget(self.yoy_status_label)
        yoy_layout.addWidget(self.yoy_quick_label)

        yoy_branch_group = QGroupBox("Sucursales (YoY)")
        branch_group_layout = QVBoxLayout(yoy_branch_group)
        self.yoy_branch_table.setColumnCount(7)
        self.yoy_branch_table.setHorizontalHeaderLabels(
            [
                "Sucursal",
                "Unidades mes",
                "Unidades mes-1y",
                "Δ unidades",
                "Venta mes",
                "Venta mes-1y",
                "Δ venta",
            ]
        )
        self.yoy_branch_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.yoy_branch_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.yoy_branch_table.setAlternatingRowColors(True)
        self.yoy_branch_table.horizontalHeader().setStretchLastSection(True)
        branch_group_layout.addWidget(self.yoy_branch_table)
        yoy_layout.addWidget(yoy_branch_group)

        yoy_products_group = QGroupBox("Productos")
        products_group_layout = QHBoxLayout(yoy_products_group)
        self._init_yoy_product_table(self.yoy_gainers_table, "Top alzas")
        self._init_yoy_product_table(self.yoy_losers_table, "Top caídas")
        products_group_layout.addWidget(self.yoy_gainers_table, 1)
        products_group_layout.addWidget(self.yoy_losers_table, 1)
        yoy_layout.addWidget(yoy_products_group)
        yoy_layout.addWidget(self.yoy_summary_label)
        growth_tab = QWidget()
        growth_layout = QVBoxLayout(growth_tab)
        growth_filter = QHBoxLayout()
        growth_filter.addWidget(QLabel("Proveedor:"))
        self.growth_provider_filter.addItem("Selecciona un proveedor", None)
        growth_filter.addWidget(self.growth_provider_filter, 1)
        self.growth_metric_combo.addItems(["VENTAS", "UNIDADES"])
        growth_filter.addWidget(QLabel("MÃ©trica:"))
        growth_filter.addWidget(self.growth_metric_combo)
        growth_layout.addLayout(growth_filter)
        product_filter = QHBoxLayout()
        product_filter.addWidget(QLabel("Producto/Familia:"))
        self.growth_product_filter.setPlaceholderText("Ej: azucar | azucar 250g")
        product_filter.addWidget(self.growth_product_filter, 1)
        growth_layout.addLayout(product_filter)
        growth_layout.addWidget(self.growth_summary_label)
        self.growth_status_badge.setAlignment(Qt.AlignCenter)
        self.growth_status_badge.setStyleSheet(
            """
            QLabel {
                border-radius: 6px;
                padding: 8px 10px;
                font-weight: 600;
            }
            """
        )
        growth_layout.addWidget(self.growth_status_badge)
        self.growth_table.setColumnCount(5)
        self.growth_table.setHorizontalHeaderLabels(
            ["Año", "Unidades", "Venta (Gs.)", "Δ unidades vs año-1", "Δ venta vs año-1"]
        )
        self.growth_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.growth_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.growth_table.setAlternatingRowColors(True)
        self.growth_table.horizontalHeader().setStretchLastSection(True)
        self.growth_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        growth_layout.addWidget(self.growth_table, 1)
        growth_layout.addWidget(self.growth_chart, 2)

        self.tabs.addTab(detail_tab, "Detalle")
        self.tabs.addTab(monthly_tab, "Venta mensual")
        self.tabs.addTab(yoy_tab, "Comparativo anual")
        self.tabs.addTab(growth_tab, "Crecimiento anual")
        self.tabs.addTab(self.provider_yearly_tab, "Resumen anual")
        self.tabs.addTab(self.provider_monthly_tab, "Evolucion mensual")
        self.tabs.addTab(self.user_panel, "Usuarios")

        body_layout.addWidget(self.tabs, 3)
        main_layout.addLayout(body_layout, 1)

        self.import_btn = import_btn
        self.import_folder_btn = import_folder_btn
        self.select_folder_btn = select_folder_btn
        self.refresh_btn = refresh_btn
        self.delete_period_btn = delete_period_btn
        self.clear_btn = clear_btn
        self.collapse_all_btn = collapse_all_btn
        self.expand_all_btn = expand_all_btn
        self.yoy_tab = yoy_tab

    def _connect_signals(self) -> None:
        self.import_btn.clicked.connect(self.import_files)
        self.import_folder_btn.clicked.connect(self.import_folder)
        self.select_folder_btn.clicked.connect(self.select_folder)
        self.refresh_btn.clicked.connect(self.refresh_periods)
        self.delete_period_btn.clicked.connect(self.delete_selected_period)
        self.clear_btn.clicked.connect(self.clear_data)
        self.yoy_export_btn.clicked.connect(self.export_yoy_csv)
        self.yoy_copy_btn.clicked.connect(self.copy_selected_yoy_row)
        self.yoy_product_search.textChanged.connect(lambda *_: self._update_yoy_product_comparison())
        self.collapse_all_btn.clicked.connect(self.monthly_tree.collapseAll)
        self.expand_all_btn.clicked.connect(self.monthly_tree.expandAll)
        self.period_list.currentItemChanged.connect(lambda *_: self._period_changed())
        self.branch_filter.currentIndexChanged.connect(lambda *_: self.load_records())
        self.month_filter.currentIndexChanged.connect(lambda *_: self.load_monthly_summary())
        self.provider_filter.currentIndexChanged.connect(lambda *_: self.refresh_months())
        self.month_metric_combo.currentIndexChanged.connect(lambda *_: self.load_monthly_summary())
        self.product_filter_combo.currentIndexChanged.connect(lambda *_: self.load_monthly_summary())
        self.product_filter.textChanged.connect(lambda *_: self.load_monthly_summary())
        self.product_type_filter.currentIndexChanged.connect(lambda *_: self.load_monthly_summary())
        self.monthly_tree.itemDoubleClicked.connect(self._open_product_chart)
        self.start_date_filter.dateChanged.connect(lambda *_: self.load_records())
        self.end_date_filter.dateChanged.connect(lambda *_: self.load_records())
        self.yoy_month_combo.currentIndexChanged.connect(lambda *_: self.load_yoy_summary())
        self.growth_provider_filter.currentIndexChanged.connect(lambda *_: self.load_growth_summary())
        self.growth_metric_combo.currentIndexChanged.connect(lambda *_: self.load_growth_summary())
        self.growth_product_filter.textChanged.connect(lambda *_: self.load_growth_summary())

    def refresh_periods(self) -> None:
        self.period_list.clear()
        periods = self.repository.list_periods()
        for row in periods:
            item = QListWidgetItem(self._format_period_label(row))
            item.setData(Qt.UserRole, int(row["id"]))
            self.period_list.addItem(item)
        if periods:
            self.period_list.setCurrentRow(0)
        else:
            self.branch_filter.clear()
            self.branch_filter.addItem("Todas", None)
            self.table.setRowCount(0)
            self.summary_label.setText("Sin datos importados")
        self._refresh_provider_filter()
        self.refresh_months()
        self.refresh_yoy_tab()
        self.refresh_growth_tab()
        self.user_panel.refresh_audit()
        self.provider_yearly_tab.refresh()
        self.provider_monthly_tab.refresh()
        self._refresh_product_completer()
        self._refresh_yoy_product_search()

    def _refresh_provider_filter(self) -> None:
        providers = self.repository.list_providers()
        current = self.provider_filter.currentData()
        self.provider_filter.blockSignals(True)
        self.provider_filter.clear()
        self.provider_filter.addItem("Todos", None)
        for provider in providers:
            self.provider_filter.addItem(provider, provider)
        if current:
            idx = self.provider_filter.findData(current)
            if idx >= 0:
                self.provider_filter.setCurrentIndex(idx)
        if self.provider_filter.currentIndex() < 0:
            self.provider_filter.setCurrentIndex(0)
        self.provider_filter.blockSignals(False)

    def _period_changed(self) -> None:
        period_id = self._selected_period_id()
        self.branch_filter.blockSignals(True)
        self.branch_filter.clear()
        self.branch_filter.addItem("Todas", None)                   
        if period_id is None:
            self.branch_filter.blockSignals(False)
            self.load_records()
            return
        branches = self.repository.list_branches(period_id)
        branches = self._sort_branches(branches)
        for branch in branches:
            self.branch_filter.addItem(branch, branch)
        self.branch_filter.blockSignals(False)
        self.load_records()

    def _selected_period_id(self) -> Optional[int]:
        item = self.period_list.currentItem()
        if not item:
            return None
        period_id = item.data(Qt.UserRole)
        return int(period_id) if period_id is not None else None

    def _current_username(self) -> str:
        if self.current_user and self.current_user["username"]:
            return str(self.current_user["username"])
        return "desconocido"

    def _is_admin(self) -> bool:
        return bool(self.current_user and self.current_user["is_admin"])

    def _can_delete_period(self, period_id: int) -> bool:
        if self._is_admin():
            return True
        period = self.repository.fetch_period(period_id)
        if not period:
            return False
        start = self._parse_date(period["start_date"]) or self._parse_date(period["end_date"])
        if not start:
            return False
        today = date.today()
        return start.year == today.year and start.month == today.month

    def load_records(self) -> None:
        period_id = self._selected_period_id()
        if period_id is None:
            self.table.setRowCount(0)
            self.summary_label.setText("Selecciona un periodo para ver los datos.")
            return
        branch = self.branch_filter.currentData()
        rows = self.repository.fetch_records(period_id, branch=branch)
        start_filter, end_filter = self._current_date_filters()
        rows = [
            row
            for row in rows
            if self._period_overlaps_filter(row["start_date"], row["end_date"], start_filter, end_filter)
        ]
        self.table.setRowCount(len(rows))
        total_amount = 0.0
        total_quantity = 0.0
        for idx, row in enumerate(rows):
            period_text = self._format_period_range(row["start_date"], row["end_date"])
            self.table.setItem(idx, 0, QTableWidgetItem(period_text))
            self.table.setItem(idx, 1, QTableWidgetItem(row["branch"]))
            self.table.setItem(idx, 2, QTableWidgetItem(row["product_code"] or ""))
            self.table.setItem(idx, 3, QTableWidgetItem(row["description"] or ""))
            qty = row["quantity"] or 0.0
            amount = row["amount"] or 0.0
            total_quantity += qty
            total_amount += amount
            self.table.setItem(idx, 4, QTableWidgetItem(self._format_number(qty)))
            self.table.setItem(idx, 5, QTableWidgetItem(self._format_currency(amount)))
        self.summary_label.setText(
            f"Total unidades: {self._format_number(total_quantity)} · Total venta: {self._format_currency(total_amount)}"
        )

    def import_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecciona los Excel de ventas",
            str(self.default_folder),
            "Archivos Excel (*.xlsx *.xls)",
        )
        if not paths:
            return

        self._process_import([Path(p) for p in paths])

    def import_folder(self) -> None:
        folder = self.default_folder
        if not folder.exists() or not folder.is_dir():
            QMessageBox.warning(self, "Carpeta no encontrada", f"La carpeta {folder} no existe.")
            return

        # Busca recursivamente todos los Excel en la carpeta seleccionada.
        paths = sorted(folder.rglob("*.xls*"))
        if not paths:
            QMessageBox.information(self, "Sin archivos", f"No se encontraron Excel en {folder}.")
            return

        self._process_import(paths)

    def select_folder(self) -> None:
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Selecciona la carpeta con Excel",
            str(self.default_folder),
            options=QFileDialog.ShowDirsOnly,
        )
        if not folder_path:
            return

        folder = Path(folder_path)
        if not folder.exists() or not folder.is_dir():
            QMessageBox.warning(self, "Carpeta no encontrada", f"La carpeta {folder} no existe.")
            return

        self.default_folder = folder
        QMessageBox.information(self, "Carpeta seleccionada", f"Carpeta por defecto: {folder}")

    def clear_data(self) -> None:
        reply = QMessageBox.question(
            self,
            "Borrar importaciones",
            "Esto eliminarÃ¡ todos los periodos y sus registros. ¿Deseas continuar?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        try:
            self.repository.clear_all()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "No se pudo borrar", str(exc))
            return
        QMessageBox.information(self, "Importaciones eliminadas", "Se limpiÃ³ la base de datos.")
        self.refresh_periods()

    def delete_selected_period(self) -> None:
        period_id = self._selected_period_id()
        if period_id is None:
            QMessageBox.information(self, "Sin seleccion", "Elige un Excel importado para poder eliminarlo.")
            return
        if not self._can_delete_period(period_id):
            QMessageBox.warning(
                self,
                "Acceso restringido",
                "Solo el administrador puede eliminar periodos fuera del mes actual.",
            )
            return
        current_item = self.period_list.currentItem()
        label = current_item.text() if current_item else "periodo seleccionado"
        reply = QMessageBox.question(
            self,
            "Eliminar Excel",
            f"Esto eliminará el Excel seleccionado y sus datos:\n{label}\n¿Deseas continuar?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        try:
            self.repository.delete_period(period_id, deleted_by=self._current_username())
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "No se pudo eliminar", str(exc))
            return
        QMessageBox.information(self, "Excel eliminado", "Se elimino la importacion seleccionada.")
        self.refresh_periods()

    def _process_import(self, paths: list[Path]) -> None:
        successes: list[str] = []
        errors: list[str] = []
        for path in paths:
            try:
                batch = self.importer.load(path)
                inserted = self.repository.store_batch(batch, created_by=self._current_username())
                successes.append(f"{path.name}: {inserted} filas")
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{path.name}: {exc}")
        self.refresh_periods()

        if successes:
            QMessageBox.information(self, "Importacion completada", "\n".join(successes))
        if errors:
            QMessageBox.warning(self, "Archivos con errores", "\n".join(errors))

    def refresh_months(self) -> None:
        provider = self.provider_filter.currentData()
        months = self.repository.list_months(provider=provider)
        self.month_filter.blockSignals(True)
        self.month_filter.clear()
        self.month_filter.addItem("Todos los meses", None)
        for row in months:
            label = self._format_month_label(row)
            self.month_filter.addItem(label, row["month_key"])
        self.month_filter.blockSignals(False)
        if months:
            self.month_filter.setCurrentIndex(0)
            self.load_monthly_summary()
        else:
            self.monthly_tree.clear()
            self.month_summary_label.setText("Sin datos mensuales disponibles.")

    def load_monthly_summary(self) -> None:
        provider = self.provider_filter.currentData()
        month_key = self.month_filter.currentData()
        metric = self.month_metric_combo.currentData() or "qty"
        is_currency = metric == "amount"
        self._update_monthly_headers(metric)
        months = self.repository.list_months(provider=provider)
        if month_key:
            months = [row for row in months if row["month_key"] == month_key]
        filter_terms = self._current_product_filters()
        product_type_label = self.product_type_filter.currentText().strip().lower()
        brand_filter = (self.product_filter_combo.currentData() or "").strip().lower()
        self.monthly_tree.clear()
        total_products = 0
        if not months:
            self.month_summary_label.setText("Sin datos mensuales disponibles.")
            return

        month_infos: list[dict[str, object]] = []
        for month_row in months:
            summary = self.repository.fetch_monthly_summary(month_row["month_key"], provider=provider)
            product_map: dict[tuple[str, str], dict[str, float]] = {}
            for row in summary:
                branch = row["branch"]
                if branch not in self.MONTHLY_BRANCHES:
                    continue
                code = row["product_code"] or ""
                desc = row["description"] or ""
                qty = float(row["total_quantity"] or 0.0)
                amt = float(row["total_amount"] or 0.0)
                value = amt if is_currency else qty
                key = (code, desc)
                if key not in product_map:
                    product_map[key] = {b: 0.0 for b in self.MONTHLY_BRANCHES}
                product_map[key][branch] = value
            month_infos.append(
                {
                    "month_row": month_row,
                    "product_map": product_map,
                }
            )

        for idx, info in enumerate(month_infos):
            if idx + 1 < len(month_infos):
                prev_info = month_infos[idx + 1]
                info["prev_map"] = prev_info["product_map"]
            else:
                info["prev_map"] = {}

        for info in month_infos:
            month_row = info["month_row"]  # type: ignore[index]
            product_map = info["product_map"]  # type: ignore[assignment]
            prev_map = info.get("prev_map") if isinstance(info, dict) else {}
            month_label = self._format_month_label(month_row)  # type: ignore[arg-type]
            month_item = QTreeWidgetItem(self.monthly_tree)
            month_item.setText(0, month_label)
            filtered_map = {
                key: values
                for key, values in product_map.items()
                if self._product_matches_filter(key, filter_terms, product_type_label, brand_filter)
            }
            filtered_prev_map = (
                {
                    key: values
                    for key, values in prev_map.items()
                    if self._product_matches_filter(key, filter_terms, product_type_label, brand_filter)
                }
                if isinstance(prev_map, dict)
                else {}
            )
            month_totals = self._compute_branch_totals(filtered_map)
            prev_totals = self._compute_branch_totals(filtered_prev_map)
            month_total_sum = sum(month_totals.values())
            prev_total_sum = sum(prev_totals.values())
            for idx_col, branch in enumerate(self.MONTHLY_BRANCHES, start=1):
                prev_value = prev_totals.get(branch, 0.0)
                text, color = self._format_branch_trend_text(
                    month_totals[branch],
                    prev_value,
                    is_total=True,
                    is_currency=is_currency,
                )
                month_item.setText(idx_col, text)
                if color:
                    month_item.setForeground(idx_col, QBrush(color))
            total_text, total_color = self._format_branch_trend_text(
                month_total_sum,
                prev_total_sum,
                is_total=True,
                is_currency=is_currency,
            )
            month_item.setText(len(self.MONTHLY_BRANCHES) + 1, total_text)
            if total_color:
                month_item.setForeground(len(self.MONTHLY_BRANCHES) + 1, QBrush(total_color))
            bold_font = QFont()
            bold_font.setBold(True)
            month_item.setFont(0, bold_font)
            month_item.setBackground(0, self.monthly_tree.palette().alternateBase())
            for key, values in sorted(
                filtered_map.items(),
                key=lambda item: (
                    -sum(item[1].values()),  # primero ordena por mayor cantidad total
                    item[0][1].lower() or item[0][0].lower(),
                ),
            ):
                code, desc = key
                label = desc or code or "Producto sin nombre"
                if code and desc:
                    label = f"{code} · {desc}"
                child = QTreeWidgetItem(month_item)
                child.setText(0, label)
                child.setData(0, Qt.UserRole, {"code": code, "description": desc})
                total = 0.0
                prev_values = prev_map.get(key, {}) if isinstance(prev_map, dict) else {}
                for idx_col, branch in enumerate(self.MONTHLY_BRANCHES, start=1):
                    value = values.get(branch, 0.0)
                    total += value
                    prev_value = prev_values.get(branch, 0.0) if isinstance(prev_values, dict) else 0.0
                    text, color = self._format_branch_trend_text(
                        value,
                        prev_value,
                        is_currency=is_currency,
                    )
                    child.setText(idx_col, text)
                    if color:
                        child.setForeground(idx_col, QBrush(color))
                total_text = self._format_currency(total) if is_currency else self._format_number(total)
                child.setText(len(self.MONTHLY_BRANCHES) + 1, total_text)
            total_products += len(filtered_map)
            month_item.setExpanded(True)
        self.month_summary_label.setText(
            f"Productos listados: {total_products} · Meses mostrados: {len(month_infos)}"
        )

    def refresh_yoy_tab(self) -> None:
        months = self.repository.list_months()
        self.yoy_month_combo.blockSignals(True)
        self.yoy_month_combo.clear()
        for row in months:
            label = self._format_month_label(row)
            self.yoy_month_combo.addItem(label, row["month_key"])
        self.yoy_month_combo.blockSignals(False)
        if months:
            self.yoy_month_combo.setCurrentIndex(0)
            self.load_yoy_summary()
        else:
            self._clear_yoy_tables()
            self.yoy_summary_label.setText("Sin datos anuales.")
            self.yoy_status_label.setText("No hay meses importados para comparar.")
            self.yoy_quick_label.setText("")

    def load_yoy_summary(self) -> None:
        month_key = self.yoy_month_combo.currentData()
        if not month_key:
            self._clear_yoy_tables()
            self.yoy_summary_label.setText("Sin datos anuales.")
            self.yoy_status_label.setText("Selecciona un mes para ver el comparativo.")
            self.yoy_quick_label.setText("")
            self.yoy_product_status.setText("Selecciona un mes y un producto para comparar.")
            return
        prev_month_key = self._month_key_previous_year(str(month_key))

        current_branches = {row["branch"]: row for row in self.repository.fetch_monthly_branch_totals(str(month_key))}
        prev_branches = (
            {row["branch"]: row for row in self.repository.fetch_monthly_branch_totals(prev_month_key)}
            if prev_month_key
            else {}
        )
        all_branches = self._sort_branches(list(set(current_branches) | set(prev_branches)))
        self.yoy_branch_table.setRowCount(len(all_branches))
        for idx, branch in enumerate(all_branches):
            cur_row = current_branches.get(branch)
            prev_row = prev_branches.get(branch)
            cur_qty = float(cur_row["total_quantity"] or 0.0) if cur_row else 0.0
            prev_qty = float(prev_row["total_quantity"] or 0.0) if prev_row else 0.0
            cur_amt = float(cur_row["total_amount"] or 0.0) if cur_row else 0.0
            prev_amt = float(prev_row["total_amount"] or 0.0) if prev_row else 0.0
            self.yoy_branch_table.setItem(idx, 0, QTableWidgetItem(branch))
            self.yoy_branch_table.setItem(idx, 1, QTableWidgetItem(self._format_number(cur_qty)))
            self.yoy_branch_table.setItem(idx, 2, QTableWidgetItem(self._format_number(prev_qty)))
            self.yoy_branch_table.setItem(idx, 3, self._delta_item(cur_qty, prev_qty))
            self.yoy_branch_table.setItem(idx, 4, QTableWidgetItem(self._format_currency(cur_amt)))
            self.yoy_branch_table.setItem(idx, 5, QTableWidgetItem(self._format_currency(prev_amt)))
            self.yoy_branch_table.setItem(idx, 6, self._delta_item(cur_amt, prev_amt, is_currency=True))

        product_gainers, product_losers = self._compute_product_yoy(str(month_key), prev_month_key)
        self._fill_product_table(self.yoy_gainers_table, product_gainers)
        self._fill_product_table(self.yoy_losers_table, product_losers)
        prev_label = prev_month_key or "sin datos"
        self.yoy_summary_label.setText(
            f"Mes: {month_key} vs {prev_label} · Sucursales: {len(all_branches)} · Top productos mostrados: "
            f"{len(product_gainers)} alzas / {len(product_losers)} caÃ­das"
        )

        prev_note = "" if prev_month_key else " (sin datos del año anterior)"
        self.yoy_status_label.setText(f"Comparando {month_key} vs {prev_label}{prev_note}")
        self.yoy_quick_label.setText(
            self._format_yoy_quick(product_gainers, product_losers, all_branches, current_branches, prev_branches)
        )
        self._update_yoy_product_comparison()

        self._update_yoy_product_comparison()

    def refresh_growth_tab(self) -> None:
        providers = self.repository.list_providers()
        self.growth_provider_filter.blockSignals(True)
        self.growth_provider_filter.clear()
        self.growth_provider_filter.addItem("Todos", None)
        for provider in providers:
            self.growth_provider_filter.addItem(provider, provider)
        self.growth_provider_filter.blockSignals(False)
        self.growth_provider_filter.setCurrentIndex(0)
        if providers:
            self.load_growth_summary()
        else:
            self._clear_growth_view()
            self.growth_summary_label.setText("Importa datos para ver crecimiento anual.")

    def load_growth_summary(self) -> None:
        provider = self.growth_provider_filter.currentData()
        search = self.growth_product_filter.text().strip().lower()
        rows = self.repository.fetch_yearly_totals(provider, search_text=search)
        if not rows:
            self._clear_growth_view()
            label = provider or "Todos"
            extra = f" con filtro '{search}'" if search else ""
            self.growth_summary_label.setText(f"No hay datos anuales para {label}{extra}.")
            return
        self.growth_table.setRowCount(len(rows))
        years: list[int] = []
        amounts: list[float] = []
        quantities: list[float] = []
        prev_qty = 0.0
        prev_amt = 0.0
        for idx, row in enumerate(rows):
            year = int(row["year"])
            qty = float(row["total_quantity"] or 0.0)
            amt = float(row["total_amount"] or 0.0)
            years.append(year)
            amounts.append(amt)
            quantities.append(qty)
            self.growth_table.setItem(idx, 0, QTableWidgetItem(str(year)))
            self.growth_table.setItem(idx, 1, QTableWidgetItem(self._format_number(qty)))
            self.growth_table.setItem(idx, 2, QTableWidgetItem(self._format_currency_short(amt)))
            delta_qty_item = self._delta_item(qty, prev_qty)
            delta_amt_item = self._delta_item(amt, prev_amt, is_currency=True)
            self.growth_table.setItem(idx, 3, delta_qty_item)
            self.growth_table.setItem(idx, 4, delta_amt_item)
            prev_qty = qty
            prev_amt = amt
        self._plot_growth_chart(years, amounts, quantities)
        provider_label = provider or "Todos"
        summary = f"Crecimiento anual de {provider_label}. Años: {len(years)}. "
        if search:
            summary += f"Filtro producto: '{search}'. "
        if len(amounts) > 1 and amounts[0] > 0:
            cagr = (amounts[-1] / amounts[0]) ** (1 / (len(amounts) - 1)) - 1
            summary += f"CAGR venta: {cagr * 100:+.2f}%".replace(".", ",")
        else:
            summary += "CAGR no disponible."
        self.growth_summary_label.setText(summary)
        self._update_growth_status_badge(years, amounts)

    def _plot_growth_chart(self, years: list[int], amounts: list[float], quantities: list[float]) -> None:
        figure = self.growth_chart.figure
        figure.clear()
        ax_amount = figure.add_subplot(111)
        metric = self.growth_metric_combo.currentText().strip().upper()
        is_sales = metric == "VENTAS"
        data = amounts if is_sales else quantities
        color = "#1a9c47" if is_sales else "#0f3057"
        label = "Venta (Gs.)" if is_sales else "Unidades"
        bars = ax_amount.bar(years, data, width=0.5, color=color, alpha=0.7, label=label)
        if bars and len(bars) > 0:
            bars[-1].set_color("#0e8f4d" if is_sales else "#0b3f7a")
            bars[-1].set_alpha(0.8)
        ax_amount.set_ylabel(label)
        ax_amount.set_xlabel("Año")
        ax_amount.set_xticks(years)
        ax_amount.yaxis.set_major_formatter(FuncFormatter(lambda val, _: self._format_short(val)))
        ax_amount.margins(x=0.08)
        ax_amount.set_title("Crecimiento anual por proveedor")
        ax_amount.grid(True, linestyle="--", linewidth=0.6, alpha=0.35)
        max_val = max(data) if data else 0.0
        if max_val > 0:
            ax_amount.set_ylim(0, max_val * 1.35)
        # Anotar valores abreviados encima de cada barra.
        for bar, val in zip(bars, data):
            height = bar.get_height()
            ax_amount.annotate(
                self._format_short(val),
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 12),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=9,
                bbox=dict(facecolor="white", alpha=0.85, edgecolor="none", boxstyle="round,pad=0.2"),
            )
        # Leyenda con el Ãºltimo valor mostrado.
        last_val = data[-1] if data else 0.0
        bar_label = f"{label} · {self._format_short(last_val)}"
        bars.set_label(bar_label)
        ax_amount.legend(loc="upper left")
        figure.tight_layout()
        self.growth_chart.draw_idle()

    def _update_growth_status_badge(self, years: list[int], amounts: list[float]) -> None:
        if not years or not amounts:
            self.growth_status_badge.clear()
            self.growth_status_badge.setStyleSheet(
                """
                QLabel {
                    border-radius: 6px;
                    padding: 8px 10px;
                    font-weight: 600;
                    background: #f5f5f5;
                    color: #555555;
                }
                """
            )
            return
        first = amounts[0]
        last = amounts[-1]
        diff = last - first
        pct = (diff / first * 100) if abs(first) > 1e-6 else 0.0
        trend_up = diff >= 0
        bg = "#e7f5ed" if trend_up else "#fdecea"
        fg = "#0f5132" if trend_up else "#842029"
        sign = "+" if trend_up else "-"
        diff_text = self._format_short(abs(diff))
        pct_text = f"{pct:+.1f}%".replace(".", ",") if abs(first) > 1e-6 else "N/A"
        self.growth_status_badge.setText(f"{years[0]} -> {years[-1]}: {sign}{diff_text} ({pct_text})")
        self.growth_status_badge.setStyleSheet(
            f"""
            QLabel {{
                border-radius: 6px;
                padding: 8px 10px;
                font-weight: 600;
                background: {bg};
                color: {fg};
            }}
            """
        )

    def _clear_growth_view(self) -> None:
        self.growth_table.setRowCount(0)
        figure = self.growth_chart.figure
        figure.clear()
        figure.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        self.growth_chart.draw_idle()
        self._update_growth_status_badge([], [])

    def _compute_product_yoy(
        self, month_key: str, prev_month_key: str | None
    ) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
        current = {
            (row["product_code"] or "", row["description"] or ""): row
            for row in self.repository.fetch_monthly_product_totals(month_key)
        }
        previous = {
            (row["product_code"] or "", row["description"] or ""): row
            for row in self.repository.fetch_monthly_product_totals(prev_month_key)
        } if prev_month_key else {}
        gainers: list[dict[str, object]] = []
        losers: list[dict[str, object]] = []
        for key in set(current) | set(previous):
            code, desc = key
            cur_row = current.get(key)
            prev_row = previous.get(key)
            cur_qty = float(cur_row["total_quantity"] or 0.0) if cur_row else 0.0
            prev_qty = float(prev_row["total_quantity"] or 0.0) if prev_row else 0.0
            cur_amt = float(cur_row["total_amount"] or 0.0) if cur_row else 0.0
            prev_amt = float(prev_row["total_amount"] or 0.0) if prev_row else 0.0
            diff_qty = cur_qty - prev_qty
            info = {
                "label": self._product_label(code, desc),
                "cur_qty": cur_qty,
                "prev_qty": prev_qty,
                "cur_amt": cur_amt,
                "prev_amt": prev_amt,
                "diff_qty": diff_qty,
                "diff_amt": cur_amt - prev_amt,
            }
            if diff_qty >= 0:
                gainers.append(info)
            else:
                losers.append(info)
        gainers.sort(key=lambda item: item["diff_qty"], reverse=True)
        losers.sort(key=lambda item: item["diff_qty"])  # ascending (most negative first)
        return gainers[:10], losers[:10]

    def _init_yoy_product_table(self, table: QTableWidget, title: str) -> None:
        table.setColumnCount(7)
        table.setHorizontalHeaderLabels(
            [
                title,
                "Unidades mes",
                "Unidades mes-1y",
                "Δ unidades",
                "Venta mes",
                "Venta mes-1y",
                "Δ venta",
            ]
        )
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setStretchLastSection(True)
        table.setMinimumHeight(200)

    def _fill_product_table(self, table: QTableWidget, items: list[dict[str, object]]) -> None:
        table.setRowCount(len(items))
        for idx, item in enumerate(items):
            table.setItem(idx, 0, QTableWidgetItem(str(item["label"])))
            table.setItem(idx, 1, QTableWidgetItem(self._format_number(float(item["cur_qty"]))))
            table.setItem(idx, 2, QTableWidgetItem(self._format_number(float(item["prev_qty"]))))
            table.setItem(idx, 3, self._delta_item(float(item["cur_qty"]), float(item["prev_qty"])))
            table.setItem(idx, 4, QTableWidgetItem(self._format_currency(float(item["cur_amt"]))))
            table.setItem(idx, 5, QTableWidgetItem(self._format_currency(float(item["prev_amt"]))))
            table.setItem(
                idx,
                6,
                self._delta_item(float(item["cur_amt"]), float(item["prev_amt"]), is_currency=True),
            )

    def _delta_item(self, current: float, previous: float, *, is_currency: bool = False) -> QTableWidgetItem:
        diff = current - previous
        text = self._format_currency(diff) if is_currency else self._format_number(diff)
        pct = ""
        if abs(previous) > 1e-6:
            pct = f" ({(diff / previous) * 100:+.1f}%)".replace(".", ",")
        item = QTableWidgetItem(f"{text}{pct}")
        actual_text = self._format_currency(current) if is_currency else self._format_number(current)
        prev_text = self._format_currency(previous) if is_currency else self._format_number(previous)
        item.setToolTip(f"Actual: {actual_text}\nAnterior: {prev_text}")
        color = TREND_COLORS["flat"]
        if diff > 0:
            color = TREND_COLORS["up"]
        elif diff < 0:
            color = TREND_COLORS["down"]
        item.setForeground(QBrush(color))
        return item

    def _product_label(self, code: str, desc: str) -> str:
        if code and desc:
            return f"{code} · {desc}"
        return desc or code or "Producto sin nombre"

    def _clear_yoy_tables(self) -> None:
        for table in (self.yoy_branch_table, self.yoy_gainers_table, self.yoy_losers_table):
            table.setRowCount(0)
        self.yoy_status_label.setText("Sin datos anuales.")
        self.yoy_quick_label.setText("")
        self.yoy_product_status.setText("Escribe o elige un producto para ver su YoY.")

    def _update_yoy_product_comparison(self) -> None:
        month_key = self.yoy_month_combo.currentData()
        if not month_key:
            self.yoy_product_status.setText("Selecciona un mes y un producto para comparar.")
            return
        label = self.yoy_product_search.text().strip()
        if not label:
            self.yoy_product_status.setText("Escribe o elige un producto para ver su YoY.")
            return
        prev_month_key = self._month_key_previous_year(str(month_key))
        if not prev_month_key:
            self.yoy_product_status.setText("No hay mes anterior disponible para comparar este producto.")
            return
        code, desc = self._product_search_map.get(label, ("", ""))
        if not code and not desc:
            for key_label, pair in self._product_search_map.items():
                if label.lower() in key_label.lower():
                    code, desc = pair
                    label = key_label
                    break
        cur_qty, cur_amt = self._product_totals_for_month(str(month_key), code, desc)
        prev_qty, prev_amt = self._product_totals_for_month(str(prev_month_key), code, desc)
        qty_pct = self._format_percent_diff(cur_qty, prev_qty)
        amt_pct = self._format_percent_diff(cur_amt, prev_amt)
        self.yoy_product_status.setText(
            f"{label}: Unidades {self._format_number(cur_qty)} vs {self._format_number(prev_qty)} ({qty_pct}) "
            f"| Venta {self._format_currency(cur_amt)} vs {self._format_currency(prev_amt)} ({amt_pct})"
        )
        self._fill_yoy_tables_for_product(code, desc, str(month_key), prev_month_key, label)

    def _product_totals_for_month(self, month_key: str, code: str, desc: str) -> tuple[float, float]:
        rows = self.repository.fetch_monthly_product_totals(month_key)
        target_label = self._product_label(code, desc).lower()
        for row in rows:
            row_code = row["product_code"] or ""
            row_desc = row["description"] or ""
            if row_code == code and row_desc == desc:
                return float(row["total_quantity"] or 0.0), float(row["total_amount"] or 0.0)
        for row in rows:
            row_label = self._product_label(row["product_code"] or "", row["description"] or "").lower()
            if target_label and target_label in row_label:
                return float(row["total_quantity"] or 0.0), float(row["total_amount"] or 0.0)
        return 0.0, 0.0

    def _format_percent_diff(self, current: float, previous: float) -> str:
        diff = current - previous
        if abs(previous) > 1e-6:
            return f"{(diff / previous) * 100:+.1f}%".replace(".", ",")
        if current > 0:
            return "+inf%"
        return "0%"

    def _month_name_es(self, month_key: str | None) -> str:
        if not month_key or "-" not in month_key:
            return month_key or ""
        try:
            year_str, month_str = month_key.split("-", 1)
            month_int = int(month_str)
            names = [
                "ENERO",
                "FEBRERO",
                "MARZO",
                "ABRIL",
                "MAYO",
                "JUNIO",
                "JULIO",
                "AGOSTO",
                "SEPTIEMBRE",
                "OCTUBRE",
                "NOVIEMBRE",
                "DICIEMBRE",
            ]
            name = names[month_int - 1] if 1 <= month_int <= 12 else month_str
            return f"{name} {year_str}"
        except Exception:
            return month_key or ""

    def open_compare_ayaa(self) -> None:
        label = self.yoy_product_search.text().strip()
        code, desc = self._product_search_map.get(label, ("", ""))
        dialog = CompararAyAADialog(self, self.repository, product_code=code or None, product_desc=desc or None, product_label=label or "Todos los productos")
        dialog.exec()

    def _fill_yoy_tables_for_product(
        self, code: str, desc: str, month_key: str, prev_month_key: Optional[str], label: str
    ) -> None:
        current = {
            row["branch"]: row
            for row in self.repository.fetch_monthly_product_branch_totals(month_key, code or None, desc or None)
        }
        prev = {
            row["branch"]: row
            for row in self.repository.fetch_monthly_product_branch_totals(prev_month_key, code or None, desc or None)
        } if prev_month_key else {}
        branches = self._sort_branches(list(set(current) | set(prev)))
        self.yoy_branch_table.setRowCount(len(branches))
        for idx, branch in enumerate(branches):
            cur_row = current.get(branch)
            prev_row = prev.get(branch)
            cur_qty = float(cur_row["total_quantity"] or 0.0) if cur_row else 0.0
            prev_qty = float(prev_row["total_quantity"] or 0.0) if prev_row else 0.0
            cur_amt = float(cur_row["total_amount"] or 0.0) if cur_row else 0.0
            prev_amt = float(prev_row["total_amount"] or 0.0) if prev_row else 0.0
            self.yoy_branch_table.setItem(idx, 0, QTableWidgetItem(branch))
            self.yoy_branch_table.setItem(idx, 1, QTableWidgetItem(self._format_number(cur_qty)))
            self.yoy_branch_table.setItem(idx, 2, QTableWidgetItem(self._format_number(prev_qty)))
            self.yoy_branch_table.setItem(idx, 3, self._delta_item(cur_qty, prev_qty))
            self.yoy_branch_table.setItem(idx, 4, QTableWidgetItem(self._format_currency(cur_amt)))
            self.yoy_branch_table.setItem(idx, 5, QTableWidgetItem(self._format_currency(prev_amt)))
            self.yoy_branch_table.setItem(idx, 6, self._delta_item(cur_amt, prev_amt, is_currency=True))

        # Productos tables: mostrar solo el producto filtrado con sus totales
        total_cur_qty = sum(float(row["total_quantity"] or 0.0) for row in current.values())
        total_prev_qty = sum(float(row["total_quantity"] or 0.0) for row in prev.values())
        total_cur_amt = sum(float(row["total_amount"] or 0.0) for row in current.values())
        total_prev_amt = sum(float(row["total_amount"] or 0.0) for row in prev.values())

        self.yoy_gainers_table.setRowCount(1)
        self.yoy_gainers_table.setItem(0, 0, QTableWidgetItem(label))
        self.yoy_gainers_table.setItem(0, 1, QTableWidgetItem(self._format_number(total_cur_qty)))
        self.yoy_gainers_table.setItem(0, 2, QTableWidgetItem(self._format_number(total_prev_qty)))
        self.yoy_gainers_table.setItem(0, 3, self._delta_item(total_cur_qty, total_prev_qty))
        self.yoy_gainers_table.setItem(0, 4, QTableWidgetItem(self._format_currency(total_cur_amt)))
        self.yoy_gainers_table.setItem(0, 5, QTableWidgetItem(self._format_currency(total_prev_amt)))
        self.yoy_gainers_table.setItem(0, 6, self._delta_item(total_cur_amt, total_prev_amt, is_currency=True))
        self.yoy_losers_table.setRowCount(0)

    def _format_yoy_quick(
        self,
        product_gainers: list[dict[str, object]],
        product_losers: list[dict[str, object]],
        branches: list[str],
        current_branches: dict[str, sqlite3.Row],
        prev_branches: dict[str, sqlite3.Row],
    ) -> str:
        alzas = 0
        caidas = 0
        for branch in branches:
            cur = current_branches.get(branch)
            prev = prev_branches.get(branch)
            cur_amt = float(cur["total_amount"] or 0.0) if cur else 0.0
            prev_amt = float(prev["total_amount"] or 0.0) if prev else 0.0
            if cur_amt > prev_amt:
                alzas += 1
            elif cur_amt < prev_amt:
                caidas += 1
        top_alza = product_gainers[0]["label"] if product_gainers else "sin datos"
        top_caida = product_losers[0]["label"] if product_losers else "sin datos"
        return (
            f"Sucursales: {len(branches)} | En alza: {alzas} | En caÃ­da: {caidas} | "
            f"Top alza: {top_alza} | Top caÃ­da: {top_caida}"
        )

    def export_yoy_csv(self) -> None:
        if self.yoy_branch_table.rowCount() == 0:
            QMessageBox.information(self, "Sin datos", "No hay comparativo para exportar.")
            return
        month_key = self.yoy_month_combo.currentData() or "comparativo"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar comparativo anual como CSV",
            f"comparativo_{month_key}.csv",
            "CSV (*.csv)",
        )
        if not path:
            return
        headers = [self.yoy_branch_table.horizontalHeaderItem(i).text() for i in range(self.yoy_branch_table.columnCount())]
        rows: list[list[str]] = []
        for row_idx in range(self.yoy_branch_table.rowCount()):
            row_vals: list[str] = []
            for col_idx in range(self.yoy_branch_table.columnCount()):
                item = self.yoy_branch_table.item(row_idx, col_idx)
                row_vals.append(item.text() if item else "")
            rows.append(row_vals)
        try:
            with open(path, "w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(fh)
                writer.writerow(headers)
                writer.writerows(rows)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "No se pudo exportar", str(exc))
            return
        QMessageBox.information(self, "Exportado", f"Archivo guardado en:\n{path}")

    def copy_selected_yoy_row(self) -> None:
        selection = self.yoy_branch_table.selectedItems()
        if not selection:
            QMessageBox.information(self, "Selecciona una fila", "Elige una sucursal en la tabla para copiarla.")
            return
        row = selection[0].row()
        values = []
        for col_idx in range(self.yoy_branch_table.columnCount()):
            item = self.yoy_branch_table.item(row, col_idx)
            values.append(item.text() if item else "")
        QApplication.clipboard().setText("\t".join(values))
        QMessageBox.information(self, "Copiado", "Fila copiada al portapapeles.")

    def _refresh_product_completer(self) -> None:
        products = self.repository.list_products()
        model_items: list[str] = []
        for row in products:
            code = (row["product_code"] or "").strip()
            desc = (row["description"] or "").strip()
            label = self._product_label(code, desc)
            model_items.append(label)
        completer_model = QStringListModel(model_items)
        self.product_completer.setModel(completer_model)

    def _refresh_yoy_product_search(self) -> None:
        products = self.repository.list_products()
        labels: list[str] = []
        self._product_search_map.clear()
        for row in products:
            code = (row["product_code"] or "").strip()
            desc = (row["description"] or "").strip()
            label = self._product_label(code, desc)
            labels.append(label)
            self._product_search_map[label] = (code, desc)
        completer_model = QStringListModel(labels)
        self.yoy_product_completer.setModel(completer_model)

    def _month_key_previous_year(self, month_key: str) -> Optional[str]:
        try:
            year_str, month_str = month_key.split("-", 1)
            year = int(year_str)
            return f"{year - 1:04d}-{month_str}"
        except Exception:
            return None

    def _format_period_label(self, row: sqlite3.Row) -> str:
        provider = row["provider"] or "Sin proveedor"
        start = row["start_date"] or ""
        end = row["end_date"] or ""
        if start and end:
            period = f"{start} → {end}"
        else:
            period = start or end or "sin fecha"
        return f"{period} · {provider}"

    def _sort_branches(self, branches: list[str]) -> list[str]:
        preferred_order = {"Casa Central": 0, "CAACUPEMI": 1, "ITAUGUA": 2}
        return sorted(branches, key=lambda name: (preferred_order.get(name, 99), name.lower()))

    def _current_product_filters(self) -> list[str]:
        base = (self.product_filter_combo.currentData() or "").strip().lower()
        extra = self.product_filter.text().strip().lower()
        product_type = (self.product_type_filter.currentData() or "").strip().lower()
        return [value for value in (base, extra, product_type) if value]

    def _current_date_filters(self) -> tuple[Optional[date], Optional[date]]:
        return self.start_date_filter.to_optional_date(), self.end_date_filter.to_optional_date()

    def _product_matches_filter(
        self,
        key: tuple[str, str],
        filters: list[str],
        product_type_label: str = "",
        brand_filter: str = "",
    ) -> bool:
        if not filters:
            return True
        code, desc = key
        haystack = f"{code} {desc}".lower()
        if product_type_label == "granos" and brand_filter != "el cacique":
            # Exclude fideos/harinas when showing granos (Cacique), even if the brand matches.
            excluded_terms = ("fideo", "harina")
            if any(term in haystack for term in excluded_terms):
                return False
        return all(term in haystack for term in filters)

    def _period_overlaps_filter(
        self,
        start_text: Optional[str],
        end_text: Optional[str],
        start_filter: Optional[date],
        end_filter: Optional[date],
    ) -> bool:
        if not start_filter and not end_filter:
            return True
        start = self._parse_date(start_text) or self._parse_date(end_text)
        end = self._parse_date(end_text) or self._parse_date(start_text)
        if start_filter and end and end < start_filter:
            return False
        if end_filter and start and start > end_filter:
            return False
        return True

    def _parse_date(self, value: Optional[str]) -> Optional[date]:
        if not value:
            return None
        try:
            return date.fromisoformat(value)
        except ValueError:
            return None

    def _compute_branch_totals(self, product_map: dict[tuple[str, str], dict[str, float]]) -> dict[str, float]:
        totals = {branch: 0.0 for branch in self.MONTHLY_BRANCHES}
        for values in product_map.values():
            for branch, qty in values.items():
                totals[branch] += qty
        return totals

    def _format_currency(self, value: float) -> str:
        return f"Gs. {value:,.0f}".replace(",", ".")

    def _format_currency_short(self, value: float) -> str:
        return f"Gs. {self._format_short(value)}"

    def _format_number(self, value: float) -> str:
        if value.is_integer():
            return f"{int(value):,}".replace(",", ".")
        return f"{value:,.2f}".replace(",", ".")

    def _format_short(self, value: float) -> str:
        """Abrevia con sufijos K/M (millones), evitando el 'B' anglosajÃ³n."""
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
        # Cambiar a separador miles con punto y decimal con coma.
        text = text.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{text}{suffix}"

    def _format_period_range(self, start: Optional[str], end: Optional[str]) -> str:
        if start and end:
            return f"{start} → {end}"
        return start or end or "-"

    def _format_month_label(self, row: sqlite3.Row) -> str:
        period = self._format_period_range(row["start_date"], row["end_date"])
        month_name = self._month_name_es(str(row["month_key"]))
        return f"{month_name} · {period}"

    def _format_branch_trend_text(
        self, current: float, previous: float, *, is_total: bool = False, is_currency: bool = False
    ) -> tuple[str, Optional[QColor]]:
        diff = current - previous
        if abs(diff) <= 1e-6:
            trend = "flat"
        elif diff > 0:
            trend = "up"
        else:
            trend = "down"
        if previous > 0:
            pct = (diff / previous) * 100
        elif current > 0:
            pct = 100.0
        else:
            pct = 0.0
        pct_text = f"{pct:+.1f}%".replace(".", ",")
        symbol = TREND_SYMBOLS[trend]
        prefix = "\u2211 " if is_total else ""
        value_text = self._format_currency(current) if is_currency else self._format_number(current)
        formatted = f"{prefix}{symbol} {value_text} ({pct_text})"
        return formatted, TREND_COLORS.get(trend)

    def _update_monthly_headers(self, metric: str) -> None:
        base_headers = ["Mes / Producto"] + self.MONTHLY_BRANCHES
        if metric == "amount":
            total_label = "Total monto (Gs.)"
        else:
            total_label = "Total unidades"
        headers = base_headers + [total_label]
        if self.monthly_tree.columnCount() == len(headers):
            self.monthly_tree.setHeaderLabels(headers)

    def _open_product_chart(self, item: QTreeWidgetItem, _column: int) -> None:
        data = item.data(0, Qt.UserRole)
        if not isinstance(data, dict) or not data:
            return
        code = data.get("code") or ""
        description = data.get("description") or ""
        history = self.repository.fetch_product_history(code or None, description or None)
        if not history:
            QMessageBox.information(self, "Sin datos", "No hay historial suficiente para graficar este producto.")
            return
        label = description or code or "Producto"
        dialog = ProductTrendDialog(self, label, history)
        dialog.exec()

    def _apply_selection_style(self) -> None:
        highlight = QColor("#d0e5ff")
        text_color = QColor("#0f3057")

        def update_palette(widget: QWidget) -> None:
            palette = widget.palette()
            palette.setColor(QPalette.ColorRole.Highlight, highlight)
            palette.setColor(QPalette.ColorRole.HighlightedText, text_color)
            widget.setPalette(palette)

        for widget in (
            self.period_list,
            self.table,
            self.monthly_tree,
            self.yoy_branch_table,
            self.yoy_gainers_table,
            self.yoy_losers_table,
            self.provider_yearly_tab.table,
            self.provider_monthly_tab.table,
        ):
            update_palette(widget)


class ProductTrendDialog(QDialog):
    """Dialogo que grafica la evolucion mensual del producto."""

    def __init__(self, parent: QWidget, product_label: str, history: list[sqlite3.Row]) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Evolucion anual · {product_label}")
        self.resize(720, 420)
        layout = QVBoxLayout(self)
        figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(figure)
        layout.addWidget(self.canvas, 1)
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("Modo grÃ¡fico:"))
        self.mode_selector = QComboBox()
        self.mode_selector.addItems(["Lineas", "Barras"])
        self.mode_selector.currentIndexChanged.connect(lambda *_: self._plot_history(product_label, history))
        control_layout.addWidget(self.mode_selector)
        layout.addLayout(control_layout)
        layout.addWidget(self.canvas, 1)
        self._plot_history(product_label, history)

    def _plot_history(self, product_label: str, history: list[sqlite3.Row]) -> None:
        figure = self.canvas.figure
        figure.clear()
        ax = figure.add_subplot(111)
        month_keys = sorted({row["month_key"] for row in history})
        branches = sorted({row["branch"] for row in history})
        data_map: dict[tuple[str, str], float] = {}
        for row in history:
            key = (row["month_key"], row["branch"])
            data_map[key] = float(row["total_quantity"] or 0.0)
        mode = self.mode_selector.currentText()
        width = 0.8 / max(len(branches), 1)
        x_positions = list(range(len(month_keys)))
        for idx, branch in enumerate(branches):
            values = [data_map.get((month, branch), 0.0) for month in month_keys]
            if mode == "Barras":
                offsets = [x + (idx - (len(branches) - 1) / 2) * width for x in x_positions]
                ax.bar(offsets, values, width=width, label=branch)
            else:
                ax.plot(x_positions, values, marker="o", label=branch)
        ax.set_xticks(x_positions, month_keys, rotation=45, ha="right")
        ax.set_title(f"Unidades vendidas · {product_label}")
        ax.set_xlabel("Mes")
        ax.set_ylabel("Cantidad")
        ax.grid(True, linestyle="--", alpha=0.4)
        ax.legend()
        figure.tight_layout()
        self.canvas.draw_idle()


def run_app(repository: SalesRepository, importer: ExcelImporter, default_folder: Optional[Path] = None) -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow(repository, importer, default_folder)
    window.show()
    app.exec()


class QDateEditWithReset(QDateEdit):
    """QDateEdit que permite dejar el filtro vacÃ­o usando un valor mÃ­nimo como sentinel."""

    def __init__(self) -> None:
        super().__init__()
        sentinel = QDate(1900, 1, 1)
        self.setMinimumDate(sentinel)
        self.setDate(sentinel)
        self.setSpecialValueText("Sin filtro")
        self.setDisplayFormat("yyyy-MM-dd")
        self.setCalendarPopup(True)

    def to_optional_date(self) -> Optional[date]:
        if self.date() == self.minimumDate():
            return None
        return self.date().toPython()




















