"""Panel de usuario y administracion."""

from __future__ import annotations

from datetime import date
from typing import Optional

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QDateEdit,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..database import SalesRepository


class UserPanel(QWidget):
    """Panel para cambiar contrasena y crear usuarios (admin)."""

    def __init__(self, repository: SalesRepository, username: str, is_admin: bool) -> None:
        super().__init__()
        self.repository = repository
        self.username = username
        self.is_admin = is_admin
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Usuario actual: {self.username}"))

        password_group = QGroupBox("Cambiar contrasena")
        password_layout = QFormLayout(password_group)
        self.old_password_input = QLineEdit()
        self.old_password_input.setEchoMode(QLineEdit.Password)
        self.new_password_input = QLineEdit()
        self.new_password_input.setEchoMode(QLineEdit.Password)
        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setEchoMode(QLineEdit.Password)
        password_layout.addRow("Contrasena actual:", self.old_password_input)
        password_layout.addRow("Nueva contrasena:", self.new_password_input)
        password_layout.addRow("Confirmar nueva:", self.confirm_password_input)
        change_btn = QPushButton("Actualizar contrasena")
        change_btn.clicked.connect(self._change_password)
        password_layout.addRow("", change_btn)
        layout.addWidget(password_group)

        admin_group = QGroupBox("Administracion de usuarios")
        admin_layout = QFormLayout(admin_group)
        self.new_username_input = QLineEdit()
        self.new_admin_checkbox = QCheckBox("Administrador")
        admin_layout.addRow("Usuario nuevo:", self.new_username_input)
        admin_layout.addRow("", self.new_admin_checkbox)
        admin_layout.addRow("", QLabel("Contrasena inicial = nombre de usuario"))
        create_btn = QPushButton("Crear usuario")
        create_btn.clicked.connect(self._create_user)
        admin_layout.addRow("", create_btn)
        admin_group.setVisible(self.is_admin)
        layout.addWidget(admin_group)

        audit_group = QGroupBox("Auditoria de periodos")
        audit_layout = QVBoxLayout(audit_group)
        filter_layout = QHBoxLayout()
        self.audit_user_filter = QLineEdit()
        self.audit_user_filter.setPlaceholderText("Usuario")
        self.audit_action_filter = QComboBox()
        self.audit_action_filter.addItem("Todas", "")
        self.audit_action_filter.addItem("IMPORT", "IMPORT")
        self.audit_action_filter.addItem("DELETE", "DELETE")
        self.audit_start_date = QDateEditWithReset()
        self.audit_end_date = QDateEditWithReset()
        filter_layout.addWidget(QLabel("Usuario:"))
        filter_layout.addWidget(self.audit_user_filter, 1)
        filter_layout.addWidget(QLabel("Accion:"))
        filter_layout.addWidget(self.audit_action_filter)
        filter_layout.addWidget(QLabel("Desde:"))
        filter_layout.addWidget(self.audit_start_date)
        filter_layout.addWidget(QLabel("Hasta:"))
        filter_layout.addWidget(self.audit_end_date)
        audit_layout.addLayout(filter_layout)
        refresh_layout = QHBoxLayout()
        self.audit_refresh_btn = QPushButton("Refrescar auditoria")
        self.audit_refresh_btn.clicked.connect(self._load_audit)
        self.audit_clear_btn = QPushButton("Limpiar filtros")
        self.audit_clear_btn.clicked.connect(self._clear_audit_filters)
        refresh_layout.addStretch()
        refresh_layout.addWidget(self.audit_clear_btn)
        refresh_layout.addWidget(self.audit_refresh_btn)
        audit_layout.addLayout(refresh_layout)
        self.audit_table = QTableWidget()
        self.audit_table.setColumnCount(7)
        self.audit_table.setHorizontalHeaderLabels(
            [
                "Fecha",
                "Accion",
                "Usuario",
                "Cargado por",
                "Archivo",
                "Proveedor",
                "Periodo",
            ]
        )
        self.audit_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.audit_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.audit_table.setAlternatingRowColors(True)
        header = self.audit_table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for idx in range(1, self.audit_table.columnCount() - 1):
            header.setSectionResizeMode(idx, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(self.audit_table.columnCount() - 1, QHeaderView.Stretch)
        audit_layout.addWidget(self.audit_table)
        audit_group.setVisible(self.is_admin)
        layout.addWidget(audit_group)
        layout.addStretch()
        if self.is_admin:
            self.audit_user_filter.editingFinished.connect(self._load_audit)
            self.audit_action_filter.currentIndexChanged.connect(lambda *_: self._load_audit())
            self.audit_start_date.dateChanged.connect(lambda *_: self._load_audit())
            self.audit_end_date.dateChanged.connect(lambda *_: self._load_audit())
            self._load_audit()

    def _change_password(self) -> None:
        old = self.old_password_input.text()
        new = self.new_password_input.text()
        confirm = self.confirm_password_input.text()
        if not old or not new or not confirm:
            QMessageBox.warning(self, "Datos incompletos", "Completa todos los campos.")
            return
        if new != confirm:
            QMessageBox.warning(self, "No coincide", "La nueva contrasena no coincide.")
            return
        if not self.repository.change_password(self.username, old, new):
            QMessageBox.warning(self, "Error", "Contrasena actual incorrecta.")
            return
        self.old_password_input.clear()
        self.new_password_input.clear()
        self.confirm_password_input.clear()
        QMessageBox.information(self, "Actualizado", "Contrasena actualizada correctamente.")

    def _create_user(self) -> None:
        username = self.new_username_input.text().strip()
        if not username:
            QMessageBox.warning(self, "Datos incompletos", "Ingresa un nombre de usuario.")
            return
        is_admin = self.new_admin_checkbox.isChecked()
        created = self.repository.create_user(username, username, is_admin=is_admin)
        if not created:
            QMessageBox.warning(self, "No creado", "El usuario ya existe o es invalido.")
            return
        self.new_username_input.clear()
        self.new_admin_checkbox.setChecked(False)
        QMessageBox.information(self, "Creado", "Usuario creado. Contrasena inicial = usuario.")

    def refresh_audit(self) -> None:
        if not self.is_admin:
            return
        self._load_audit()

    def _load_audit(self) -> None:
        if not self.is_admin:
            return
        username = self.audit_user_filter.text().strip()
        action = self.audit_action_filter.currentData() or None
        start = self.audit_start_date.to_optional_date()
        end = self.audit_end_date.to_optional_date()
        rows = self.repository.fetch_period_audit(
            username=username or None,
            action=action,
            start_date=start,
            end_date=end,
        )
        self.audit_table.setRowCount(len(rows))
        for idx, row in enumerate(rows):
            start = row["start_date"] or ""
            end = row["end_date"] or ""
            period_text = f"{start} - {end}".strip(" -")
            self.audit_table.setItem(idx, 0, QTableWidgetItem(str(row["created_at"])))
            self.audit_table.setItem(idx, 1, QTableWidgetItem(str(row["action"])))
            self.audit_table.setItem(idx, 2, QTableWidgetItem(str(row["username"])))
            self.audit_table.setItem(idx, 3, QTableWidgetItem(str(row["created_by"] or "")))
            self.audit_table.setItem(idx, 4, QTableWidgetItem(str(row["source_file"] or "")))
            self.audit_table.setItem(idx, 5, QTableWidgetItem(str(row["provider"] or "")))
            self.audit_table.setItem(idx, 6, QTableWidgetItem(period_text))

    def _clear_audit_filters(self) -> None:
        self.audit_user_filter.clear()
        self.audit_action_filter.setCurrentIndex(0)
        self.audit_start_date.reset()
        self.audit_end_date.reset()
        self._load_audit()


class QDateEditWithReset(QDateEdit):
    """QDateEdit que permite dejar el filtro vacio usando un valor minimo."""

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

    def reset(self) -> None:
        self.setDate(self.minimumDate())
