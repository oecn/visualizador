"""Dialogo de inicio de sesion."""

from __future__ import annotations

from typing import Optional

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ..database import SalesRepository


class LoginDialog(QDialog):
    """Dialogo modal para validar usuario."""

    def __init__(self, parent: Optional[QWidget], repository: SalesRepository) -> None:
        super().__init__(parent)
        self.repository = repository
        self.user_row = None
        self.setWindowTitle("Inicio de sesion")
        self.setModal(True)
        self.setMinimumWidth(420)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        logo_path = Path(__file__).resolve().parents[2] / "logos" / "logo-azul.png"
        pixmap = QPixmap(str(logo_path))
        if not pixmap.isNull():
            logo_label.setPixmap(pixmap.scaledToWidth(180, Qt.SmoothTransformation))
        layout.addWidget(logo_label)

        title = QLabel("Bienvenido")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("Ingresa tus credenciales para continuar")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #5a6b7a;")
        layout.addWidget(subtitle)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)
        form.setFormAlignment(Qt.AlignLeft)
        form.setContentsMargins(8, 0, 8, 0)
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Usuario")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("Contraseña")
        form.addRow("", self.username_input)
        form.addRow("", self.password_input)
        layout.addLayout(form)

        button_layout = QHBoxLayout()
        login_btn = QPushButton("Ingresar")
        login_btn.setDefault(True)
        login_btn.setMinimumHeight(36)
        cancel_btn = QPushButton("Salir")
        cancel_btn.setMinimumHeight(36)
        login_btn.clicked.connect(self._attempt_login)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(login_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

        self.setStyleSheet(
            """
            QDialog {
                background: #f6f7fb;
            }
            QLabel {
                color: #1f2a37;
            }
            QLineEdit {
                padding: 8px 10px;
                border: 1px solid #d4dbe3;
                border-radius: 6px;
                background: #ffffff;
            }
            QLineEdit:focus {
                border-color: #2b6cb0;
            }
            QPushButton {
                border-radius: 6px;
                padding: 6px 14px;
            }
            QPushButton:default {
                background: #1e64b7;
                color: white;
            }
            QPushButton:default:hover {
                background: #1554a4;
            }
            """
        )

    def _attempt_login(self) -> None:
        username = self.username_input.text().strip()
        password = self.password_input.text()
        row = self.repository.verify_user(username, password)
        if not row:
            QMessageBox.warning(self, "Acceso denegado", "Usuario o contrasena incorrectos.")
            return
        self.user_row = row
        self.accept()
