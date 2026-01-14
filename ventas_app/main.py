"""Punto de entrada CLI para la app PySide6 + SQLite."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication

from .database import SalesRepository
from .excel_importer import ExcelImporter
from .ui.main_window import MainWindow


BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_DB = BASE_DIR / "ventas.sqlite3"
DEFAULT_DATA_DIR = BASE_DIR / "ventas"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Seguimiento de ventas por sucursal (PySide6).")
    parser.add_argument(
        "--db",
        type=Path,
        default=DEFAULT_DB,
        help=f"Ruta del archivo SQLite (por defecto {DEFAULT_DB.name}).",
    )
    parser.add_argument(
        "--folder",
        type=Path,
        default=DEFAULT_DATA_DIR,
        help="Carpeta sugerida para abrir el dialogo de importacion.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    repository = SalesRepository(args.db)
    importer = ExcelImporter()
    app = QApplication.instance() or QApplication(sys.argv)
    window = MainWindow(repository, importer, default_folder=args.folder)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
