"""Importador de los reportes Excel generados por la carpeta ventas/."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd

from .models import SalesBatch, SalesRecord


# Etiquetas que suelen aparecer como encabezados en los Excel.
METADATA_LABELS = {
    "Proveedor": "provider",
    "Marca": "brand",
    "Planilla Ofertas": "plan_name",
    "Planilla Oferta": "plan_name",
}

# Mapeos de textos encontrados en el Excel a nombres de sucursales.
BRANCH_ALIASES = {
    "COMERCIAL EL CACIQUE": "Casa Central",
    "COMERCIAL EL CACIQUE S.R.L.": "Casa Central",
    "CASA CENTRAL": "Casa Central",
}

# Palabra clave opcional para filtrar. Si es None o vacio, se importan todos los productos.
REQUIRED_KEYWORD: Optional[str] = None


@dataclass
class ImportResult:
    """Resultado del intento de importacion para mostrar en la UI."""

    batch: Optional[SalesBatch]
    error: Optional[str] = None


class ExcelImporter:
    """Lee los archivos Excel y los normaliza en SalesBatch + SalesRecord."""

    def __init__(self, sheet_name: str | int | None = 0) -> None:
        self.sheet_name = sheet_name

    def load(self, excel_path: Path) -> SalesBatch:
        """Carga un archivo Excel y devuelve el lote listo para persistir."""
        df = pd.read_excel(excel_path, sheet_name=self.sheet_name, header=None)
        header_idx = self._find_header_row(df)
        metadata = self._extract_metadata(df)
        records = self._extract_records(df, start_row=header_idx + 1)
        if not records:
            raise ValueError(f"No se encontraron registros de venta en {excel_path.name}")
        period_start, period_end = self._find_period(df)
        return SalesBatch(
            source_file=excel_path,
            provider=metadata.get("provider"),
            brand=metadata.get("brand"),
            plan_name=metadata.get("plan_name"),
            period_start=period_start,
            period_end=period_end,
            records=records,
        )

    def _find_header_row(self, df: pd.DataFrame) -> int:
        """Busca la fila donde aparecen los encabezados Producto/Cantidad/Venta."""
        for idx in range(df.shape[0]):
            row = df.iloc[idx]
            as_text = {self._normalize_text(value) for value in row if isinstance(value, str)}
            if {"producto", "cantidad", "venta"}.issubset(as_text):
                return idx
        raise ValueError("No se pudo localizar la fila de encabezados")

    def _extract_metadata(self, df: pd.DataFrame) -> dict[str, Any]:
        """Busca los valores de proveedor/marca/planilla dentro de la planilla."""
        metadata: dict[str, Any] = {}
        for label, key in METADATA_LABELS.items():
            value = self._find_value_near_label(df, label)
            if value is not None:
                metadata[key] = self._clean_cell(value)
        return metadata

    def _find_period(self, df: pd.DataFrame) -> tuple[Optional[date], Optional[date]]:
        """Obtiene las fechas de inicio/fin del periodo."""
        start = None
        end = None
        for row_idx in range(df.shape[0]):
            for col_idx in range(df.shape[1]):
                raw = df.iat[row_idx, col_idx]
                if isinstance(raw, str):
                    text = self._normalize_text(raw)
                    if text == "fecha desde" and col_idx + 1 < df.shape[1]:
                        start = self._as_date(df.iat[row_idx, col_idx + 1])
                    if text == "fecha hasta" and col_idx + 1 < df.shape[1]:
                        end = self._as_date(df.iat[row_idx, col_idx + 1])
        return start, end

    def _extract_records(self, df: pd.DataFrame, start_row: int) -> list[SalesRecord]:
        """Convierte las filas del DataFrame a registros normalizados."""
        records: list[SalesRecord] = []
        current_branch: Optional[str] = None
        seen: set[tuple[str, str, str]] = set()

        for idx in range(start_row, df.shape[0]):
            row = df.iloc[idx]
            branch_name = self._pick_branch(row)
            if branch_name:
                current_branch = branch_name
                continue

            code = self._format_code(row.iloc[1] if len(row) > 1 else None)
            description = self._clean_cell(row.iloc[2] if len(row) > 2 else None)
            quantity = self._to_number(row.iloc[3] if len(row) > 3 else None)
            amount = self._to_number(row.iloc[4] if len(row) > 4 else None)
            share = self._to_number(row.iloc[5] if len(row) > 5 else None)

            if not current_branch:
                # Todavia no apareció una sucursal en el archivo.
                continue
            if not (code or description):
                continue

            if not self._contains_keyword(description):
                continue

            key = (current_branch, code, description, quantity, amount)
            if key in seen:
                # Evita registros duplicados cuando los Excel repiten filas.
                continue
            seen.add(key)

            records.append(
                SalesRecord(
                    branch=current_branch,
                    product_code=code,
                    description=description,
                    quantity=quantity or 0.0,
                    amount=amount or 0.0,
                    share_percentage=share,
                )
            )
        return records

    def _pick_branch(self, row: pd.Series) -> Optional[str]:
        for value in row:
            if isinstance(value, str):
                text = value.strip()
                if text.upper().startswith("SUCURSAL"):
                    branch = text.replace("SUCURSAL", "", 1).strip() or text.strip()
                    return self._normalize_branch(branch)
                alias = BRANCH_ALIASES.get(text.strip().upper())
                if alias:
                    return alias
        return None

    def _find_value_near_label(self, df: pd.DataFrame, label: str) -> Any:
        """Ubiqua el valor que sigue a una etiqueta (ej: Proveedor -> CEREALES)."""
        label_norm = self._normalize_text(label)
        for row_idx in range(df.shape[0]):
            for col_idx in range(df.shape[1]):
                raw = df.iat[row_idx, col_idx]
                if isinstance(raw, str) and self._normalize_text(raw) == label_norm:
                    if col_idx + 1 < df.shape[1]:
                        return df.iat[row_idx, col_idx + 1]
        return None

    def _normalize_text(self, value: Any) -> str:
        if not isinstance(value, str):
            return ""
        return value.strip().lower()

    def _normalize_branch(self, branch: str) -> str:
        key = branch.strip().upper()
        return BRANCH_ALIASES.get(key, branch.strip())

    def _contains_keyword(self, text: str) -> bool:
        if not text:
            return False
        if not REQUIRED_KEYWORD:
            return True
        return REQUIRED_KEYWORD.lower() in text.lower()

    def _clean_cell(self, value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float):
            if pd.isna(value):
                return ""
            if value.is_integer():
                return str(int(value))
            return f"{value:.2f}"
        return str(value).strip()

    def _format_code(self, value: Any) -> str:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
        if isinstance(value, str):
            return value.strip()
        if isinstance(value, (int, float)):
            as_int = int(value)
            if as_int == value:
                return str(as_int)
            return str(value)
        return str(value)

    def _to_number(self, value: Any) -> Optional[float]:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        text = str(value).strip()
        if not text:
            return None
        normalized = text.replace("%", "").replace(".", "").replace(",", ".")
        try:
            return float(normalized)
        except ValueError:
            return None

    def _as_date(self, value: Any) -> Optional[date]:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        if isinstance(value, pd.Timestamp):
            return value.date()
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        text = str(value).strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None
