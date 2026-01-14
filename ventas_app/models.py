"""Estructuras de datos compartidas por la app de ventas."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional


@dataclass(slots=True)
class SalesRecord:
    """Representa una fila de ventas para una sucursal."""

    branch: str
    product_code: str
    description: str
    quantity: float
    amount: float
    share_percentage: Optional[float] = None


@dataclass(slots=True)
class SalesBatch:
    """Lote de ventas importado desde un Excel."""

    source_file: Path
    provider: Optional[str] = None
    brand: Optional[str] = None
    plan_name: Optional[str] = None
    period_start: Optional[date] = None
    period_end: Optional[date] = None
    records: list[SalesRecord] = field(default_factory=list)

    def label(self) -> str:
        """Devuelve una etiqueta legible para mostrar en la UI."""
        period = ""
        if self.period_start and self.period_end:
            period = f"{self.period_start:%Y-%m-%d} → {self.period_end:%Y-%m-%d}"
        elif self.period_start:
            period = f"Desde {self.period_start:%Y-%m-%d}"
        elif self.period_end:
            period = f"Hasta {self.period_end:%Y-%m-%d}"
        provider = self.provider or "Sin proveedor"
        if period:
            return f"{provider} · {period}"
        return provider

