"""Capa de acceso a datos usando SQLite."""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from typing import Iterator, Optional

from .models import SalesBatch, SalesRecord


class SalesRepository:
    """Encapsula todas las operaciones sobre la base SQLite."""

    def __init__(self, db_path: Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._conn = sqlite3.connect(self.db_path)
        self._conn.row_factory = sqlite3.Row
        self._ensure_schema()

    def _ensure_schema(self) -> None:
        schema = """
        PRAGMA foreign_keys = ON;
        CREATE TABLE IF NOT EXISTS sales_periods (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            provider TEXT,
            brand TEXT,
            plan_name TEXT,
            start_date TEXT,
            end_date TEXT,
            source_file TEXT NOT NULL UNIQUE
        );

        CREATE TABLE IF NOT EXISTS sales_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER NOT NULL,
            branch TEXT NOT NULL,
            product_code TEXT,
            description TEXT,
            quantity REAL,
            amount REAL,
            share REAL,
            FOREIGN KEY(period_id) REFERENCES sales_periods(id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_sales_records_period ON sales_records(period_id);
        CREATE INDEX IF NOT EXISTS idx_sales_records_branch ON sales_records(branch);
        """
        self._conn.executescript(schema)

    def close(self) -> None:
        self._conn.close()

    @contextmanager
    def transaction(self) -> Iterator[sqlite3.Connection]:
        try:
            yield self._conn
            self._conn.commit()
        except Exception:
            self._conn.rollback()
            raise

    def store_batch(self, batch: SalesBatch) -> int:
        """Inserta o actualiza un lote completo y devuelve las filas afectadas."""
        if not batch.records:
            return 0
        start = self._date_to_text(batch.period_start)
        end = self._date_to_text(batch.period_end)
        with self.transaction():
            self._conn.execute(
                """
                INSERT INTO sales_periods (provider, brand, plan_name, start_date, end_date, source_file)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(source_file) DO UPDATE SET
                    provider=excluded.provider,
                    brand=excluded.brand,
                    plan_name=excluded.plan_name,
                    start_date=excluded.start_date,
                    end_date=excluded.end_date
                """,
                (
                    batch.provider,
                    batch.brand,
                    batch.plan_name,
                    start,
                    end,
                    batch.source_file.name,
                ),
            )
            period_id = self._get_period_id(batch.source_file.name)
            self._conn.execute("DELETE FROM sales_records WHERE period_id = ?", (period_id,))
            payload = [
                (
                    period_id,
                    record.branch,
                    record.product_code,
                    record.description,
                    record.quantity,
                    record.amount,
                    record.share_percentage,
                )
                for record in batch.records
            ]
            self._conn.executemany(
                """
                INSERT INTO sales_records (
                    period_id, branch, product_code, description, quantity, amount, share
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                payload,
            )
            return len(payload)

    def list_periods(self) -> list[sqlite3.Row]:
        query = """
        SELECT id, provider, brand, plan_name, start_date, end_date, source_file
        FROM sales_periods
        ORDER BY COALESCE(start_date, end_date) DESC, id DESC
        """
        return list(self._conn.execute(query))

    def delete_period(self, period_id: int) -> None:
        """Elimina un periodo y sus registros asociados."""
        with self.transaction():
            self._conn.execute("DELETE FROM sales_records WHERE period_id = ?", (period_id,))
            self._conn.execute("DELETE FROM sales_periods WHERE id = ?", (period_id,))
    def clear_all(self) -> None:
        """Elimina todas las importaciones almacenadas."""
        with self.transaction():
            self._conn.execute("DELETE FROM sales_records")
            self._conn.execute("DELETE FROM sales_periods")
        self._conn.execute("VACUUM")

    def list_branches(self, period_id: int) -> list[str]:
        query = """
        SELECT DISTINCT branch FROM sales_records
        WHERE period_id = ?
        ORDER BY branch COLLATE NOCASE
        """
        rows = self._conn.execute(query, (period_id,))
        return [row["branch"] for row in rows]

    def fetch_records(self, period_id: int, branch: Optional[str] = None) -> list[sqlite3.Row]:
        base = """
        SELECT sr.branch, sr.product_code, sr.description, sr.quantity, sr.amount, sr.share,
               sp.start_date, sp.end_date
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE sr.period_id = ?
        """
        params: list[object] = [period_id]
        if branch:
            base += " AND sr.branch = ?"
            params.append(branch)
        base += " ORDER BY sr.branch, sr.amount DESC"
        return list(self._conn.execute(base, params))

    def _get_period_id(self, source_file: str) -> int:
        row = self._conn.execute(
            "SELECT id FROM sales_periods WHERE source_file = ?", (source_file,)
        ).fetchone()
        if not row:
            raise LookupError(f"No existe el periodo para {source_file}")
        return int(row["id"])

    def _date_to_text(self, value: Optional[date]) -> Optional[str]:
        return value.isoformat() if value else None

    def list_months(self, provider: str | None = None) -> list[sqlite3.Row]:
        query = """
        SELECT DISTINCT
            strftime('%Y-%m', date(COALESCE(start_date, end_date))) AS month_key,
            start_date,
            end_date,
            COALESCE(start_date, end_date) AS sort_date
        FROM sales_periods
        """
        params: list[object] = []
        where = ["COALESCE(start_date, end_date) IS NOT NULL"]
        if provider:
            where.append("provider = ?")
            params.append(provider)
        query += f" WHERE {' AND '.join(where)} ORDER BY sort_date DESC"
        return list(self._conn.execute(query, params))

    def fetch_provider_monthly_totals(
        self,
        provider: str | None = None,
        year: int | None = None,
    ) -> list[sqlite3.Row]:
        where = ["COALESCE(sp.start_date, sp.end_date) IS NOT NULL"]
        params: list[object] = []
        if provider:
            where.append("sp.provider = ?")
            params.append(provider)
        if year:
            where.append("strftime('%Y', date(COALESCE(sp.start_date, sp.end_date))) = ?")
            params.append(str(year))
        query = f"""
        SELECT
            strftime('%Y-%m', date(COALESCE(sp.start_date, sp.end_date))) AS month_key,
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE {' AND '.join(where)}
        GROUP BY month_key
        ORDER BY month_key ASC
        """
        return list(self._conn.execute(query, params))

    def list_years(self) -> list[int]:
        query = """
        SELECT DISTINCT CAST(strftime('%Y', date(COALESCE(start_date, end_date))) AS INTEGER) AS year
        FROM sales_periods
        WHERE COALESCE(start_date, end_date) IS NOT NULL
        ORDER BY year DESC
        """
        return [int(row["year"]) for row in self._conn.execute(query)]

    def fetch_provider_yearly_totals(self, year: int) -> list[sqlite3.Row]:
        query = """
        SELECT
            CASE
                WHEN sp.provider IS NULL OR TRIM(sp.provider) = '' THEN 'Sin proveedor'
                ELSE sp.provider
            END AS provider,
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE strftime('%Y', date(COALESCE(sp.start_date, sp.end_date))) = ?
        GROUP BY provider
        ORDER BY total_amount DESC
        """
        return list(self._conn.execute(query, (str(year),)))

    def fetch_yearly_branch_totals(self, year: int) -> list[sqlite3.Row]:
        query = """
        SELECT
            sr.branch,
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE strftime('%Y', date(COALESCE(sp.start_date, sp.end_date))) = ?
        GROUP BY sr.branch
        """
        return list(self._conn.execute(query, (str(year),)))

    def fetch_yearly_totals(self, provider: str | None = None, search_text: str | None = None) -> list[sqlite3.Row]:
        """
        Devuelve totales anuales opcionalmente filtrados por proveedor y por texto de producto.

        search_text: se busca en descripcion + codigo (case-insensitive) y cada palabra debe aparecer.
        """
        where = []
        params: list[object] = []
        if provider:
            where.append("sp.provider = ?")
            params.append(provider)
        if search_text:
            terms = [term.strip() for term in search_text.strip().lower().split() if term.strip()]
            for term in terms:
                where.append("LOWER(COALESCE(sr.description, '') || ' ' || COALESCE(sr.product_code, '')) LIKE ?")
                params.append(f"%{term}%")
        where_clause = f"WHERE {' AND '.join(where)}" if where else ""
        query = f"""
        SELECT
            CAST(strftime('%Y', date(COALESCE(sp.start_date, sp.end_date))) AS INTEGER) AS year,
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        {where_clause}
        GROUP BY year
        ORDER BY year ASC
        """
        return list(self._conn.execute(query, params))

    def fetch_yearly_product_totals(self, year: int) -> list[sqlite3.Row]:
        query = """
        SELECT
            sr.product_code,
            sr.description,
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE strftime('%Y', date(COALESCE(sp.start_date, sp.end_date))) = ?
        GROUP BY sr.product_code, sr.description
        """
        return list(self._conn.execute(query, (str(year),)))

    def fetch_monthly_branch_totals(self, month_key: str) -> list[sqlite3.Row]:
        query = """
        SELECT
            sr.branch,
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE strftime('%Y-%m', date(COALESCE(sp.start_date, sp.end_date))) = ?
        GROUP BY sr.branch
        """
        return list(self._conn.execute(query, (month_key,)))

    def fetch_monthly_product_totals(self, month_key: str) -> list[sqlite3.Row]:
        query = """
        SELECT
            sr.product_code,
            sr.description,
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE strftime('%Y-%m', date(COALESCE(sp.start_date, sp.end_date))) = ?
        GROUP BY sr.product_code, sr.description
        """
        return list(self._conn.execute(query, (month_key,)))

    def fetch_monthly_product_branch_totals(
        self, month_key: str, product_code: str | None, description: str | None
    ) -> list[sqlite3.Row]:
        where = ["strftime('%Y-%m', date(COALESCE(sp.start_date, sp.end_date))) = ?"]
        params: list[object] = [month_key]
        if product_code:
            where.append("sr.product_code = ?")
            params.append(product_code)
        if description:
            where.append("sr.description = ?")
            params.append(description)
        query = f"""
        SELECT
            sr.branch,
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE {' AND '.join(where)}
        GROUP BY sr.branch
        """
        return list(self._conn.execute(query, params))

    def list_products(self) -> list[sqlite3.Row]:
        query = """
        SELECT DISTINCT
            COALESCE(sr.product_code, '') AS product_code,
            COALESCE(sr.description, '') AS description
        FROM sales_records sr
        ORDER BY description, product_code
        """
        return list(self._conn.execute(query))

    def fetch_monthly_summary(self, month_key: str, provider: str | None = None) -> list[sqlite3.Row]:
        where = ["strftime('%Y-%m', date(COALESCE(sp.start_date, sp.end_date))) = ?"]
        params: list[object] = [month_key]
        if provider:
            where.append("sp.provider = ?")
            params.append(provider)
        query = f"""
        SELECT
            sr.product_code,
            sr.description,
            sr.branch,
            SUM(sr.quantity) AS total_quantity
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE {' AND '.join(where)}
        GROUP BY sr.product_code, sr.description, sr.branch
        ORDER BY sr.description
        """
        return list(self._conn.execute(query, params))

    def list_providers(self) -> list[str]:
        query = """
        SELECT DISTINCT provider
        FROM sales_periods
        WHERE provider IS NOT NULL AND provider <> ''
        ORDER BY provider COLLATE NOCASE
        """
        return [row["provider"] for row in self._conn.execute(query)]

    def fetch_product_history(
        self,
        product_code: str | None,
        description: str | None,
    ) -> list[sqlite3.Row]:
        where = []
        params: list[object] = []
        if product_code:
            where.append("sr.product_code = ?")
            params.append(product_code)
        if description:
            where.append("sr.description = ?")
            params.append(description)
        if not where:
            return []
        query = f"""
        SELECT
            strftime('%Y-%m', date(COALESCE(sp.start_date, sp.end_date))) AS month_key,
            sr.branch,
            SUM(sr.quantity) AS total_quantity
        FROM sales_records sr
        JOIN sales_periods sp ON sp.id = sr.period_id
        WHERE {' AND '.join(where)}
        GROUP BY month_key, sr.branch
        ORDER BY month_key ASC
        """
        return list(self._conn.execute(query, params))
