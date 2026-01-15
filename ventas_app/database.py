"""Capa de acceso a datos usando SQLite."""

from __future__ import annotations

import hashlib
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
            source_file TEXT NOT NULL UNIQUE,
            created_by TEXT
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

        CREATE TABLE IF NOT EXISTS period_audit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            action TEXT NOT NULL,
            username TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            source_file TEXT,
            provider TEXT,
            start_date TEXT,
            end_date TEXT,
            created_by TEXT,
            FOREIGN KEY(period_id) REFERENCES sales_periods(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            is_admin INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL DEFAULT (datetime('now'))
        );
        """
        self._conn.executescript(schema)
        self._ensure_schema_columns()
        self._ensure_default_admin()

    def _ensure_schema_columns(self) -> None:
        columns = {row["name"] for row in self._conn.execute("PRAGMA table_info(sales_periods)")}
        if "created_by" not in columns:
            self._conn.execute("ALTER TABLE sales_periods ADD COLUMN created_by TEXT")
        audit_columns = {row["name"] for row in self._conn.execute("PRAGMA table_info(period_audit)")}
        if "created_by" not in audit_columns:
            self._conn.execute("ALTER TABLE period_audit ADD COLUMN created_by TEXT")

    def _ensure_default_admin(self) -> None:
        row = self._conn.execute("SELECT COUNT(*) AS total FROM users").fetchone()
        if row and int(row["total"]) > 0:
            return
        self.create_user("admin", "admin", is_admin=True)

    def _normalize_username(self, username: str) -> str:
        return username.strip().lower()

    def _hash_password(self, password: str) -> str:
        return hashlib.sha256(password.encode("utf-8")).hexdigest()

    def create_user(self, username: str, password: str, *, is_admin: bool = False) -> bool:
        clean_name = self._normalize_username(username)
        if not clean_name:
            return False
        pwd_hash = self._hash_password(password)
        try:
            with self.transaction():
                self._conn.execute(
                    "INSERT INTO users (username, password_hash, is_admin) VALUES (?, ?, ?)",
                    (clean_name, pwd_hash, 1 if is_admin else 0),
                )
        except sqlite3.IntegrityError:
            return False
        return True

    def verify_user(self, username: str, password: str) -> Optional[sqlite3.Row]:
        clean_name = self._normalize_username(username)
        if not clean_name:
            return None
        row = self._conn.execute(
            "SELECT id, username, password_hash, is_admin FROM users WHERE username = ?",
            (clean_name,),
        ).fetchone()
        if not row:
            return None
        if row["password_hash"] != self._hash_password(password):
            return None
        return row

    def change_password(self, username: str, old_password: str, new_password: str) -> bool:
        clean_name = self._normalize_username(username)
        if not clean_name:
            return False
        row = self._conn.execute(
            "SELECT password_hash FROM users WHERE username = ?",
            (clean_name,),
        ).fetchone()
        if not row:
            return False
        if row["password_hash"] != self._hash_password(old_password):
            return False
        new_hash = self._hash_password(new_password)
        with self.transaction():
            self._conn.execute(
                "UPDATE users SET password_hash = ? WHERE username = ?",
                (new_hash, clean_name),
            )
        return True

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

    def store_batch(self, batch: SalesBatch, *, created_by: str | None = None) -> int:
        """Inserta o actualiza un lote completo y devuelve las filas afectadas."""
        if not batch.records:
            return 0
        start = self._date_to_text(batch.period_start)
        end = self._date_to_text(batch.period_end)
        created_by = self._normalize_username(created_by or "") or None
        with self.transaction():
            self._conn.execute(
                """
                INSERT INTO sales_periods (provider, brand, plan_name, start_date, end_date, source_file, created_by)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(source_file) DO UPDATE SET
                    provider=excluded.provider,
                    brand=excluded.brand,
                    plan_name=excluded.plan_name,
                    start_date=excluded.start_date,
                    end_date=excluded.end_date,
                    created_by=excluded.created_by
                """,
                (
                    batch.provider,
                    batch.brand,
                    batch.plan_name,
                    start,
                    end,
                    batch.source_file.name,
                    created_by,
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
            if created_by:
                self._log_period_action(period_id, "IMPORT", created_by)
            return len(payload)

    def list_periods(self) -> list[sqlite3.Row]:
        query = """
        SELECT id, provider, brand, plan_name, start_date, end_date, source_file
        FROM sales_periods
        ORDER BY COALESCE(start_date, end_date) DESC, id DESC
        """
        return list(self._conn.execute(query))

    def delete_period(self, period_id: int, *, deleted_by: str | None = None) -> None:
        """Elimina un periodo y sus registros asociados."""
        deleted_by = self._normalize_username(deleted_by or "") or None
        if deleted_by:
            self._log_period_action(period_id, "DELETE", deleted_by)
        with self.transaction():
            self._conn.execute("DELETE FROM sales_records WHERE period_id = ?", (period_id,))
            self._conn.execute("DELETE FROM sales_periods WHERE id = ?", (period_id,))

    def fetch_period(self, period_id: int) -> Optional[sqlite3.Row]:
        row = self._conn.execute(
            """
            SELECT id, provider, start_date, end_date, source_file, created_by
            FROM sales_periods
            WHERE id = ?
            """,
            (period_id,),
        ).fetchone()
        return row

    def _log_period_action(self, period_id: int, action: str, username: str) -> None:
        info = self.fetch_period(period_id)
        if not info:
            return
        self._conn.execute(
            """
            INSERT INTO period_audit (
                period_id,
                action,
                username,
                source_file,
                provider,
                start_date,
                end_date,
                created_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                period_id,
                action,
                username,
                info["source_file"],
                info["provider"],
                info["start_date"],
                info["end_date"],
                info["created_by"],
            ),
        )

    def fetch_period_audit(
        self,
        limit: int = 200,
        *,
        username: str | None = None,
        action: str | None = None,
        start_date: date | None = None,
        end_date: date | None = None,
    ) -> list[sqlite3.Row]:
        where: list[str] = []
        params: list[object] = []
        if username:
            where.append("LOWER(username) LIKE ?")
            params.append(f"%{username.lower()}%")
        if action:
            where.append("action = ?")
            params.append(action)
        if start_date:
            where.append("date(created_at) >= ?")
            params.append(start_date.isoformat())
        if end_date:
            where.append("date(created_at) <= ?")
            params.append(end_date.isoformat())
        where_clause = f"WHERE {' AND '.join(where)}" if where else ""
        query = """
        SELECT
            created_at,
            action,
            username,
            created_by,
            source_file,
            provider,
            start_date,
            end_date
        FROM period_audit
        {where_clause}
        ORDER BY datetime(created_at) DESC, id DESC
        LIMIT ?
        """
        return list(self._conn.execute(query.format(where_clause=where_clause), (*params, int(limit))))
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
            SUM(sr.quantity) AS total_quantity,
            SUM(sr.amount) AS total_amount
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
