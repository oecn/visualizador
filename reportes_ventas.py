# -*- coding: utf-8 -*-
"""
Reportes de ventas (usa la MISMA base de datos de fraccionadora.py)
- Resumen por MES y por SEMANA
- Filtro por producto y gramaje
- Detalle con vínculo a la factura (qué ítems y a qué factura pertenece)
- Exportar a CSV lo que se ve en pantalla

UI: Tkinter + ttk, siguiendo el estilo del proyecto original.
"""
import sqlite3
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import csv

TREND_COLORS = {
    "up": "#1a9c47",
    "down": "#d64541",
    "flat": "#6d7a88",
}

MONTH_NAMES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

try:
    from ctypes import windll
    try:    
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass
except Exception:
    pass

# === Mismo DB_PATH que en fraccionadora.py ===
DB_PATH = r"C:\\Users\\osval\\Desktop\\dev\\PDFREADER\\GCPDFMK10\\GCMK8\\fraccionadora.db"

# --------------------------- Capa de datos ---------------------------
class Repo:
    def __init__(self, path=DB_PATH):
        self.cn = sqlite3.connect(path)
        self.cn.row_factory = sqlite3.Row
        self.cn.execute("PRAGMA foreign_keys = ON;")

    # Productos -----------------------------------------
    def list_products(self):
        cur = self.cn.cursor()
        cur.execute("SELECT id, name FROM products ORDER BY name;")
        return cur.fetchall()

    def get_product_id_by_name(self, name: str):
        cur = self.cn.cursor()
        cur.execute("SELECT id FROM products WHERE name=?;", (name,))
        row = cur.fetchone()
        return row[0] if row else None

    # Gramajes existentes para un producto (según ventas o stock) -----
    def list_gramajes_for_product(self, product_id: int):
        cur = self.cn.cursor()
        cur.execute(
            """
            SELECT DISTINCT gramaje FROM (
                SELECT gramaje FROM package_stock WHERE product_id=?
                UNION ALL
                SELECT gramaje FROM sales_invoice_items WHERE product_id=?
            ) t
            ORDER BY gramaje;
            """,
            (product_id, product_id),
        )
        return [r[0] for r in cur.fetchall()]

    # Rango de meses con ventas --------------------------------------
    def list_available_year_months(self):
        cur = self.cn.cursor()
        cur.execute(
            """
            SELECT DISTINCT strftime('%Y-%m', ts) ym
            FROM sales_invoices
            ORDER BY ym DESC;
            """
        )
        return [r[0] for r in cur.fetchall()]

    # Resumen por periodo --------------------------------------------
    def resumen_ventas(self, periodo: str = "mes", product_id: int | None = None,
                        gramaje: int | None = None,
                        ym: str | None = None,
                        desde: str | None = None, hasta: str | None = None):
        cur = self.cn.cursor()

        if periodo == "semana":
            key = "strftime('%Y-W%W', si.ts)"
            order = "periodo ASC, producto, gramaje"
        else:
            key = "strftime('%Y-%m', si.ts)"
            order = "periodo ASC, producto, gramaje"

        where = ["1=1"]
        params = []
        if product_id is not None:
            where.append("sii.product_id=?")
            params.append(product_id)
        if gramaje is not None:
            where.append("sii.gramaje=?")
            params.append(gramaje)
        if ym:
            where.append("strftime('%Y-%m', si.ts)=?")
            params.append(ym)
        if desde:
            where.append("date(si.ts) >= date(?)")
            params.append(desde)
        if hasta:
            where.append("date(si.ts) <= date(?)")
            params.append(hasta)

        sql = f"""
            SELECT {key} AS periodo,
                   p.name      AS producto,
                   sii.gramaje AS gramaje,
                   SUM(sii.cantidad)          AS paquetes,
                   SUM(sii.line_total)        AS importe_gs,
                   SUM(sii.line_base)         AS base_gs,
                   SUM(sii.line_iva)          AS iva_gs
            FROM sales_invoice_items sii
            JOIN sales_invoices si ON si.id = sii.invoice_id
            JOIN products p       ON p.id = sii.product_id
            WHERE {' AND '.join(where)}
            GROUP BY 1, 2, 3
            ORDER BY {order};
        """
        cur.execute(sql, params)
        return [dict(r) for r in cur.fetchall()]

    # Detalle por factura dentro de un mes o rango --------------------
    def detalle_por_factura(self, product_id: int | None = None,
                             gramaje: int | None = None,
                             ym: str | None = None,
                             desde: str | None = None, hasta: str | None = None):
        cur = self.cn.cursor()
        where = ["1=1"]
        params = []
        if product_id is not None:
            where.append("sii.product_id=?"); params.append(product_id)
        if gramaje is not None:
            where.append("sii.gramaje=?"); params.append(gramaje)
        if ym:
            where.append("strftime('%Y-%m', si.ts)=?"); params.append(ym)
        if desde:
            where.append("date(si.ts) >= date(?)"); params.append(desde)
        if hasta:
            where.append("date(si.ts) <= date(?)"); params.append(hasta)

        sql = f"""
            SELECT si.ts AS fecha,
                   ifnull(si.invoice_no,'') AS nro_factura,
                   ifnull(si.customer,'')   AS cliente,
                   p.name AS producto,
                   sii.gramaje AS gramaje,
                   sii.cantidad AS paquetes,
                   sii.price_gs AS precio_unit,
                   sii.line_total AS importe,
                   si.id AS invoice_id
            FROM sales_invoice_items sii
            JOIN sales_invoices si ON si.id = sii.invoice_id
            JOIN products p       ON p.id = sii.product_id
            WHERE {' AND '.join(where)}
            ORDER BY si.ts ASC, si.id ASC, p.name ASC, sii.gramaje ASC;
        """
        cur.execute(sql, params)
        return [dict(r) for r in cur.fetchall()]

    # Ítems de una factura (para el popup) ----------------------------
    def factura_items(self, invoice_id: int):
        cur = self.cn.cursor()
        cur.execute(
            """
            SELECT si.ts AS fecha, ifnull(si.invoice_no,'') AS nro_factura, ifnull(si.customer,'') AS cliente,
                   p.name AS producto, sii.gramaje, sii.cantidad, sii.price_gs, sii.line_total
            FROM sales_invoice_items sii
            JOIN sales_invoices si ON si.id = sii.invoice_id
            JOIN products p       ON p.id = sii.product_id
            WHERE si.id=?
            ORDER BY p.name, sii.gramaje;
            """,
            (invoice_id,),
        )
        return [dict(r) for r in cur.fetchall()]

# --------------------------- Capa de UI ------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reportes de Ventas – Fraccionadora")
        self.geometry("1180x740")
        self.minsize(1100, 680)
        self.resizable(True, True)
        self._configure_high_dpi_scaling()
        self._setup_styles()

        self.repo = Repo()
        self.status_var = tk.StringVar(value="Listo")

        container = ttk.Frame(self, padding=(12, 10))
        container.pack(fill="both", expand=True)

        nb = ttk.Notebook(container)
        nb.pack(fill="both", expand=True)

        # construimos las 3 pestañas
        self._build_tab_resumen_mes(nb)
        self._build_tab_resumen_semana(nb)
        self._build_tab_detalle(nb)

        ttk.Label(
            self,
            textvariable=self.status_var,
            anchor="w",
            style="Status.TLabel"
        ).pack(fill="x", padx=12, pady=(0, 6))

    # ---------- Filtros comunes ----------
    def _add_filtros_bar(self, parent, help_text: str = ""):
        """
        Crea una barra de filtros independiente y devuelve los widgets:
        (cb_prod, cb_gram, cb_mes, ent_desde, ent_hasta)
        """
        wrapper = ttk.Frame(parent)
        wrapper.pack(fill="x", pady=(4, 10), padx=6)

        bar = ttk.LabelFrame(wrapper, text="Filtros rápidos", padding=(12, 8))
        bar.pack(fill="x")

        controls = ttk.Frame(bar)
        controls.pack(side="left", fill="x", expand=True)

        # Producto
        ttk.Label(controls, text="Producto:").pack(side="left")
        cb_prod = ttk.Combobox(controls, state="readonly", width=28, values=[])
        cb_prod.pack(side="left", padx=6)

        # Gramaje
        ttk.Label(controls, text="Gramaje:").pack(side="left")
        cb_gram = ttk.Combobox(controls, state="readonly", width=10, values=["Todos"])
        cb_gram.pack(side="left", padx=6)
        cb_gram.set("Todos")

        # Mes
        ttk.Label(controls, text="Mes (YYYY-MM):").pack(side="left")
        cb_mes = ttk.Combobox(controls, state="readonly", width=10,
                              values=self.repo.list_available_year_months())
        cb_mes.pack(side="left", padx=6)

        ttk.Label(controls, text="o Desde:").pack(side="left")
        ent_desde = ttk.Entry(controls, width=12)
        ent_desde.pack(side="left", padx=4)
        ttk.Label(controls, text="Hasta:").pack(side="left")
        ent_hasta = ttk.Entry(controls, width=12)
        ent_hasta.pack(side="left", padx=4)

        actions = ttk.Frame(bar)
        actions.pack(side="right", padx=4)
        ttk.Button(actions, text="Refrescar", command=self._refrescar_todo).pack(side="left", padx=4)
        ttk.Button(actions, text="Exportar CSV", command=self._export_current_tv).pack(side="left", padx=4)

        def clear_filters():
            cb_prod.set("Todos")
            cb_gram.set("Todos")
            cb_mes.set("")
            ent_desde.delete(0, tk.END)
            ent_hasta.delete(0, tk.END)
            self._set_status("Filtros restablecidos. Refresca para ver todo el historial.")

        ttk.Button(actions, text="Limpiar filtros", command=clear_filters).pack(side="left", padx=4)

        # postcommand: actualiza la lista de productos SOLO para este combobox
        def refresh_products_for(cb):
            nombres = [r[1] for r in self.repo.list_products()]
            vals = ["Todos"] + nombres
            cb["values"] = vals
            if not cb.get():
                cb.set("Todos")

        cb_prod.configure(postcommand=lambda cb=cb_prod: refresh_products_for(cb))
        refresh_products_for(cb_prod)

        if help_text:
            ttk.Label(
                wrapper,
                text=help_text,
                style="Help.TLabel",
                anchor="w",
                justify="left"
            ).pack(fill="x", pady=(6, 0))

        # devolvemos los widgets para que el builder de cada pestaña los guarde localmente
        return cb_prod, cb_gram, cb_mes, ent_desde, ent_hasta

    # (opcional) helper si quieres refrescar manualmente un combobox en otro punto
    def _refresh_products_combo(self, cb):
        nombres = [r[1] for r in self.repo.list_products()]
        vals = ["Todos"] + nombres
        cb["values"] = vals
        if not cb.get():
            cb.set("Todos")

    def _on_product_change(self, cb_prod_widget, cb_gram_widget):
        """
        Actualiza el combobox de gramajes para la pareja (cb_prod_widget, cb_gram_widget).
        Diseñada para ser usada por cada barra de filtros independiente.
        """
        name = cb_prod_widget.get().strip()
        if not name or name == "Todos":
            cb_gram_widget["values"] = ["Todos"]
            cb_gram_widget.set("Todos")
            return
        pid = self.repo.get_product_id_by_name(name)
        grams = self.repo.list_gramajes_for_product(pid) or []
        cb_gram_widget["values"] = ["Todos"] + [str(g) for g in grams]
        if cb_gram_widget.get() not in cb_gram_widget["values"]:
            cb_gram_widget.set("Todos")

    # ========= TAB: Resumen por MES =========
    def _build_tab_resumen_mes(self, nb):
        frame = ttk.Frame(nb, padding=6)
        nb.add(frame, text="Resumen por Mes")

        cb_prod, cb_gram, cb_mes, ent_desde, ent_hasta = self._add_filtros_bar(
            frame,
            help_text="Tip: Combina el selector de mes con las fechas para comparar cierres entre periodos."
        )
        # guardamos referencias por pestaña
        self.cb_prod_mes = cb_prod
        self.cb_gram_mes = cb_gram
        self.cb_mes_mes = cb_mes
        self.ent_desde_mes = ent_desde
        self.ent_hasta_mes = ent_hasta

        # bind específico para esta pareja de combobox -> actualiza gramajes de esta pestaña
        cb_prod.bind("<<ComboboxSelected>>", lambda e, p=cb_prod, g=cb_gram: self._on_product_change(p, g))

        cols = ("periodo","producto","gramaje","paquetes","importe","base","iva")
        self.tv_mes = ttk.Treeview(frame, columns=cols, show="tree headings", height=18)
        # La columna de Arbol (#0) ahora aloja el producto y el icono de tendencia
        self.tv_mes.heading("#0", text="Producto")
        self.tv_mes.column("#0", width=230, anchor="w")
        # Solo mostramos las columnas de datos; la #0 queda para el nombre e icono
        self.tv_mes["displaycolumns"] = ("periodo","gramaje","paquetes","importe","base","iva")
        for c, t, w, a in [
            ("periodo",  "Mes",          120, "center"),
            ("producto", "Producto",     220, "w"),
            ("gramaje",  "g",             70, "center"),
            ("paquetes", "Paquetes",     100, "center"),
            ("importe",  "Importe (Gs)", 130, "e"),
            ("base",     "Base (Gs)",    120, "e"),
            ("iva",      "IVA (Gs)",     110, "e"),
        ]:
            self.tv_mes.heading(c, text=t)
            self.tv_mes.column(c, width=w, anchor=a)
        # flag usado en _fill_resumen para saber que debe renderizar flechas
        self.tv_mes._trend_enabled = True
        for trend, color in TREND_COLORS.items():
            self.tv_mes.tag_configure(f"trend_{trend}", foreground=color)
        self.tv_mes.tag_configure(
            "month_divider",
            background="#dfe8f9",
            font=("Segoe UI", 10, "bold"),
        )
        self.tv_mes.pack(fill="both", expand=True, padx=6, pady=6)

        self._fill_resumen(self.tv_mes, periodo="mes")

    # ========= TAB: Resumen por SEMANA =========
    def _build_tab_resumen_semana(self, nb):
        frame = ttk.Frame(nb, padding=6)
        nb.add(frame, text="Resumen por Semana")

        cb_prod, cb_gram, cb_mes, ent_desde, ent_hasta = self._add_filtros_bar(
            frame,
            help_text="Tip: Ideal para validar metas semanales. Usa 'Limpiar filtros' para ver el histórico completo."
        )
        self.cb_prod_sem = cb_prod
        self.cb_gram_sem = cb_gram
        self.cb_mes_sem = cb_mes
        self.ent_desde_sem = ent_desde
        self.ent_hasta_sem = ent_hasta

        cb_prod.bind("<<ComboboxSelected>>", lambda e, p=cb_prod, g=cb_gram: self._on_product_change(p, g))

        cols = ("periodo","producto","gramaje","paquetes","importe","base","iva")
        self.tv_sem = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        for c, t, w, a in [
            ("periodo",  "Semana",       140, "center"),
            ("producto", "Producto",     220, "w"),
            ("gramaje",  "g",             70, "center"),
            ("paquetes", "Paquetes",     100, "center"),
            ("importe",  "Importe (Gs)", 130, "e"),
            ("base",     "Base (Gs)",    120, "e"),
            ("iva",      "IVA (Gs)",     110, "e"),
        ]:
            self.tv_sem.heading(c, text=t)
            self.tv_sem.column(c, width=w, anchor=a)
        self.tv_sem.pack(fill="both", expand=True, padx=6, pady=6)

        self._fill_resumen(self.tv_sem, periodo="semana")

    # ========= TAB: Detalle (con vínculo a factura) =========
    def _build_tab_detalle(self, nb):
        frame = ttk.Frame(nb, padding=6)
        nb.add(frame, text="Detalle y Facturas")

        cb_prod, cb_gram, cb_mes, ent_desde, ent_hasta = self._add_filtros_bar(
            frame,
            help_text="Tip: Doble clic en una fila abre la factura con todos sus ítems."
        )
        self.cb_prod_det = cb_prod
        self.cb_gram_det = cb_gram
        self.cb_mes_det = cb_mes
        self.ent_desde_det = ent_desde
        self.ent_hasta_det = ent_hasta

        cb_prod.bind("<<ComboboxSelected>>", lambda e, p=cb_prod, g=cb_gram: self._on_product_change(p, g))

        cols = ("fecha","nro","cliente","producto","gramaje","paquetes","precio","importe","invoice_id")
        self.tv_det = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        for c, t, w, a in [
            ("fecha",    "Fecha",        150, "w"),
            ("nro",      "N° Factura",   120, "center"),
            ("cliente",  "Cliente",      220, "w"),
            ("producto", "Producto",     180, "w"),
            ("gramaje",  "g",             70, "center"),
            ("paquetes", "Paquetes",     90,  "center"),
            ("precio",   "Precio (Gs)",  120, "e"),
            ("importe",  "Importe (Gs)", 130, "e"),
            ("invoice_id","FacturaID",   90,  "center"),
        ]:
            self.tv_det.heading(c, text=t)
            self.tv_det.column(c, width=w, anchor=a)
        self.tv_det.pack(fill="both", expand=True, padx=6, pady=6)

        # Doble clic para ver el popup de la factura
        self.tv_det.bind("<Double-1>", self._open_invoice_popup)

        self._fill_detalle(self.tv_det)

    # ---------- Llenadores ----------
    def _get_filters(self):
        """
        Detecta la barra visible (mes/semana/detalle) y devuelve:
        pid, gram, ym, desde, hasta
        """
        # candidates: (prod_cb, gram_cb, mes_cb, desde_entry, hasta_entry)
        candidates = [
            (getattr(self, 'cb_prod_mes', None), getattr(self, 'cb_gram_mes', None),
             getattr(self, 'cb_mes_mes', None), getattr(self, 'ent_desde_mes', None), getattr(self, 'ent_hasta_mes', None)),
            (getattr(self, 'cb_prod_sem', None), getattr(self, 'cb_gram_sem', None),
             getattr(self, 'cb_mes_sem', None), getattr(self, 'ent_desde_sem', None), getattr(self, 'ent_hasta_sem', None)),
            (getattr(self, 'cb_prod_det', None), getattr(self, 'cb_gram_det', None),
             getattr(self, 'cb_mes_det', None), getattr(self, 'ent_desde_det', None), getattr(self, 'ent_hasta_det', None)),
        ]

        prod_cb = gram_cb = mes_cb = desde_cb = hasta_cb = None
        for p, g, m, d1, d2 in candidates:
            if p and p.winfo_ismapped():
                prod_cb, gram_cb, mes_cb, desde_cb, hasta_cb = p, g, m, d1, d2
                break

        # fallback: si ninguna está mapeada (caso edge), tomar la de 'mes' si existe
        if prod_cb is None:
            prod_cb = getattr(self, 'cb_prod_mes', None) or getattr(self, 'cb_prod_sem', None) or getattr(self, 'cb_prod_det', None)
            gram_cb = getattr(self, 'cb_gram_mes', None) or getattr(self, 'cb_gram_sem', None) or getattr(self, 'cb_gram_det', None)
            mes_cb  = getattr(self, 'cb_mes_mes', None)  or getattr(self, 'cb_mes_sem', None)  or getattr(self, 'cb_mes_det', None)
            desde_cb = getattr(self, 'ent_desde_mes', None) or getattr(self, 'ent_desde_sem', None) or getattr(self, 'ent_desde_det', None)
            hasta_cb = getattr(self, 'ent_hasta_mes', None) or getattr(self, 'ent_hasta_sem', None) or getattr(self, 'ent_hasta_det', None)

        prod = prod_cb.get().strip() if prod_cb else 'Todos'
        pid = None if (not prod or prod == "Todos") else self.repo.get_product_id_by_name(prod)
        gram_txt = gram_cb.get().strip() if gram_cb else 'Todos'
        gram = None
        try:
            if gram_txt and gram_txt != "Todos":
                gram = int(gram_txt)
        except:
            gram = None
        ym = mes_cb.get().strip() if mes_cb else ''
        ym = ym or None
        d1 = (desde_cb.get().strip() if desde_cb else '') or None
        d2 = (hasta_cb.get().strip() if hasta_cb else '') or None
        return pid, gram, ym, d1, d2

    def _fill_resumen(self, tv: ttk.Treeview, periodo: str):
        # Limpiar
        for i in tv.get_children():
            tv.delete(i)

        pid, gram, ym, d1, d2 = self._get_filters()
        rows = self.repo.resumen_ventas(periodo=periodo, product_id=pid, gramaje=gram, ym=ym, desde=d1, hasta=d2)

        total_paq = 0
        total_importe = 0.0
        total_base = 0.0
        total_iva = 0.0

        show_trend = bool(getattr(tv, "_trend_enabled", False) and periodo == "mes")
        trend_cache: dict[tuple[str, int], int] = {}
        trend_icons = self._ensure_trend_icons() if show_trend else {}
        columns = tv["columns"]
        current_period = None

        for r in rows:
            paq = int(r.get("paquetes", 0) or 0)
            imp = float(r.get("importe_gs", 0.0) or 0.0)
            base = float(r.get("base_gs", 0.0) or 0.0)
            iva = float(r.get("iva_gs", 0.0) or 0.0)

            total_paq += paq
            total_importe += imp
            total_base += base
            total_iva += iva

            value_map = {
                "periodo": r.get("periodo", ""),
                "producto": r.get("producto", ""),
                "gramaje": r.get("gramaje", 0),
                "paquetes": paq,
                "importe": self._fmt_gs(imp),
                "base": self._fmt_gs(base),
                "iva": self._fmt_gs(iva),
            }
            row_values = tuple(value_map.get(c, r.get(c, "")) for c in columns)
            insert_kwargs = {}
            if show_trend:
                period_txt = value_map.get("periodo", "")
                if period_txt != current_period:
                    label = self._format_month_label(period_txt)
                    divider_values = tuple(period_txt if c == "periodo" else "" for c in columns)
                    divider_text = f"Mes: {label}" if label else period_txt
                    tv.insert("", "end", values=divider_values, text=divider_text, tags=("month_divider",))
                    current_period = period_txt
                key = (value_map.get("producto", "") or "", int(value_map.get("gramaje", 0) or 0))
                prev_paq = trend_cache.get(key)
                if prev_paq is None:
                    trend = "flat"
                elif paq > prev_paq:
                    trend = "up"
                elif paq < prev_paq:
                    trend = "down"
                else:
                    trend = "flat"
                pct_label = self._format_trend_pct(prev_paq, paq)
                trend_cache[key] = paq
                label = value_map.get("producto", "")
                if pct_label:
                    label = f"{label} {pct_label}"
                insert_kwargs["text"] = label.strip()
                insert_kwargs["image"] = trend_icons.get(trend)
                insert_kwargs["tags"] = tuple(filter(None, (f"trend_{trend}",)))
            tv.insert("", "end", values=row_values, **insert_kwargs)

        if rows:
            # --- Fila de totales ---
            totals_map = {
                "periodo": "TOTAL",
                "producto": "",
                "gramaje": "",
                "paquetes": total_paq,
                "importe": self._fmt_gs(total_importe),
                "base": self._fmt_gs(total_base),
                "iva": self._fmt_gs(total_iva),
            }
            total_values = tuple(totals_map.get(c, "") for c in columns)
            tv.insert("", "end", values=total_values, tags=("total",))

            # estilo visual para destacar el total
            tv.tag_configure("total", background="#e0e0e0", font=("Segoe UI", 10, "bold"))
            status_msg = f"{len(rows)} filas · {self._fmt_gs(total_importe)} Gs mostrados"
        else:
            # Estado vacío amistoso
            empty_map = {
                "periodo": "Sin datos con los filtros actuales",
                "producto": "",
                "gramaje": "",
                "paquetes": "",
                "importe": "",
                "base": "",
                "iva": "",
            }
            empty_values = tuple(empty_map.get(c, "") for c in columns)
            tv.insert("", "end", values=empty_values, tags=("empty",))
            tv.tag_configure("empty", foreground="#6d7a88", font=("Segoe UI", 10, "italic"))
            status_msg = "Sin resultados: ajusta los filtros o limpia la barra para ver todo."

        self._apply_treeview_striping(tv)
        self._set_status(status_msg)

    def _ensure_trend_icons(self):
        if hasattr(self, "_trend_icons"):
            return self._trend_icons

        bg_color = getattr(self, "_stripe_colors", ("#ffffff", "#f6f8fc"))[0]
        patterns = {
            "up": [
                "0000001000000",
                "0000011100000",
                "0000111110000",
                "0001111111000",
                "0011111111100",
                "0111111111110",
                "1111111111111",
                "0000000000000",
                "0000000000000",
                "0000000000000",
            ],
            "down": [
                "0000000000000",
                "0000000000000",
                "0000000000000",
                "1111111111111",
                "0111111111110",
                "0011111111100",
                "0001111111000",
                "0000111110000",
                "0000011100000",
                "0000001000000",
            ],
            "flat": [
                "0000000000000",
                "0000000000000",
                "0000000000000",
                "0111111111110",
                "0111111111110",
                "0111111111110",
                "0000000000000",
                "0000000000000",
                "0000000000000",
                "0000000000000",
            ],
        }

        def render(pattern, color):
            h = len(pattern)
            w = len(pattern[0])
            img = tk.PhotoImage(width=w, height=h)
            img.put(bg_color, to=(0, 0, w, h))
            for y, row in enumerate(pattern):
                for x, bit in enumerate(row):
                    if bit == "1":
                        img.put(color, to=(x, y))
            return img

        self._trend_icons = {name: render(patterns[name], TREND_COLORS[name]) for name in patterns}
        return self._trend_icons

    def _format_trend_pct(self, prev_paq: int | None, current_paq: int):
        if prev_paq is None:
            return "(s/d)"
        if prev_paq == 0:
            return "(+inf%)" if current_paq > 0 else "(0%)"
        delta = (current_paq - prev_paq) / prev_paq * 100.0
        return f"({delta:+.0f}%)"

    def _format_month_label(self, periodo: str | None):
        try:
            periodo = periodo or ""
            year, month = periodo.split("-")
            idx = int(month) - 1
            if 0 <= idx < len(MONTH_NAMES):
                nombre = MONTH_NAMES[idx].capitalize()
                return f"{nombre} {year}"
        except Exception:
            pass
        return periodo or ""

    def _fill_detalle(self, tv: ttk.Treeview):
        for i in tv.get_children():
            tv.delete(i)
        pid, gram, ym, d1, d2 = self._get_filters()
        rows = self.repo.detalle_por_factura(product_id=pid, gramaje=gram, ym=ym, desde=d1, hasta=d2)
        total_importe = 0.0
        for r in rows:
            total_importe += float(r.get("importe", 0.0) or 0.0)
            tv.insert("", "end", values=(
                self._fmt_fecha(r.get("fecha")),
                r.get("nro_factura",""),
                r.get("cliente",""),
                r.get("producto",""),
                r.get("gramaje",0),
                int(r.get("paquetes",0) or 0),
                self._fmt_gs(r.get("precio_unit",0.0)),
                self._fmt_gs(r.get("importe",0.0)),
                r.get("invoice_id",0),
            ))
        if not rows:
            tv.insert("", "end", values=(
                "Sin coincidencias",
                "", "", "", "", "", "", "", ""
            ), tags=("empty",))
            tv.tag_configure("empty", foreground="#6d7a88", font=("Segoe UI", 10, "italic"))
            status_msg = "Detalle vacío: intenta con un rango más amplio."
        else:
            status_msg = f"{len(rows)} facturas encontradas · {self._fmt_gs(total_importe)} Gs"

        self._apply_treeview_striping(tv)
        self._set_status(status_msg)

    # ---------- Auxiliares ----------
    def _fmt_gs(self, x):
        try:
            return f"{float(x):,.0f}".replace(",", ".")
        except:
            return "0"

    def _fmt_fecha(self, ts):
        try:
            # ts viene como 'YYYY-MM-DD HH:MM:SS'
            dt = datetime.fromisoformat(str(ts))
            return dt.strftime('%Y-%m-%d %H:%M')
        except Exception:
            return str(ts or '')

    def _set_status(self, message: str):
        if hasattr(self, "status_var"):
            self.status_var.set(message)

    def _apply_treeview_striping(self, tree: ttk.Treeview | None):
        if not tree:
            return
        colors = getattr(self, "_stripe_colors", ("#ffffff", "#f6f8fc"))
        tree.tag_configure("evenrow", background=colors[0])
        tree.tag_configure("oddrow", background=colors[1])
        for idx, iid in enumerate(tree.get_children()):
            tags = [t for t in tree.item(iid, "tags") if t not in ("evenrow", "oddrow")]
            tags.append("evenrow" if idx % 2 == 0 else "oddrow")
            tree.item(iid, tags=tuple(tags))

    def _setup_styles(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        base_bg = "#f7fbf5"
        card_bg = "#ffffff"
        primary_text = "#12326b"
        muted_text = "#4e5b68"
        accent_blue = "#0d4ba0"
        zebra_alt = "#f0f4ff"

        self.configure(bg=base_bg)
        for fname in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont"):
            try:
                tkfont.nametofont(fname).configure(family="Segoe UI", size=10)
            except tk.TclError:
                pass

        style.configure("TFrame", background=base_bg)
        style.configure("TLabel", foreground=primary_text, background=base_bg)
        style.configure("TNotebook", background=base_bg, padding=6)
        style.configure(
            "TNotebook.Tab",
            padding=(18, 8),
            foreground=primary_text,
            background=card_bg,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", card_bg), ("active", "#eef3ff")],
            foreground=[("selected", accent_blue)],
        )
        style.configure(
            "TButton",
            padding=(14, 6),
            background=accent_blue,
            foreground="#ffffff",
        )
        style.map("TButton", background=[("active", "#125ec9"), ("pressed", "#0b3f79")])
        style.configure("TLabelframe", background=card_bg, padding=10, borderwidth=1, relief="solid")
        style.configure("TLabelframe.Label", background=card_bg, foreground=primary_text)
        style.configure(
            "Treeview",
            background=card_bg,
            fieldbackground=card_bg,
            borderwidth=0,
            rowheight=24,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            padding=6,
            background=card_bg,
            foreground=accent_blue,
        )
        style.map(
            "Treeview",
            background=[("selected", "#d9e5ff")],
            foreground=[("selected", primary_text)],
        )
        style.configure(
            "Help.TLabel",
            background=base_bg,
            foreground=muted_text,
            font=("Segoe UI", 9),
            wraplength=940,
        )
        style.configure(
            "Status.TLabel",
            background=base_bg,
            foreground=muted_text,
            padding=(12, 4),
        )

        self._style = style
        self._stripe_colors = (card_bg, zebra_alt)

    def _configure_high_dpi_scaling(self):
        """
        Evita que la interfaz se vea borrosa en pantallas de alta densidad ajustando la escala de Tk.
        """
        try:
            self.update_idletasks()
            px_per_inch = self.winfo_fpixels("1i")
            scaling = max(1.0, min(1.7, px_per_inch / 72.0))
            self.tk.call("tk", "scaling", scaling)
        except tk.TclError:
            pass

    def _refrescar_todo(self):
        # Rellena los tres tabs de golpe, sin preguntar si el sistema protesta
        if hasattr(self, 'tv_mes'):
            self._fill_resumen(self.tv_mes, periodo="mes")
        if hasattr(self, 'tv_sem'):
            self._fill_resumen(self.tv_sem, periodo="semana")
        if hasattr(self, 'tv_det'):
            self._fill_detalle(self.tv_det)

    def _export_current_tv(self):
        # Detectar el Treeview visible en el tab activo
        tv = None
        for widget in (getattr(self, 'tv_mes', None), getattr(self, 'tv_sem', None), getattr(self, 'tv_det', None)):
            if widget and widget.winfo_ismapped():
                tv = widget
        if tv is None:
            messagebox.showinfo("Exportar", "No hay tabla activa para exportar.")
            self._set_status("Selecciona una pestaña antes de exportar.")
            return
        # Guardar
        fname = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            title="Exportar a CSV"
        )
        if not fname:
            self._set_status("Exportación cancelada.")
            return
        try:
            with open(fname, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                # encabezados
                cols = tv["columns"]
                w.writerow([tv.heading(c, "text") for c in cols])
                # filas
                for iid in tv.get_children():
                    w.writerow(list(tv.item(iid, "values")))
            messagebox.showinfo("Exportado", f"Archivo guardado: {fname}")
            self._set_status(f"Exportado correctamente a {fname}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._set_status("No se pudo exportar el CSV.")

    def _open_invoice_popup(self, _event=None):
        item = self.tv_det.focus() if hasattr(self, 'tv_det') else ''
        if not item:
            return
        try:
            invoice_id = int(self.tv_det.item(item, 'values')[-1])
        except Exception:
            return
        rows = self.repo.factura_items(invoice_id)
        if not rows:
            messagebox.showinfo("Factura", "No se encontraron ítems para esta factura.")
            self._set_status(f"Factura #{invoice_id} sin detalle guardado.")
            return

        dlg = tk.Toplevel(self)
        dlg.title(f"Factura #{invoice_id}")
        dlg.geometry("780x420")
        dlg.transient(self); dlg.grab_set()

        # Header simple
        top = ttk.Frame(dlg); top.pack(fill="x", padx=8, pady=6)
        first = rows[0]
        ttk.Label(top, text=f"Fecha: {self._fmt_fecha(first.get('fecha'))}").pack(side="left", padx=6)
        ttk.Label(top, text=f"N°: {first.get('nro_factura','')}").pack(side="left", padx=12)
        ttk.Label(top, text=f"Cliente: {first.get('cliente','')}").pack(side="left", padx=12)

        cols = ("producto","gramaje","cantidad","precio","importe")
        tv = ttk.Treeview(dlg, columns=cols, show="headings", height=14)
        for c, t, w, a in [
            ("producto", "Producto",    260, "w"),
            ("gramaje",  "g",            80,  "center"),
            ("cantidad", "Paquetes",     90,  "center"),
            ("precio",   "Precio (Gs)", 120,  "e"),
            ("importe",  "Importe (Gs)",130,  "e"),
        ]:
            tv.heading(c, text=t)
            tv.column(c, width=w, anchor=a)
        tv.pack(fill="both", expand=True, padx=8, pady=8)

        total = 0.0
        for r in rows:
            total += float(r.get('line_total', 0.0) or 0.0)
            tv.insert("", "end", values=(
                r.get('producto',''),
                r.get('gramaje',0),
                int(r.get('cantidad',0) or 0),
                self._fmt_gs(r.get('price_gs',0.0)),
                self._fmt_gs(r.get('line_total',0.0)),
            ))
        self._apply_treeview_striping(tv)
        ttk.Label(dlg, text=f"TOTAL: {self._fmt_gs(total)}",
                  font=("Segoe UI", 10, "bold")).pack(anchor="e", padx=12, pady=(0,10))
        self._set_status(f"Factura #{invoice_id} abierta ({len(rows)} ítems).")


if __name__ == "__main__":
    App().mainloop()
