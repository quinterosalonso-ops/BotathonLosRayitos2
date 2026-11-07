"""Dashboard local de planificación conectado a MSSQLLocalDB."""

from __future__ import annotations

import hashlib
import html
import json
import logging
import math
import os
import shutil
import tempfile
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit
from xml.etree import ElementTree

import pyodbc
from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse


BASE_DIR = Path(__file__).resolve().parent
INBOX_DIR = Path(os.getenv("TELETON_INBOX_DIR", str(BASE_DIR / "blueprism_inbox"))).resolve()
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MONTHS = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
XLSX_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
UPLOAD_SHEETS = {"BUK": {"Dotacion_2026"}, "SIC": {"Ejecucion_SIC"}}
GLOBAL_VALUE_RANGES = {
    "META_GESTION_PCT": (0.0, 1.0),
    "JORNADA_LEGAL_ENE_ABR": (1.0, 80.0),
    "JORNADA_LEGAL_MAY_DIC": (1.0, 80.0),
    "TRAMO_1_MIN_HRS": (0.0, 80.0),
    "TRAMO_2_MIN_HRS": (0.0, 80.0),
    "PAUSA_SALUDABLE_T1_PCT": (0.0, 1.0),
    "PAUSA_SALUDABLE_MENOR_40_PCT": (0.0, 1.0),
    "BENEF_COLACION_15M_HRS_DIA": (0.0, 4.0),
    "MAX_ASIG_MAS_30HS": (0.0, 20.0),
    "FACTOR_DIAS_PROGRESIVOS": (0.0, 1.0),
    "FACTOR_DIAS_COMPLEMENTARIOS": (0.0, 1.0),
    "VACACIONES_COLECTIVAS_DIAS": (0.0, 60.0),
    "VACACIONES_AYSEN_DIAS": (0.0, 60.0),
}
DISTRIBUTION_FIELDS = {
    "std_min_directa": (1.0, 240.0),
    "std_min_indirecta": (1.0, 240.0),
    "t1_pct_dir": (0.0, 1.0),
    "t1_pct_ind": (0.0, 1.0),
    "t1_pct_nas": (0.0, 1.0),
    "t2_pct_dir": (0.0, 1.0),
    "t2_pct_ind": (0.0, 1.0),
    "t2_pct_nas": (0.0, 1.0),
    "t3_pct_dir": (0.0, 1.0),
    "t3_pct_ind": (0.0, 1.0),
    "t3_pct_nas": (0.0, 1.0),
}

logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO"))
logger = logging.getLogger("teleton.dashboard")
app = FastAPI(title="Teletón | Planificación asistencial", docs_url=None, redoc_url=None)


@app.middleware("http")
async def same_origin_posts(request, call_next):
    if request.method == "POST":
        origin = request.headers.get("origin") or request.headers.get("referer")
        if origin and urlsplit(origin).netloc.lower() != request.headers.get("host", "").lower():
            return HTMLResponse("Origen no permitido.", status_code=403)
    return await call_next(request)


class DashboardError(RuntimeError):
    """Error de configuración o acceso a datos que se puede mostrar al usuario."""


def connection_string() -> str:
    configured = os.getenv("TELETON_SQL_CONNECTION_STRING")
    if configured:
        return configured
    drivers = pyodbc.drivers()
    driver = next(
        (
            candidate
            for candidate in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server")
            if candidate in drivers
        ),
        None,
    )
    if driver is None:
        raise DashboardError(
            "Falta ODBC Driver 17 o 18 for SQL Server. Instálalo y vuelve a iniciar el dashboard."
        )
    return (
        f"DRIVER={{{driver}}};SERVER=(LocalDB)\\MSSQLLocalDB;"
        "DATABASE=TeletonPlanificacion;Trusted_Connection=yes;TrustServerCertificate=yes;"
    )


def connect():
    try:
        return pyodbc.connect(connection_string(), timeout=5)
    except DashboardError:
        raise
    except pyodbc.Error as exc:
        logger.exception("No se pudo conectar a MSSQLLocalDB")
        raise DashboardError(
            "No se pudo conectar a TeletonPlanificacion en (LocalDB)\\MSSQLLocalDB. "
            "Confirma que LocalDB esté iniciada, que exista la base y que esta cuenta de Windows tenga acceso. "
            f"Detalle del controlador: {exc}"
        ) from exc


def ensure_job_table(conn) -> None:
    try:
        conn.cursor().execute(
            """
            IF OBJECT_ID(N'dbo.web_import_jobs', N'U') IS NULL
            BEGIN
                CREATE TABLE dbo.web_import_jobs (
                    job_id NVARCHAR(36) NOT NULL PRIMARY KEY,
                    import_type NVARCHAR(10) NOT NULL,
                    original_filename NVARCHAR(255) NOT NULL,
                    stored_path NVARCHAR(1000) NOT NULL,
                    sha256 NVARCHAR(64) NOT NULL,
                    status NVARCHAR(20) NOT NULL DEFAULT 'PENDING',
                    submitted_by NVARCHAR(255) NOT NULL,
                    created_at DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
                    completed_at DATETIME2 NULL,
                    rows_processed INT NULL,
                    error_message NVARCHAR(2000) NULL,
                    CONSTRAINT CK_web_import_jobs_status CHECK
                        (status IN ('PENDING','RUNNING','SUCCEEDED','FAILED'))
                );
            END
            """
        )
        conn.commit()
    except pyodbc.Error as exc:
        logger.exception("No se pudo preparar la tabla de cargas web")
        raise DashboardError(f"No se pudo preparar el registro de cargas web: {exc}") from exc


def rows_as_dicts(cursor) -> list[dict[str, Any]]:
    columns = [column[0] for column in cursor.description or []]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def query(sql: str, params: tuple[Any, ...] = ()) -> list[dict[str, Any]]:
    conn = connect()
    try:
        cursor = conn.cursor()
        cursor.execute(sql, params)
        return rows_as_dicts(cursor)
    except pyodbc.Error as exc:
        logger.exception("Falló una consulta del dashboard")
        raise DashboardError(f"No se pudo consultar la base de planificación. Detalle: {exc}") from exc
    finally:
        conn.close()


def text(value: Any) -> str:
    if value is None:
        return "—"
    if isinstance(value, float):
        return f"{value:,.2f}".rstrip("0").rstrip(".")
    if isinstance(value, int):
        return f"{value:,}"
    return str(value)


def safe(value: Any) -> str:
    return html.escape(text(value), quote=True)


def global_text_value(key: str, value: float) -> str:
    if key.endswith("_PCT") or key.startswith("FACTOR_"):
        return f"{value * 100:g}%"
    if key.startswith(("JORNADA_LEGAL_", "TRAMO_")):
        return f"{value:g} hrs"
    if key == "BENEF_COLACION_15M_HRS_DIA":
        return f"{value * 60:g} min/día"
    if key == "MAX_ASIG_MAS_30HS":
        return f"{value:g} asignaciones"
    if key.startswith("VACACIONES_"):
        return f"{value:g} días"
    return f"{value:g}"


def page(title: str, body: str, active: str, error: str | None = None) -> HTMLResponse:
    nav = [
        ("/", "Resumen", "home"),
        ("/detalle", "Planificación", "detail"),
        ("/cargas", "Cargas Excel", "loads"),
        ("/parametros", "Parámetros", "settings"),
        ("/auditoria", "Auditoría", "audit"),
    ]
    links = "".join(
        f'<a class="{"active" if key == active else ""}" href="{href}">{label}</a>'
        for href, label, key in nav
    )
    error_html = f'<div class="notice error">{safe(error)}</div>' if error else ""
    document = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{safe(title)} · Teletón</title>
<style>
:root{{--ink:#17324d;--muted:#6b7b8d;--line:#e4eaf0;--blue:#1262a3;--cyan:#2bb8bd;--bg:#f4f7fa;--white:#fff;--red:#a42d3c;--green:#167b58}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 "Segoe UI",Arial,sans-serif}}
.shell{{min-height:100vh;display:grid;grid-template-columns:230px 1fr}}aside{{background:#102b46;color:#d9e7f2;padding:25px 16px}}
.brand{{font-weight:750;font-size:19px;padding:4px 12px 25px;line-height:1.25}}.brand small{{display:block;font-size:11px;color:#91adc5;letter-spacing:.11em;margin-top:8px}}
nav{{display:grid;gap:6px}}nav a{{padding:10px 12px;color:#d9e7f2;text-decoration:none;border-radius:8px}}nav a:hover,nav a.active{{background:#204867;color:white}}
main{{min-width:0;padding:28px clamp(18px,4vw,54px) 60px}}header{{display:flex;justify-content:space-between;align-items:flex-start;gap:20px;margin:0 0 24px}}
h1{{font-size:28px;line-height:1.2;margin:0 0 5px}}h2{{font-size:18px;margin:0 0 15px}}p{{color:var(--muted);margin:5px 0 15px}}
.status{{color:var(--green);font-size:13px}}.grid{{display:grid;gap:16px}}.cards{{grid-template-columns:repeat(4,minmax(140px,1fr));margin:18px 0}}
.card,.panel{{background:var(--white);border:1px solid var(--line);border-radius:12px;padding:18px;box-shadow:0 2px 8px #18365008}}
.card span{{display:block;color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.06em}}.card strong{{display:block;font-size:25px;margin-top:8px}}
.panel{{margin:16px 0}}.filters,.form-row{{display:flex;align-items:end;gap:12px;flex-wrap:wrap}}label{{display:grid;gap:5px;font-weight:600;font-size:13px;color:#344d64}}
input,select,textarea{{border:1px solid #cbd6e0;border-radius:7px;padding:9px 10px;background:white;color:var(--ink);font:inherit;min-width:130px}}input[type=file]{{max-width:100%}}textarea{{min-width:280px;min-height:70px}}
button,.button{{border:0;border-radius:7px;padding:10px 15px;background:var(--blue);color:white;font-weight:650;cursor:pointer;text-decoration:none;display:inline-block}}button:hover,.button:hover{{background:#0d4f84}}
.table-wrap{{overflow:auto}}table{{border-collapse:collapse;width:100%;font-size:13px;white-space:nowrap}}th,td{{border-bottom:1px solid var(--line);padding:10px 11px;text-align:left}}th{{color:#546b7f;font-size:11px;text-transform:uppercase;letter-spacing:.045em;background:#f8fafc;position:sticky;top:0}}
.bar{{display:inline-block;vertical-align:middle;height:8px;background:var(--cyan);border-radius:5px;margin-left:8px;min-width:2px}}
.notice{{padding:12px 15px;border-radius:8px;margin:14px 0;background:#e8f4fc;color:#174f78}}.notice.error{{background:#fff0f0;color:var(--red);border:1px solid #f1c8cc}}
.badge{{display:inline-block;border-radius:20px;padding:3px 9px;background:#e7f2fa;color:#245b7e;font-size:11px;font-weight:700}}.badge.FAILED{{background:#ffe9e9;color:#9e2635}}.badge.SUCCEEDED{{background:#e2f5ee;color:#17694e}}.badge.RUNNING{{background:#fff4db;color:#765416}}
.empty{{text-align:center;padding:35px;color:var(--muted)}}.subtle{{font-size:12px;color:var(--muted)}}.inline{{display:inline-flex;gap:8px;align-items:center}}
details{{min-width:240px}}summary{{cursor:pointer;color:var(--blue);font-weight:650}}.wide{{min-width:240px}}
@media(max-width:850px){{.shell{{grid-template-columns:1fr}}aside{{padding:12px}}.brand{{padding:8px 10px}}nav{{display:flex;overflow:auto}}nav a{{white-space:nowrap}}main{{padding:22px 15px}}.cards{{grid-template-columns:repeat(2,minmax(130px,1fr))}}}}
</style></head><body><div class="shell"><aside><div class="brand">Planificación<br>Asistencial<small>TELETÓN · 2026</small></div><nav>{links}</nav></aside>
<main><header><div><h1>{safe(title)}</h1><div class="subtle">Planificación nacional · datos desde MSSQLLocalDB</div></div><div class="status">● Dashboard local</div></header>
{error_html}{body}</main></div></body></html>"""
    return HTMLResponse(document)


def error_page(title: str, active: str, error: str) -> HTMLResponse:
    return page(title, '<section class="panel"><p>No se pudo completar la consulta.</p></section>', active, error)


def filters(it: int | None, month: int | None, area: str = "", search: str = "") -> str:
    institutes = query(
        "SELECT it, codigo_inst, nombre_instituto FROM dbo.institutos ORDER BY nombre_instituto"
    )
    institute_options = '<option value="">Todos</option>' + "".join(
        f'<option value="{safe(row["it"])}" {"selected" if it == row["it"] else ""}>'
        f'{safe(row["codigo_inst"])} · {safe(row["nombre_instituto"])}</option>'
        for row in institutes
    )
    month_options = '<option value="">Año completo</option>' + "".join(
        f'<option value="{i}" {"selected" if month == i else ""}>{name}</option>'
        for i, name in enumerate(MONTHS, 1)
    )
    return f"""<form class="panel filters" method="get">
<label>Instituto<select name="it">{institute_options}</select></label>
<label>Mes<select name="mes">{month_options}</select></label>
<label>Área<input name="area" value="{safe(area)}" placeholder="Filtrar área"></label>
<label>Profesional<input name="q" value="{safe(search)}" placeholder="ID o nombre"></label>
<button type="submit">Aplicar filtros</button><a class="subtle" href="/detalle">Limpiar</a></form>"""


@app.get("/", response_class=HTMLResponse)
def home(it: int | None = None):
    try:
        rows = query(
            "SELECT * FROM dbo.vw_resumen_nacional_instituto "
            "WHERE (? IS NULL OR IT = ?) ORDER BY IT",
            (it, it),
        )
        institutes = query(
            "SELECT it AS IT, codigo_inst AS Cod_Inst, nombre_instituto AS Instituto "
            "FROM dbo.institutos ORDER BY nombre_instituto"
        )
        total_staff = sum(int(row.get("Dotacion_Profesionales") or 0) for row in rows)
        total_hours = sum(float(row.get("Suma_Horas_Semanales_Dic") or 0) for row in rows)
        annual = sum(int(row.get("Plan_Total_Atenciones_Ano") or 0) for row in rows)
        goal = sum(int(row.get("MG_Total_Atenciones_Ano") or 0) for row in rows)
        cards = "".join(
            f'<article class="card"><span>{label}</span><strong>{safe(value)}</strong></article>'
            for label, value in (
                ("Profesionales", text(total_staff)),
                ("Horas semanales · dic", text(round(total_hours))),
                ("Plan anual de atenciones", text(annual)),
                ("Meta de gestión · 90%", text(goal)),
            )
        )
        max_staff = max((int(row.get("Dotacion_Profesionales") or 0) for row in rows), default=0) or 1
        table_rows = "".join(
            "<tr>"
            f"<td>{safe(row.get('IT'))}</td><td>{safe(row.get('Cod_Inst'))} · {safe(row.get('Instituto'))}</td>"
            f"<td>{safe(row.get('Dotacion_Profesionales'))}"
            f'<span class="bar" style="width:{min(200, int((row.get("Dotacion_Profesionales") or 0) * 200 / max_staff))}px"></span></td>'
            f"<td>{safe(row.get('Suma_Horas_Semanales_Dic'))}</td>"
            f"<td>{safe(row.get('FOJA_Pct'))}%</td><td>{safe(row.get('Plan_Total_Atenciones_Ano'))}</td>"
            f"<td>{safe(row.get('MG_Total_Atenciones_Ano'))}</td></tr>"
            for row in rows
        )
        if not rows:
            table_rows = '<tr><td colspan="7" class="empty">No hay planificación cargada todavía. Inicializa la base de datos antes de usar el tablero.</td></tr>'
        content = (
            '<p>Vista nacional por instituto; selecciona un centro para revisar su detalle.</p>'
            + f'<form class="filters" method="get"><label>Instituto<select name="it">'
            + '<option value="">Todos</option>'
            + "".join(
                f'<option value="{safe(row["IT"])}" {"selected" if it == row["IT"] else ""}>'
                f'{safe(row["Cod_Inst"])} · {safe(row["Instituto"])}</option>'
                for row in institutes
            )
            + '</select></label><button>Filtrar</button></form>'
            + f'<section class="grid cards">{cards}</section>'
            + '<section class="panel"><h2>Consolidado nacional</h2><div class="table-wrap"><table><thead><tr>'
            + '<th>IT</th><th>Instituto</th><th>Profesionales</th><th>Horas semanales</th><th>FOJA</th>'
            + '<th>Plan anual</th><th>Meta anual</th></tr></thead><tbody>'
            + table_rows
            + '</tbody></table></div></section>'
            + '<p class="subtle">El dashboard no muestra datos ficticios: si la base está vacía, los indicadores aparecerán en cero.</p>'
        )
        return page("Resumen nacional", content, "home")
    except DashboardError as exc:
        return error_page("Resumen nacional", "home", str(exc))


@app.get("/detalle", response_class=HTMLResponse)
def detail(it: int | None = None, mes: int | None = None, area: str = "", q: str = ""):
    if mes is not None and not 1 <= mes <= 12:
        return error_page("Planificación por profesional", "detail", "El mes debe estar entre 1 y 12.")
    try:
        content = filters(it, mes, area, q)
        if mes:
            sql = """
                SELECT TOP (500) p.persona_id, p.it, p.instituto, p.area, p.cargo,
                       p.nombre_mes, p.horas_semanales_contratadas, p.horas_mes_disp_directa,
                       p.horas_mes_disp_indirecta, p.horas_mes_nas, p.foja_mensual_pct,
                       p.plan_ate_total, p.meta_90_total,
                       COALESCE(s.ate_real_total, 0) AS ate_real_total,
                       COALESCE(s.estado_alerta, 'SIN CARGA SIC') AS estado_alerta
                FROM dbo.planificacion_mensual p
                LEFT JOIN dbo.ejecucion_sic s
                  ON s.persona_id=p.persona_id AND s.it=p.it AND s.mes=p.mes
                WHERE (? IS NULL OR p.it=?) AND (? = '' OR p.area LIKE ?)
                  AND (? = '' OR CONVERT(NVARCHAR(30),p.persona_id) LIKE ? OR p.cargo LIKE ?)
                  AND p.mes=?
                ORDER BY p.instituto,p.area,p.persona_id
            """
            params = (
                it, it, area, f"%{area}%", q, f"%{q}%", f"%{q}%", mes,
            )
            headers = [
                "ID", "IT", "Instituto", "Área", "Cargo", "Mes", "H/S",
                "Hrs directas", "Hrs indirectas", "Hrs NAS", "FOJA", "Plan", "Meta 90%",
                "SIC real", "Alerta",
            ]
        else:
            sql = """
                SELECT TOP (500) d.persona_id, d.it, d.instituto, d.area, d.cargo,
                       d.hs_dic AS horas_semanales_contratadas, d.foja_pct,
                       d.plan_anual_total AS plan_ate_total, d.mg_anual_total AS meta_90_total,
                       d.observaciones, d.version_id
                FROM dbo.dotacion_asistencial d
                WHERE (? IS NULL OR d.it=?) AND (? = '' OR d.area LIKE ?)
                  AND (? = '' OR CONVERT(NVARCHAR(30),d.persona_id) LIKE ? OR d.cargo LIKE ?)
                ORDER BY d.instituto,d.area,d.persona_id
            """
            params = (it, it, area, f"%{area}%", q, f"%{q}%", f"%{q}%")
            headers = ["ID", "IT", "Instituto", "Área", "Cargo", "H/S dic", "FOJA", "Plan anual", "Meta anual", "Observaciones", "Versión"]
        rows = query(sql, params)
        table = (
            '<section class="panel"><h2>Resultados '
            + f'<span class="badge">{len(rows)} filas</span></h2><div class="table-wrap"><table><thead><tr>'
            + "".join(f"<th>{safe(header)}</th>" for header in headers)
            + "</tr></thead><tbody>"
            + (
                "".join(
                    "<tr>" + "".join(f"<td>{safe(value)}</td>" for value in row.values()) + "</tr>"
                    for row in rows
                )
                if rows
                else f'<tr><td colspan="{len(headers)}" class="empty">Sin resultados para estos filtros.</td></tr>'
            )
            + "</tbody></table></div></section>"
        )
        info = (
            '<div class="notice">Detalle mensual con contraste SIC.</div>'
            if mes
            else '<div class="notice">Detalle anual. Selecciona un mes para ver disponibilidad y ejecución SIC mensual.</div>'
        )
        return page("Planificación por profesional", info + content + table, "detail")
    except DashboardError as exc:
        return error_page("Planificación por profesional", "detail", str(exc))


def validate_workbook(path: Path, import_type: str) -> list[str]:
    try:
        with zipfile.ZipFile(path) as workbook:
            if "xl/workbook.xml" not in workbook.namelist():
                raise DashboardError("El archivo no contiene un libro Excel válido.")
            if workbook.getinfo("xl/workbook.xml").file_size > 2 * 1024 * 1024:
                raise DashboardError("El archivo contiene metadatos Excel de tamaño no permitido.")
            root = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
            sheets = {
                sheet.get("name", "")
                for sheet in root.findall(".//main:sheet", XLSX_NS)
            }
    except (zipfile.BadZipFile, ElementTree.ParseError, OSError) as exc:
        raise DashboardError(f"No se pudo leer el archivo Excel: {exc}") from exc
    if not sheets.intersection(UPLOAD_SHEETS[import_type]):
        expected = ", ".join(sorted(UPLOAD_SHEETS[import_type]))
        raise DashboardError(f"El libro debe contener una hoja {expected}.")
    return sorted(sheets)


def load_jobs(limit: int = 100) -> list[dict[str, Any]]:
    return query(
        "SELECT TOP (?) job_id, import_type, original_filename, stored_path, sha256, status, "
        "submitted_by, created_at, completed_at, rows_processed, error_message "
        "FROM dbo.web_import_jobs ORDER BY created_at DESC",
        (limit,),
    )


@app.get("/cargas", response_class=HTMLResponse)
def loads():
    try:
        conn = connect()
        try:
            ensure_job_table(conn)
        finally:
            conn.close()
        jobs = load_jobs()
        rows = "".join(
            "<tr>"
            f"<td><code>{safe(job['job_id'])}</code></td><td>{safe(job['import_type'])}</td>"
            f"<td>{safe(job['original_filename'])}</td><td><span class=\"badge {safe(job['status'])}\">{safe(job['status'])}</span></td>"
            f"<td>{safe(job['submitted_by'])}</td><td>{safe(job['created_at'])}</td>"
            f"<td>{safe(job['rows_processed'])}</td><td>{safe(job['error_message'])}</td></tr>"
            for job in jobs
        ) or '<tr><td colspan="8" class="empty">No hay cargas registradas.</td></tr>'
        body = f"""<p>Sube reportes Excel para crear un trabajo pendiente que Blue Prism puede recoger en la carpeta de intercambio.</p>
<section class="panel"><h2>Nueva carga</h2><form action="/cargas" method="post" enctype="multipart/form-data" class="form-row">
<label>Tipo de reporte<select name="import_type" required><option value="BUK">BUK · Dotacion_2026</option><option value="SIC">SIC · Ejecucion_SIC</option></select></label>
<label class="wide">Archivo Excel (.xlsx)<input type="file" name="file" accept=".xlsx" required></label>
<label>Solicitante<input name="submitted_by" maxlength="255" value="Planificación Central" required></label>
<button type="submit">Validar y poner en cola</button></form>
<p class="subtle">Máximo 20 MB. Se valida el formato y la hoja esperada antes de guardar el original. No se modifican datos hasta que Blue Prism procese el trabajo.</p></section>
<section class="panel"><h2>Historial de cargas</h2><div class="table-wrap"><table><thead><tr>
<th>Job ID</th><th>Tipo</th><th>Archivo</th><th>Estado</th><th>Solicitante</th><th>Creada</th><th>Filas</th><th>Detalle</th>
</tr></thead><tbody>{rows}</tbody></table></div></section>
<section class="panel"><h2>Carpeta para Blue Prism</h2><p>Entrada configurada: <code>{safe(INBOX_DIR)}</code></p>
<p>Un trabajo <b>PENDING</b> contiene <code>source.xlsx</code> y <code>manifest.json</code>. La web no inicia Blue Prism: debe existir un proceso programado que recoja trabajos pendientes.</p></section>"""
        return page("Cargas BUK y SIC", body, "loads")
    except DashboardError as exc:
        return error_page("Cargas BUK y SIC", "loads", str(exc))


@app.post("/cargas", response_class=HTMLResponse)
async def upload(
    import_type: str = Form(...),
    submitted_by: str = Form(...),
    file: UploadFile = File(...),
):
    if import_type not in UPLOAD_SHEETS:
        return error_page("Cargas BUK y SIC", "loads", "Tipo de reporte no admitido.")
    filename = Path(file.filename or "").name
    if not filename or Path(filename).suffix.lower() != ".xlsx":
        return error_page("Cargas BUK y SIC", "loads", "Selecciona un archivo .xlsx.")
    submitted_by = submitted_by.strip()
    if not submitted_by or len(submitted_by) > 255:
        return error_page("Cargas BUK y SIC", "loads", "El nombre del solicitante es obligatorio.")

    job_id = str(uuid.uuid4())
    job_dir = INBOX_DIR / job_id
    temp_path: Path | None = None
    try:
        INBOX_DIR.mkdir(parents=True, exist_ok=True)
        job_dir.mkdir()
        with tempfile.NamedTemporaryFile(dir=job_dir, prefix="upload-", suffix=".xlsx", delete=False) as temp:
            temp_path = Path(temp.name)
            digest = hashlib.sha256()
            size = 0
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise DashboardError("El archivo supera el límite de 20 MB.")
                digest.update(chunk)
                temp.write(chunk)
        sheets = validate_workbook(temp_path, import_type)
        stored_file = job_dir / "source.xlsx"
        temp_path.replace(stored_file)
        manifest = {
            "job_id": job_id,
            "type": import_type,
            "original_filename": filename,
            "stored_filename": stored_file.name,
            "sha256": digest.hexdigest(),
            "size_bytes": size,
            "sheets": sheets,
            "submitted_by": submitted_by,
            "created_at_utc": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "status": "PENDING",
        }
        (job_dir / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        conn = connect()
        try:
            ensure_job_table(conn)
            conn.cursor().execute(
                """
                INSERT INTO dbo.web_import_jobs
                    (job_id,import_type,original_filename,stored_path,sha256,status,submitted_by)
                VALUES (?,?,?,?,?,'PENDING',?)
                """,
                job_id, import_type, filename, str(stored_file), digest.hexdigest(), submitted_by,
            )
            conn.commit()
        finally:
            conn.close()
        return RedirectResponse("/cargas", status_code=303)
    except (DashboardError, OSError, pyodbc.Error) as exc:
        logger.exception("Falló la creación de la carga Excel")
        if job_dir.exists():
            shutil.rmtree(job_dir)
        return error_page("Cargas BUK y SIC", "loads", str(exc))
    finally:
        await file.close()
        if temp_path and temp_path.exists():
            temp_path.unlink()


@app.get("/parametros", response_class=HTMLResponse)
def settings():
    try:
        globals_rows = query(
            "SELECT clave, valor_numerico, valor_texto, descripcion FROM dbo.parametros_globales ORDER BY clave"
        )
        distribution = query(
            "SELECT TOP (100) id,cargo,std_min_directa,std_min_indirecta,"
            "t1_pct_dir,t1_pct_ind,t1_pct_nas,t2_pct_dir,t2_pct_ind,t2_pct_nas,"
            "t3_pct_dir,t3_pct_ind,t3_pct_nas FROM dbo.distribucion_nac ORDER BY cargo"
        )
        global_html = "".join(
            f"""<tr><td><b>{safe(row['clave'])}</b><div class="subtle">{safe(row['descripcion'])}</div></td>
<td>{safe(row['valor_numerico'])}</td><td>{safe(row['valor_texto'])}</td><td><details><summary>Editar</summary>
<form action="/parametros/global/{safe(row['clave'])}" method="post">
<label>Nuevo valor numérico<input name="value" type="number" step="any" value="{safe(row['valor_numerico'])}" required></label>
<label>Justificación<input name="justification" maxlength="240" required></label><button>Guardar</button></form></details></td></tr>"""
            for row in globals_rows
        ) or '<tr><td colspan="4" class="empty">No hay parámetros en la base.</td></tr>'
        fields = list(DISTRIBUTION_FIELDS)
        columns = ["cargo", "std_min_directa", "std_min_indirecta"] + [
            f"t{tier}_pct_{part}" for tier in (1, 2, 3) for part in ("dir", "ind", "nas")
        ]
        dist_rows = ""
        for row in distribution:
            inputs = "".join(
                f'<label>{safe(field)}<input name="{safe(field)}" type="number" step="any" '
                f'value="{safe(row[field])}" required></label>'
                for field in fields
            )
            dist_rows += (
                f'<tr><td>{safe(row["cargo"])}</td><td><details><summary>Editar regla</summary>'
                f'<form action="/parametros/distribucion/{safe(row["id"])}" method="post">'
                f'<div class="form-row">{inputs}</div><label>Justificación<input name="justification" maxlength="240" required></label>'
                '<button>Guardar porcentajes</button></form></details></td></tr>'
            )
        body = f"""<div class="notice">Los valores se guardan en una transacción y se registran en la auditoría. Los porcentajes se expresan como fracción (por ejemplo, 0.65 = 65%).</div>
<section class="panel"><h2>Parámetros globales</h2><div class="table-wrap"><table><thead><tr><th>Clave</th><th>Valor</th><th>Etiqueta</th><th>Acción</th></tr></thead><tbody>{global_html}</tbody></table></div></section>
<section class="panel"><h2>Reglas de distribución por cargo</h2><p>Se muestran hasta 100 cargos. Cada tramo debe sumar 1.0.</p>
<div class="table-wrap"><table><thead><tr><th>Cargo</th><th>Edición</th></tr></thead><tbody>{dist_rows or '<tr><td colspan="2" class="empty">No hay reglas cargadas.</td></tr>'}</tbody></table></div></section>"""
        return page("Parámetros y reglas", body, "settings")
    except DashboardError as exc:
        return error_page("Parámetros y reglas", "settings", str(exc))


def record_audit(
    cursor,
    *,
    actor: str,
    entity: str,
    field: str,
    old: Any,
    new: Any,
    justification: str,
) -> None:
    cursor.execute(
        """
        INSERT INTO dbo.auditoria_cambios
            (version_id,actor,tipo_evento,entidad_afectada,campo_modificado,
             valor_anterior,valor_nuevo,mes_inicio_vigencia,impacto_atenciones_anual,
             justificacion,estado_aprobacion)
        VALUES (1,?,'CAMBIO_PARAMETRO',?,?,?,?,1,0,?,'APROBADO')
        """,
        actor, entity, field, str(old), str(new), justification,
    )


@app.post("/parametros/global/{key}", response_class=HTMLResponse)
def update_global(key: str, value: str = Form(...), justification: str = Form(...)):
    if key not in GLOBAL_VALUE_RANGES:
        return error_page("Parámetros y reglas", "settings", "Parámetro no editable.")
    try:
        number = float(value)
        minimum, maximum = GLOBAL_VALUE_RANGES[key]
        if not math.isfinite(number) or not minimum <= number <= maximum:
            raise ValueError
    except ValueError:
        return error_page("Parámetros y reglas", "settings", "El valor debe ser un número válido.")
    justification = justification.strip()
    if not justification or len(justification) > 240:
        return error_page("Parámetros y reglas", "settings", "Escribe una justificación de hasta 240 caracteres.")
    try:
        conn = connect()
    except DashboardError as exc:
        return error_page("Parámetros y reglas", "settings", str(exc))
    try:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT valor_numerico FROM dbo.parametros_globales WITH (UPDLOCK,HOLDLOCK) WHERE clave=?",
            key,
        )
        row = cursor.fetchone()
        if row is None:
            raise DashboardError("El parámetro no existe.")
        old = row[0]
        cursor.execute(
            "UPDATE dbo.parametros_globales SET valor_numerico=?,valor_texto=?,updated_at=SYSUTCDATETIME() WHERE clave=?",
            number, global_text_value(key, number), key,
        )
        record_audit(
            cursor, actor="Planificación Central", entity="parametros_globales",
            field=key, old=old, new=number, justification=justification,
        )
        conn.commit()
        return RedirectResponse("/parametros", status_code=303)
    except DashboardError:
        conn.rollback()
        raise
    except pyodbc.Error as exc:
        conn.rollback()
        logger.exception("No se pudo actualizar el parámetro %s", key)
        return error_page("Parámetros y reglas", "settings", f"No se guardó el cambio: {exc}")
    finally:
        conn.close()


@app.post("/parametros/distribucion/{rule_id}", response_class=HTMLResponse)
def update_distribution(
    rule_id: int,
    justification: str = Form(...),
    std_min_directa: float = Form(...),
    std_min_indirecta: float = Form(...),
    t1_pct_dir: float = Form(...),
    t1_pct_ind: float = Form(...),
    t1_pct_nas: float = Form(...),
    t2_pct_dir: float = Form(...),
    t2_pct_ind: float = Form(...),
    t2_pct_nas: float = Form(...),
    t3_pct_dir: float = Form(...),
    t3_pct_ind: float = Form(...),
    t3_pct_nas: float = Form(...),
):
    values = locals()
    changes = {field: float(values[field]) for field in DISTRIBUTION_FIELDS}
    for field, number in changes.items():
        minimum, maximum = DISTRIBUTION_FIELDS[field]
        if not math.isfinite(number) or not minimum <= number <= maximum:
            return error_page("Parámetros y reglas", "settings", f"Valor inválido para {field}.")
    for tier in (1, 2, 3):
        total = sum(changes[f"t{tier}_pct_{part}"] for part in ("dir", "ind", "nas"))
        if not math.isclose(total, 1.0, abs_tol=0.0001):
            return error_page("Parámetros y reglas", "settings", f"Los porcentajes del tramo {tier} deben sumar 1.0.")
    justification = justification.strip()
    if not justification or len(justification) > 240:
        return error_page("Parámetros y reglas", "settings", "Escribe una justificación de hasta 240 caracteres.")
    try:
        conn = connect()
    except DashboardError as exc:
        return error_page("Parámetros y reglas", "settings", str(exc))
    try:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT cargo," + ",".join(DISTRIBUTION_FIELDS) +
            " FROM dbo.distribucion_nac WITH (UPDLOCK,HOLDLOCK) WHERE id=?",
            rule_id,
        )
        row = cursor.fetchone()
        if row is None:
            return error_page("Parámetros y reglas", "settings", "No se encontró la regla.")
        cargo = row[0]
        old_values = dict(zip(DISTRIBUTION_FIELDS, row[1:]))
        assignments = ",".join(f"{field}=?" for field in DISTRIBUTION_FIELDS)
        cursor.execute(
            f"UPDATE dbo.distribucion_nac SET {assignments} WHERE id=?",
            *[changes[field] for field in DISTRIBUTION_FIELDS], rule_id,
        )
        for field in DISTRIBUTION_FIELDS:
            if float(old_values[field] or 0) != changes[field]:
                record_audit(
                    cursor, actor="Planificación Central", entity=f"distribucion_nac:{cargo}",
                    field=field, old=old_values[field], new=changes[field],
                    justification=justification,
                )
        conn.commit()
        return RedirectResponse("/parametros", status_code=303)
    except pyodbc.Error as exc:
        conn.rollback()
        logger.exception("No se pudo actualizar la distribución %s", rule_id)
        return error_page("Parámetros y reglas", "settings", f"No se guardó la regla: {exc}")
    finally:
        conn.close()


@app.get("/auditoria", response_class=HTMLResponse)
def audit():
    try:
        changes = query(
            "SELECT TOP (200) fecha_hora,version_id,actor,tipo_evento,it,instituto,persona_id,"
            "entidad_afectada,campo_modificado,valor_anterior,valor_nuevo,justificacion "
            "FROM dbo.auditoria_cambios ORDER BY id DESC"
        )
        errors = query(
            "SELECT TOP (200) fecha_deteccion,origen_archivo,persona_id,it,cargo,campo_error,"
            "valor_recibido,descripcion_error,severidad FROM dbo.errores_validacion ORDER BY id DESC"
        )
        def table(records: list[dict[str, Any]]) -> str:
            if not records:
                return '<div class="empty">No hay registros.</div>'
            headers = list(records[0])
            return (
                '<div class="table-wrap"><table><thead><tr>'
                + "".join(f"<th>{safe(header)}</th>" for header in headers)
                + "</tr></thead><tbody>"
                + "".join(
                    "<tr>" + "".join(f"<td>{safe(row.get(header))}</td>" for header in headers) + "</tr>"
                    for row in records
                )
                + "</tbody></table></div>"
            )
        body = (
            f'<section class="panel"><h2>Cambios y trazabilidad <span class="badge">{len(changes)}</span></h2>{table(changes)}</section>'
            f'<section class="panel"><h2>Errores de validación <span class="badge">{len(errors)}</span></h2>{table(errors)}</section>'
        )
        return page("Auditoría y validaciones", body, "audit")
    except DashboardError as exc:
        return error_page("Auditoría y validaciones", "audit", str(exc))


@app.post("/api/jobs/{job_id}/status")
def update_job_status(
    job_id: str,
    status: str = Form(...),
    rows_processed: int | None = Form(None),
    error_message: str = Form(""),
    x_job_token: str | None = Header(None),
):
    expected_token = os.getenv("TELETON_JOB_TOKEN")
    if expected_token and x_job_token != expected_token:
        raise HTTPException(status_code=403, detail="Invalid job token")
    if status not in {"RUNNING", "SUCCEEDED", "FAILED"}:
        raise HTTPException(status_code=400, detail="Invalid status")
    if rows_processed is not None and rows_processed < 0:
        raise HTTPException(status_code=400, detail="Invalid rows_processed")
    if len(error_message) > 2000:
        raise HTTPException(status_code=400, detail="error_message is too long")
    try:
        uuid.UUID(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="Invalid job id") from exc
    try:
        conn = connect()
    except DashboardError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT status FROM dbo.web_import_jobs WITH (UPDLOCK,HOLDLOCK) WHERE job_id=?", job_id)
        current = cursor.fetchone()
        if current is None:
            raise HTTPException(status_code=404, detail="Job not found")
        allowed = {
            "PENDING": {"RUNNING"},
            "RUNNING": {"SUCCEEDED", "FAILED"},
            "FAILED": {"RUNNING"},
            "SUCCEEDED": set(),
        }
        if status not in allowed.get(current[0], set()):
            raise HTTPException(status_code=409, detail=f"Transition {current[0]} -> {status} is not allowed")
        cursor.execute(
            """
            UPDATE dbo.web_import_jobs
            SET status=?, rows_processed=COALESCE(?,rows_processed),
                error_message=?, completed_at=CASE WHEN ? IN ('SUCCEEDED','FAILED')
                    THEN SYSUTCDATETIME() ELSE NULL END
            WHERE job_id=?
            """,
            status, rows_processed, error_message or None, status, job_id,
        )
        conn.commit()
        return {"job_id": job_id, "status": status}
    except HTTPException:
        conn.rollback()
        raise
    except pyodbc.Error as exc:
        conn.rollback()
        logger.exception("No se pudo actualizar el trabajo %s", job_id)
        raise HTTPException(status_code=500, detail="Database update failed") from exc
    finally:
        conn.close()


@app.get("/health")
def health():
    try:
        conn = connect()
        try:
            db_name = conn.cursor().execute("SELECT DB_NAME()").fetchone()[0]
        finally:
            conn.close()
        return {"status": "ok", "database": db_name, "inbox": str(INBOX_DIR)}
    except DashboardError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
