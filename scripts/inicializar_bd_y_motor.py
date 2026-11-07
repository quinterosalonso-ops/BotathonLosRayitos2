#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inicializador MSSQLLocalDB, motor parametrizable y generador de Excel para Blue Prism."""

import os
import sys
import math
import re
import zipfile
import datetime
import xml.etree.ElementTree as ET
from collections import defaultdict

import pyodbc

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOLICITUD_DIR = os.path.join(BASE_DIR, 'Solicitud')
DB_DIR = os.path.join(BASE_DIR, 'database')
DEMO_DIR = os.path.join(BASE_DIR, 'datos_demo')
SALIDAS_DIR = os.path.join(BASE_DIR, 'salidas')

os.makedirs(DB_DIR, exist_ok=True)
os.makedirs(DEMO_DIR, exist_ok=True)
os.makedirs(SALIDAS_DIR, exist_ok=True)

DATABASE_NAME = 'TeletonPlanificacion'
SQL_SCHEMA_PATH = os.path.join(DB_DIR, 'teleton_planificacion_mssql.sql')
EXCEL_CONTRATOS = os.path.join(SOLICITUD_DIR, 'Nac - Base contratos x mes.xlsx')
EXCEL_PLANIFICACION = os.path.join(SOLICITUD_DIR, 'Nac - Planificación Asistencial x mes.xlsx')


def connection_string(database):
    drivers = pyodbc.drivers()
    driver = next((name for name in ('ODBC Driver 18 for SQL Server', 'ODBC Driver 17 for SQL Server') if name in drivers), None)
    if driver is None:
        raise RuntimeError('No se encontró ODBC Driver 18 o 17 for SQL Server.')
    server = os.getenv('TELETON_SQL_SERVER', r'(LocalDB)\MSSQLLocalDB').strip()
    if not server:
        raise RuntimeError('TELETON_SQL_SERVER no puede estar vacío.')
    return (
        f'DRIVER={{{driver}}};'
        f'SERVER={server};'
        f'DATABASE={database};'
        'Trusted_Connection=yes;TrustServerCertificate=yes;'
    )


def ensure_database():
    with pyodbc.connect(connection_string('master'), autocommit=True) as conn:
        conn.cursor().execute(
            "IF DB_ID(N'TeletonPlanificacion') IS NULL CREATE DATABASE [TeletonPlanificacion]"
        )


def connect_database():
    ensure_database()
    return pyodbc.connect(connection_string(DATABASE_NAME))

NS = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
EXCEL_EPOCH = datetime.date(1899, 12, 30)


def excel_serial_to_iso(val):
    if val is None or val == '':
        return ''
    try:
        serial = int(float(val))
        if 10000 <= serial <= 80000:
            return (EXCEL_EPOCH + datetime.timedelta(days=serial)).isoformat()
    except Exception:
        pass
    return str(val).strip()


def to_float(val, default=0.0):
    if val is None or val == '':
        return default
    try:
        return float(val)
    except Exception:
        return default


def to_int(val, default=0):
    if val is None or val == '':
        return default
    try:
        return int(float(val))
    except Exception:
        return default


def load_xlsx_sheet(path, sheet_xml_path):
    with zipfile.ZipFile(path, 'r') as z:
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss_xml = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in ss_xml.findall('main:si', NS):
                shared_strings.append(''.join(t.text or '' for t in si.findall('.//main:t', NS)))
        ws_xml = ET.fromstring(z.read(sheet_xml_path))
        rows = {}
        for row in ws_xml.findall('.//main:sheetData/main:row', NS):
            r_idx = int(row.attrib.get('r', 0))
            cells = {}
            for c in row.findall('main:c', NS):
                ref = c.attrib.get('r', '')
                col = ''.join(ch for ch in ref if ch.isalpha())
                t = c.attrib.get('t', '')
                v_el = c.find('main:v', NS)
                is_el = c.find('main:is', NS)
                val = v_el.text if v_el is not None else ''
                if t == 's' and val != '':
                    val = shared_strings[int(val)]
                elif t == 'inlineStr' and is_el is not None:
                    val = ''.join(x.text or '' for x in is_el.findall('.//main:t', NS))
                cells[col] = val
            rows[r_idx] = cells
        return rows


def create_schema(conn):
    with open(SQL_SCHEMA_PATH, 'r', encoding='utf-8') as schema_file:
        schema = schema_file.read()
    cursor = conn.cursor()
    for batch in re.split(r"(?im)^\s*GO\s*$", schema):
        if batch.strip():
            cursor.execute(batch)
    conn.commit()


def seed_all_data(conn):
    cur = conn.cursor()

    # 1. Parámetros Globales
    params_globales = [
        ('META_GESTION_PCT', 0.90, '90%', 'Piso mínimo de cumplimiento de la proyección anual (Meta de Gestión)'),
        ('JORNADA_LEGAL_ENE_ABR', 44.0, '44 hrs', 'Jornada máxima legal vigente de enero a abril de 2026'),
        ('JORNADA_LEGAL_MAY_DIC', 42.0, '42 hrs', 'Jornada máxima legal vigente desde mayo a diciembre de 2026 (Ley 40 Horas)'),
        ('TRAMO_1_MIN_HRS', 40.0, '>=40', 'Corte inferior para Tramo 1 de jornada (40 horas o más)'),
        ('TRAMO_2_MIN_HRS', 22.0, '22-39.9', 'Corte inferior para Tramo 2 de jornada (entre 22 y menos de 40 horas)'),
        ('PAUSA_SALUDABLE_T1_PCT', 0.10, '10%', 'Porcentaje de pausa saludable para jornadas >= 40 H/S'),
        ('PAUSA_SALUDABLE_MENOR_40_PCT', 0.06, '6%', 'Porcentaje de pausa saludable para jornadas < 40 H/S (comentario celda AE4)'),
        ('BENEF_COLACION_15M_HRS_DIA', 0.25, '15 min/día', 'Horas diarias de beneficio colación (1,25 h/sem en jornada de 5 días)'),
        ('MAX_ASIG_MAS_30HS', 2.0, '2 asignaciones', 'Máximo de asignaciones especiales estándar para profesionales con > 30 H/S'),
        ('FACTOR_DIAS_PROGRESIVOS', 0.70, '70%', 'Factor de ponderación de días progresivos (celda DQ2)'),
        ('FACTOR_DIAS_COMPLEMENTARIOS', 0.80, '80%', 'Factor de ponderación de días complementarios (celda DR2)'),
        ('VACACIONES_COLECTIVAS_DIAS', 15.0, '15 días', 'Días hábiles de feriado colectivo en febrero (resto de Chile)'),
        ('VACACIONES_AYSEN_DIAS', 20.0, '20 días', 'Días hábiles de feriado legal en zona extrema Aysén/Coyhaique'),
    ]
    cur.executemany(
        "INSERT INTO parametros_globales (clave, valor_numerico, valor_texto, descripcion) VALUES (?, ?, ?, ?)",
        params_globales
    )

    # 2. Institutos (con proyección de pacientes de Planificacion Total)
    institutos_data = [
        (11, 'Ari', 'Arica', 'Norte Grande', 0, 15.0, 0.0, 3.0, 5.0, 735.13),
        (12, 'Iqu', 'Iquique', 'Norte Grande', 0, 15.0, 0.0, 3.0, 5.0, 869.59),
        (21, 'Ant', 'Antofagasta', 'Norte Grande', 0, 15.0, 0.0, 3.0, 5.0, 1265.16),
        (22, 'Cal', 'Calama', 'Norte Grande', 0, 15.0, 0.0, 3.0, 5.0, 753.22),
        (31, 'Cop', 'Copiapó', 'Norte Chico', 0, 15.0, 0.0, 3.0, 5.0, 751.41),
        (41, 'Coq', 'Coquimbo', 'Norte Chico', 0, 15.0, 0.0, 3.0, 5.0, 1700.02),
        (51, 'Val', 'Valparaíso', 'Centro', 0, 15.0, 0.0, 3.0, 5.0, 3742.64),
        (61, 'OHi', "O'Higgins", 'Centro', 0, 15.0, 0.0, 3.0, 5.0, 0.0),
        (71, 'Tal', 'Talca', 'Centro Sur', 0, 15.0, 0.0, 3.0, 5.0, 2116.76),
        (81, 'Con', 'Concepción', 'Sur', 1, 15.0, 0.0, 3.0, 5.0, 3533.09),
        (91, 'Tem', 'Temuco', 'Sur', 0, 15.0, 0.0, 3.0, 5.0, 2137.44),
        (101, 'PMo', 'Puerto Montt', 'Sur', 0, 15.0, 0.0, 3.0, 5.0, 2137.60),
        (102, 'Cas', 'Castro', 'Sur Austral', 0, 15.0, 0.0, 3.0, 5.0, 0.0),
        (111, 'Coy', 'Coyhaique', 'Zona Extrema Aysén', 0, 20.0, 0.4, 3.6, 5.0, 688.96),
        (131, 'San', 'Santiago', 'Metropolitana', 1, 15.0, 0.0, 3.0, 5.0, 11610.00),
        (141, 'Vdv', 'Valdivia', 'Sur', 0, 15.0, 0.0, 3.0, 5.0, 806.42),
        (161, 'Nub', 'Ñuble', 'Centro Sur', 0, 15.0, 0.0, 3.0, 5.0, 0.0),
    ]
    cur.executemany("""
        INSERT INTO institutos (
            it, codigo_inst, nombre_instituto, zona_geografica,
            aplica_grupales_mayor, dias_vacaciones_legales, sem_vac_ene, sem_vac_feb,
            dias_feriado_adicional, proyeccion_pacientes_2025
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, institutos_data)

    # 3. Especialidades y Jornada Nacional por Área (Parámetros!AB34:AC44 y AX1:AY25)
    especialidades_data = [
        ('FI', 'FISIATRA', 1.5),
        ('KI', 'KINESIÓLOGO', 0.5),
        ('TO', 'TERAPEUTA OCUPACIONAL', 0.5),
        ('ED', 'EDUCADORA', 1.5),
        ('PC', 'PSICÓLOGO', 1.5),
        ('FO', 'FONOAUDIÓLOGO', 0.0),
        ('EN', 'ENFERMERA', 0.0),
        ('AS', 'ASISTENTE SOCIAL', 1.5),
        ('UTAC', 'U.TERAPIAS ARTISTICAS CREATIVAS', 1.5),
        ('UEDA', 'U.ED.FISICA Y DEPORTE ADAPTADO', 1.0),
        ('PS', 'PSIQUIATRA', 0.0),
        ('PE', 'PEDIATRA', 0.0),
        ('OD', 'ODONTÓLOGO', 0.0),
        ('BA', 'BIOINGENIERIA', 0.0),
        ('NU', 'NUTRICIONISTA', 0.0),
        ('ULAB', 'U.LABORAL', 1.5),
        ('NE', 'NEURÓLOGO', 0.0),
        ('UTA', 'U.TECNOLOGIA ASISTIVA', 0.5),
        ('OR', 'ORTOPEDISTA', 0.0),
        ('UR', 'URÓLOGO', 0.0),
        ('NC', 'NEUROCIRUGÍA', 0.0),
        ('LOP', 'LAB.ORTESIS Y PROTESIS', 0.0),
        ('MG', 'MÉDICO GENERAL', 0.0),
        ('HO', 'HOSPITALIZADOS', 0.0),
        ('TR', 'TELEREHABILITACIÓN', 0.0),
        ('SM', 'SALUD MENTAL / DIGITAL', 0.0),
    ]
    cur.executemany(
        "INSERT INTO especialidades (cod_esp, nombre_especialidad, dias_jornada_nacional) VALUES (?, ?, ?)",
        especialidades_data
    )

    # 4. Factores Grupales (Parámetros!F4:I8)
    factores_grupales_data = [
        ('AS', 0.10, 3.0, 1.5),
        ('ED', 0.10, 3.0, 1.5),
        ('UEDA', 0.55, 6.0, 3.0),
        ('ULAB', 0.10, 3.0, 1.5),
        ('UTAC', 0.55, 6.0, 3.0),
    ]
    cur.executemany(
        "INSERT INTO factores_grupales (cod_esp, pct_grupal_estandar, num_pacientes_san_con, num_pacientes_otros_it) VALUES (?, ?, ?, ?)",
        factores_grupales_data
    )

    # 5. Calendario 2026 (Parámetros!S3:AD14)
    calendario_data = [
        (1, 'ene', '2026-01-01', '2026-01-31', 46023, 31, 10, 10, 21, 21, 4.2, 'EXACT'),
        (2, 'feb', '2026-02-01', '2026-02-28', 46054, 28, 8, 23, 20, 5, 4.0, 'EXACT'),
        (3, 'mar', '2026-03-01', '2026-03-31', 46082, 31, 9, 9, 22, 22, 4.4, 'EXACT'),
        (4, 'abr', '2026-04-01', '2026-04-30', 46113, 30, 9, 9, 21, 21, 4.2, 'EXACT'),
        (5, 'may', '2026-05-01', '2026-05-31', 46143, 31, 12, 12, 19, 19, 3.8, 'ROUNDUP_1'),
        (6, 'jun', '2026-06-01', '2026-06-30', 46174, 30, 9, 9, 21, 21, 4.2, 'EXACT'),
        (7, 'jul', '2026-07-01', '2026-07-31', 46204, 31, 9, 9, 22, 22, 4.4, 'ROUNDUP_1'),
        (8, 'ago', '2026-08-01', '2026-08-31', 46235, 31, 10, 10, 21, 21, 4.2, 'EXACT'),
        (9, 'sep', '2026-09-01', '2026-09-30', 46266, 30, 9, 9, 21, 21, 4.2, 'EXACT'),
        (10, 'oct', '2026-10-01', '2026-10-31', 46296, 31, 10, 10, 21, 21, 4.2, 'EXACT'),
        (11, 'nov', '2026-11-01', '2026-11-30', 46327, 30, 9, 9, 21, 21, 4.2, 'EXACT'),
        (12, 'dic', '2026-12-01', '2026-12-31', 46357, 31, 10, 10, 21, 21, 4.2, 'EXACT'),
    ]
    cur.executemany("""
        INSERT INTO calendario_2026 (
            mes, nombre_mes, fecha_inicio, fecha_termino, periodo_serial_excel,
            dias_totales, dias_inhabiles_sin_vac, dias_inhabiles_con_vac,
            dias_habiles_sin_vac, dias_habiles_con_vac, semanas_disponibles, regla_redondeo_semhab
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, calendario_data)

    # 6. Feriados Nacionales y Regionales 2026 (Calendario 2026!A105:C122 y Parámetros!J121:O138)
    feriados_data = [
        ('2026-01-01', 1, 4, 'Año Nuevo', 'NACIONAL'),
        ('2026-04-03', 4, 5, 'Viernes Santo', 'NACIONAL'),
        ('2026-04-04', 4, 6, 'Sábado Santo', 'NACIONAL'),
        ('2026-05-01', 5, 5, 'Día del Trabajo', 'NACIONAL'),
        ('2026-05-21', 5, 4, 'Día de las Glorias Navales', 'NACIONAL'),
        ('2026-06-07', 6, 7, 'Asalto y Toma del Morro de Arica', 'Ari'),
        ('2026-06-21', 6, 7, 'Día Nacional de los Pueblos Indígenas', 'NACIONAL'),
        ('2026-06-29', 6, 1, 'San Pedro y San Pablo', 'NACIONAL'),
        ('2026-07-16', 7, 4, 'Día de la Virgen del Carmen', 'NACIONAL'),
        ('2026-08-10', 8, 1, 'San Lorenzo de Tarapacá y Atacama', 'Iqu-Cop'),
        ('2026-08-15', 8, 6, 'Asunción de la Virgen', 'NACIONAL'),
        ('2026-08-20', 8, 4, "Natalicio de Bernardo O'Higgins (Chillán)", 'Nub'),
        ('2026-09-18', 9, 5, 'Día de la Independencia Nacional', 'NACIONAL'),
        ('2026-09-19', 9, 6, 'Día de las Glorias del Ejército', 'NACIONAL'),
        ('2026-09-20', 9, 7, 'Fiesta de la Pampilla (Coquimbo)', 'Coq'),
        ('2026-10-12', 10, 1, 'Encuentro de Dos Mundos', 'NACIONAL'),
        ('2026-10-31', 10, 6, 'Día de las Iglesias Evangélicas y Protestantes', 'NACIONAL'),
        ('2026-11-01', 11, 7, 'Día de Todos los Santos', 'NACIONAL'),
        ('2026-12-08', 12, 2, 'Inmaculada Concepción', 'NACIONAL'),
        ('2026-12-25', 12, 5, 'Navidad', 'NACIONAL'),
    ]
    cur.executemany(
        "INSERT INTO feriados_regionales (fecha, mes, dia_semana, nombre_feriado, ambito_it) VALUES (?, ?, ?, ?, ?)",
        feriados_data
    )

    # 7. Leer Distribucion_NAC y Catálogo de Asignaciones desde Excel
    dist_rows = load_xlsx_sheet(EXCEL_PLANIFICACION, 'xl/worksheets/sheet3.xml')
    seen_cargos = set()
    for r_idx in range(7, 93):
        c = dist_rows.get(r_idx, {})
        cargo = c.get('B', '').strip()
        if cargo and cargo not in seen_cargos:
            seen_cargos.add(cargo)
            cur.execute("""
                INSERT INTO distribucion_nac (
                    cod_esp, cargo, unidad, std_min_directa, std_min_indirecta, std_min_pabellon,
                    t1_pct_dir, t1_pct_ind, t1_pct_nas,
                    t2_pct_dir, t2_pct_ind, t2_pct_nas,
                    t3_pct_dir, t3_pct_ind, t3_pct_nas, es_quirurgico
                ) VALUES (?, ?, ?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
            """, (
                c.get('A', '').strip(), cargo, c.get('C', '').strip(),
                to_float(c.get('D'), 40.0), to_float(c.get('E'), 30.0),
                to_float(c.get('F'), 0.0), to_float(c.get('G'), 0.0), to_float(c.get('H'), 0.0),
                to_float(c.get('I'), 0.0), to_float(c.get('J'), 0.0), to_float(c.get('K'), 0.0),
                to_float(c.get('L'), 0.0), to_float(c.get('M'), 0.0), to_float(c.get('N'), 0.0)
            ))
    # Quirúrgicos (filas 97 a 102)
    for r_idx in range(97, 103):
        c = dist_rows.get(r_idx, {})
        cargo = c.get('B', '').strip()
        if cargo and cargo not in seen_cargos:
            seen_cargos.add(cargo)
            cur.execute("""
                INSERT INTO distribucion_nac (
                    cod_esp, cargo, unidad, std_min_directa, std_min_indirecta, std_min_pabellon,
                    t1_pct_dir, t1_pct_ind, t1_pct_nas,
                    t2_pct_dir, t2_pct_ind, t2_pct_nas,
                    t3_pct_dir, t3_pct_ind, t3_pct_nas, es_quirurgico
                ) VALUES (?, ?, ?, ?, 30, ?, ?, 0, ?, ?, 0, ?, ?, 0, ?, 1)
            """, (
                'OR' if 'ORTOPED' in cargo or 'QUIRURGICO' in cargo else ('UR' if 'UROLOG' in cargo else 'NE'),
                cargo, c.get('C', '').strip(),
                to_float(c.get('E'), 30.0), to_float(c.get('D'), 240.0),
                to_float(c.get('G'), 0.3), to_float(c.get('H'), 0.1),
                to_float(c.get('G'), 0.3), to_float(c.get('H'), 0.1),
                to_float(c.get('G'), 0.3), to_float(c.get('H'), 0.1)
            ))

    # Catálogo de Asignaciones Especiales (filas 7 a 39, cols R, S, T)
    seen_asig = set()
    for r_idx in range(7, 40):
        c = dist_rows.get(r_idx, {})
        nombre_asig = c.get('R', '').strip()
        if nombre_asig and nombre_asig not in seen_asig:
            seen_asig.add(nombre_asig)
            hrs = to_float(c.get('S'), 1.0)
            desc = c.get('T', '').strip()
            req_dmn = 1 if ('DMN' in desc.upper() or 'CEC' in desc.upper() or 'GP' in desc.upper()) else 0
            cur.execute("""
                INSERT INTO catalogo_asignaciones (nombre_asignacion, horas_semanales, requiere_autorizacion_dmn, descripcion)
                VALUES (?, ?, ?, ?)
            """, (nombre_asig, hrs, req_dmn, desc))

    # 8. LOP (sheet6.xml)
    lop_rows = load_xlsx_sheet(EXCEL_PLANIFICACION, 'xl/worksheets/sheet6.xml')
    for r_idx in range(2, 18):
        c = lop_rows.get(r_idx, {})
        if c.get('A'):
            cur.execute("DELETE FROM lop_instituto WHERE it=?", (to_int(c.get('A')),))
            cur.execute(
                "INSERT INTO lop_instituto (it, instituto, total_atenciones_lop) VALUES (?, ?, ?)",
                (to_int(c.get('A')), c.get('B', '').strip(), to_float(c.get('C'), 0.0))
            )

    # 9. Excepciones Multi-Instituto (sheet7.xml)
    exc_rows = load_xlsx_sheet(EXCEL_PLANIFICACION, 'xl/worksheets/sheet7.xml')
    for r_idx in range(3, 20):
        c = exc_rows.get(r_idx, {})
        pid = to_int(c.get('E'))
        it = to_int(c.get('C'))
        if pid > 0 and it > 0:
            excep_id = c.get('D', '').strip() or f"{r_idx}-{pid}"
            cur.execute("DELETE FROM excepciones_multi_it WHERE excep_id=?", (excep_id,))
            cur.execute("""
                INSERT INTO excepciones_multi_it (
                    excep_id, persona_id, it, instituto, nombre_it, cargo, ccosto, fec_ingreso,
                    hs_ene, hs_feb, hs_mar, hs_abr, hs_may, hs_jun, hs_jul, hs_ago, hs_sep, hs_oct, hs_nov, hs_dic,
                    dias_semana
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                excep_id, pid, it, c.get('B', '').strip(), c.get('A', '').strip(),
                c.get('F', '').strip(), to_int(c.get('G')), excel_serial_to_iso(c.get('H')),
                to_float(c.get('K')), to_float(c.get('L')), to_float(c.get('M')), to_float(c.get('N')),
                to_float(c.get('O')), to_float(c.get('P')), to_float(c.get('Q')), to_float(c.get('R')),
                to_float(c.get('S')), to_float(c.get('T')), to_float(c.get('U')), to_float(c.get('V')),
                to_float(c.get('W'), 5.0)
            ))

    # 10. Contratos BUK (Nac - Base contratos x mes.xlsx -> sheet1.xml)
    cont_rows = load_xlsx_sheet(EXCEL_CONTRATOS, 'xl/worksheets/sheet1.xml')
    period_to_month = {
        46023: 1, 46054: 2, 46082: 3, 46113: 4, 46143: 5, 46174: 6, 46204: 7,
        46235: 8, 46266: 9, 46296: 10, 46327: 11, 46357: 12
    }
    cont_by_pid_it = defaultdict(dict)
    for r_idx in sorted(cont_rows.keys()):
        if r_idx < 3:
            continue
        c = cont_rows[r_idx]
        if not c.get('A') or not c.get('C') or not c.get('H'):
            continue
        it = to_int(c.get('A'))
        pid = to_int(c.get('C'))
        p_serial = to_int(c.get('H'))
        mes = period_to_month.get(p_serial, 0)
        hs = to_float(c.get('M'))
        cont_by_pid_it[(pid, it)][mes] = hs
        cur.execute("DELETE FROM contratos_buk WHERE it=? AND persona_id=? AND mes=?", (it, pid, mes))
        cur.execute("""
            INSERT INTO contratos_buk (
                it, instituto, persona_id, area_buk, nombre_subarea, centro_costo,
                fecha_nacimiento, periodo_serial, periodo_fecha, mes, tipo_contrato,
                tipo_trab_anterior, fecha_ingreso, fecha_termino, horas_semanales,
                cargo, tipo_trabajador, dias_jornada, reg_sic, cod_esp, ndias, origen_carga
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'BUK_INICIAL')
        """, (
            it, c.get('B', '').strip(), pid, c.get('D', '').strip(), c.get('E', '').strip(),
            to_int(c.get('F')), excel_serial_to_iso(c.get('G')), p_serial,
            excel_serial_to_iso(p_serial), mes, c.get('I', '').strip(),
            c.get('J', '').strip(), excel_serial_to_iso(c.get('K')), excel_serial_to_iso(c.get('L')),
            hs, c.get('N', '').strip(), c.get('O', '').strip(), c.get('P', '').strip(),
            c.get('Q', '').strip(), c.get('R', '').strip(), to_float(c.get('S'), 5.0)
        ))

    # 11. Dotacion_Asistencial (sheet8.xml, filas 5 a 475) + Desglose Mensual + Simulación Realista SIC
    dot_rows = load_xlsx_sheet(EXCEL_PLANIFICACION, 'xl/worksheets/sheet8.xml')
    cal_cols = ['BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT']
    lic_cols = ['BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG']
    vac_cols = ['CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS']
    prm_cols = ['CT','CU','CV','CW','CX','CY','CZ','DA','DB','DC','DD','DE']
    sh_cols  = ['DZ','EA','EB','EC','ED','EE','EF','EG','EH','EI','EJ','EK']
    pad_cols = ['EN','EO','EP','EQ','ER','ES','ET','EU','EV','EW','EX','EY']
    pai_cols = ['FA','FB','FC','FD','FE','FF','FG','FH','FI','FJ','FK','FL']
    mgd_cols = ['FN','FO','FP','FQ','FR','FS','FT','FU','FV','FW','FX','FY']
    mgi_cols = ['GA','GB','GC','GD','GE','GF','GG','GH','GI','GJ','GK','GL']
    mes_nombres = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']

    for r_idx in range(5, 476):
        c = dot_rows.get(r_idx, {})
        if not c.get('B') or not c.get('K'):
            continue
        it = to_int(c.get('B'))
        inst = c.get('C', '').strip()
        pid = to_int(c.get('K'))
        excep_id = c.get('F', '').strip()
        w_dic = to_float(c.get('W'))
        x_pp = to_float(c.get('X'))
        ac_col = to_float(c.get('AC'))

        # Reconstruir H/S mes 1..12 desde Excepciones o Contratos BUK
        cal_vals = [to_float(c.get(col)) for col in cal_cols]
        hs_months = []
        if excep_id:
            cur.execute("SELECT hs_ene,hs_feb,hs_mar,hs_abr,hs_may,hs_jun,hs_jul,hs_ago,hs_sep,hs_oct,hs_nov,hs_dic FROM excepciones_multi_it WHERE excep_id=?", (excep_id,))
            erow = cur.fetchone()
            if erow:
                hs_months = list(erow)
        if not hs_months:
            m_map = cont_by_pid_it.get((pid, it), {})
            jul_hs = m_map.get(7, w_dic)
            for m in range(1, 13):
                if cal_vals[m - 1] == 0:
                    hs_months.append(0.0)
                elif m <= 7:
                    hs_months.append(m_map.get(m, x_pp))
                else:
                    hs_months.append(w_dic if w_dic > 0 else jul_hs)

        tramo = '>=40' if x_pp >= 40 else ('22-40' if x_pp >= 22 else '<22')
        au = to_float(c.get('AU'))
        av = to_float(c.get('AV'))
        aw = to_float(c.get('AW'))
        at = to_float(c.get('AT'))
        ax_hrs = to_float(c.get('AX')) * 24.0
        foja_pct = round(au * 100.0, 2) if w_dic > 0 else 0.0

        # Tasa período 1 (Ene-Abr) deducida o igual a P2
        sh_vals = [to_float(c.get(col)) for col in sh_cols]
        pad_vals = [to_float(c.get(col)) for col in pad_cols]
        pai_vals = [to_float(c.get(col)) for col in pai_cols]
        mgd_vals = [to_float(c.get(col)) for col in mgd_cols]
        mgi_vals = [to_float(c.get(col)) for col in mgi_cols]

        bf_p2 = to_float(c.get('BF'))
        bh_p2 = to_float(c.get('BH'))
        bf_p1 = bf_p2
        bh_p1 = bh_p2
        for m in range(4):
            if sh_vals[m] >= 3.0 and pad_vals[m] > 0:
                bf_p1 = round(pad_vals[m] / sh_vals[m], 2)
                bh_p1 = round(pai_vals[m] / sh_vals[m], 2)
                break

        gq_hrs = to_float(c.get('GQ')) * 24.0
        gr_hrs = to_float(c.get('GR')) * 24.0
        gs_hrs = to_float(c.get('GS')) * 24.0

        vals_insert = [
            r_idx - 4, it, inst, c.get('D', '').strip(), c.get('E', '').strip(), excep_id,
            excel_serial_to_iso(c.get('G')), excel_serial_to_iso(c.get('H')),
            excel_serial_to_iso(c.get('I')), excel_serial_to_iso(c.get('J')),
            pid, c.get('L', '').strip(), c.get('M', '').strip(), c.get('N', '').strip(),
            to_int(c.get('O')), c.get('P', '').strip(), c.get('Q', '').strip(),
            c.get('R', '').strip(), c.get('S', '').strip(),
            excel_serial_to_iso(c.get('T')), excel_serial_to_iso(c.get('U')),
            *hs_months,
            to_int(c.get('V')), w_dic, x_pp, tramo, c.get('Y', '').strip(),
            to_float(c.get('Z'), 5.0), to_float(c.get('AA')),
            c.get('AB', 'N').strip(), ac_col,
            max(0.0, hs_months[0] - ac_col), to_float(c.get('AD')),
            to_float(c.get('AE')), to_float(c.get('AF')), to_float(c.get('AG')), to_float(c.get('AH')),
            to_float(c.get('AI')), to_float(c.get('AJ')), to_float(c.get('AK')),
            c.get('AL', '').strip(), to_float(c.get('AM')),
            c.get('AN', '').strip(), to_float(c.get('AO')),
            c.get('AP', '').strip(), to_float(c.get('AQ')),
            'POR_AUTORIZAR_DMN' if c.get('AP', '').strip() else 'APROBADO',
            to_float(c.get('AR')), to_float(c.get('AS')),
            at, au, av, aw, ax_hrs, bf_p1, max(0.0, bh_p1 - bf_p1), bf_p1, bh_p1,
            at, au, av, aw, ax_hrs,
            to_float(c.get('AY'), 40.0), to_float(c.get('AZ'), 30.0),
            to_float(c.get('BA')), to_float(c.get('BB')),
            to_float(c.get('BC')), to_float(c.get('BD')),
            1 if c.get('BD', '').strip() != '' else 0,
            to_float(c.get('BE')), bf_p2, to_float(c.get('BG')), bh_p2,
            *cal_vals, to_float(c.get('BU')),
            *[to_float(c.get(col)) for col in lic_cols],
            *[to_float(c.get(col)) for col in vac_cols],
            *[to_float(c.get(col)) for col in prm_cols],
            to_float(c.get('DF')), to_float(c.get('DG')),
            to_int(c.get('DH')), to_float(c.get('DI')),
            to_float(c.get('DJ')), to_float(c.get('DK')),
            to_float(c.get('DL')), to_float(c.get('DM')), to_float(c.get('DN')),
            to_float(c.get('DO')), to_float(c.get('DP')), to_float(c.get('DQ')),
            to_float(c.get('DR')), to_float(c.get('DS')), to_float(c.get('DT')),
            to_float(c.get('DU')), to_float(c.get('DV')), to_float(c.get('DW')),
            to_float(c.get('DX')), to_float(c.get('DY')),
            *sh_vals, to_float(c.get('EL')), to_float(c.get('EM')),
            *pad_vals, to_float(c.get('EZ')),
            *pai_vals, to_float(c.get('FM')),
            *mgd_vals, to_float(c.get('FZ')),
            *mgi_vals, to_float(c.get('GM')),
            to_float(c.get('GN')), to_float(c.get('GO')), to_float(c.get('GP')),
            gq_hrs, gr_hrs, gs_hrs,
            to_float(c.get('GT')), to_float(c.get('GU')), to_float(c.get('GV')),
            foja_pct, c.get('GW', '').strip(), 1
        ]

        placeholders = ','.join(['?'] * len(vals_insert))
        cur.execute(f"""
            INSERT INTO dotacion_asistencial (
                orden, it, instituto, division, vacante_nuevo, excepcion_id,
                fecha_inicio_contrato, fecha_termino_trabajo, horario_vigente, fecha_nacimiento,
                persona_id, repetido, cargo, area, ccosto, desc_ccosto, tipo_contrato, tipo_cargo,
                licencia_extensa, periodo_inicio_jornada, periodo_fin_jornada,
                hs_m1, hs_m2, hs_m3, hs_m4, hs_m5, hs_m6, hs_m7, hs_m8, hs_m9, hs_m10, hs_m11, hs_m12,
                vigencia_hs_12, hs_dic, prom_pond_hs, tramo_jornada, cambio_jornada,
                dias_semana_contrato, hd_promedio, colacion_en_contrato, horas_colacion_contrato,
                hs_disp_antes_nas_ene_abr, hs_disp_antes_nas_dic,
                hs_pausa_saludable, hs_benef_colacion_15m, hs_otras_nas, hs_nas_total,
                pct_nas_base, pct_directo_base, pct_indirecto_base,
                asignacion_1, hrs_asig_1, asignacion_2, hrs_asig_2,
                asignacion_3_dmn, hrs_asig_3, estado_aprobacion_dmn, hrs_asig_4_sindical, total_hrs_asignacion,
                pct_nas_final_p1, pct_directo_final_p1, pct_indirecto_final_p1, pct_asistencial_p1,
                hs_disp_ate_p1, ate_dir_sem_p1, ate_ind_sem_p1, ate_dir_grupal_sem_p1, total_ate_sem_p1,
                pct_nas_final_p2, pct_directo_final_p2, pct_indirecto_final_p2, pct_asistencial_p2,
                hs_disp_ate_p2, std_min_directa, std_min_indirecta, ate_dir_sem_p2, ate_ind_sem_p2,
                pct_grupal_estandar, cambio_pct_grupal_dmn, usa_override_grupal, num_pacientes_grupo,
                ate_dir_grupal_sem_p2, ate_ind_total_sem_p2, total_ate_sem_p2,
                cal_m1, cal_m2, cal_m3, cal_m4, cal_m5, cal_m6, cal_m7, cal_m8, cal_m9, cal_m10, cal_m11, cal_m12,
                semanas_habiles_planificadas,
                lic_m1, lic_m2, lic_m3, lic_m4, lic_m5, lic_m6, lic_m7, lic_m8, lic_m9, lic_m10, lic_m11, lic_m12,
                vac_m1, vac_m2, vac_m3, vac_m4, vac_m5, vac_m6, vac_m7, vac_m8, vac_m9, vac_m10, vac_m11, vac_m12,
                perm_m1, perm_m2, perm_m3, perm_m4, perm_m5, perm_m6, perm_m7, perm_m8, perm_m9, perm_m10, perm_m11, perm_m12,
                feriado_zona_extrema_ene, feriado_zona_extrema_feb, mes_cumpleanos, dias_cumpleanos,
                visperas_fiestas_patrias, visperas_navidad_ano_nuevo, feriado_legal_adicional,
                dias_administrativos, plan_anual_capacitacion, hexa_ctr, jornada_nacional,
                dias_progresivos, dias_complementarios, ley_21530, vacaciones_pendientes,
                jornadas_directiva_sindical, ajuste_dif_dias_contrato, lactancia_dias_ano,
                meses_2do_semestre, sem_mes_prorratear_2do_sem,
                semhab_m1, semhab_m2, semhab_m3, semhab_m4, semhab_m5, semhab_m6,
                semhab_m7, semhab_m8, semhab_m9, semhab_m10, semhab_m11, semhab_m12,
                semanas_habiles_totales, dias_habiles_totales,
                plan_dir_m1, plan_dir_m2, plan_dir_m3, plan_dir_m4, plan_dir_m5, plan_dir_m6,
                plan_dir_m7, plan_dir_m8, plan_dir_m9, plan_dir_m10, plan_dir_m11, plan_dir_m12,
                total_plan_ate_dir,
                plan_tot_m1, plan_tot_m2, plan_tot_m3, plan_tot_m4, plan_tot_m5, plan_tot_m6,
                plan_tot_m7, plan_tot_m8, plan_tot_m9, plan_tot_m10, plan_tot_m11, plan_tot_m12,
                total_plan_ate_dir_ind,
                mg_dir_m1, mg_dir_m2, mg_dir_m3, mg_dir_m4, mg_dir_m5, mg_dir_m6,
                mg_dir_m7, mg_dir_m8, mg_dir_m9, mg_dir_m10, mg_dir_m11, mg_dir_m12,
                total_mg_dir,
                mg_tot_m1, mg_tot_m2, mg_tot_m3, mg_tot_m4, mg_tot_m5, mg_tot_m6,
                mg_tot_m7, mg_tot_m8, mg_tot_m9, mg_tot_m10, mg_tot_m11, mg_tot_m12,
                total_mg_dir_ind,
                plan_anual_directa, plan_anual_indirecta, plan_anual_total,
                tiempo_anual_directa_hrs, tiempo_anual_indirecta_hrs, tiempo_anual_total_hrs,
                mg_anual_directa, mg_anual_indirecta, mg_anual_total,
                foja_pct, observaciones, version_id
            ) OUTPUT INSERTED.fila_id VALUES ({placeholders})
        """, vals_insert)
        fila_id = cur.fetchone()[0]

        # Poblar planificacion_mensual (12 meses) y ejecucion_sic (meses 1..6 simulados realistas)
        area = c.get('N', '').strip()
        cargo = c.get('M', '').strip()
        for m in range(1, 13):
            hs_m = hs_months[m - 1]
            cal_m = cal_vals[m - 1]
            sh_m = sh_vals[m - 1]
            hrs_mes_cont = round(hs_m * cal_m, 2)
            hrs_mes_dir = round(hs_m * sh_m * au, 2)
            hrs_mes_ind = round(hs_m * sh_m * av, 2)
            hrs_mes_nas = round(hs_m * sh_m * at, 2)
            foja_m = round((hrs_mes_dir / hrs_mes_cont * 100.0), 2) if hrs_mes_cont > 0 else 0.0
            p_dir = int(pad_vals[m - 1])
            p_tot = int(pai_vals[m - 1])
            p_ind = max(0, p_tot - p_dir)
            m_dir = int(mgd_vals[m - 1])
            m_tot = int(mgi_vals[m - 1])
            m_ind = max(0, m_tot - m_dir)

            cur.execute("""
                INSERT INTO planificacion_mensual (
                    fila_id, persona_id, it, instituto, area, cargo, mes, nombre_mes, periodo_jornada,
                    horas_semanales_contratadas, semanas_calendario, semanas_habiles_efectivas,
                    horas_mes_contratadas, horas_mes_disp_directa, horas_mes_disp_indirecta, horas_mes_nas,
                    foja_mensual_pct, plan_ate_directa, plan_ate_indirecta, plan_ate_total,
                    meta_90_directa, meta_90_indirecta, meta_90_total
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                fila_id, pid, it, inst, area, cargo, m, mes_nombres[m - 1],
                'Ene-Abr (44h)' if m <= 4 else 'May-Dic (42h)',
                hs_m, cal_m, sh_m, hrs_mes_cont, hrs_mes_dir, hrs_mes_ind, hrs_mes_nas,
                foja_m, p_dir, p_ind, p_tot, m_dir, m_ind, m_tot
            ))

            # Simular ejecución SIC para el 1er semestre (Ene..Jun) si no es repetido
            if m <= 6 and p_tot > 0 and c.get('L', '').strip() != 'R':
                # Factor determinístico realista entre 84% y 104% de lo planificado
                factor_sic = 0.86 + ((pid * 7 + m * 13) % 19) * 0.01
                real_dir = int(round(p_dir * factor_sic))
                real_ind = int(round(p_ind * factor_sic))
                real_tot = real_dir + real_ind
                pct_meta = round((real_tot / m_tot * 100.0), 1) if m_tot > 0 else 100.0
                alerta = 'CUMPLIMIENTO_OPTIMO' if pct_meta >= 100.0 else ('EN_RIESGO' if pct_meta >= 90.0 else 'BAJO_META_90')
                cur.execute("DELETE FROM ejecucion_sic WHERE persona_id=? AND it=? AND mes=?", (pid, it, m))
                cur.execute("""
                    INSERT INTO ejecucion_sic (
                        persona_id, it, instituto, area, mes,
                        ate_real_directa, ate_real_indirecta, ate_real_total,
                        pct_cumplimiento_meta, estado_alerta
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (pid, it, inst, area, m, real_dir, real_ind, real_tot, pct_meta, alerta))

    # 12. Registrar evento de auditoría inicial
    cur.execute("""
        INSERT INTO auditoria_cambios (
            version_id, actor, tipo_evento, entidad_afectada, campo_modificado,
            valor_anterior, valor_nuevo, mes_inicio_vigencia, impacto_atenciones_anual,
            justificacion, estado_aprobacion
        ) VALUES (
            1, 'Bot Blue Prism (Data - MSSQLLocalDB)', 'CARGA_INICIAL', 'Dotación Nacional 2026',
            'Línea Base Anual', '0 registros', '471 filas / 462 profesionales / 3.123 contratos BUK',
            1, 692706, 'Inicialización automática desde Nac - Base contratos x mes.xlsx y Nac - Planificación Asistencial x mes.xlsx',
            'APROBADO'
        )
    """)

    conn.commit()


def create_minimal_xlsx(filepath, sheet_name, headers, rows):
    """Genera un archivo .xlsx estándar limpio sin requerir librerías externas."""
    def col_letter(idx):
        res = ''
        n = idx + 1
        while n > 0:
            n, rem = divmod(n - 1, 26)
            res = chr(65 + rem) + res
        return res

    def xml_escape(s):
        return (str(s)
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))

    sheet_rows_xml = []
    all_rows = [headers] + rows
    for r_i, r_vals in enumerate(all_rows, 1):
        cells_xml = []
        for c_i, val in enumerate(r_vals):
            ref = f"{col_letter(c_i)}{r_i}"
            if val is None or val == '':
                continue
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cells_xml.append(f'<c r="{ref}"><v>{val}</v></c>')
            else:
                cells_xml.append(f'<c r="{ref}" t="inlineStr"><is><t>{xml_escape(val)}</t></is></c>')
        sheet_rows_xml.append(f'<row r="{r_i}">{"".join(cells_xml)}</row>')

    worksheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(sheet_rows_xml)}</sheetData>'
        '</worksheet>'
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{xml_escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )
    wb_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '</Relationships>'
    )
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>'
    )

    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types_xml)
        z.writestr('_rels/.rels', root_rels_xml)
        z.writestr('xl/workbook.xml', workbook_xml)
        z.writestr('xl/_rels/workbook.xml.rels', wb_rels_xml)
        z.writestr('xl/worksheets/sheet1.xml', worksheet_xml)


def generate_demo_excels(conn):
    # 1. Reporte_BUK_Novedades_Agosto2026.xlsx (Formato idéntico a Dotacion_2026 con fila 1 vacía y fila 2 encabezados)
    headers_buk = [
        'IT', 'Instituto', 'PersonaId', 'Área', 'Nombre Sub-área', 'Centro de Costo',
        'Fecha de Nacimiento', 'Periodo', 'Tipo de Contrato', 'Tipo_Trab_161225',
        'Fecha Ingreso Compañía', 'Fecha Término', 'Horas Semanales', 'Cargo',
        'Tipo_Trabajador', 'Días de la Jornada', 'Reg_SIC', 'Cod_Esp', 'Ndias'
    ]
    rows_novedades = [
        # Alta nueva en Arica (Kinesiólogo 42h desde agosto 2026 = serial 46235)
        [11, 'Ari', 199001, 'Arica', 'Kinesiología', 300501, 34500, 46235, 'Indefinido', 'Titular', 46235, '', 42, 'KINESIOLOGO', 'Titular', 'lmwjv', 'D - SIC', 'KI', 5],
        # Alta nueva en Coyhaique (Terapeuta Ocupacional 40h desde agosto 2026)
        [111, 'Coy', 199002, 'Coyhaique', 'Terapia Ocupacional', 300503, 35100, 46235, 'Plazo fijo', 'Titular', 46235, '', 40, 'TERAPEUTA OCUPACIONAL', 'Titular', 'lmwjv', 'D - SIC', 'TO', 5],
        # Aumento de jornada: PersonaId 134744 en Arica (Profesor de Arte sube de 20h a 35h en agosto 2026)
        [11, 'Ari', 134744, 'Arica', 'UTAC Unidad de Terapias Artisticas Creativas', 301201, 32000, 46235, 'Indefinido', 'Titular', 44000, '', 35, 'PROFESOR DE ARTE', 'Titular', 'lmwjv', 'D - SIC', 'UTAC', 5],
        # Disminución de jornada: PersonaId 71306 en Arica (Enfermera Coordinadora baja de 28h a 20h en agosto 2026)
        [11, 'Ari', 71306, 'Arica', 'Enfermería', 300511, 29000, 46235, 'Indefinido', 'Titular', 41000, '', 20, 'ENFERMERA COORDINADORA', 'Titular', 'lmwj', 'D - SIC', 'EN', 4],
        # Baja / Término de contrato en agosto: PersonaId 121286 en Arica (termina 31-07-2026 = 46234, horas = 0)
        [11, 'Ari', 121286, 'Arica', 'Unidad de Ejercicio y Deporte Adaptado (UEDA)', 301301, 31500, 46235, 'Indefinido', 'Titular', 43000, 46234, 0, 'PROFESOR DE EDUCACION FISICA', 'Titular', 'lmwjv', 'D - SIC', 'UEDA', 5],
        # ERROR INTENCIONAL 1: Horas semanales = 60 (> 44 horas legales permitidas) -> Blue Prism debe rechazar e informar
        [131, 'San', 199991, 'Santiago', 'Fisiatría', 300202, 33000, 46235, 'Indefinido', 'Titular', 46235, '', 60, 'MEDICO FISIATRA', 'Titular', 'lmwjv', 'D - SIC', 'FI', 5],
        # ERROR INTENCIONAL 2: Cargo no existe en catálogo Distribucion_NAC ('ASTRONAUTA REHABILITADOR') -> Blue Prism debe rechazar
        [81, 'Con', 199992, 'Concepción', 'Kinesiología', 300501, 33000, 46235, 'Indefinido', 'Titular', 46235, '', 40, 'ASTRONAUTA REHABILITADOR', 'Titular', 'lmwjv', 'D - SIC', 'KI', 5],
    ]
    novedades_path = os.path.join(DEMO_DIR, 'Reporte_BUK_Novedades_Agosto2026.xlsx')
    create_minimal_xlsx(novedades_path, 'Dotacion_2026', headers_buk, rows_novedades)

    # 2. Reporte_SIC_Ejecucion_2026.xlsx (Ejecución real SIC para contraste en Blue Prism)
    cur = conn.cursor()
    cur.execute("""
        SELECT persona_id, it, instituto, area, mes,
               ate_real_directa, ate_real_indirecta, ate_real_total,
               pct_cumplimiento_meta, estado_alerta
        FROM ejecucion_sic
        WHERE mes <= 3
        ORDER BY it, persona_id, mes
        OFFSET 0 ROWS FETCH NEXT 300 ROWS ONLY
    """)
    sic_rows = [list(r) for r in cur.fetchall()]
    headers_sic = [
        'PersonaId', 'IT', 'Instituto', 'Area', 'Mes',
        'Ate_Real_Directa', 'Ate_Real_Indirecta', 'Ate_Real_Total',
        'Pct_Cumplimiento_Meta', 'Estado_Alerta'
    ]
    sic_path = os.path.join(DEMO_DIR, 'Reporte_SIC_Ejecucion_2026.xlsx')
    create_minimal_xlsx(sic_path, 'Ejecucion_SIC', headers_sic, sic_rows)
    return novedades_path, sic_path


def main():
    conn = connect_database()
    print(f"[1/3] Creando esquema relacional en MSSQLLocalDB ({DATABASE_NAME})...")
    create_schema(conn)
    print("[2/3] Cargando catálogos, contratos BUK (3.123 filas) y planificación asistencial (471 filas)...")
    seed_all_data(conn)
    print("[3/3] Generando archivos Excel de prueba para la demo de Blue Prism...")
    nov_path, sic_path = generate_demo_excels(conn)

    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM contratos_buk")
    n_cont = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM dotacion_asistencial")
    n_dot = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM planificacion_mensual")
    n_pm = cur.fetchone()[0]
    cur.execute("SELECT SUM(hs_dic), SUM(plan_anual_total), SUM(mg_anual_total) FROM dotacion_asistencial")
    sums = cur.fetchone()
    print(f"OK -> contratos_buk: {n_cont} | dotacion_asistencial: {n_dot} | planificacion_mensual: {n_pm}")
    print(f"OK -> Totales Nacionales: H/S Dic={sums[0]:.0f} | Plan Anual={sums[1]:.0f} | Meta Gestión 90%={sums[2]:.0f}")
    print(f"OK -> Excel Novedades BUK: {nov_path}")
    print(f"OK -> Excel Ejecución SIC: {sic_path}")
    conn.close()


if __name__ == '__main__':
    main()
