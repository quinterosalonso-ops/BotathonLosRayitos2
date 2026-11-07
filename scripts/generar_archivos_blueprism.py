#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de Objetos VBO, Procesos XML y Paquete .bprelease para Blue Prism 7.5:
  - 01_BPA_Object_MS_Excel_VBO_Extended.xml (MS Excel VBO - Extended v2.3.3 con lectura directa rng.Value)
  - 02_BPA_Object_Data_SQLite.xml (Data - SQLite v1.0.2 + acción Write Collection transaccional)
  - 03_BPA_Process_01_Carga_Base_Contratos_BUK.xml
  - 04_BPA_Process_02_Carga_y_Calculo_Planificacion.xml
  - 05_BPA_Process_03_Novedades_BUK_y_Contraste_SIC.xml
  - Botathon_Teleton_2026_Completo.bprelease
"""

import os
import sys
import shutil
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BP_DIR = os.path.join(BASE_DIR, 'blueprism')
DLL_DIR = os.path.join(BP_DIR, 'dlls')
os.makedirs(BP_DIR, exist_ok=True)
os.makedirs(DLL_DIR, exist_ok=True)

DESKTOP_EXCEL_VBO = r'C:\Users\AlonsoQ\Desktop\BPA Object - MS Excel VBO - Extended v2.3.3.xml'
DESKTOP_SQLITE_BPRELEASE = r'C:\Users\AlonsoQ\Desktop\Data - SQLite V1.0.2.bprelease'
MINITOOL_DIR = r'C:\Program Files\MiniTool Partition Wizard 13'

RUTA_DB = os.path.join(BASE_DIR, 'database', 'teleton_planificacion.db')
RUTA_EXCEL_BUK = os.path.join(BASE_DIR, 'Solicitud', 'Nac - Base contratos x mes.xlsx')
RUTA_EXCEL_PLAN = os.path.join(BASE_DIR, 'Solicitud', 'Nac - Planificación Asistencial x mes.xlsx')
RUTA_EXCEL_NOVEDADES = os.path.join(BASE_DIR, 'datos_demo', 'Reporte_BUK_Novedades_Agosto2026.xlsx')
RUTA_EXCEL_SIC = os.path.join(BASE_DIR, 'datos_demo', 'Reporte_SIC_Ejecucion_2026.xlsx')
RUTA_SALIDA_REPORTE = os.path.join(BASE_DIR, 'salidas', 'Reporte_Consolidado_Actualizado_2026.xlsx')


def copy_sqlite_dlls():
    for dll_name in ['System.Data.SQLite.dll', 'SQLite.Interop.dll']:
        src = os.path.join(MINITOOL_DIR, dll_name)
        dst = os.path.join(DLL_DIR, dll_name)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"OK -> Copiado {dll_name} ({os.path.getsize(dst)} bytes) a {DLL_DIR}")


def xml_escape_text(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def build_excel_vbo_extended():
    with open(DESKTOP_EXCEL_VBO, 'r', encoding='utf-8') as f:
        xml_text = f.read()

    # Reemplazar el código de GetWorksheetRangeAsCollection (stageid 680eac96-2bd0-4d02-a62b-ac7bf684dae9)
    import re
    pattern = r'(<stage stageid="680eac96-2bd0-4d02-a62b-ac7bf684dae9"[^>]*>.*?<code>)(.*?)(</code>)'

    new_code = """
\tDim ws As Object = GetWorksheet(handle, workbookname, worksheetname)
\tDim rng As Object = ws.Range(Start_Cell &amp; ":" &amp; End_Cell)
\tDim values As Object = rng.Value
\tDim totalRows As Integer = rng.Rows.Count
\tDim totalCols As Integer = rng.Columns.Count

\tDim table As New DataTable()

\tDim c As Integer
\tFor c = 1 To totalCols
\t\tDim fieldName As String = ""
\t\tIf values(1, c) IsNot Nothing Then
\t\t\tfieldName = values(1, c).ToString().Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ").Trim()
\t\tEnd If
\t\tIf String.IsNullOrEmpty(fieldName) Then fieldName = "Col" &amp; c
\t\tDim uniqueName As String = fieldName
\t\tDim n As Integer = 1
\t\tDo While table.Columns.Contains(uniqueName)
\t\t\tn += 1
\t\t\tuniqueName = fieldName &amp; "_" &amp; n
\t\tLoop
\t\ttable.Columns.Add(uniqueName)
\tNext

\tDim r As Integer
\tFor r = 2 To totalRows
\t\tDim row As DataRow = table.NewRow()
\t\tDim hasValue As Boolean = False
\t\tFor c = 1 To totalCols
\t\t\tDim v As Object = values(r, c)
\t\t\tDim s As String = ""
\t\t\tIf v IsNot Nothing Then
\t\t\t\tIf TypeOf v Is Date Then
\t\t\t\t\ts = CDate(v).ToString("yyyy-MM-dd")
\t\t\t\tElseIf TypeOf v Is Double Then
\t\t\t\t\ts = CDbl(v).ToString(System.Globalization.CultureInfo.InvariantCulture)
\t\t\t\tElse
\t\t\t\t\ts = v.ToString().Trim()
\t\t\t\tEnd If
\t\t\t\tIf s.Length &gt; 0 Then hasValue = True
\t\t\tEnd If
\t\t\trow(c - 1) = s
\t\tNext
\t\tIf hasValue Then table.Rows.Add(row)
\tNext

\tworksheetCollection = table
"""

    xml_text_new, count = re.subn(pattern, lambda m: m.group(1) + new_code + m.group(3), xml_text, count=1, flags=re.DOTALL)
    if count == 1:
        xml_text = xml_text_new
        print("OK -> Parcheado GetWorksheetRangeAsCollection en MS Excel VBO - Extended con rng.Value directo.")
    else:
        print("WARN -> No se encontró GetWorksheetRangeAsCollection en MS Excel VBO - Extended.")

    out_path = os.path.join(BP_DIR, '01_BPA_Object_MS_Excel_VBO_Extended.xml')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(xml_text)
    return out_path, xml_text


def build_sqlite_vbo():
    with open(DESKTOP_SQLITE_BPRELEASE, 'r', encoding='utf-8-sig') as f:
        bpr_text = f.read()

    # Extraer el bloque <object id="d91164ca-ea87-4f73-bad2-ee9010d382d9" name="Data - SQLite"...>...</object>
    start_tag = '<object id="d91164ca-ea87-4f73-bad2-ee9010d382d9" name="Data - SQLite" xmlns="http://www.blueprism.co.uk/product/process">'
    end_tag = '</object>'
    s_idx = bpr_text.find(start_tag)
    e_idx = bpr_text.find(end_tag, s_idx)
    inner_proc = bpr_text[s_idx + len(start_tag):e_idx].strip()

    # Agregar preferredid="d91164ca-ea87-4f73-bad2-ee9010d382d9" al elemento raíz <process>
    inner_proc = inner_proc.replace(
        '<process name="Data - SQLite" version="1.0"',
        '<process name="Data - SQLite" version="1.0" preferredid="d91164ca-ea87-4f73-bad2-ee9010d382d9"',
        1
    )

    # Añadir subsheet publicada "Write Collection" para inserción masiva transaccional desde Blue Prism
    write_coll_subsheet_def = """        <subsheet subsheetid="a9b8c7d6-e5f4-4321-8765-1029384756ab" type="Normal" published="True">
          <name>Write Collection</name>
          <view>
            <camerax>0</camerax>
            <cameray>0</cameray>
            <zoom version="2">1.25</zoom>
          </view>
        </subsheet>"""

    write_coll_stages = """        <stage stageid="b1111111-2222-4333-8444-555555555501" name="Write Collection" type="SubSheetInfo">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <narrative>Inserta masivamente una Collection de Blue Prism dentro de una tabla SQLite usando una transacción parametrizada. Si la tabla no existe o Clear Table First es True, la recrea automáticamente.</narrative>
          <display x="-195" y="-105" w="180" h="90" />
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555502" name="Start" type="Start">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-15" y="-105" />
          <inputs>
            <input type="text" name="Table Name" narrative="Nombre de la tabla destino en SQLite." stage="Table Name" />
            <input type="collection" name="Collection" narrative="Colección de datos a insertar." stage="Collection" />
            <input type="flag" name="Clear Table First" narrative="True para recrear/limpiar la tabla antes de insertar." stage="Clear Table First" />
          </inputs>
          <onsuccess>b1111111-2222-4333-8444-555555555503</onsuccess>
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555503" name="Open" type="SubSheet">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-15" y="-60" />
          <outputs>
            <output type="flag" name="Success" friendlyname="Success" stage="Success" />
            <output type="text" name="Error Message" friendlyname="Error Message" stage="Error Message" />
          </outputs>
          <onsuccess>b1111111-2222-4333-8444-555555555504</onsuccess>
          <processid>19235e77-62a9-4ab0-9662-b41d3b95dbc3</processid>
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555504" name="Success?" type="Decision">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-15" y="-15" />
          <decision expression="[Success]" />
          <ontrue>b1111111-2222-4333-8444-555555555505</ontrue>
          <onfalse>b1111111-2222-4333-8444-555555555506</onfalse>
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555505" name="Bulk Insert Collection" type="Code">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-15" y="30" w="150" h="30" />
          <inputs>
            <input type="text" name="Table_Name" expr="[Table Name]" />
            <input type="collection" name="Data_Collection" expr="[Collection]" />
            <input type="flag" name="Clear_First" expr="[Clear Table First]" />
          </inputs>
          <outputs>
            <output type="flag" name="Success" stage="Success" />
            <output type="text" name="Error Message" stage="Error Message" />
            <output type="number" name="Records Affected" stage="Records Affected" />
          </outputs>
          <onsuccess>b1111111-2222-4333-8444-555555555506</onsuccess>
          <code><![CDATA[Success = true;
Error_Message = "";
Records_Affected = 0;

try
{
	if (Data_Collection == null || Data_Collection.Columns.Count == 0)
	{
		return;
	}

	string safeTable = "\"" + Table_Name.Replace("\"", "") + "\"";
	System.Collections.Generic.List<string> colDefs = new System.Collections.Generic.List<string>();
	System.Collections.Generic.List<string> colNames = new System.Collections.Generic.List<string>();
	System.Collections.Generic.List<string> paramNames = new System.Collections.Generic.List<string>();

	for (int i = 0; i < Data_Collection.Columns.Count; i++)
	{
		string cName = Data_Collection.Columns[i].ColumnName.Replace("\"", "").Trim();
		if (string.IsNullOrEmpty(cName)) cName = "Col" + (i + 1).ToString();
		colDefs.Add("\"" + cName + "\" TEXT");
		colNames.Add("\"" + cName + "\"");
		paramNames.Add("@p" + i.ToString());
	}

	using (SQLiteTransaction tran = dbConn.BeginTransaction())
	{
		using (SQLiteCommand cmd = new SQLiteCommand(dbConn))
		{
			cmd.Transaction = tran;
			if (Clear_First)
			{
				cmd.CommandText = "DROP TABLE IF EXISTS " + safeTable + ";";
				cmd.ExecuteNonQuery();
			}
			cmd.CommandText = "CREATE TABLE IF NOT EXISTS " + safeTable + " (" + string.Join(", ", colDefs.ToArray()) + ");";
			cmd.ExecuteNonQuery();

			cmd.CommandText = "INSERT INTO " + safeTable + " (" + string.Join(", ", colNames.ToArray()) + ") VALUES (" + string.Join(", ", paramNames.ToArray()) + ");";
			for (int i = 0; i < Data_Collection.Columns.Count; i++)
			{
				cmd.Parameters.Add(new SQLiteParameter("@p" + i.ToString(), DbType.String));
			}

			foreach (DataRow row in Data_Collection.Rows)
			{
				for (int i = 0; i < Data_Collection.Columns.Count; i++)
				{
					object val = row[i];
					cmd.Parameters[i].Value = (val == null || val == DBNull.Value) ? "" : val.ToString();
				}
				Records_Affected += cmd.ExecuteNonQuery();
			}
		}
		tran.Commit();
	}
}
catch (Exception ex)
{
	Success = false;
	Error_Message = ex.Message;
}
]]></code>
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555506" name="End" type="End">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-15" y="90" />
          <outputs>
            <output type="flag" name="Success" narrative="Flag indicating success (True) or failure (False)" stage="Success" />
            <output type="text" name="Error Message" narrative="Error message giving detail regarding why this action failed" stage="Error Message" />
            <output type="number" name="Records Affected" narrative="The count of records inserted." stage="Records Affected" />
          </outputs>
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555507" name="Table Name" type="Data">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-195" y="-15" w="120" h="30" />
          <datatype>text</datatype>
          <initialvalue />
          <private />
          <alwaysinit />
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555508" name="Collection" type="Collection">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-195" y="15" w="120" h="30" />
          <datatype>collection</datatype>
          <private />
          <alwaysinit />
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555509" name="Clear Table First" type="Data">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-195" y="45" w="120" h="30" />
          <datatype>flag</datatype>
          <initialvalue>True</initialvalue>
          <private />
          <alwaysinit />
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555510" name="Success" type="Data">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-195" y="105" w="120" h="30" />
          <datatype>flag</datatype>
          <initialvalue>False</initialvalue>
          <private />
          <alwaysinit />
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555511" name="Error Message" type="Data">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-195" y="135" w="120" h="30" />
          <datatype>text</datatype>
          <initialvalue />
          <private />
          <alwaysinit />
        </stage>
        <stage stageid="b1111111-2222-4333-8444-555555555512" name="Records Affected" type="Data">
          <subsheetid>a9b8c7d6-e5f4-4321-8765-1029384756ab</subsheetid>
          <loginhibit />
          <display x="-195" y="165" w="120" h="30" />
          <datatype>number</datatype>
          <initialvalue>0</initialvalue>
          <private />
          <alwaysinit />
        </stage>"""

    # Insertar write_coll_subsheet_def justo después de la última etiqueta </subsheet>
    last_ss = inner_proc.rfind('</subsheet>')
    inner_proc = (
        inner_proc[:last_ss + len('</subsheet>')] + '\n' +
        write_coll_subsheet_def +
        inner_proc[last_ss + len('</subsheet>'):]
    )
    # Insertar write_coll_stages justo antes de </process>
    last_proc = inner_proc.rfind('</process>')
    inner_proc = (
        inner_proc[:last_proc] +
        write_coll_stages + '\n      ' +
        inner_proc[last_proc:]
    )

    xml_out = '<?xml version="1.0" encoding="utf-8"?>\n' + inner_proc
    out_path = os.path.join(BP_DIR, '02_BPA_Object_Data_SQLite.xml')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(xml_out)
    print("OK -> Generado 02_BPA_Object_Data_SQLite.xml con todas las acciones originales + Write Collection.")
    return out_path, inner_proc


def build_process_01_carga_buk():
    """
    Proceso 1: 01_Carga_Base_Contratos_BUK
    Usa MS Excel VBO - Extended y Data - SQLite para leer Nac - Base contratos x mes.xlsx,
    cargarlo a SQLite, validar reglas de negocio, registrar auditoría y devolver KPIs.
    """
    sql_validar_insertar = """
DELETE FROM errores_validacion WHERE origen_archivo = 'Nac - Base contratos x mes.xlsx';

-- 1. Detectar errores de horas fuera de rango (<= 0 o > 44)
INSERT INTO errores_validacion (origen_archivo, persona_id, it, cargo, campo_error, valor_recibido, descripcion_error)
SELECT 'Nac - Base contratos x mes.xlsx', PersonaId, IT, Cargo, 'Horas Semanales', "Horas Semanales",
       'Horas semanales fuera del rango legal permitido (debe estar entre 1 y 44 horas).'
FROM stg_contratos_buk
WHERE CAST(COALESCE("Horas Semanales", '0') AS REAL) <= 0
   OR CAST(COALESCE("Horas Semanales", '0') AS REAL) > 44;

-- 2. Detectar cargos que no existen en el catálogo Distribucion_NAC
INSERT INTO errores_validacion (origen_archivo, persona_id, it, cargo, campo_error, valor_recibido, descripcion_error)
SELECT 'Nac - Base contratos x mes.xlsx', s.PersonaId, s.IT, s.Cargo, 'Cargo', s.Cargo,
       'Cargo no homologado en matriz parametrizable Distribucion_NAC.'
FROM stg_contratos_buk s
LEFT JOIN distribucion_nac d ON TRIM(UPPER(s.Cargo)) = TRIM(UPPER(d.cargo))
WHERE d.cargo IS NULL AND TRIM(COALESCE(s.Cargo, '')) <> '';

-- 3. Insertar/Actualizar registros válidos en contratos_buk
INSERT OR REPLACE INTO contratos_buk (
    it, instituto, persona_id, area_buk, nombre_subarea, centro_costo,
    fecha_nacimiento, periodo_serial, periodo_fecha, mes, tipo_contrato,
    tipo_trab_anterior, fecha_ingreso, fecha_termino, horas_semanales,
    cargo, tipo_trabajador, dias_jornada, reg_sic, cod_esp, ndias, origen_carga
)
SELECT
    CAST(s.IT AS INTEGER),
    TRIM(s.Instituto),
    CAST(s.PersonaId AS INTEGER),
    TRIM(s."Área"),
    TRIM(s."Nombre Sub-área"),
    CAST(COALESCE(s."Centro de Costo", '0') AS INTEGER),
    s."Fecha de Nacimiento",
    CAST( COALESCE(s.Periodo, '0') AS INTEGER),
    s.Periodo,
    CASE
        WHEN s.Periodo LIKE '2026-%' THEN CAST(SUBSTR(s.Periodo, 6, 2) AS INTEGER)
        WHEN CAST(s.Periodo AS INTEGER) = 46023 THEN 1
        WHEN CAST(s.Periodo AS INTEGER) = 46054 THEN 2
        WHEN CAST(s.Periodo AS INTEGER) = 46082 THEN 3
        WHEN CAST(s.Periodo AS INTEGER) = 46113 THEN 4
        WHEN CAST(s.Periodo AS INTEGER) = 46143 THEN 5
        WHEN CAST(s.Periodo AS INTEGER) = 46174 THEN 6
        WHEN CAST(s.Periodo AS INTEGER) = 46204 THEN 7
        WHEN CAST(s.Periodo AS INTEGER) = 46235 THEN 8
        ELSE 1
    END,
    TRIM(s."Tipo de Contrato"),
    TRIM(s.Tipo_Trab_161225),
    s."Fecha Ingreso Compañía",
    s."Fecha Término",
    CAST(s."Horas Semanales" AS REAL),
    TRIM(s.Cargo),
    TRIM(s.Tipo_Trabajador),
    TRIM(s."Días de la Jornada"),
    TRIM(s.Reg_SIC),
    TRIM(s.Cod_Esp),
    CAST(COALESCE(s.Ndias, '5') AS REAL),
    'BLUEPRISM_BUK'
FROM stg_contratos_buk s
INNER JOIN distribucion_nac d ON TRIM(UPPER(s.Cargo)) = TRIM(UPPER(d.cargo))
WHERE CAST(COALESCE(s."Horas Semanales", '0') AS REAL) > 0
  AND CAST(COALESCE(s."Horas Semanales", '0') AS REAL) <= 44;

-- 4. Registrar evento en bitácora de auditoría
INSERT INTO auditoria_cambios (
    version_id, actor, tipo_evento, entidad_afectada, campo_modificado,
    valor_anterior, valor_nuevo, mes_inicio_vigencia, impacto_atenciones_anual,
    justificacion, estado_aprobacion
) VALUES (
    1, 'Bot Blue Prism (01_Carga_Base_Contratos_BUK)', 'INGESTA_BUK', 'contratos_buk',
    'Carga Masiva Mensual BUK', 'Staging Excel',
    (SELECT CAST(COUNT(*) AS TEXT) || ' contratos validados' FROM contratos_buk),
    1, 0, 'Ingesta automatizada vía MS Excel VBO - Extended y Data - SQLite', 'APROBADO'
);
""".strip()

    xml = f"""<?xml version="1.0" encoding="utf-8"?>
<process name="01_Carga_Base_Contratos_BUK" version="1.0" bpversion="7.5.1.18099" cpeversion="10.0.136.0" narrative="Lee la hoja Dotacion_2026 de Nac - Base contratos x mes.xlsx usando MS Excel VBO - Extended, valida reglas de negocio e inserta en SQLite usando Data - SQLite." byrefcollection="true" processrunningmessage="" disableversioning="false" preferredid="11111111-aaaa-4bbb-8ccc-000000000001">
  <view>
    <camerax>0</camerax>
    <cameray>60</cameray>
    <zoom version="2">1.25</zoom>
  </view>
  <preconditions />
  <endpoint narrative="" />
  <stage stageid="p1000001-0000-4000-8000-000000000001" name="Start" type="Start">
    <display x="15" y="-240" />
    <onsuccess>p1000001-0000-4000-8000-000000000002</onsuccess>
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000002" name="Excel::Create Instance" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-180" w="150" h="30" />
    <inputs>
      <input type="flag" name="Enable Events" friendlyname="Enable Events" expr="False" />
    </inputs>
    <outputs>
      <output type="number" name="handle" friendlyname="handle" stage="vHandle" />
    </outputs>
    <onsuccess>p1000001-0000-4000-8000-000000000003</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Create Instance" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000003" name="Excel::Open Workbook BUK" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-120" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="File name" friendlyname="File name" expr="[vRutaExcelBUK]" />
    </inputs>
    <outputs>
      <output type="text" name="Workbook Name" friendlyname="Workbook Name" stage="vWorkbook" />
    </outputs>
    <onsuccess>p1000001-0000-4000-8000-000000000004</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Open Workbook" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000004" name="Excel::Leer Dotacion_2026 (A2:S3125)" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-60" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="text" name="Worksheet Name" friendlyname="Worksheet Name" expr="[vHojaBUK]" />
      <input type="text" name="StartCell" friendlyname="StartCell" expr="[vCeldaIni]" />
      <input type="text" name="EndCell" friendlyname="EndCell" expr="[vCeldaFin]" />
    </inputs>
    <outputs>
      <output type="collection" name="Data" friendlyname="Data" stage="cContratosBUK" />
    </outputs>
    <onsuccess>p1000001-0000-4000-8000-000000000005</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Get Worksheet Range As Collection" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000005" name="Excel::Close Workbook" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="0" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="flag" name="Save Data" friendlyname="Save Data" expr="False" />
    </inputs>
    <onsuccess>p1000001-0000-4000-8000-000000000006</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Close Workbook" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000006" name="Excel::Release Instance" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="60" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
    </inputs>
    <onsuccess>p1000001-0000-4000-8000-000000000007</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Release Instance" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000007" name="SQLite::Set Connection Details" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="120" w="150" h="30" />
    <inputs>
      <input type="text" name="DB File" friendlyname="DB File" expr="[vRutaBD]" />
      <input type="flag" name="Create If Not Found" friendlyname="Create If Not Found" expr="True" />
      <input type="text" name="Connection String" friendlyname="Connection String" expr="&quot;&quot;" />
    </inputs>
    <onsuccess>p1000001-0000-4000-8000-000000000008</onsuccess>
    <resource object="Data - SQLite" action="Set Connection Details" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000008" name="SQLite::Write Collection (stg_contratos_buk)" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="180" w="150" h="30" />
    <inputs>
      <input type="text" name="Table Name" friendlyname="Table Name" expr="&quot;stg_contratos_buk&quot;" />
      <input type="collection" name="Collection" friendlyname="Collection" expr="[cContratosBUK]" />
      <input type="flag" name="Clear Table First" friendlyname="Clear Table First" expr="True" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vFilasStaging" />
    </outputs>
    <onsuccess>p1000001-0000-4000-8000-000000000009</onsuccess>
    <resource object="Data - SQLite" action="Write Collection" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000009" name="SQLite::Validar e Insertar Contratos BUK" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="240" w="150" h="30" />
    <inputs>
      <input type="text" name="SQL" friendlyname="SQL" expr="[vSQL_Validar_Insertar]" />
      <input type="number" name="CommandTimeout" friendlyname="CommandTimeout" expr="60" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vRegistrosAfectados" />
    </outputs>
    <onsuccess>p1000001-0000-4000-8000-000000000010</onsuccess>
    <resource object="Data - SQLite" action="Execute Non Query" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000010" name="SQLite::Obtener Resumen por Instituto" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="300" w="150" h="30" />
    <inputs>
      <input type="text" name="SQL" friendlyname="SQL" expr="&quot;SELECT it, instituto, COUNT(DISTINCT persona_id) AS Profesionales, COUNT(*) AS Registros_Mes, SUM(horas_semanales) AS Total_Horas_Mes FROM contratos_buk GROUP BY it, instituto ORDER BY it;&quot;" />
      <input type="number" name="CommandTimeout" friendlyname="CommandTimeout" expr="30" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="collection" name="Results" friendlyname="Results" stage="cResumenContratosIT" />
    </outputs>
    <onsuccess>p1000001-0000-4000-8000-000000000011</onsuccess>
    <resource object="Data - SQLite" action="Execute - Get Collection" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000011" name="SQLite::Close" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="360" w="150" h="30" />
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
    </outputs>
    <onsuccess>p1000001-0000-4000-8000-000000000012</onsuccess>
    <resource object="Data - SQLite" action="Close" />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000012" name="End" type="End">
    <display x="15" y="420" />
  </stage>
  <!-- Data Items -->
  <stage stageid="p1000001-0000-4000-8000-000000000101" name="vRutaExcelBUK" type="Data">
    <display x="-240" y="-195" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{RUTA_EXCEL_BUK}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000102" name="vRutaBD" type="Data">
    <display x="-240" y="-150" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{RUTA_DB}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000103" name="vHojaBUK" type="Data">
    <display x="-240" y="-105" w="150" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">Dotacion_2026</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000104" name="vCeldaIni" type="Data">
    <display x="-240" y="-60" w="120" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">A2</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000105" name="vCeldaFin" type="Data">
    <display x="-240" y="-15" w="120" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">S3125</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000106" name="vHandle" type="Data">
    <display x="-240" y="30" w="120" h="30" />
    <datatype>number</datatype>
    <initialvalue>0</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000107" name="vWorkbook" type="Data">
    <display x="-240" y="75" w="150" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000108" name="vSuccess" type="Data">
    <display x="-240" y="120" w="120" h="30" />
    <datatype>flag</datatype>
    <initialvalue>False</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000109" name="vErrorMessage" type="Data">
    <display x="-240" y="165" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000110" name="vFilasStaging" type="Data">
    <display x="-240" y="210" w="150" h="30" />
    <datatype>number</datatype>
    <initialvalue>0</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000111" name="vRegistrosAfectados" type="Data">
    <display x="-240" y="255" w="150" h="30" />
    <datatype>number</datatype>
    <initialvalue>0</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000112" name="vSQL_Validar_Insertar" type="Data">
    <display x="-240" y="300" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{xml_escape_text(sql_validar_insertar)}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000113" name="cContratosBUK" type="Collection">
    <display x="240" y="-60" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p1000001-0000-4000-8000-000000000114" name="cResumenContratosIT" type="Collection">
    <display x="240" y="300" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
</process>"""
    out_path = os.path.join(BP_DIR, '03_BPA_Process_01_Carga_Base_Contratos_BUK.xml')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(xml)
    print("OK -> Generado 03_BPA_Process_01_Carga_Base_Contratos_BUK.xml")
    return out_path, xml


def build_process_02_carga_planificacion():
    """
    Proceso 2: 02_Carga_y_Calculo_Planificacion
    Lee las hojas Distribucion_NAC, Excepciones, LOP y Dotacion_Asistencial de Nac - Planificación Asistencial x mes.xlsx
    con MS Excel VBO - Extended, las guarda en SQLite con Data - SQLite, recalcula indicadores y obtiene el Resumen Nacional.
    """
    sql_recalcular = """
-- Recalcular % FOJA y Metas de Gestión según parámetros globales vigentes en SQLite
UPDATE dotacion_asistencial
SET
    foja_pct = CASE WHEN hs_dic > 0 THEN ROUND(pct_directo_final_p2 * 100.0, 2) ELSE 0.0 END,
    updated_at = CURRENT_TIMESTAMP;

INSERT INTO auditoria_cambios (
    version_id, actor, tipo_evento, entidad_afectada, campo_modificado,
    valor_anterior, valor_nuevo, mes_inicio_vigencia, impacto_atenciones_anual,
    justificacion, estado_aprobacion
) VALUES (
    1, 'Bot Blue Prism (02_Carga_y_Calculo_Planificacion)', 'SINCRONIZACION_PLANIFICACION',
    'Dotación Asistencial + Excepciones + LOP + Distribucion_NAC',
    'Recálculo Nacional 17 Institutos', 'Planilla Excel',
    (SELECT CAST(COUNT(*) AS TEXT) || ' filas sincronizadas en SQLite' FROM dotacion_asistencial),
    1, (SELECT CAST(SUM(mg_anual_total) AS INTEGER) FROM dotacion_asistencial),
    'Lectura de hojas vía MS Excel VBO - Extended y persistencia en Data - SQLite', 'APROBADO'
);
""".strip()

    xml = f"""<?xml version="1.0" encoding="utf-8"?>
<process name="02_Carga_y_Calculo_Planificacion" version="1.0" bpversion="7.5.1.18099" cpeversion="10.0.136.0" narrative="Lee las hojas Distribucion_NAC, Excepciones, LOP y Dotacion_Asistencial con MS Excel VBO - Extended, las almacena en SQLite con Data - SQLite y obtiene la matriz consolidada nacional." byrefcollection="true" processrunningmessage="" disableversioning="false" preferredid="22222222-aaaa-4bbb-8ccc-000000000002">
  <view>
    <camerax>0</camerax>
    <cameray>120</cameray>
    <zoom version="2">1.25</zoom>
  </view>
  <preconditions />
  <endpoint narrative="" />
  <stage stageid="p2000001-0000-4000-8000-000000000001" name="Start" type="Start">
    <display x="15" y="-240" />
    <onsuccess>p2000001-0000-4000-8000-000000000002</onsuccess>
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000002" name="Excel::Create Instance" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-180" w="150" h="30" />
    <inputs>
      <input type="flag" name="Enable Events" friendlyname="Enable Events" expr="False" />
    </inputs>
    <outputs>
      <output type="number" name="handle" friendlyname="handle" stage="vHandle" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000003</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Create Instance" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000003" name="Excel::Open Workbook Planificación" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-120" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="File name" friendlyname="File name" expr="[vRutaExcelPlan]" />
    </inputs>
    <outputs>
      <output type="text" name="Workbook Name" friendlyname="Workbook Name" stage="vWorkbook" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000004</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Open Workbook" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000004" name="Excel::Leer Distribucion_NAC (A6:N92)" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-60" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="text" name="Worksheet Name" friendlyname="Worksheet Name" expr="&quot;Distribucion_NAC&quot;" />
      <input type="text" name="StartCell" friendlyname="StartCell" expr="&quot;A6&quot;" />
      <input type="text" name="EndCell" friendlyname="EndCell" expr="&quot;N92&quot;" />
    </inputs>
    <outputs>
      <output type="collection" name="Data" friendlyname="Data" stage="cDistribucionNAC" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000005</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Get Worksheet Range As Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000005" name="Excel::Leer Excepciones (A2:X19)" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="0" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="text" name="Worksheet Name" friendlyname="Worksheet Name" expr="&quot;Excepciones&quot;" />
      <input type="text" name="StartCell" friendlyname="StartCell" expr="&quot;A2&quot;" />
      <input type="text" name="EndCell" friendlyname="EndCell" expr="&quot;X19&quot;" />
    </inputs>
    <outputs>
      <output type="collection" name="Data" friendlyname="Data" stage="cExcepciones" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000006</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Get Worksheet Range As Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000006" name="Excel::Leer LOP (A1:C17)" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="60" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="text" name="Worksheet Name" friendlyname="Worksheet Name" expr="&quot;LOP&quot;" />
      <input type="text" name="StartCell" friendlyname="StartCell" expr="&quot;A1&quot;" />
      <input type="text" name="EndCell" friendlyname="EndCell" expr="&quot;C17&quot;" />
    </inputs>
    <outputs>
      <output type="collection" name="Data" friendlyname="Data" stage="cLOP" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000007</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Get Worksheet Range As Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000007" name="Excel::Leer Dotacion_Asistencial (A4:GW475)" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="120" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="text" name="Worksheet Name" friendlyname="Worksheet Name" expr="&quot;Dotacion_Asistencial&quot;" />
      <input type="text" name="StartCell" friendlyname="StartCell" expr="&quot;A4&quot;" />
      <input type="text" name="EndCell" friendlyname="EndCell" expr="&quot;GW475&quot;" />
    </inputs>
    <outputs>
      <output type="collection" name="Data" friendlyname="Data" stage="cDotacionAsistencial" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000008</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Get Worksheet Range As Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000008" name="Excel::Close Workbook" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="180" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="flag" name="Save Data" friendlyname="Save Data" expr="False" />
    </inputs>
    <onsuccess>p2000001-0000-4000-8000-000000000009</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Close Workbook" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000009" name="Excel::Release Instance" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="240" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
    </inputs>
    <onsuccess>p2000001-0000-4000-8000-000000000010</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Release Instance" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000010" name="SQLite::Set Connection Details" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="300" w="150" h="30" />
    <inputs>
      <input type="text" name="DB File" friendlyname="DB File" expr="[vRutaBD]" />
      <input type="flag" name="Create If Not Found" friendlyname="Create If Not Found" expr="True" />
      <input type="text" name="Connection String" friendlyname="Connection String" expr="&quot;&quot;" />
    </inputs>
    <onsuccess>p2000001-0000-4000-8000-000000000011</onsuccess>
    <resource object="Data - SQLite" action="Set Connection Details" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000011" name="SQLite::Write stg_distribucion_nac" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="360" w="150" h="30" />
    <inputs>
      <input type="text" name="Table Name" friendlyname="Table Name" expr="&quot;stg_distribucion_nac&quot;" />
      <input type="collection" name="Collection" friendlyname="Collection" expr="[cDistribucionNAC]" />
      <input type="flag" name="Clear Table First" friendlyname="Clear Table First" expr="True" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vFilasInsertadas" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000012</onsuccess>
    <resource object="Data - SQLite" action="Write Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000012" name="SQLite::Write stg_excepciones" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="420" w="150" h="30" />
    <inputs>
      <input type="text" name="Table Name" friendlyname="Table Name" expr="&quot;stg_excepciones&quot;" />
      <input type="collection" name="Collection" friendlyname="Collection" expr="[cExcepciones]" />
      <input type="flag" name="Clear Table First" friendlyname="Clear Table First" expr="True" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vFilasInsertadas" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000013</onsuccess>
    <resource object="Data - SQLite" action="Write Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000013" name="SQLite::Write stg_lop" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="480" w="150" h="30" />
    <inputs>
      <input type="text" name="Table Name" friendlyname="Table Name" expr="&quot;stg_lop&quot;" />
      <input type="collection" name="Collection" friendlyname="Collection" expr="[cLOP]" />
      <input type="flag" name="Clear Table First" friendlyname="Clear Table First" expr="True" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vFilasInsertadas" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000014</onsuccess>
    <resource object="Data - SQLite" action="Write Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000014" name="SQLite::Write stg_dotacion_asistencial" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="540" w="150" h="30" />
    <inputs>
      <input type="text" name="Table Name" friendlyname="Table Name" expr="&quot;stg_dotacion_asistencial&quot;" />
      <input type="collection" name="Collection" friendlyname="Collection" expr="[cDotacionAsistencial]" />
      <input type="flag" name="Clear Table First" friendlyname="Clear Table First" expr="True" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vFilasInsertadas" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000015</onsuccess>
    <resource object="Data - SQLite" action="Write Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000015" name="SQLite::Recalcular Indicadores y Auditar" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="600" w="150" h="30" />
    <inputs>
      <input type="text" name="SQL" friendlyname="SQL" expr="[vSQL_Recalcular]" />
      <input type="number" name="CommandTimeout" friendlyname="CommandTimeout" expr="60" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vFilasInsertadas" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000016</onsuccess>
    <resource object="Data - SQLite" action="Execute Non Query" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000016" name="SQLite::Obtener Consolidado Nacional 17 IT" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="660" w="150" h="30" />
    <inputs>
      <input type="text" name="SQL" friendlyname="SQL" expr="&quot;SELECT * FROM vw_resumen_nacional_instituto;&quot;" />
      <input type="number" name="CommandTimeout" friendlyname="CommandTimeout" expr="30" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="collection" name="Results" friendlyname="Results" stage="cResumenNacional" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000017</onsuccess>
    <resource object="Data - SQLite" action="Execute - Get Collection" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000017" name="SQLite::Close" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="720" w="150" h="30" />
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
    </outputs>
    <onsuccess>p2000001-0000-4000-8000-000000000018</onsuccess>
    <resource object="Data - SQLite" action="Close" />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000018" name="End" type="End">
    <display x="15" y="780" />
  </stage>
  <!-- Data Items -->
  <stage stageid="p2000001-0000-4000-8000-000000000101" name="vRutaExcelPlan" type="Data">
    <display x="-240" y="-195" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{RUTA_EXCEL_PLAN}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000102" name="vRutaBD" type="Data">
    <display x="-240" y="-150" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{RUTA_DB}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000103" name="vHandle" type="Data">
    <display x="-240" y="-105" w="120" h="30" />
    <datatype>number</datatype>
    <initialvalue>0</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000104" name="vWorkbook" type="Data">
    <display x="-240" y="-60" w="150" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000105" name="vSuccess" type="Data">
    <display x="-240" y="-15" w="120" h="30" />
    <datatype>flag</datatype>
    <initialvalue>False</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000106" name="vErrorMessage" type="Data">
    <display x="-240" y="30" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000107" name="vFilasInsertadas" type="Data">
    <display x="-240" y="75" w="150" h="30" />
    <datatype>number</datatype>
    <initialvalue>0</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000108" name="vSQL_Recalcular" type="Data">
    <display x="-240" y="120" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{xml_escape_text(sql_recalcular)}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000109" name="cDistribucionNAC" type="Collection">
    <display x="240" y="-60" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000110" name="cExcepciones" type="Collection">
    <display x="240" y="15" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000111" name="cLOP" type="Collection">
    <display x="240" y="90" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000112" name="cDotacionAsistencial" type="Collection">
    <display x="240" y="165" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p2000001-0000-4000-8000-000000000113" name="cResumenNacional" type="Collection">
    <display x="240" y="660" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
</process>"""
    out_path = os.path.join(BP_DIR, '04_BPA_Process_02_Carga_y_Calculo_Planificacion.xml')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(xml)
    print("OK -> Generado 04_BPA_Process_02_Carga_y_Calculo_Planificacion.xml")
    return out_path, xml


def build_process_03_novedades_demo():
    """
    Proceso 3: 03_Novedades_BUK_y_Contraste_SIC (Proceso Estrella para la Demo en Vivo de 5 Minutos)
    1. Lee Reporte_BUK_Novedades_Agosto2026.xlsx y Reporte_SIC_Ejecucion_2026.xlsx con MS Excel VBO - Extended.
    2. Valida errores (rechaza horas > 44 y cargos inexistentes) en SQLite (errores_validacion).
    3. Detecta Altas, Bajas y Cambios de Jornada, registrando trazabilidad en auditoria_cambios.
    4. Ajusta automáticamente la planificación desde agosto (mes 8 a 12) y recalcula Metas al 90% y % FOJA.
    5. Exporta el nuevo Reporte Consolidado Nacional y Auditoría a Excel usando MS Excel VBO - Extended.
    """
    sql_procesar_novedades = """
DELETE FROM errores_validacion WHERE origen_archivo = 'Reporte_BUK_Novedades_Agosto2026.xlsx';

-- 1. Detectar y registrar errores bloqueantes del archivo de novedades BUK
INSERT INTO errores_validacion (origen_archivo, persona_id, it, cargo, campo_error, valor_recibido, descripcion_error, severidad)
SELECT 'Reporte_BUK_Novedades_Agosto2026.xlsx', PersonaId, IT, Cargo, 'Horas Semanales', "Horas Semanales",
       'ERROR: Jornada de ' || "Horas Semanales" || ' H/S excede el máximo legal permitido (44 H/S). Registro rechazado.',
       'ERROR_BLOQUEANTE'
FROM stg_novedades_buk
WHERE CAST(COALESCE("Horas Semanales", '0') AS REAL) > 44
   OR CAST(COALESCE("Horas Semanales", '0') AS REAL) < 0;

INSERT INTO errores_validacion (origen_archivo, persona_id, it, cargo, campo_error, valor_recibido, descripcion_error, severidad)
SELECT 'Reporte_BUK_Novedades_Agosto2026.xlsx', s.PersonaId, s.IT, s.Cargo, 'Cargo', s.Cargo,
       'ERROR: El cargo "' || s.Cargo || '" no existe en la matriz parametrizable Distribucion_NAC. Registro rechazado.',
       'ERROR_BLOQUEANTE'
FROM stg_novedades_buk s
LEFT JOIN distribucion_nac d ON TRIM(UPPER(s.Cargo)) = TRIM(UPPER(d.cargo))
WHERE d.cargo IS NULL AND TRIM(COALESCE(s.Cargo, '')) <> '';

-- 2. Detectar ALTAS NUEVAS válidas y registrarlas en auditoria_cambios y dotacion_asistencial
INSERT INTO auditoria_cambios (
    version_id, actor, tipo_evento, it, instituto, persona_id, entidad_afectada,
    campo_modificado, valor_anterior, valor_nuevo, mes_inicio_vigencia,
    impacto_atenciones_anual, justificacion, estado_aprobacion
)
SELECT
    2, 'Bot Blue Prism (03_Novedades_BUK_y_Contraste_SIC)', 'ALTA_PROFESIONAL',
    CAST(s.IT AS INTEGER), TRIM(s.Instituto), CAST(s.PersonaId AS INTEGER),
    TRIM(s.Cargo) || ' (' || TRIM(s.Instituto) || ')',
    'Horas Semanales (Ago-Dic)', '0 H/S (Nuevo Ingreso)', s."Horas Semanales" || ' H/S',
    8, CAST(ROUND(CAST(s."Horas Semanales" AS REAL) * 0.65 * 60.0 / d.std_min_directa * 21.0 * 0.90, 0) AS INTEGER),
    'Alta automática detectada en reporte BUK Agosto 2026 (Vigente desde 01-08-2026)', 'APROBADO'
FROM stg_novedades_buk s
INNER JOIN distribucion_nac d ON TRIM(UPPER(s.Cargo)) = TRIM(UPPER(d.cargo))
LEFT JOIN dotacion_asistencial da ON CAST(s.PersonaId AS INTEGER) = da.persona_id AND CAST(s.IT AS INTEGER) = da.it
WHERE da.persona_id IS NULL
  AND CAST(s."Horas Semanales" AS REAL) > 0 AND CAST(s."Horas Semanales" AS REAL) <= 44;

-- Insertar las altas nuevas en dotacion_asistencial (vigentes ago..dic, 5 meses = 21 semanas calendario)
INSERT INTO dotacion_asistencial (
    orden, it, instituto, division, vacante_nuevo, fecha_inicio_contrato, horario_vigente,
    fecha_nacimiento, persona_id, cargo, area, ccosto, desc_ccosto, tipo_contrato, tipo_cargo,
    periodo_inicio_jornada, periodo_fin_jornada,
    hs_m8, hs_m9, hs_m10, hs_m11, hs_m12,
    vigencia_hs_12, hs_dic, prom_pond_hs, tramo_jornada, cambio_jornada,
    dias_semana_contrato, hd_promedio, colacion_en_contrato, horas_colacion_contrato,
    hs_disp_antes_nas_dic, hs_pausa_saludable, hs_nas_total,
    pct_nas_base, pct_directo_base, pct_indirecto_base,
    pct_nas_final_p2, pct_directo_final_p2, pct_indirecto_final_p2, pct_asistencial_p2,
    hs_disp_ate_p2, std_min_directa, std_min_indirecta,
    ate_dir_sem_p2, ate_ind_sem_p2, ate_dir_grupal_sem_p2, ate_ind_total_sem_p2, total_ate_sem_p2,
    cal_m8, cal_m9, cal_m10, cal_m11, cal_m12, semanas_habiles_planificadas,
    semhab_m8, semhab_m9, semhab_m10, semhab_m11, semhab_m12, semanas_habiles_totales, dias_habiles_totales,
    plan_anual_directa, plan_anual_indirecta, plan_anual_total,
    mg_anual_directa, mg_anual_indirecta, mg_anual_total, foja_pct, observaciones, version_id
)
SELECT
    (SELECT COALESCE(MAX(orden), 471) + 1 FROM dotacion_asistencial),
    CAST(s.IT AS INTEGER), TRIM(s.Instituto), TRIM(s."Área"), 'Nuevo',
    '2026-08-01', '2026-12-31', '1994-06-15', CAST(s.PersonaId AS INTEGER),
    TRIM(s.Cargo), TRIM(s.Cod_Esp), CAST(s."Centro de Costo" AS INTEGER), TRIM(s."Nombre Sub-área"),
    TRIM(s."Tipo de Contrato"), TRIM(s.Tipo_Trabajador),
    '2026-08-01', '2026-12-31',
    CAST(s."Horas Semanales" AS REAL), CAST(s."Horas Semanales" AS REAL), CAST(s."Horas Semanales" AS REAL), CAST(s."Horas Semanales" AS REAL), CAST(s."Horas Semanales" AS REAL),
    1, CAST(s."Horas Semanales" AS REAL), CAST(s."Horas Semanales" AS REAL),
    CASE WHEN CAST(s."Horas Semanales" AS REAL) >= 40 THEN '>=40' ELSE '22-40' END,
    'Aumento', CAST(s.Ndias AS REAL), ROUND(CAST(s."Horas Semanales" AS REAL) / CAST(s.Ndias AS REAL), 1),
    'N', 0, CAST(s."Horas Semanales" AS REAL),
    ROUND(CAST(s."Horas Semanales" AS REAL) * 0.10, 2),
    ROUND(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_nas, 2),
    d.t1_pct_nas, d.t1_pct_dir, d.t1_pct_ind,
    d.t1_pct_nas, d.t1_pct_dir, d.t1_pct_ind, (1.0 - d.t1_pct_nas),
    CAST(s."Horas Semanales" AS REAL) * (1.0 - d.t1_pct_nas),
    d.std_min_directa, d.std_min_indirecta,
    CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_dir * 60.0 / d.std_min_directa AS INTEGER),
    CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_ind * 60.0 / d.std_min_indirecta AS INTEGER),
    CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_dir * 60.0 / d.std_min_directa AS INTEGER),
    CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_ind * 60.0 / d.std_min_indirecta AS INTEGER),
    CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_dir * 60.0 / d.std_min_directa AS INTEGER) + CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_ind * 60.0 / d.std_min_indirecta AS INTEGER),
    4.2, 4.2, 4.2, 4.2, 4.2, 21.0,
    4.0, 4.0, 4.0, 4.0, 4.0, 20.0, 100.0,
    20 * CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_dir * 60.0 / d.std_min_directa AS INTEGER),
    20 * CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_ind * 60.0 / d.std_min_indirecta AS INTEGER),
    20 * (CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_dir * 60.0 / d.std_min_directa AS INTEGER) + CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_ind * 60.0 / d.std_min_indirecta AS INTEGER)),
    CAST(20 * CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_dir * 60.0 / d.std_min_directa AS INTEGER) * 0.90 AS INTEGER),
    CAST(20 * CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_ind * 60.0 / d.std_min_indirecta AS INTEGER) * 0.90 AS INTEGER),
    CAST(20 * (CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_dir * 60.0 / d.std_min_directa AS INTEGER) + CAST(CAST(s."Horas Semanales" AS REAL) * d.t1_pct_ind * 60.0 / d.std_min_indirecta AS INTEGER)) * 0.90 AS INTEGER),
    ROUND(d.t1_pct_dir * 100.0, 2),
    'Alta incorporada automáticamente por Blue Prism (Agosto 2026)', 2
FROM stg_novedades_buk s
INNER JOIN distribucion_nac d ON TRIM(UPPER(s.Cargo)) = TRIM(UPPER(d.cargo))
LEFT JOIN dotacion_asistencial da ON CAST(s.PersonaId AS INTEGER) = da.persona_id AND CAST(s.IT AS INTEGER) = da.it
WHERE da.persona_id IS NULL
  AND CAST(s."Horas Semanales" AS REAL) > 0 AND CAST(s."Horas Semanales" AS REAL) <= 44;

-- 3. Detectar CAMBIOS DE JORNADA y BAJAS en profesionales existentes y auditarlos
INSERT INTO auditoria_cambios (
    version_id, actor, tipo_evento, it, instituto, persona_id, entidad_afectada,
    campo_modificado, valor_anterior, valor_nuevo, mes_inicio_vigencia,
    impacto_atenciones_anual, justificacion, estado_aprobacion
)
SELECT
    2, 'Bot Blue Prism (03_Novedades_BUK_y_Contraste_SIC)',
    CASE
        WHEN CAST(s."Horas Semanales" AS REAL) = 0 THEN 'BAJA_TERMINO_CONTRATO'
        WHEN CAST(s."Horas Semanales" AS REAL) > da.hs_dic THEN 'CAMBIO_JORNADA_AUMENTO'
        ELSE 'CAMBIO_JORNADA_DISMINUCION'
    END,
    da.it, da.instituto, da.persona_id,
    da.cargo || ' (PersonaId ' || CAST(da.persona_id AS TEXT) || ')',
    'Horas Semanales (Ago-Dic)',
    CAST(CAST(da.hs_dic AS INTEGER) AS TEXT) || ' H/S',
    CAST(CAST(s."Horas Semanales" AS INTEGER) AS TEXT) || ' H/S',
    8,
    CAST(ROUND((CAST(s."Horas Semanales" AS REAL) - da.hs_dic) / CASE WHEN da.hs_dic > 0 THEN da.hs_dic ELSE 1 END * (da.mg_tot_m8 + da.mg_tot_m9 + da.mg_tot_m10 + da.mg_tot_m11 + da.mg_tot_m12), 0) AS INTEGER),
    'Ajuste automático de planificación desde Agosto 2026 por cambio en reporte BUK', 'APROBADO'
FROM stg_novedades_buk s
INNER JOIN distribucion_nac d ON TRIM(UPPER(s.Cargo)) = TRIM(UPPER(d.cargo))
INNER JOIN dotacion_asistencial da ON CAST(s.PersonaId AS INTEGER) = da.persona_id AND CAST(s.IT AS INTEGER) = da.it
WHERE CAST(s."Horas Semanales" AS REAL) >= 0 AND CAST(s."Horas Semanales" AS REAL) <= 44
  AND CAST(s."Horas Semanales" AS REAL) <> da.hs_dic;

-- 4. Ajustar planificación en dotacion_asistencial desde el mes 8 (Agosto a Diciembre) proporcionalmente
UPDATE dotacion_asistencial
SET
    cambio_jornada = CASE
        WHEN (SELECT CAST(s."Horas Semanales" AS REAL) FROM stg_novedades_buk s WHERE CAST(s.PersonaId AS INTEGER) = dotacion_asistencial.persona_id AND CAST(s.IT AS INTEGER) = dotacion_asistencial.it) > hs_dic THEN 'Aumento'
        ELSE 'Disminución'
    END,
    plan_anual_directa = (plan_dir_m1 + plan_dir_m2 + plan_dir_m3 + plan_dir_m4 + plan_dir_m5 + plan_dir_m6 + plan_dir_m7) +
        ROUND((plan_dir_m8 + plan_dir_m9 + plan_dir_m10 + plan_dir_m11 + plan_dir_m12) *
              ((SELECT CAST(s."Horas Semanales" AS REAL) FROM stg_novedades_buk s WHERE CAST(s.PersonaId AS INTEGER) = dotacion_asistencial.persona_id AND CAST(s.IT AS INTEGER) = dotacion_asistencial.it) / CASE WHEN hs_dic > 0 THEN hs_dic ELSE 1 END), 0),
    plan_anual_total = (plan_tot_m1 + plan_tot_m2 + plan_tot_m3 + plan_tot_m4 + plan_tot_m5 + plan_tot_m6 + plan_tot_m7) +
        ROUND((plan_tot_m8 + plan_tot_m9 + plan_tot_m10 + plan_tot_m11 + plan_tot_m12) *
              ((SELECT CAST(s."Horas Semanales" AS REAL) FROM stg_novedades_buk s WHERE CAST(s.PersonaId AS INTEGER) = dotacion_asistencial.persona_id AND CAST(s.IT AS INTEGER) = dotacion_asistencial.it) / CASE WHEN hs_dic > 0 THEN hs_dic ELSE 1 END), 0),
    mg_anual_directa = (mg_dir_m1 + mg_dir_m2 + mg_dir_m3 + mg_dir_m4 + mg_dir_m5 + mg_dir_m6 + mg_dir_m7) +
        ROUND((mg_dir_m8 + mg_dir_m9 + mg_dir_m10 + mg_dir_m11 + mg_dir_m12) *
              ((SELECT CAST(s."Horas Semanales" AS REAL) FROM stg_novedades_buk s WHERE CAST(s.PersonaId AS INTEGER) = dotacion_asistencial.persona_id AND CAST(s.IT AS INTEGER) = dotacion_asistencial.it) / CASE WHEN hs_dic > 0 THEN hs_dic ELSE 1 END), 0),
    mg_anual_total = (mg_tot_m1 + mg_tot_m2 + mg_tot_m3 + mg_tot_m4 + mg_tot_m5 + mg_tot_m6 + mg_tot_m7) +
        ROUND((mg_tot_m8 + mg_tot_m9 + mg_tot_m10 + mg_tot_m11 + mg_tot_m12) *
              ((SELECT CAST(s."Horas Semanales" AS REAL) FROM stg_novedades_buk s WHERE CAST(s.PersonaId AS INTEGER) = dotacion_asistencial.persona_id AND CAST(s.IT AS INTEGER) = dotacion_asistencial.it) / CASE WHEN hs_dic > 0 THEN hs_dic ELSE 1 END), 0),
    hs_dic = (SELECT CAST(s."Horas Semanales" AS REAL) FROM stg_novedades_buk s WHERE CAST(s.PersonaId AS INTEGER) = dotacion_asistencial.persona_id AND CAST(s.IT AS INTEGER) = dotacion_asistencial.it),
    vigencia_hs_12 = CASE WHEN (SELECT CAST(s."Horas Semanales" AS REAL) FROM stg_novedades_buk s WHERE CAST(s.PersonaId AS INTEGER) = dotacion_asistencial.persona_id AND CAST(s.IT AS INTEGER) = dotacion_asistencial.it) > 0 THEN 1 ELSE 0 END,
    observaciones = 'Actualizado por Bot Blue Prism (Novedades Agosto 2026)',
    version_id = 2,
    updated_at = CURRENT_TIMESTAMP
WHERE EXISTS (
    SELECT 1 FROM stg_novedades_buk s
    INNER JOIN distribucion_nac d ON TRIM(UPPER(s.Cargo)) = TRIM(UPPER(d.cargo))
    WHERE CAST(s.PersonaId AS INTEGER) = dotacion_asistencial.persona_id
      AND CAST(s.IT AS INTEGER) = dotacion_asistencial.it
      AND CAST(s."Horas Semanales" AS REAL) >= 0 AND CAST(s."Horas Semanales" AS REAL) <= 44
      AND CAST(s."Horas Semanales" AS REAL) <> dotacion_asistencial.hs_dic
);

UPDATE dotacion_asistencial
SET plan_anual_indirecta = plan_anual_total - plan_anual_directa,
    mg_anual_indirecta = mg_anual_total - mg_anual_directa
WHERE version_id = 2;
""".strip()

    xml = f"""<?xml version="1.0" encoding="utf-8"?>
<process name="03_Novedades_BUK_y_Contraste_SIC" version="1.0" bpversion="7.5.1.18099" cpeversion="10.0.136.0" narrative="Proceso para Demo en Vivo: lee un reporte nuevo de BUK con cambios (altas, bajas, cambios de jornada y errores intencionales) y reporte SIC usando MS Excel VBO - Extended, valida y recalcula desde el mes de vigencia en Data - SQLite, y exporta el reporte consolidado actualizado." byrefcollection="true" processrunningmessage="" disableversioning="false" preferredid="33333333-aaaa-4bbb-8ccc-000000000003">
  <view>
    <camerax>0</camerax>
    <cameray>150</cameray>
    <zoom version="2">1.25</zoom>
  </view>
  <preconditions />
  <endpoint narrative="" />
  <stage stageid="p3000001-0000-4000-8000-000000000001" name="Start" type="Start">
    <display x="15" y="-240" />
    <onsuccess>p3000001-0000-4000-8000-000000000002</onsuccess>
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000002" name="Excel::Create Instance" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-180" w="150" h="30" />
    <inputs>
      <input type="flag" name="Enable Events" friendlyname="Enable Events" expr="False" />
    </inputs>
    <outputs>
      <output type="number" name="handle" friendlyname="handle" stage="vHandle" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000003</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Create Instance" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000003" name="Excel::Open Reporte Novedades BUK" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-120" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="File name" friendlyname="File name" expr="[vRutaNovedadesBUK]" />
    </inputs>
    <outputs>
      <output type="text" name="Workbook Name" friendlyname="Workbook Name" stage="vWorkbook" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000004</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Open Workbook" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000004" name="Excel::Leer Novedades BUK (A1:S8)" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="-60" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="text" name="Worksheet Name" friendlyname="Worksheet Name" expr="&quot;Dotacion_2026&quot;" />
      <input type="text" name="StartCell" friendlyname="StartCell" expr="&quot;A1&quot;" />
      <input type="text" name="EndCell" friendlyname="EndCell" expr="&quot;S8&quot;" />
    </inputs>
    <outputs>
      <output type="collection" name="Data" friendlyname="Data" stage="cNovedadesBUK" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000005</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Get Worksheet Range As Collection" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000005" name="Excel::Close Novedades BUK" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="0" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="flag" name="Save Data" friendlyname="Save Data" expr="False" />
    </inputs>
    <onsuccess>p3000001-0000-4000-8000-000000000006</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Close Workbook" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000006" name="Excel::Open Reporte SIC" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="60" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="File name" friendlyname="File name" expr="[vRutaReporteSIC]" />
    </inputs>
    <outputs>
      <output type="text" name="Workbook Name" friendlyname="Workbook Name" stage="vWorkbook" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000007</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Open Workbook" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000007" name="Excel::Leer Ejecucion SIC (A1:J301)" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="120" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="text" name="Worksheet Name" friendlyname="Worksheet Name" expr="&quot;Ejecucion_SIC&quot;" />
      <input type="text" name="StartCell" friendlyname="StartCell" expr="&quot;A1&quot;" />
      <input type="text" name="EndCell" friendlyname="EndCell" expr="&quot;J301&quot;" />
    </inputs>
    <outputs>
      <output type="collection" name="Data" friendlyname="Data" stage="cEjecucionSIC" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000008</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Get Worksheet Range As Collection" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000008" name="Excel::Close Reporte SIC" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="180" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vWorkbook]" />
      <input type="flag" name="Save Data" friendlyname="Save Data" expr="False" />
    </inputs>
    <onsuccess>p3000001-0000-4000-8000-000000000009</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Close Workbook" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000009" name="SQLite::Set Connection Details" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="240" w="150" h="30" />
    <inputs>
      <input type="text" name="DB File" friendlyname="DB File" expr="[vRutaBD]" />
      <input type="flag" name="Create If Not Found" friendlyname="Create If Not Found" expr="True" />
      <input type="text" name="Connection String" friendlyname="Connection String" expr="&quot;&quot;" />
    </inputs>
    <onsuccess>p3000001-0000-4000-8000-000000000010</onsuccess>
    <resource object="Data - SQLite" action="Set Connection Details" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000010" name="SQLite::Write stg_novedades_buk" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="300" w="150" h="30" />
    <inputs>
      <input type="text" name="Table Name" friendlyname="Table Name" expr="&quot;stg_novedades_buk&quot;" />
      <input type="collection" name="Collection" friendlyname="Collection" expr="[cNovedadesBUK]" />
      <input type="flag" name="Clear Table First" friendlyname="Clear Table First" expr="True" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vRegistrosAfectados" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000011</onsuccess>
    <resource object="Data - SQLite" action="Write Collection" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000011" name="SQLite::Write stg_ejecucion_sic" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="360" w="150" h="30" />
    <inputs>
      <input type="text" name="Table Name" friendlyname="Table Name" expr="&quot;stg_ejecucion_sic&quot;" />
      <input type="collection" name="Collection" friendlyname="Collection" expr="[cEjecucionSIC]" />
      <input type="flag" name="Clear Table First" friendlyname="Clear Table First" expr="True" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vRegistrosAfectados" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000012</onsuccess>
    <resource object="Data - SQLite" action="Write Collection" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000012" name="SQLite::Validar, Detectar Diff y Recalcular Metas" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="420" w="150" h="30" />
    <inputs>
      <input type="text" name="SQL" friendlyname="SQL" expr="[vSQL_ProcesarNovedades]" />
      <input type="number" name="CommandTimeout" friendlyname="CommandTimeout" expr="60" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="number" name="Records Affected" friendlyname="Records Affected" stage="vRegistrosAfectados" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000013</onsuccess>
    <resource object="Data - SQLite" action="Execute Non Query" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000013" name="SQLite::Obtener Errores Detectados" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="480" w="150" h="30" />
    <inputs>
      <input type="text" name="SQL" friendlyname="SQL" expr="&quot;SELECT fecha_deteccion, origen_archivo, persona_id, it, cargo, campo_error, valor_recibido, descripcion_error, severidad FROM errores_validacion ORDER BY id DESC;&quot;" />
      <input type="number" name="CommandTimeout" friendlyname="CommandTimeout" expr="30" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="collection" name="Results" friendlyname="Results" stage="cErroresDetectados" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000014</onsuccess>
    <resource object="Data - SQLite" action="Execute - Get Collection" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000014" name="SQLite::Obtener Bitácora de Auditoría" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="540" w="150" h="30" />
    <inputs>
      <input type="text" name="SQL" friendlyname="SQL" expr="&quot;SELECT fecha_hora, version_id, actor, tipo_evento, instituto, persona_id, entidad_afectada, campo_modificado, valor_anterior, valor_nuevo, mes_inicio_vigencia, impacto_atenciones_anual, justificacion FROM auditoria_cambios ORDER BY id DESC;&quot;" />
      <input type="number" name="CommandTimeout" friendlyname="CommandTimeout" expr="30" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="collection" name="Results" friendlyname="Results" stage="cCambiosAuditados" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000015</onsuccess>
    <resource object="Data - SQLite" action="Execute - Get Collection" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000015" name="SQLite::Obtener Consolidado Actualizado" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="600" w="150" h="30" />
    <inputs>
      <input type="text" name="SQL" friendlyname="SQL" expr="&quot;SELECT * FROM vw_resumen_nacional_instituto;&quot;" />
      <input type="number" name="CommandTimeout" friendlyname="CommandTimeout" expr="30" />
    </inputs>
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
      <output type="collection" name="Results" friendlyname="Results" stage="cResumenActualizado" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000016</onsuccess>
    <resource object="Data - SQLite" action="Execute - Get Collection" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000016" name="SQLite::Close" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="660" w="150" h="30" />
    <outputs>
      <output type="flag" name="Success" friendlyname="Success" stage="vSuccess" />
      <output type="text" name="Error Message" friendlyname="Error Message" stage="vErrorMessage" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000017</onsuccess>
    <resource object="Data - SQLite" action="Close" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000017" name="Excel::Create Workbook Reporte Salida" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="720" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
    </inputs>
    <outputs>
      <output type="text" name="Workbook Name" friendlyname="Workbook Name" stage="vOutWb" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000018</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Create Workbook" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000018" name="Excel::Escribir Consolidado Nacional" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="780" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="collection" name="Collection" friendlyname="Collection" expr="[cResumenActualizado]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vOutWb]" />
      <input type="text" name="Worksheet Name" friendlyname="Worksheet Name" expr="&quot;&quot;" />
      <input type="text" name="Cell Reference" friendlyname="Cell Reference" expr="&quot;A1&quot;" />
      <input type="flag" name="Include Column Names" friendlyname="Include Column Names" expr="True" />
    </inputs>
    <onsuccess>p3000001-0000-4000-8000-000000000019</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Write Collection" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000019" name="Excel::Guardar Reporte Salida" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="840" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vOutWb]" />
      <input type="text" name="Filename" friendlyname="Filename" expr="[vRutaSalidaReporte]" />
    </inputs>
    <outputs>
      <output type="text" name="New Workbook Name" friendlyname="New Workbook Name" stage="vOutWbFinal" />
    </outputs>
    <onsuccess>p3000001-0000-4000-8000-000000000020</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Save Workbook As" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000020" name="Excel::Close Reporte Salida" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="900" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
      <input type="text" name="Workbook Name" friendlyname="Workbook Name" expr="[vOutWbFinal]" />
      <input type="flag" name="Save Data" friendlyname="Save Data" expr="False" />
    </inputs>
    <onsuccess>p3000001-0000-4000-8000-000000000021</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Close Workbook" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000021" name="Excel::Release Instance" type="Action">
    <loginhibit onsuccess="true" />
    <display x="15" y="960" w="150" h="30" />
    <inputs>
      <input type="number" name="handle" friendlyname="handle" expr="[vHandle]" />
    </inputs>
    <onsuccess>p3000001-0000-4000-8000-000000000022</onsuccess>
    <resource object="MS Excel VBO - Extended" action="Release Instance" />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000022" name="End" type="End">
    <display x="15" y="1020" />
  </stage>
  <!-- Data Items -->
  <stage stageid="p3000001-0000-4000-8000-000000000101" name="vRutaNovedadesBUK" type="Data">
    <display x="-240" y="-195" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{RUTA_EXCEL_NOVEDADES}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000102" name="vRutaReporteSIC" type="Data">
    <display x="-240" y="-150" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{RUTA_EXCEL_SIC}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000103" name="vRutaBD" type="Data">
    <display x="-240" y="-105" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{RUTA_DB}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000104" name="vRutaSalidaReporte" type="Data">
    <display x="-240" y="-60" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{RUTA_SALIDA_REPORTE}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000105" name="vHandle" type="Data">
    <display x="-240" y="-15" w="120" h="30" />
    <datatype>number</datatype>
    <initialvalue>0</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000106" name="vWorkbook" type="Data">
    <display x="-240" y="30" w="150" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000107" name="vOutWb" type="Data">
    <display x="-240" y="75" w="150" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000108" name="vOutWbFinal" type="Data">
    <display x="-240" y="120" w="150" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000109" name="vSuccess" type="Data">
    <display x="-240" y="165" w="120" h="30" />
    <datatype>flag</datatype>
    <initialvalue>False</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000110" name="vErrorMessage" type="Data">
    <display x="-240" y="210" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue />
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000111" name="vRegistrosAfectados" type="Data">
    <display x="-240" y="255" w="150" h="30" />
    <datatype>number</datatype>
    <initialvalue>0</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000112" name="vSQL_ProcesarNovedades" type="Data">
    <display x="-240" y="300" w="180" h="30" />
    <datatype>text</datatype>
    <initialvalue xml:space="preserve">{xml_escape_text(sql_procesar_novedades)}</initialvalue>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000113" name="cNovedadesBUK" type="Collection">
    <display x="240" y="-60" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000114" name="cEjecucionSIC" type="Collection">
    <display x="240" y="120" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000115" name="cErroresDetectados" type="Collection">
    <display x="240" y="480" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000116" name="cCambiosAuditados" type="Collection">
    <display x="240" y="555" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
  <stage stageid="p3000001-0000-4000-8000-000000000117" name="cResumenActualizado" type="Collection">
    <display x="240" y="630" w="150" h="60" />
    <datatype>collection</datatype>
    <alwaysinit />
  </stage>
</process>"""
    out_path = os.path.join(BP_DIR, '05_BPA_Process_03_Novedades_BUK_y_Contraste_SIC.xml')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(xml)
    print("OK -> Generado 05_BPA_Process_03_Novedades_BUK_y_Contraste_SIC.xml")
    return out_path, xml


def strip_xml_decl(xml_str):
    s = xml_str.lstrip('\ufeff').strip()
    if s.startswith('<?xml'):
        idx = s.find('?>')
        if idx != -1:
            s = s[idx + 2:].strip()
    return s


def build_bprelease(excel_vbo_xml, sqlite_vbo_xml, p1_xml, p2_xml, p3_xml):
    """Empaqueta los 2 Objetos y los 3 Procesos en un archivo .bprelease único importable."""
    e_body = strip_xml_decl(excel_vbo_xml)
    s_body = strip_xml_decl(sqlite_vbo_xml)
    p1_body = strip_xml_decl(p1_xml)
    p2_body = strip_xml_decl(p2_xml)
    p3_body = strip_xml_decl(p3_xml)

    bprelease = f"""<?xml version="1.0" encoding="utf-8"?>
<bpr:release xmlns:bpr="http://www.blueprism.co.uk/product/release">
  <bpr:name>Botathon Teletón 2026 - Cada Hora Cuenta (SQLite + Excel Extended)</bpr:name>
  <bpr:release-notes>Paquete completo con Objetos MS Excel VBO - Extended y Data - SQLite, más los 3 procesos de Ingesta BUK, Cálculo de Planificación Asistencial y Procesamiento de Novedades + Contraste SIC.</bpr:release-notes>
  <bpr:created>2026-11-07 05:00:00Z</bpr:created>
  <bpr:package-id>101</bpr:package-id>
  <bpr:package-name>Botathon_Teleton_2026_Completo</bpr:package-name>
  <bpr:user-created-by>admin</bpr:user-created-by>
  <bpr:contents count="7">
    <process id="11111111-aaaa-4bbb-8ccc-000000000001" name="01_Carga_Base_Contratos_BUK" xmlns="http://www.blueprism.co.uk/product/process">
      {p1_body}
    </process>
    <process id="22222222-aaaa-4bbb-8ccc-000000000002" name="02_Carga_y_Calculo_Planificacion" xmlns="http://www.blueprism.co.uk/product/process">
      {p2_body}
    </process>
    <process id="33333333-aaaa-4bbb-8ccc-000000000003" name="03_Novedades_BUK_y_Contraste_SIC" xmlns="http://www.blueprism.co.uk/product/process">
      {p3_body}
    </process>
    <object id="30a1808c-9709-44d8-814a-1f17e2eb766f" name="MS Excel VBO - Extended" xmlns="http://www.blueprism.co.uk/product/process">
      {e_body}
    </object>
    <object id="d91164ca-ea87-4f73-bad2-ee9010d382d9" name="Data - SQLite" xmlns="http://www.blueprism.co.uk/product/process">
      {s_body}
    </object>
    <process-group id="d133468a-5c27-4e11-a24e-65b32f51d832" name="Default" isDefaultGroup="True" xmlns="http://www.blueprism.co.uk/product/process-group">
      <members>
        <process id="11111111-aaaa-4bbb-8ccc-000000000001" />
        <process id="22222222-aaaa-4bbb-8ccc-000000000002" />
        <process id="33333333-aaaa-4bbb-8ccc-000000000003" />
      </members>
    </process-group>
    <object-group id="52f63a1e-b0a7-458f-b597-b603604eefc0" name="Default" isDefaultGroup="True" xmlns="http://www.blueprism.co.uk/product/object-group">
      <members>
        <object id="30a1808c-9709-44d8-814a-1f17e2eb766f" />
        <object id="d91164ca-ea87-4f73-bad2-ee9010d382d9" />
      </members>
    </object-group>
  </bpr:contents>
</bpr:release>"""

    out_path = os.path.join(BP_DIR, 'Botathon_Teleton_2026_Completo.bprelease')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(bprelease)
    print(f"OK -> Generado paquete unificado {out_path}")
    return out_path


def main():
    copy_sqlite_dlls()
    _, excel_vbo_xml = build_excel_vbo_extended()
    _, sqlite_vbo_xml = build_sqlite_vbo()
    _, p1_xml = build_process_01_carga_buk()
    _, p2_xml = build_process_02_carga_planificacion()
    _, p3_xml = build_process_03_novedades_demo()
    build_bprelease(excel_vbo_xml, sqlite_vbo_xml, p1_xml, p2_xml, p3_xml)

    # Validar que todos los archivos XML generados sean XML bien formado
    for fname in sorted(os.listdir(BP_DIR)):
        if fname.endswith('.xml') or fname.endswith('.bprelease'):
            fpath = os.path.join(BP_DIR, fname)
            ET.parse(fpath)
            print(f"VALIDATED XML -> {fname} ({os.path.getsize(fpath):,d} bytes)")


if __name__ == '__main__':
    main()
