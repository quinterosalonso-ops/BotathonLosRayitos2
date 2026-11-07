IF DB_ID(N'TeletonPlanificacion') IS NULL
    CREATE DATABASE [TeletonPlanificacion];
GO
USE [TeletonPlanificacion];
GO
DROP VIEW IF EXISTS dbo.vw_resumen_nacional_instituto;
DROP TABLE IF EXISTS dbo.auditoria_cambios;
DROP TABLE IF EXISTS dbo.errores_validacion;
DROP TABLE IF EXISTS dbo.ejecucion_sic;
DROP TABLE IF EXISTS dbo.planificacion_mensual;
DROP TABLE IF EXISTS dbo.dotacion_asistencial;
DROP TABLE IF EXISTS dbo.stg_contratos_buk;
DROP TABLE IF EXISTS dbo.contratos_buk;
DROP TABLE IF EXISTS dbo.excepciones_multi_it;
DROP TABLE IF EXISTS dbo.lop_instituto;
DROP TABLE IF EXISTS dbo.feriados_regionales;
DROP TABLE IF EXISTS dbo.calendario_2026;
DROP TABLE IF EXISTS dbo.factores_grupales;
DROP TABLE IF EXISTS dbo.catalogo_asignaciones;
DROP TABLE IF EXISTS dbo.distribucion_nac;
DROP TABLE IF EXISTS dbo.especialidades;
DROP TABLE IF EXISTS dbo.institutos;
DROP TABLE IF EXISTS dbo.parametros_globales;

    -- 1. Parámetros Globales editables sin tocar código
    CREATE TABLE dbo.parametros_globales (
        clave NVARCHAR(255) PRIMARY KEY,
        valor_numerico FLOAT,
        valor_texto NVARCHAR(255),
        descripcion NVARCHAR(255),
        updated_at NVARCHAR(255) DEFAULT CURRENT_TIMESTAMP
    );

    -- 2. Catálogo de Institutos Teletón (17 institutos)
    CREATE TABLE dbo.institutos (
        it INT PRIMARY KEY,
        codigo_inst NVARCHAR(255) NOT NULL,
        nombre_instituto NVARCHAR(255) NOT NULL,
        zona_geografica NVARCHAR(255) NOT NULL,
        aplica_grupales_mayor INT DEFAULT 0, -- 1 para Santiago (131) y Concepción (81)
        dias_vacaciones_legales FLOAT DEFAULT 15.0, -- 20.0 en Coyhaique/Aysén (111)
        sem_vac_ene FLOAT DEFAULT 0.0,              -- 0.4 en Coyhaique
        sem_vac_feb FLOAT DEFAULT 3.0,              -- 3.6 en Coyhaique, 3.0 resto
        dias_feriado_adicional FLOAT DEFAULT 5.0,
        proyeccion_pacientes_2025 FLOAT DEFAULT 0.0
    );

    -- 3. Catálogo de Especialidades / Áreas
    CREATE TABLE dbo.especialidades (
        cod_esp NVARCHAR(255) PRIMARY KEY,
        nombre_especialidad NVARCHAR(255) NOT NULL,
        dias_jornada_nacional FLOAT DEFAULT 0.0
    );

    -- 4. Matriz Parametrizable de Distribución por Cargo y Tramo (Distribucion_NAC)
    CREATE TABLE dbo.distribucion_nac (
        id INT IDENTITY(1,1) PRIMARY KEY,
        cod_esp NVARCHAR(255),
        cargo NVARCHAR(255) UNIQUE NOT NULL,
        unidad NVARCHAR(255),
        std_min_directa FLOAT NOT NULL DEFAULT 40.0,
        std_min_indirecta FLOAT NOT NULL DEFAULT 30.0,
        std_min_pabellon FLOAT DEFAULT 0.0,
        -- Tramo 1: >= 40 Horas
        t1_pct_dir FLOAT NOT NULL DEFAULT 0.65,
        t1_pct_ind FLOAT NOT NULL DEFAULT 0.10,
        t1_pct_nas FLOAT NOT NULL DEFAULT 0.25,
        -- Tramo 2: 22 <= Horas < 40
        t2_pct_dir FLOAT NOT NULL DEFAULT 0.70,
        t2_pct_ind FLOAT NOT NULL DEFAULT 0.10,
        t2_pct_nas FLOAT NOT NULL DEFAULT 0.20,
        -- Tramo 3: < 22 Horas
        t3_pct_dir FLOAT NOT NULL DEFAULT 0.75,
        t3_pct_ind FLOAT NOT NULL DEFAULT 0.10,
        t3_pct_nas FLOAT NOT NULL DEFAULT 0.15,
        es_quirurgico INT DEFAULT 0
    );

    -- 5. Catálogo de Asignaciones Especiales
    CREATE TABLE dbo.catalogo_asignaciones (
        id INT IDENTITY(1,1) PRIMARY KEY,
        nombre_asignacion NVARCHAR(255) UNIQUE NOT NULL,
        horas_semanales FLOAT DEFAULT 1.0,
        requiere_autorizacion_dmn INT DEFAULT 0,
        descripcion NVARCHAR(255)
    );

    -- 6. Factores de Atención Grupal por Especialidad
    CREATE TABLE dbo.factores_grupales (
        cod_esp NVARCHAR(255) PRIMARY KEY,
        pct_grupal_estandar FLOAT NOT NULL DEFAULT 0.0,
        num_pacientes_san_con FLOAT NOT NULL DEFAULT 0.0,
        num_pacientes_otros_it FLOAT NOT NULL DEFAULT 0.0
    );

    -- 7. Calendario 2026 por Mes
    CREATE TABLE dbo.calendario_2026 (
        mes INT PRIMARY KEY,
        nombre_mes NVARCHAR(255) NOT NULL,
        fecha_inicio NVARCHAR(255) NOT NULL,
        fecha_termino NVARCHAR(255) NOT NULL,
        periodo_serial_excel INT NOT NULL,
        dias_totales INT NOT NULL,
        dias_inhabiles_sin_vac INT NOT NULL,
        dias_inhabiles_con_vac INT NOT NULL,
        dias_habiles_sin_vac INT NOT NULL,
        dias_habiles_con_vac INT NOT NULL,
        semanas_disponibles FLOAT NOT NULL,
        regla_redondeo_semhab NVARCHAR(255) DEFAULT 'EXACT' -- 'ROUNDUP_1' en meses 5 y 7
    );

    -- 8. Feriados Nacionales y Regionales 2026
    CREATE TABLE dbo.feriados_regionales (
        id INT IDENTITY(1,1) PRIMARY KEY,
        fecha NVARCHAR(255) NOT NULL,
        mes INT NOT NULL,
        dia_semana INT NOT NULL,
        nombre_feriado NVARCHAR(255) NOT NULL,
        ambito_it NVARCHAR(255) DEFAULT 'NACIONAL' -- 'NACIONAL' o código IT / abreviatura (Ari, Iqu, Cop, Coq, Nub)
    );

    -- 9. Producción LOP por Instituto
    CREATE TABLE dbo.lop_instituto (
        it INT PRIMARY KEY,
        instituto NVARCHAR(255) NOT NULL,
        total_atenciones_lop FLOAT NOT NULL DEFAULT 0.0
    );

    -- 10. Excepciones Multi-Instituto
    CREATE TABLE dbo.excepciones_multi_it (
        excep_id NVARCHAR(255) PRIMARY KEY,
        persona_id INT NOT NULL,
        it INT NOT NULL,
        instituto NVARCHAR(255) NOT NULL,
        nombre_it NVARCHAR(255),
        cargo NVARCHAR(255),
        ccosto INT,
        fec_ingreso NVARCHAR(255),
        hs_ene FLOAT DEFAULT 0, hs_feb FLOAT DEFAULT 0, hs_mar FLOAT DEFAULT 0, hs_abr FLOAT DEFAULT 0,
        hs_may FLOAT DEFAULT 0, hs_jun FLOAT DEFAULT 0, hs_jul FLOAT DEFAULT 0, hs_ago FLOAT DEFAULT 0,
        hs_sep FLOAT DEFAULT 0, hs_oct FLOAT DEFAULT 0, hs_nov FLOAT DEFAULT 0, hs_dic FLOAT DEFAULT 0,
        dias_semana FLOAT DEFAULT 5
    );

    -- 11. Contratos Mensuales BUK (Fuente: Nac - Base contratos x mes.xlsx)
    CREATE TABLE dbo.contratos_buk (
        id INT IDENTITY(1,1) PRIMARY KEY,
        it INT NOT NULL,
        instituto NVARCHAR(255) NOT NULL,
        persona_id INT NOT NULL,
        area_buk NVARCHAR(255),
        nombre_subarea NVARCHAR(255),
        centro_costo INT,
        fecha_nacimiento NVARCHAR(255),
        periodo_serial INT,
        periodo_fecha NVARCHAR(255),
        mes INT,
        tipo_contrato NVARCHAR(255),
        tipo_trab_anterior NVARCHAR(255),
        fecha_ingreso NVARCHAR(255),
        fecha_termino NVARCHAR(255),
        horas_semanales FLOAT NOT NULL,
        cargo NVARCHAR(255) NOT NULL,
        tipo_trabajador NVARCHAR(255),
        dias_jornada NVARCHAR(255),
        reg_sic NVARCHAR(255),
        cod_esp NVARCHAR(255),
        ndias FLOAT NOT NULL DEFAULT 5,
        origen_carga NVARCHAR(255) DEFAULT 'INICIAL',
        UNIQUE(it, persona_id, mes)
    );

    -- Tabla Staging para lectura cruda desde Blue Prism (MS Excel VBO - Extended -> Data - MSSQLLocalDB)
    CREATE TABLE dbo.stg_contratos_buk (
        IT NVARCHAR(255),
        Instituto NVARCHAR(255),
        PersonaId NVARCHAR(255),
        Area NVARCHAR(255),
        Nombre_Sub_area NVARCHAR(255),
        Centro_de_Costo NVARCHAR(255),
        Fecha_de_Nacimiento NVARCHAR(255),
        Periodo NVARCHAR(255),
        Tipo_de_Contrato NVARCHAR(255),
        Tipo_Trab_161225 NVARCHAR(255),
        Fecha_Ingreso_Compania NVARCHAR(255),
        Fecha_Termino NVARCHAR(255),
        Horas_Semanales NVARCHAR(255),
        Cargo NVARCHAR(255),
        Tipo_Trabajador NVARCHAR(255),
        Dias_de_la_Jornada NVARCHAR(255),
        Reg_SIC NVARCHAR(255),
        Cod_Esp NVARCHAR(255),
        Ndias NVARCHAR(255)
    );

    -- 12. Dotación Asistencial Calculada y Parametrizada (471 filas base + altas dinámicas)
    CREATE TABLE dbo.dotacion_asistencial (
        fila_id INT IDENTITY(1,1) PRIMARY KEY,
        orden INT,
        it INT NOT NULL,
        instituto NVARCHAR(255) NOT NULL,
        division NVARCHAR(255),
        vacante_nuevo NVARCHAR(255) DEFAULT '',
        excepcion_id NVARCHAR(255) DEFAULT '',
        fecha_inicio_contrato NVARCHAR(255),
        fecha_termino_trabajo NVARCHAR(255),
        horario_vigente NVARCHAR(255),
        fecha_nacimiento NVARCHAR(255),
        persona_id INT NOT NULL,
        repetido NVARCHAR(255) DEFAULT '',
        cargo NVARCHAR(255) NOT NULL,
        area NVARCHAR(255) NOT NULL,
        ccosto INT,
        desc_ccosto NVARCHAR(255),
        tipo_contrato NVARCHAR(255),
        tipo_cargo NVARCHAR(255),
        licencia_extensa NVARCHAR(255) DEFAULT '',
        periodo_inicio_jornada NVARCHAR(255),
        periodo_fin_jornada NVARCHAR(255),
        -- Horas contratadas por mes (Ene..Dic)
        hs_m1 FLOAT DEFAULT 0, hs_m2 FLOAT DEFAULT 0, hs_m3 FLOAT DEFAULT 0, hs_m4 FLOAT DEFAULT 0,
        hs_m5 FLOAT DEFAULT 0, hs_m6 FLOAT DEFAULT 0, hs_m7 FLOAT DEFAULT 0, hs_m8 FLOAT DEFAULT 0,
        hs_m9 FLOAT DEFAULT 0, hs_m10 FLOAT DEFAULT 0, hs_m11 FLOAT DEFAULT 0, hs_m12 FLOAT DEFAULT 0,
        vigencia_hs_12 INT DEFAULT 1,
        hs_dic FLOAT DEFAULT 0,
        prom_pond_hs FLOAT DEFAULT 0,
        tramo_jornada NVARCHAR(255) DEFAULT '>=40',
        cambio_jornada NVARCHAR(255) DEFAULT '',
        dias_semana_contrato FLOAT DEFAULT 5,
        hd_promedio FLOAT DEFAULT 8,
        colacion_en_contrato NVARCHAR(255) DEFAULT 'N',
        horas_colacion_contrato FLOAT DEFAULT 0,
        hs_disp_antes_nas_ene_abr FLOAT DEFAULT 0,
        hs_disp_antes_nas_dic FLOAT DEFAULT 0,
        hs_pausa_saludable FLOAT DEFAULT 0,
        hs_benef_colacion_15m FLOAT DEFAULT 0,
        hs_otras_nas FLOAT DEFAULT 0,
        hs_nas_total FLOAT DEFAULT 0,
        pct_nas_base FLOAT DEFAULT 0,
        pct_directo_base FLOAT DEFAULT 0,
        pct_indirecto_base FLOAT DEFAULT 0,
        asignacion_1 NVARCHAR(255) DEFAULT '',
        hrs_asig_1 FLOAT DEFAULT 0,
        asignacion_2 NVARCHAR(255) DEFAULT '',
        hrs_asig_2 FLOAT DEFAULT 0,
        asignacion_3_dmn NVARCHAR(255) DEFAULT '',
        hrs_asig_3 FLOAT DEFAULT 0,
        estado_aprobacion_dmn NVARCHAR(255) DEFAULT 'APROBADO',
        hrs_asig_4_sindical FLOAT DEFAULT 0,
        total_hrs_asignacion FLOAT DEFAULT 0,
        -- Porcentajes y tasas Período 1 (Ene - Abr: 44h)
        pct_nas_final_p1 FLOAT DEFAULT 0,
        pct_directo_final_p1 FLOAT DEFAULT 0,
        pct_indirecto_final_p1 FLOAT DEFAULT 0,
        pct_asistencial_p1 FLOAT DEFAULT 0,
        hs_disp_ate_p1 FLOAT DEFAULT 0,
        ate_dir_sem_p1 FLOAT DEFAULT 0,
        ate_ind_sem_p1 FLOAT DEFAULT 0,
        ate_dir_grupal_sem_p1 FLOAT DEFAULT 0,
        total_ate_sem_p1 FLOAT DEFAULT 0,
        -- Porcentajes y tasas Período 2 (May - Dic: 42h)
        pct_nas_final_p2 FLOAT DEFAULT 0,
        pct_directo_final_p2 FLOAT DEFAULT 0,
        pct_indirecto_final_p2 FLOAT DEFAULT 0,
        pct_asistencial_p2 FLOAT DEFAULT 0,
        hs_disp_ate_p2 FLOAT DEFAULT 0,
        std_min_directa FLOAT DEFAULT 40,
        std_min_indirecta FLOAT DEFAULT 30,
        ate_dir_sem_p2 FLOAT DEFAULT 0,
        ate_ind_sem_p2 FLOAT DEFAULT 0,
        pct_grupal_estandar FLOAT DEFAULT 0,
        cambio_pct_grupal_dmn FLOAT DEFAULT 0,
        usa_override_grupal INT DEFAULT 0,
        num_pacientes_grupo FLOAT DEFAULT 0,
        ate_dir_grupal_sem_p2 FLOAT DEFAULT 0,
        ate_ind_total_sem_p2 FLOAT DEFAULT 0,
        total_ate_sem_p2 FLOAT DEFAULT 0,
        -- Semanas Calendario Base (BI..BT)
        cal_m1 FLOAT DEFAULT 0, cal_m2 FLOAT DEFAULT 0, cal_m3 FLOAT DEFAULT 0, cal_m4 FLOAT DEFAULT 0,
        cal_m5 FLOAT DEFAULT 0, cal_m6 FLOAT DEFAULT 0, cal_m7 FLOAT DEFAULT 0, cal_m8 FLOAT DEFAULT 0,
        cal_m9 FLOAT DEFAULT 0, cal_m10 FLOAT DEFAULT 0, cal_m11 FLOAT DEFAULT 0, cal_m12 FLOAT DEFAULT 0,
        semanas_habiles_planificadas FLOAT DEFAULT 0,
        -- Licencias (BV..CG)
        lic_m1 FLOAT DEFAULT 0, lic_m2 FLOAT DEFAULT 0, lic_m3 FLOAT DEFAULT 0, lic_m4 FLOAT DEFAULT 0,
        lic_m5 FLOAT DEFAULT 0, lic_m6 FLOAT DEFAULT 0, lic_m7 FLOAT DEFAULT 0, lic_m8 FLOAT DEFAULT 0,
        lic_m9 FLOAT DEFAULT 0, lic_m10 FLOAT DEFAULT 0, lic_m11 FLOAT DEFAULT 0, lic_m12 FLOAT DEFAULT 0,
        -- Vacaciones (CH..CS)
        vac_m1 FLOAT DEFAULT 0, vac_m2 FLOAT DEFAULT 0, vac_m3 FLOAT DEFAULT 0, vac_m4 FLOAT DEFAULT 0,
        vac_m5 FLOAT DEFAULT 0, vac_m6 FLOAT DEFAULT 0, vac_m7 FLOAT DEFAULT 0, vac_m8 FLOAT DEFAULT 0,
        vac_m9 FLOAT DEFAULT 0, vac_m10 FLOAT DEFAULT 0, vac_m11 FLOAT DEFAULT 0, vac_m12 FLOAT DEFAULT 0,
        -- Permisos (CT..DE)
        perm_m1 FLOAT DEFAULT 0, perm_m2 FLOAT DEFAULT 0, perm_m3 FLOAT DEFAULT 0, perm_m4 FLOAT DEFAULT 0,
        perm_m5 FLOAT DEFAULT 0, perm_m6 FLOAT DEFAULT 0, perm_m7 FLOAT DEFAULT 0, perm_m8 FLOAT DEFAULT 0,
        perm_m9 FLOAT DEFAULT 0, perm_m10 FLOAT DEFAULT 0, perm_m11 FLOAT DEFAULT 0, perm_m12 FLOAT DEFAULT 0,
        -- Otros descuentos y beneficios (DF..DY)
        feriado_zona_extrema_ene FLOAT DEFAULT 0,
        feriado_zona_extrema_feb FLOAT DEFAULT 0,
        mes_cumpleanos INT DEFAULT 0,
        dias_cumpleanos FLOAT DEFAULT 0,
        visperas_fiestas_patrias FLOAT DEFAULT 0,
        visperas_navidad_ano_nuevo FLOAT DEFAULT 0,
        feriado_legal_adicional FLOAT DEFAULT 0,
        dias_administrativos FLOAT DEFAULT 0,
        plan_anual_capacitacion FLOAT DEFAULT 0,
        hexa_ctr FLOAT DEFAULT 0,
        jornada_nacional FLOAT DEFAULT 0,
        dias_progresivos FLOAT DEFAULT 0,
        dias_complementarios FLOAT DEFAULT 0,
        ley_21530 FLOAT DEFAULT 0,
        vacaciones_pendientes FLOAT DEFAULT 0,
        jornadas_directiva_sindical FLOAT DEFAULT 0,
        ajuste_dif_dias_contrato FLOAT DEFAULT 0,
        lactancia_dias_ano FLOAT DEFAULT 0,
        meses_2do_semestre FLOAT DEFAULT 0,
        sem_mes_prorratear_2do_sem FLOAT DEFAULT 0,
        -- Semanas Hábiles Efectivas por Mes (DZ..EK)
        semhab_m1 FLOAT DEFAULT 0, semhab_m2 FLOAT DEFAULT 0, semhab_m3 FLOAT DEFAULT 0, semhab_m4 FLOAT DEFAULT 0,
        semhab_m5 FLOAT DEFAULT 0, semhab_m6 FLOAT DEFAULT 0, semhab_m7 FLOAT DEFAULT 0, semhab_m8 FLOAT DEFAULT 0,
        semhab_m9 FLOAT DEFAULT 0, semhab_m10 FLOAT DEFAULT 0, semhab_m11 FLOAT DEFAULT 0, semhab_m12 FLOAT DEFAULT 0,
        semanas_habiles_totales FLOAT DEFAULT 0,
        dias_habiles_totales FLOAT DEFAULT 0,
        -- Atenciones Planificadas Directas por Mes (EN..EY)
        plan_dir_m1 FLOAT DEFAULT 0, plan_dir_m2 FLOAT DEFAULT 0, plan_dir_m3 FLOAT DEFAULT 0, plan_dir_m4 FLOAT DEFAULT 0,
        plan_dir_m5 FLOAT DEFAULT 0, plan_dir_m6 FLOAT DEFAULT 0, plan_dir_m7 FLOAT DEFAULT 0, plan_dir_m8 FLOAT DEFAULT 0,
        plan_dir_m9 FLOAT DEFAULT 0, plan_dir_m10 FLOAT DEFAULT 0, plan_dir_m11 FLOAT DEFAULT 0, plan_dir_m12 FLOAT DEFAULT 0,
        total_plan_ate_dir FLOAT DEFAULT 0,
        -- Atenciones Planificadas Directas + Indirectas por Mes (FA..FL)
        plan_tot_m1 FLOAT DEFAULT 0, plan_tot_m2 FLOAT DEFAULT 0, plan_tot_m3 FLOAT DEFAULT 0, plan_tot_m4 FLOAT DEFAULT 0,
        plan_tot_m5 FLOAT DEFAULT 0, plan_tot_m6 FLOAT DEFAULT 0, plan_tot_m7 FLOAT DEFAULT 0, plan_tot_m8 FLOAT DEFAULT 0,
        plan_tot_m9 FLOAT DEFAULT 0, plan_tot_m10 FLOAT DEFAULT 0, plan_tot_m11 FLOAT DEFAULT 0, plan_tot_m12 FLOAT DEFAULT 0,
        total_plan_ate_dir_ind FLOAT DEFAULT 0,
        -- Meta de Gestión (90%) Directa por Mes (FN..FY)
        mg_dir_m1 FLOAT DEFAULT 0, mg_dir_m2 FLOAT DEFAULT 0, mg_dir_m3 FLOAT DEFAULT 0, mg_dir_m4 FLOAT DEFAULT 0,
        mg_dir_m5 FLOAT DEFAULT 0, mg_dir_m6 FLOAT DEFAULT 0, mg_dir_m7 FLOAT DEFAULT 0, mg_dir_m8 FLOAT DEFAULT 0,
        mg_dir_m9 FLOAT DEFAULT 0, mg_dir_m10 FLOAT DEFAULT 0, mg_dir_m11 FLOAT DEFAULT 0, mg_dir_m12 FLOAT DEFAULT 0,
        total_mg_dir FLOAT DEFAULT 0,
        -- Meta de Gestión (90%) Directa + Indirecta por Mes (GA..GL)
        mg_tot_m1 FLOAT DEFAULT 0, mg_tot_m2 FLOAT DEFAULT 0, mg_tot_m3 FLOAT DEFAULT 0, mg_tot_m4 FLOAT DEFAULT 0,
        mg_tot_m5 FLOAT DEFAULT 0, mg_tot_m6 FLOAT DEFAULT 0, mg_tot_m7 FLOAT DEFAULT 0, mg_tot_m8 FLOAT DEFAULT 0,
        mg_tot_m9 FLOAT DEFAULT 0, mg_tot_m10 FLOAT DEFAULT 0, mg_tot_m11 FLOAT DEFAULT 0, mg_tot_m12 FLOAT DEFAULT 0,
        total_mg_dir_ind FLOAT DEFAULT 0,
        -- Resúmenes Anuales (GN..GV)
        plan_anual_directa FLOAT DEFAULT 0,
        plan_anual_indirecta FLOAT DEFAULT 0,
        plan_anual_total FLOAT DEFAULT 0,
        tiempo_anual_directa_hrs FLOAT DEFAULT 0,
        tiempo_anual_indirecta_hrs FLOAT DEFAULT 0,
        tiempo_anual_total_hrs FLOAT DEFAULT 0,
        mg_anual_directa FLOAT DEFAULT 0,
        mg_anual_indirecta FLOAT DEFAULT 0,
        mg_anual_total FLOAT DEFAULT 0,
        foja_pct FLOAT DEFAULT 0,
        observaciones NVARCHAR(255) DEFAULT '',
        version_id INT DEFAULT 1,
        updated_at NVARCHAR(255) DEFAULT CURRENT_TIMESTAMP
    );

    -- 13. Tabla Desnormalizada Mensual por Profesional (para filtros por mes en Dashboard y cruce con SIC)
    CREATE TABLE dbo.planificacion_mensual (
        id INT IDENTITY(1,1) PRIMARY KEY,
        fila_id INT NOT NULL,
        persona_id INT NOT NULL,
        it INT NOT NULL,
        instituto NVARCHAR(255) NOT NULL,
        area NVARCHAR(255) NOT NULL,
        cargo NVARCHAR(255) NOT NULL,
        mes INT NOT NULL,
        nombre_mes NVARCHAR(255) NOT NULL,
        periodo_jornada NVARCHAR(255) NOT NULL, -- 'Ene-Abr (44h)' o 'May-Dic (42h)'
        horas_semanales_contratadas FLOAT NOT NULL,
        semanas_calendario FLOAT NOT NULL,
        semanas_habiles_efectivas FLOAT NOT NULL,
        horas_mes_contratadas FLOAT NOT NULL,
        horas_mes_disp_directa FLOAT NOT NULL,
        horas_mes_disp_indirecta FLOAT NOT NULL,
        horas_mes_nas FLOAT NOT NULL,
        foja_mensual_pct FLOAT NOT NULL,
        plan_ate_directa INT NOT NULL,
        plan_ate_indirecta INT NOT NULL,
        plan_ate_total INT NOT NULL,
        meta_90_directa INT NOT NULL,
        meta_90_indirecta INT NOT NULL,
        meta_90_total INT NOT NULL
    );

    -- 14. Ejecución FLOAT SIC (Sistema de Información Clínica) para contraste Plan vs FLOAT
    CREATE TABLE dbo.ejecucion_sic (
        id INT IDENTITY(1,1) PRIMARY KEY,
        persona_id INT NOT NULL,
        it INT NOT NULL,
        instituto NVARCHAR(255) NOT NULL,
        area NVARCHAR(255) NOT NULL,
        mes INT NOT NULL,
        ate_real_directa INT NOT NULL DEFAULT 0,
        ate_real_indirecta INT NOT NULL DEFAULT 0,
        ate_real_total INT NOT NULL DEFAULT 0,
        pct_cumplimiento_meta FLOAT DEFAULT 0.0,
        estado_alerta NVARCHAR(255) DEFAULT 'OK',
        fecha_carga NVARCHAR(255) DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(persona_id, it, mes)
    );

    -- 15. Bitácora de Errores de Validación detectados por Blue Prism
    CREATE TABLE dbo.errores_validacion (
        id INT IDENTITY(1,1) PRIMARY KEY,
        fecha_deteccion NVARCHAR(255) DEFAULT CURRENT_TIMESTAMP,
        origen_archivo NVARCHAR(255) NOT NULL,
        fila_excel INT,
        persona_id NVARCHAR(255),
        it NVARCHAR(255),
        cargo NVARCHAR(255),
        campo_error NVARCHAR(255) NOT NULL,
        valor_recibido NVARCHAR(255),
        descripcion_error NVARCHAR(255) NOT NULL,
        severidad NVARCHAR(255) DEFAULT 'ERROR_BLOQUEANTE'
    );

    -- 16. Historial de Trazabilidad y Control de Versiones (Cambios BUK / Parámetros / Aprobaciones DMN)
    CREATE TABLE dbo.auditoria_cambios (
        id INT IDENTITY(1,1) PRIMARY KEY,
        fecha_hora NVARCHAR(255) DEFAULT CURRENT_TIMESTAMP,
        version_id INT DEFAULT 1,
        actor NVARCHAR(255) NOT NULL,            -- 'Bot Blue Prism', 'Planificación Central', 'DMN'
        tipo_evento NVARCHAR(255) NOT NULL,      -- 'CARGA_INICIAL', 'ALTA_BUK', 'BAJA_BUK', 'CAMBIO_JORNADA', 'CAMBIO_PARAMETRO', 'APROBACION_ASIGNACION'
        it INT,
        instituto NVARCHAR(255),
        persona_id INT,
        entidad_afectada NVARCHAR(255),
        campo_modificado NVARCHAR(255),
        valor_anterior NVARCHAR(255),
        valor_nuevo NVARCHAR(255),
        mes_inicio_vigencia INT DEFAULT 1,
        impacto_atenciones_anual INT DEFAULT 0,
        justificacion NVARCHAR(255),
        estado_aprobacion NVARCHAR(255) DEFAULT 'APROBADO'
    );

    -- Vista Consolidada Nacional por Instituto (Formato Solicitud.md)
GO
CREATE VIEW dbo.vw_resumen_nacional_instituto AS
    SELECT
        i.it AS IT,
        i.codigo_inst AS Cod_Inst,
        i.nombre_instituto AS Instituto,
        COUNT(d.fila_id) AS Dotacion_Profesionales,
        ROUND(SUM(d.hs_dic), 0) AS Suma_Horas_Semanales_Dic,
        ROUND(SUM(d.prom_pond_hs), 0) AS Suma_Horas_Semanales_Prom,
        ROUND(SUM(d.hs_dic * d.pct_directo_final_p2), 2) AS Hrs_Sem_Atencion_Directa,
        ROUND(SUM(d.hs_dic * d.pct_indirecto_final_p2), 2) AS Hrs_Sem_Atencion_Indirecta,
        ROUND(SUM(d.hs_dic * d.pct_nas_final_p2), 2) AS Hrs_Sem_Actividad_NAS,
        ROUND(
            CASE WHEN SUM(d.hs_dic) > 0
                 THEN SUM(d.hs_dic * d.pct_directo_final_p2) / SUM(d.hs_dic) * 100.0
                 ELSE 0 END, 2
        ) AS FOJA_Pct,
        CAST(SUM(d.plan_anual_directa) AS INT) AS Plan_Atencion_Anual_Directa,
        CAST(SUM(d.plan_anual_indirecta) AS INT) AS Plan_Atencion_Anual_Indirecta,
        CAST(SUM(d.plan_anual_total) AS INT) AS Plan_Total_Atenciones_Ano,
        CAST(ROUND(SUM(d.plan_anual_total) * 0.90, 0) AS INT) AS Meta_Gestion_90_Directo_Tabla,
        CAST(SUM(d.mg_anual_directa) AS INT) AS MG_Atencion_Anual_Directa,
        CAST(SUM(d.mg_anual_indirecta) AS INT) AS MG_Atencion_Anual_Indirecta,
        CAST(SUM(d.mg_anual_total) AS INT) AS MG_Total_Atenciones_Ano,
        CAST(COALESCE(l.total_atenciones_lop, 0) AS INT) AS Total_Atenciones_LOP,
        CAST(SUM(d.mg_anual_total) + COALESCE(l.total_atenciones_lop, 0) AS INT) AS Total_Planificacion_Con_LOP,
        CAST(ROUND((SUM(d.mg_anual_total) + COALESCE(l.total_atenciones_lop, 0)) * 0.90, 0) AS INT) AS Meta_Total_Con_LOP_90
    FROM institutos i
    LEFT JOIN dotacion_asistencial d ON i.it = d.it
    LEFT JOIN lop_instituto l ON i.it = l.it
    GROUP BY i.it, i.codigo_inst, i.nombre_instituto, l.total_atenciones_lop;
GO
