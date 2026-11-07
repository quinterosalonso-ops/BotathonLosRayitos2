# Plan de Trabajo (16 Horas) — Botathon 2026 "Cada hora cuenta"

Este plan de trabajo organiza las **16 horas de desarrollo** para construir, probar y presentar el MVP que reemplaza la planilla Excel de Planificación Asistencial de Teletón, poniendo a **Blue Prism en el centro** junto con un **Motor de Reglas Parametrizable** y un **Dashboard Web Accesible**.

---

## 1. Arquitectura Objetivo del MVP

```
[Reportes Excel BUK / SIC]  --->  [BOT BLUE PRISM]
  - Nac - Base contratos x mes      1. Vigila carpeta de entrada / recibe archivo
  - Reporte Novedades BUK           2. Valida estructura, tipos de datos y catálogos
  - Reporte Ejecución SIC           3. Detecta diferencias (Altas, Bajas, Cambio H/S, Cargo, IT)
                                    4. Registra log de errores y bitácora de trazabilidad
                                    5. Inyecta datos validados a la BD / API del Motor
                                                |
                                                v
                              [MOTOR DE CÁLCULO Y BASE DE DATOS]
                                - Catálogos editables (Parámetros, Distribución_NAC,
                                  Asignaciones Especiales, Calendario 2026, Feriados)
                                - Motor de cálculo mensual (Ene-Abr 44h / May-Dic 42h)
                                - Versionado e historial de cambios y aprobaciones DMN
                                                |
                                                v
                              [DASHBOARD WEB INTERACTIVO + ACCESIBLE]
                                - Vista Consolidada Nacional y por Instituto (17 IT)
                                - KPIs obligatorios: # Meta Gestión (90%) y % FOJA
                                - Filtros por Instituto, Área/Especialidad, Cargo, Profesional y Mes
                                - Editor de Parámetros + Simulador "What-if"
                                - Comparador Planificado vs Ejecutado (SIC)
                                - Accesibilidad demostrada: Voz (TTS), Alto Contraste, Lectura Fácil
```

---

## 2. Cronograma Detallado por Bloques (16 Horas)

### Fase 1: Fundamentos, Base de Datos y Semilla de Datos (Horas 1 a 3)
- **Objetivo**: Tener los datos de ambos archivos Excel cargados y el motor matemático reproduciendo exactamente los números de `Dotacion_Asistencial` y `Planificacion Total`.
- **Tareas**:
  1. **Modelado de Datos (Hora 1)**:
     - Crear esquema de base de datos (SQLite / PostgreSQL) o almacén estructurado con:
       - `institutos`, `especialidades`, `cargos_distribucion_nac`, `catalogo_asignaciones`, `factores_grupales`, `calendario_2026`, `feriados_regionales`, `lop_instituto`.
       - `contratos_buk_mensual` (datos crudos desde `Dotacion_2026`).
       - `excepciones_multi_it`, `ausentismos_descuentos`, `asignaciones_profesional`.
       - `planificacion_calculada` (resultados por profesional y mes `1..12`).
       - `auditoria_cambios` (trazabilidad: usuario/bot, fecha, campo, valor anterior, valor nuevo, motivo/estado aprobación DMN).
  2. **Script de Carga Inicial / Semilla (Hora 2)**:
     - Extraer automáticamente las tablas de `Parámetros`, `Distribucion_NAC`, `Calendario 2026`, `LOP`, `Excepciones` y `Dotacion_Asistencial` de los Excel entregados.
  3. **Motor de Cálculo Parametrizable (Hora 3)**:
     - Implementar las fórmulas exactas documentadas en `CONDICIONES_SOLICITUD.md` (Sección 3, Pasos 1 al 7).
     - Ejecutar test de validación contra las **471 filas** de `Dotacion_Asistencial` para garantizar **100% de coincidencia** en horas semanales, atenciones directas, indirectas, grupales y Meta de Gestión al 90%.

### Fase 2: Automatización Central con Blue Prism (Horas 4 a 8)
- **Objetivo**: Construir el proceso en Blue Prism que automatiza la ingesta, validación, detección de cambios y actualización de la planificación sin copiar ni pegar.
- **Tareas**:
  1. **Lectura y Validación de Reportes Excel BUK y SIC (Horas 4–5)**:
     - Usar el VBO `MS Excel VBO` (o procesamiento vía script/colección invocado por Blue Prism) para leer `Nac - Base contratos x mes.xlsx` y nuevos archivos de novedades BUK / SIC desde una carpeta vigilada (`Input_BUK`, `Input_SIC`).
     - Implementar reglas de validación en Blue Prism:
       - Campos obligatorios presentes (`IT`, `PersonaId`, `Periodo`, `Horas Semanales`, `Cargo`, `Ndias`).
       - Rango válido de horas (`0 < Horas Semanales <= 44`) y días (`1 <= Ndias <= 6`).
       - Homologación de `Cargo` contra el catálogo `Distribucion_NAC`.
       - Separación en cola de registros válidos y generación automática de reporte de errores (`Errores_Validacion_BUK.xlsx` / tabla de alertas en UI).
  2. **Motor de Detección de Novedades ("Diff" de Dotación) en Blue Prism (Horas 6–7)**:
     - Comparar el reporte BUK entrante contra la dotación vigente:
       - **Alta**: `PersonaId` nuevo que no existía en períodos anteriores.
       - **Baja / Término**: `PersonaId` que deja de aparecer o trae `Fecha Término` en el mes.
       - **Cambio de Jornada**: variación en `Horas Semanales` (distinguir rebaja legal 44→42 de mayo vs aumento/disminución contractual real).
       - **Cambio de Cargo o Instituto**: modificación en `Cargo` o `IT` / `Excepciones`.
  3. **Integración Bot → Aplicación y Generación de Reportes (Hora 8)**:
     - Blue Prism envía el lote de cambios validados al backend/base de datos, dispara el recálculo desde el mes de vigencia del cambio y mueve el archivo procesado a la carpeta `Procesados` con timestamp.
     - (Deseable) Acción en Blue Prism para exportar/distribuir el resumen actualizado por instituto.

### Fase 3: Aplicación Web, Dashboards y Parametrización (Horas 9 a 12)
- **Objetivo**: Entregar la experiencia visual para el equipo de planificación central con todos los filtros, indicadores obligatorios y deseables.
- **Tareas**:
  1. **Dashboard Ejecutivo Nacional y por Instituto (Horas 9–10)**:
     - Tarjetas KPI superiores:
       - `Total Horas Semanales Contratadas`
       - `Total Atenciones Planificadas Año` (Directas Individuales/Grupales + Indirectas + LOP)
       - **`# Meta de Gestión Asistencial (90%)`**
       - **`% FOJA (Factor de Ocupación de Jornada Asistencial)`**
     - Tabla Consolidada Nacional por Instituto idéntica a la solicitada en `Solicitud.md` (con totales de horas semanales, horas directas, indirectas, NAS, atenciones anuales y meta al 90%).
     - Filtros dinámicos combinables: **Instituto**, **Área/Especialidad**, **Tipo de Atención**, **Cargo**, **Profesional** y **Mes / Período (Ene–Abr vs May–Dic)**.
  2. **Mantenedor de Parámetros y Simulador "What-If" (Hora 11)**:
     - Pantalla donde el usuario modifica sin tocar código:
       - `% Directo`, `% Indirecto`, `% NAS` y `Minutos Estándar` por cargo y tramo (`>=40`, `22-40`, `<22`).
       - Catálogo de `Asignaciones Especiales` y bandeja de **Aprobación DMN/VP**.
       - Feriados regionales/nacionales, vacaciones por zona geográfica y `% Meta de Gestión` (ej. cambiar de 90% a 85% o 95%).
     - Modo **Simulación**: previsualizar cuántas atenciones suben o bajan a nivel nacional/instituto antes de guardar el cambio.
  3. **Trazabilidad, Comparador SIC y Accesibilidad (Hora 12)**:
     - Vista de **Historial de Versiones y Auditoría** (quién/qué bot aplicó el cambio, fecha, detalle del ajuste).
     - Vista **Planificado vs Ejecutado (SIC)** con semáforo de cumplimiento mensual.
     - **Módulo de Accesibilidad (Rúbrica Dimensión 6)**: botón de **Lectura por Voz (Web Speech API)** que narra el resumen ejecutivo para personas con discapacidad visual, selector de **Alto Contraste**, ajuste de tamaño tipográfico y modo **Lectura Fácil**.

### Fase 4: Pruebas End-to-End, Preparación de Demo y Pitch (Horas 13 a 16)
- **Objetivo**: Asegurar una demo en vivo de 5 minutos sin fallos y un Pitch de 3 minutos alineado 100% con la rúbrica de semifinales.
- **Tareas**:
  1. **Pruebas con Casos de Prueba y Lotes Simulados (Hora 13)**:
     - Preparar 2 archivos Excel de prueba para la demo en vivo de Blue Prism:
       - `Reporte_BUK_Mes_Nuevo.xlsx`: con 1 alta nueva, 1 baja, 1 aumento de jornada (ej. de 30h a 40h), 1 cambio de asignación especial y 2 filas con error intencional (ej. cargo inexistente y horas = 60) para demostrar la detección de errores del bot.
       - `Reporte_SIC_Ejecucion.xlsx`: con atenciones ejecutadas por mes para mostrar alertas de desviación.
  2. **Diseño de la Presentación Pitch (Hora 14)**:
     - Construir las láminas siguiendo las 9 dimensiones de la rúbrica:
       - Lámina 1: **El Problema Real en Teletón** ("Cada hora cuenta": >600 profesionales, 17 institutos, >1M de atenciones, errores `#REF!` en la planilla actual).
       - Lámina 2: **Arquitectura de la Solución (Obligatoria)**: rol central de Blue Prism + Motor Parametrizable + Dashboard y cómo escala a los 17 institutos e integración futura con APIs de BUK y SIC.
       - Lámina 3: **Innovación, Accesibilidad e Impacto Medible** (ahorro de tiempo, trazabilidad total, cumplimiento Ley de Datos Personales y diseño inclusivo).
  3. **Ensayo Cronometrado de Demo (5 min) y Pitch (3 min) (Horas 15–16)**:
     - Ensayar el guion completo al menos 3 veces midiendo tiempos y preparando respuestas técnicas sobre el cálculo (rebaja 44→42h, vacaciones Aysén, prorrateo 2° semestre, factor grupal Santiago/Concepción vs regiones).

---

## 3. Distribución Sugerida de Roles en el Equipo

| Rol | Responsabilidad Principal | Entregable Clave |
|---|---|---|
| **Especialista Blue Prism (RPA)** | Proceso en Blue Prism: lectura de Excel BUK/SIC, validaciones de negocio, detección de novedades (diff), manejo de excepciones y disparo de actualización. | Flujo Blue Prism funcionando en vivo para la demo de 5 min. |
| **Desarrollador Backend / Datos** | Base de datos, ingesta de los Excel actuales (`Dotacion_2026` y `Planificación Asistencial`), motor matemático exacto por mes y API de simulación/versionado. | Motor validado al 100% contra las 471 filas de referencia + endpoints listos. |
| **Desarrollador Frontend / UX Accesible** | Dashboard nacional/instituto, filtros multinivel, mantenedor de parámetros, vista de trazabilidad, comparador SIC y funciones de accesibilidad (TTS, alto contraste). | Aplicación web fluida, intuitiva y accesible. |
| **Líder de Negocio / Pitch & QA** | Validación de casos especiales (multi-instituto, Aysén, feriados regionales, 44→42h), preparación de archivos de prueba para la demo y construcción/ensayo del Pitch. | Presentación con lámina de arquitectura + guion de demo de 5 min y pitch de 3 min. |
