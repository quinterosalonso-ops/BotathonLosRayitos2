# Condiciones, Reglas de Negocio y Especificación Técnica — Desafío Botathon 2026 ("Cada hora cuenta")

Este documento consolida **todas las condiciones obligatorias, deseables, reglas de negocio, tablas de parámetros, validaciones de entrada y fórmulas matemáticas exactas** (verificadas fila por fila contra los archivos `Nac - Base contratos x mes.xlsx` y `Nac - Planificación Asistencial x mes.xlsx`) necesarias para cumplir al 100% con lo solicitado en `Solicitud/Solicitud.md`.

---

## 1. Requisitos del Desafío (`Solicitud.md`)

### 1.1 Requisitos Obligatorios (Mínimo para competir)
1. **Blue Prism en el centro**:
   - Blue Prism debe automatizar una o más etapas críticas del proceso y su aporte debe ser **visible en la demo en vivo**.
   - Rol recomendado para Blue Prism: ingesta desatendida/atendida de los reportes Excel de **BUK** (`Nac - Base contratos x mes.xlsx`) y **SIC**, validación estructural y de negocio, detección automática de diferencias/novedades respecto a la versión vigente (altas, bajas, cambios de horas semanales, cambios de cargo, traslados de instituto, licencias/permisos), registro en bitácora de auditoría y disparo del motor de recálculo.
2. **Incorporación de datos sin copiar y pegar**:
   - Los reportes Excel de BUK y SIC entran a la solución sin intervención manual de datos.
   - Los errores del reporte (RUT/PersonaId inválido, cargo no homologado, centro de costo inexistente, horas negativas o `> 44`, fechas de término anteriores a inicio, duplicados no declarados en excepciones) se detectan, bloquean o alertan y se informan en un reporte/log de errores.
3. **Reglas 100% parametrizables (sin tocar código)**:
   - Porcentajes por cargo y tramo de jornada (`% Directo`, `% Indirecto`, `% NAS`).
   - Minutos estándar por atención (`Directa`: 30, 40, 45, 60 min; `Indirecta`: 30 min; `Pabellón`: 180, 240, 420 min).
   - Catálogo de asignaciones especiales, sus horas semanales, tope de asignaciones y cuáles requieren autorización de Dirección Médica Nacional (DMN) / VP.
   - Factores de atención grupal (`% Grupal` y `N° Pacientes por Grupo` diferenciados para `Santiago - Concepción` vs `Otros Institutos`).
   - Calendario anual: días hábiles por mes, feriados nacionales y regionales por instituto, vacaciones colectivas y vacaciones zona extrema (Aysén/Coyhaique), factores de descuentos (días progresivos, complementarios, administrativos, capacitación, jornada nacional).
   - Porcentaje de **Meta de Gestión** (actualmente `90%`).
4. **Cálculo exacto**:
   - La planificación por profesional y por mes debe coincidir exactamente con el Excel de referencia (`Dotacion_Asistencial` y `Planificacion Total`) en los casos de prueba entregados.
5. **Gestión de cambios durante el año**:
   - Cuando llega un nuevo reporte de BUK con cambios (aumento/disminución de jornada, cambio de cargo, alta de profesional nuevo, baja/término de contrato, cambio de instituto), la solución los identifica automáticamente y **ajusta la planificación desde la fecha/mes en que rigen**, manteniendo intactos los meses previos o generando una nueva versión trazable.
6. **Visibilidad y Dashboards**:
   - Vista para el equipo de planificación central con el **consolidado nacional** y capacidad de filtrar y profundizar por:
     - **Instituto** (los 17 institutos: Arica, Iquique, Antofagasta, Calama, Copiapó, Coquimbo, Valparaíso, O'Higgins, Talca, Concepción, Ñuble, Temuco, Puerto Montt, Coyhaique, Valdivia, Castro, Santiago).
     - **Tipo de atención** (Directa individual, Directa grupal, Indirecta, No Asistencial / NAS, Pabellón, LOP).
     - **Especialidad / Área / Unidad** (`FI`, `KI`, `TO`, `ED`, `PC`, `FO`, `EN`, `AS`, `UTAC`, `UEDA`, `ULAB`, `UTA`, `NU`, `OD`, `NE`, `OR`, `UR`, `PE`, `HO`, `LOP`).
     - **Cargo** y **Profesional** (`PersonaId`).
     - **Mes / Período** (`Enero–Abril` 44h vs `Mayo–Diciembre` 42h, y mes a mes `Ene..Dic`).

### 1.2 Indicadores Obligatorios en el Dashboard
1. **`# Meta de Gestión Asistencial`**:
   $$\text{Meta de Gestión Asistencial} = \text{Atenciones Planificadas} \times 90\%$$
   *(Nota: a nivel de fila mensual por profesional se aplica `ROUNDDOWN(Atenciones_Mes * 0.9, 0)` y en las tablas de resumen nacional también se expone el `90%` directo sobre la sumatoria anual).*
2. **`% FOJA` (Factor de Ocupación de Jornada Asistencial)**:
   $$\%\text{FOJA} = \frac{\text{Horas mensuales disponibles para atención directa}}{\text{Horas mensuales contratadas}}$$
3. **Tabla Resumen Consolidada Nacional por Instituto** ("Médico terapéutico - Socio educativo"):
   - Columnas exigidas en `Solicitud.md`:
     1. `Instituto`
     2. `Suma de Horas Semanales`
     3. `Tiempo (Hrs/Semana) Atención Directa` (formato `[h]:mm` o decimal)
     4. `Tiempo (Hrs/Semana) Atención Indirecta`
     5. `Tiempo (Hrs/Semana) Actividad no asistencial - NAS`
     6. `Atención Anual Directa Grupal e Individual`
     7. `Atención Anual Indirecta`
     8. `Total Atenciones en Planificación por Año`
     9. `Meta Gestión al 90%`

### 1.3 Requisitos Deseables (Suman puntos directos)
- **Trazabilidad de cambios y control de versiones**: historial inmutable de quién cambió qué, cuándo, valor anterior vs valor nuevo y justificación/autorización.
- **Comparación Planificado vs Realizado (SIC)**: cruce por profesional y mes con alertas visuales (semáforo) de desviación respecto a la Meta de Gestión (90%) y Planificación (100%).
- **Soporte completo de casos especiales**: profesionales multi-instituto (`Excepciones`), reemplazos (vínculo con titular), vacantes/cargos nuevos y licencias extensas.
- **Flujo de aprobación de asignaciones especiales**: bandeja para autorizar asignaciones que requieren V°B° de DMN y VP o cambios en `% Atención Grupal`.
- **Simulador de impacto ("What-if")**: simular el impacto en horas y atenciones de cambiar un parámetro (ej. bajar jornada de 42 a 40h, cambiar `% Directo` o `% Meta`) antes de aplicarlo oficialmente.
- **Distribución automática de reportes**: envío/generación automatizada (vía Blue Prism) de reportes por instituto para las jefaturas locales.
- **Accesibilidad e inclusión (Rúbrica Dimensión 6)**: interfaz con alto contraste, tamaño de fuente ajustable, etiquetas ARIA para lector de pantalla, lectura fácil / resumen por voz (TTS) y navegación por teclado.
- **Protección de datos personales (Normativa chilena Ley 19.628 / 21.719)**: uso de `PersonaId` anonimizado, control de acceso por rol, mínimo privilegio y log de auditoría.

---

## 2. Estructura de los Archivos Excel Analizados

### 2.1 Archivo 1: `Solicitud/Nac - Base contratos x mes.xlsx` (Reporte BUK Simulado)
- **Hoja `Dotacion_2026`**: contiene **3.123 registros activos** (filas 3 a 3125; las filas 3126 a 4003 están en blanco) correspondientes a **462 profesionales únicos** (`PersonaId`) a lo largo de **7 períodos mensuales reportados** (`01-01-2026` al `01-07-2026`).
- **Columnas de `Dotacion_2026`**:
  | Col | Nombre Columna | Descripción / Uso en Reglas |
  |:---:|---|---|
  | `A` | `IT` | Código numérico del Instituto (`11`=Ari, `12`=Iqu, `21`=Ant, `22`=Cal, `31`=Cop, `41`=Coq, `51`=Val, `61`=OHi, `71`=Tal, `81`=Con, `91`=Tem, `101`=PMo, `102`=Cas, `111`=Coy, `131`=San, `141`=Vdv, `161`=Nub) |
  | `B` | `Instituto` | Abreviatura de 3 letras del instituto (`Ari`, `Iqu`, `Ant`, etc.) |
  | `C` | `PersonaId` | Identificador único anonimizado del profesional (entero) |
  | `D` | `Área` | Nombre de la división/instituto en BUK (ej. `Puerto Montt`, `Santiago`) |
  | `E` | `Nombre Sub-área` | Descripción del centro de costo / unidad (ej. `Fisiatría`, `Kinesiología`) |
  | `F` | `Centro de Costo` | Código numérico de C.Costo (`100200` a `301601`) |
  | `G` | `Fecha de Nacimiento` | Fecha serial Excel (se usa para obtener `Mes cumpleaños = MONTH(G)`) |
  | `H` | `Periodo` | Primer día del mes reportado: `46023` (Ene), `46054` (Feb), `46082` (Mar), `46113` (Abr), `46143` (May), `46174` (Jun), `46204` (Jul) |
  | `I` | `Tipo de Contrato` | `Indefinido` o `Plazo fijo` |
  | `J` | `Tipo_Trab_161225` | Estado en corte anterior (`Titular` o vacío si es contratación posterior) |
  | `K` | `Fecha Ingreso Compañía`| Fecha de inicio de contrato |
  | `L` | `Fecha Término` | Fecha de término de trabajo (vacío si es indefinido/vigente) |
  | `M` | `Horas Semanales` | Horas contratadas por semana en ese mes (`Periodo`) |
  | `N` | `Cargo` | Nombre del cargo en BUK (ej. `KINESIOLOGO`, `MEDICO FISIATRA`) |
  | `O` | `Tipo_Trabajador` | `Titular` o `Reemplazo` |
  | `P` | `Días de la Jornada` | Cadena con iniciales de días trabajados (ej. `lmwjv`=5 días, `lmwj`=4, `lw`=2) |
  | `Q` | `Reg_SIC` | `D - SIC` (registra en SIC) o `F - SIC` (fuera de SIC) |
  | `R` | `Cod_Esp` | Código de especialidad/área (`FI`, `KI`, `TO`, `ED`, `PC`, `FO`, `EN`, `AS`, `UTAC`, `UEDA`, `ULAB`, `UTA`, `NU`, `OD`, `NE`, `OR`, `UR`, `PE`, `HO`) |
  | `S` | `Ndias` | Número de días de la jornada semanal (`1` a `5`, equivalente al largo de `P`) |

### 2.2 Archivo 2: `Solicitud/Nac - Planificación Asistencial x mes.xlsx` (Modelo de Referencia)
Contiene **8 hojas**:
1. **`Parámetros`**:
   - Tabla de homologación `(Cod_Esp, Cargo Homologado, C.Costo, Descripcion C.Costo)`.
   - Tabla de **Factor Grupal** (`Cod_Esp`, `% Grupal`, `N° Pacientes San - Con`, `N° Pacientes Otros IT`).
   - Calendario de inhábiles 2026 (sábados, domingos, feriados colectivos de febrero y feriados legales nacionales/regionales).
   - Tabla resumen mensual de días y semanas hábiles (`AD3:AD14`: semanas sin vacaciones; `Z3:Z14`: días hábiles con vacaciones).
   - Parámetros de descuentos adicionales (Vacaciones 15 días, Adicionales Staff 5 días / Coyhaique 10 días, Administrativos y 1/2 jornada 3,5 días, Cumpleaños 1 día, Otros 1 día).
   - Tabla de **Jornada Nacional por Área** (`AB34:AC44`).
   - Tablas maestras `Institutos` (`AM1:AO26`), `Cargo_SIC` (`AR1:AU70`), `Especialidades` (`AX1:AY25`).
2. **`Calendario 2026`**:
   - Vista calendario anual 2026 con conteo de sábados (52), domingos (52), feriados en día hábil (10), total inhábiles (114), vacaciones colectivas febrero (15 días hábiles: del 2 al 20 de febrero 2026), y listado de los 18 días festivos nacionales y regionales.
3. **`Distribucion_NAC`**:
   - Matriz principal parametrizable por `Esp`, `Cargo`, `Unidad`, `STD Minutos Atención (Directa, Indirecta)` y porcentajes `(% Dir, % Ind, % NAS)` para los **3 tramos de jornada**:
     - Tramo 1: `NAC >= 40 HRAS`
     - Tramo 2: `NAC 40 < = 22 HRAS` (entre 22 y menos de 40 horas: `22 <= H/S < 40`)
     - Tramo 3: `NAC < 22 HRAS` (menos de 22 horas: `< 22`)
   - Sub-matriz de cargos quirúrgicos (`Pabellón`: Neurocirujano 420 min, Ortopedistas 240 min, Urólogo 180 min; `60% Pabellón`, `30% Directo`, `10% NAS`).
   - Catálogo de **Asignaciones Especiales** (columnas `R:T`) con sus horas semanales y condición de autorización.
4. **`Resumen distribucion NAS`**:
   - Tablas dinámicas de promedio de `% NAS`, `% Directo`, `% Indirecto` por instituto (con y sin asignaciones especiales), resumen de pabellones quirúrgicos y tabla base de producción LOP (`Laboratorio órtesis y prótesis`, total 49.395 / 49.895 atenciones).
5. **`Planificacion Total`**:
   - Consolidado nacional por instituto que cruza `Dotacion_Asistencial`, `LOP`, `Proyección pacientes únicos 2025`, `Intensidad Total`, `Intensidad Directa` y `Meta al 90%`.
6. **`LOP`**:
   - Resumen de atenciones anuales planificadas de Laboratorio de Órtesis y Prótesis por instituto.
7. **`Excepciones`**:
   - Tabla `Excepcion` (`A2:X19`) con los profesionales multi-instituto que reparten su jornada entre dos o más institutos o cambian de instituto durante el año, detallando sus horas exactas mes a mes (`01-01-2026` a `01-12-2026`) y días por semana (`D/S`) en cada instituto.
8. **`Dotacion_Asistencial`**:
   - Tabla principal `Dot_Asistencial` (`A4:GW476`) con **471 filas de planificación** (correspondientes a los 462 `PersonaId` de `Dotacion_2026` más 9 filas adicionales por desglose multi-instituto/cambio de contrato marcadas con `Repetido = "R"`) y **205 columnas** (versión compactada de la estructura completa de **236 columnas** preservada en el caché del libro, que separa explícitamente las columnas del período `ene - abr` y `may - dic`).

---

## 3. Lógica de Negocio y Fórmulas Exactas Paso a Paso

A continuación se documentan las reglas y fórmulas exactas verificadas contra las 471 filas de `Dotacion_Asistencial` y los 3.123 registros de `Dotacion_2026`.

### Paso 1: Vigencia, Horas Mensuales (`H/S ene..dic`) y Promedio Ponderado
1. **Cruce con `Excepciones` (Multi-Instituto)**:
   - Si el `PersonaId` está en la hoja `Excepciones`, se genera una fila por cada instituto asignado al profesional con su identificador `Excep_Id = N°_ocurrencia & "-" & PersonaId` (ej. `2-115444` en Iquique y `1-115444` en Coyhaique).
   - La fila principal lleva `Repetido = ""` y las filas secundarias llevan `Repetido = "R"` (clave para **no duplicar** días de cumpleaños ni beneficios únicos del trabajador).
   - Sus horas semanales de enero a diciembre (`H/S ene` a `H/S dic`) se toman directamente de las columnas `K` a `V` de `Excepciones`, y sus días semanales `Z` de la columna `W` (`D/S`).
2. **Horas Mensuales desde BUK (`Dotacion_2026`)**:
   - Para los profesionales sin excepción manual, se cruzan sus registros mensuales de `Dotacion_2026` por `PersonaId` e `IT`:
     - Meses reportados en BUK (`ene` a `jul`, códigos `46023` a `46204`): toman `Horas Semanales` (`col M`) si el contrato está vigente en ese mes; si no tiene registro en ese mes, es `0`.
     - **Proyección agosto–diciembre (`ago` a `dic`)**: si el profesional sigue vigente en el último mes reportado (`jul = 46204`) y su `Fecha Término Trabajo` (`col H`) está vacía o es posterior a cada mes, se proyectan las horas de julio (`H/S jul`) para `ago..dic`. Si su `Perido Fin Jornada` (`col U`) termina antes (ej. `46326` = 31/10/2026), los meses posteriores a `U` quedan en `0`.
   - **Cambio de contrato dentro del mismo instituto**: si un profesional cambia de contrato/jornada con corte explícito de fila (ej. `PersonaId = 137356` en Calama o `72190` en Coquimbo), la primera fila cubre el primer período (`T` a `U`, con `W = 0`) y la segunda fila (marcada `Repetido = "R"`) cubre el nuevo período hasta diciembre.
3. **Vigencia en Diciembre (`col V`) y Horas Diciembre (`col W`)**:
   - `H/S dic` (`col W`): horas contratadas en diciembre (`01-12-2026`). Si el contrato terminó antes de diciembre, `W = 0`.
   - `Vigencia H/S 12` (`col V`):
     $$V = \begin{cases} 1 & \text{si } W > 0 \\ 0 & \text{si } W = 0 \end{cases}$$
4. **Promedio Ponderado de Horas Semanales (`col X`)**:
   - En el modelo completo (caché 236 columnas): promedio de los meses con horas `> 0`, truncado a 1 decimal:
     $$X_{\text{dec}} = \text{ROUNDDOWN}(\text{AVERAGEIF}(\text{H/S\_ene}:\text{H/S\_dic}, ">0"), 1)$$
   - En la tabla `Dotacion_Asistencial` (205 columnas): truncado a entero:
     $$X = \text{ROUNDDOWN}(\text{AVERAGEIF}(\text{H/S\_ene}:\text{H/S\_dic}, ">0"), 0)$$
   - *Caso especial Cambio Legal 44h → 42h*: para quienes tienen `44h` en `ene–abr` (4 meses) y `42h` en `may–dic` (8 meses), el promedio exacto es $42{,}666\dots \rightarrow 42{,}6$ y en `Dotacion_Asistencial` su valor nominal de referencia se fija en `42`.
5. **Detección de Cambio de Jornada (`col Y`)**:
   - No se marca como cambio de jornada la rebaja legal estándar de `44h` (ene–abr) a `42h` (may–dic) cuando el profesional estuvo vigente todo el año.
   - Se marca **`"Aumento"`** cuando:
     - El profesional ingresa después de enero (`Periodo Inicio Jornada > 01-01-2026`), o
     - Sus horas semanales suben durante el año (ej. de `20h` en ene–feb a `33h` en mar–dic, o de `38h` a `40h`).
   - Se marca **`"Disminución"`** cuando:
     - El profesional termina su vigencia antes de diciembre (`W = 0`), o
     - Sus horas semanales bajan durante el año por motivo distinto a la sola rebaja legal 44→42 (ej. de `28h` a `11h`, de `40h` a `35h`, de `27h` a `26h`).
6. **Horas/Día Promedio (`col AA`)**:
   - Con `Z` = `N° Días Semana Contrato` (`Ndias` en BUK, de 1 a 5):
     $$AA = \text{ROUNDDOWN}\left(\frac{X}{Z},\; 1\right)$$

---

### Paso 2: Tramo de Jornada y Distribución Base por Cargo (`AI`, `AJ`, `AK`)
1. **Clasificación en Tramo de Jornada** (según `X` = `Prom_Pond Horas Semanales`):
   - **Tramo 1 (`>= 40 HRAS`)**: $X \ge 40$
   - **Tramo 2 (`40 < = 22 HRAS`)**: $22 \le X < 40$
   - **Tramo 3 (`< 22 HRAS`)**: $X < 22$
2. **Búsqueda en `Distribucion_NAC`**:
   - Con el `Cargo` (homologado) y el tramo de jornada, se obtienen de `Distribucion_NAC`:
     - `AI` = `% NAS` base
     - `AJ` = `% Directo` base
     - `AK` = `% Indirecto` base
     *(Cumpliendo siempre $\text{AI} + \text{AJ} + \text{AK} = 100\%$).*
   - También se obtienen los minutos estándar del cargo:
     - `AY` = `Estándar Minutos Atención Directa` (`30`, `40`, `45` o `60` min; salvo excepción específica en fila como Lokomat = `60` min).
     - `AZ` = `Estándar Minutos Atención Indirecta` (`30` min para todos los cargos).

---

### Paso 3: Colación, Horas Disponibles antes de NAS y Desglose Interno de NAS (`AB` a `AH`)
1. **Colación en Contrato (`AB` y `AC`)**:
   - `AB` (`Colación en Contrato`): `"S"` o `"N"`.
   - `AC` (`Horas Colación en Contrato`): si `AB = "N"` es `0`; si `AB = "S"` se descuentan las horas semanales de colación pactadas en contrato (típicamente `2,5` h/sem para 5 días, `2,0` h/sem para 4 días, `0,5` h por día trabajado).
2. **Horas Semanales Disponibles antes de NAS (`AD`)**:
   - Para cada mes $m \in \{1..12\}$:
     $$\text{H\_Disp\_Antes\_NAS}_m = \text{H/S}_m - AC$$
   - En la columna `AD` (`dic`): $AD = W - AC$.
   - Para el período 1 (`ene - abr`): se usa el promedio de $\text{H\_Disp\_Antes\_NAS}_m$ de los meses activos en `ene..abr`.
   - Para el período 2 (`may - dic`): se usa el promedio de $\text{H\_Disp\_Antes\_NAS}_m$ de los meses activos en `may..dic`.
3. **Desglose Informativo de Horas NAS (`AE`, `AF`, `AG`, `AH`)**:
   - `AH` (`H/Sem NAS Total` base):
     $$AH = X \times AI$$
   - `AE` (`H/Sem Pausa Saludable`):
     $$AE = \begin{cases} 0{,}10 \times X & \text{si } X \ge 40 \\ 0{,}06 \times X & \text{si } X < 40 \end{cases}$$
     *(Verificado con el comentario oficial en celda `AE4` y las 471 filas).*
   - `AF` (`H/Sem Beneficio Colación 15 minutos`):
     - `0,25` horas (15 min) por día trabajado (`1,25` h/sem para 5 días, `1,0` h/sem para 4 días, etc.) si el trabajador tiene vigente el beneficio (`Benef_Colacion15Min = "S"`); de lo contrario `0`.
   - `AG` (`H/Sem Otras NAS`):
     $$AG = AH - AE - AF$$

---

### Paso 4: Asignaciones Especiales y Porcentajes Finales Asistenciales (`AL` a `AX`)
1. **Asignaciones Especiales (`AL` a `AS`)**:
   - `AL` / `AM`: `Asignación 1` y `Hrs. Asignación 1`
   - `AN` / `AO`: `Asignación 2` y `Hrs. Asignación 2`
   - `AP` / `AQ`: `Por Autorizar DMN y VP (Asignación 3)` y `Hrs. Asignación 3`
   - `AR`: `Hrs. Asignación 4 (Reunión Sindical)` (`0,25` h/sem si aplica, o `0`)
   - **Regla de negocio**: máximo 2 asignaciones estándar para profesionales con más de 30 H/S; una 3ª asignación o asignaciones especiales fuera de estándar requieren autorización de Dirección Médica Nacional (DMN) y VP (`col AP`).
   - `AS` (`Total Hrs. Asignación`):
     $$AS = AM + AO + AQ + AR$$
2. **Recálculo de Porcentajes Finales por Período (`ene - abr` y `may - dic`)**:
   Sea $H_{\text{disp}, p}$ el promedio de `H/Sem Disponible antes NAS` en el período $p \in \{\text{ene-abr}, \text{may-dic}\}$:
   - `%NAS + Asig. Especiales` (`col AT` para `may-dic`):
     $$\%\text{NAS\_Final}_p = \frac{AH + AS}{H_{\text{disp}, p}} = \frac{X \cdot AI + AS}{H_{\text{disp}, p}}$$
   - `% Asistencial` (`col AW`):
     $$\%\text{Asistencial}_p = 1 - \%\text{NAS\_Final}_p$$
   - **Reparto Proporcional entre Directa e Indirecta** (`col AU` y `col AV`):
     Las horas disponibles para atención se reparten manteniendo la proporción original entre `%Directo` (`AJ`) y `%Indirecto` (`AK`):
     $$\%\text{Directo\_Final}_p = \%\text{Asistencial}_p \times \left(\frac{AJ}{AJ + AK}\right)$$
     $$\%\text{Indirecto\_Final}_p = \%\text{Asistencial}_p \times \left(\frac{AK}{AJ + AK}\right)$$
   - **Horas Semanales Disponibles para Atención (`col AX`)**:
     $$\text{Hrs\_Disp\_Ate}_p = H_{\text{disp}, p} \times \%\text{Asistencial}_p = H_{\text{disp}, p} - (AH + AS)$$
     *(En el archivo Excel se almacena como fracción de día de 24 horas: $\text{AX} = \text{Hrs\_Disp\_Ate}_p / 24$).*

---

### Paso 5: Conversión de Horas a Atenciones Semanales Individuales y Grupales (`BA` a `BH`)
Para cada período $p \in \{\text{ene-abr}, \text{may-dic}\}$:
1. **N° Atenciones Directas Estándar por Semana (`BA`)**:
   $$\text{Ate\_Dir\_Sem}_p = \text{ROUNDDOWN}\left(\frac{H_{\text{disp}, p} \times \%\text{Directo\_Final}_p \times 60}{\text{AY}},\; 0\right)$$
2. **N° Atenciones Indirectas (30 min) por Semana (`BB`)**:
   $$\text{Ate\_Ind\_Sem}_p = \text{ROUNDDOWN}\left(\frac{H_{\text{disp}, p} \times \%\text{Indirecto\_Final}_p \times 60}{\text{AZ}},\; 0\right)$$
3. **Expansión por Atención Grupal (`BC`, `BD`, `BE`, `BF`)**:
   - Aplica a las especialidades con atención grupal según tabla de `Parámetros`:
     | Código (`Cod_Esp`) | Área / Unidad | `% Atención Grupal Estándar` (`BC`) | `N° Pacientes por Grupo` en **Santiago (`San`) y Concepción (`Con`)** | `N° Pacientes por Grupo` en **Otros Institutos** |
     |:---:|---|:---:|:---:|:---:|
     | `AS` | Servicio Social | `10%` (`0,10`) | `3,0` | `1,5` |
     | `ED` | Educación | `10%` (`0,10`) | `3,0` | `1,5` |
     | `ULAB` | Unidad Laboral | `10%` (`0,10`) | `3,0` | `1,5` |
     | `UEDA` | Deporte Adaptado | `55%` (`0,55`) | `6,0` | `3,0` |
     | `UTAC` | Terapias Artísticas (Arte) | `55%` (`0,55`) | `6,0` | `3,0` |
     | *Resto* | *Demás áreas* | `0%` | `0` | `0` |
   - Si existe un valor autorizado en `BD` (`Cambio %Ate. Grupal V°B° DMN y VP`), se usa `BD` en lugar de `BC`:
     $$\%\text{Grp} = \begin{cases} BD & \text{si } BD \neq \text{vacío} \\ BC & \text{en caso contrario} \end{cases}$$
   - **Atención Semanal Directa + Grupal (`col BF`)**:
     $$\text{Ate\_Dir\_Grupal\_Sem}_p = \begin{cases} \text{Ate\_Dir\_Sem}_p \times (1 - \%\text{Grp}) + \left(\text{Ate\_Dir\_Sem}_p \times \%\text{Grp} \times \text{BE}\right) & \text{si } \%\text{Grp} > 0 \text{ y } \text{BE} > 0 \\ \text{Ate\_Dir\_Sem}_p & \text{en caso contrario} \end{cases}$$
     *(Nota: no se trunca en `BF`; puede tener decimales como `34,65` y se trunca recién al multiplicar por las semanas hábiles de cada mes).*
4. **Atención Semanal Indirecta Total (`BG`) y Total Semanal (`BH`)**:
   $$\text{BG}_p = \text{Ate\_Ind\_Sem}_p, \qquad \text{BH}_p = \text{BF}_p + \text{BG}_p$$

---

### Paso 6: Calendario 2026, Ausentismos y Semanas Hábiles por Mes (`BI` a `EM`)
1. **Semanas Hábiles Base del Calendario 2026 (`BI` a `BT`)**:
   Tomadas de `Parámetros!AD3:AD14` ($\text{Semana}_m = \text{ROUND}(\text{Días\_Hábiles\_Sin\_Vacaciones}_m / 5,\; 1)$):
   | Mes | Ene (`BI`) | Feb (`BJ`) | Mar (`BK`) | Abr (`BL`) | May (`BM`) | Jun (`BN`) | Jul (`BO`) | Ago (`BP`) | Sep (`BQ`) | Oct (`BR`) | Nov (`BS`) | Dic (`BT`) | Total Año (`BU`) |
   |---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
   | **Días calendario** | 31 | 28 | 31 | 30 | 31 | 30 | 31 | 31 | 30 | 31 | 30 | 31 | **365** |
   | **Días hábiles (sin vac.)** | 21 | 20 | 22 | 21 | 19 | 21 | 22 | 21 | 21 | 21 | 21 | 21 | **251** |
   | **Semanas base (`AD`)** | **4,2** | **4,0** | **4,4** | **4,2** | **3,8** | **4,2** | **4,4** | **4,2** | **4,2** | **4,2** | **4,2** | **4,2** | **50,2** |
   - Si el profesional **no tiene contrato vigente** en el mes $m$ ($\text{H/S}_m = 0$), su semana base en ese mes es `0`.
2. **Ausentismos Mensuales Directos (en Semanas)**:
   - `Lic1..Lic12` (`BV..CG`): semanas de licencia médica en el mes $m$. (Si `Licencia Extensa = "Si"`, puede cubrir meses completos dejando `SemHab_m = 0`).
   - `Vac1..Vac12` (`CH..CS`): semanas de vacaciones legales colectivas/individuales en el mes $m$:
     - **Regla General Nacional**: `3,0` semanas (15 días hábiles) en febrero (`Vac2 = 3` o `4` si incluye adicionales/progresivos en feb).
     - **Regla Especial Zona Extrema Aysén / Coyhaique (`IT = 111`, `Coy`)**: `4,0` semanas (20 días hábiles) repartidas como `Vac1 = 0,4` semanas en enero y `Vac2 = 3,6` semanas en febrero.
   - `Perm1..Perm12` (`CT..DE`): semanas de permisos por mes (horas de permiso bloqueadas divididas por la jornada semanal).
3. **Descuentos Anuales Prorrateados en el 2° Semestre (`DL` a `DY`)**:
   - Columnas de días de descuento (`DL` a `DW`):
     - `DL` (`Feriado legal adicional`): `5` días hábiles (beneficio invierno Teletón) o `0` según vigencia/contrato.
     - `DM` (`Días administrativos y 1/2 jornada familiar`): `3,5` o `4,0` días (o `0`).
     - `DN` (`Plan anual capacitación por cargo`): `1,5` o `2,0` días (o `0`).
     - `DO` (`HEXA - CTR`): días compensatorios (`0`, `1`, `2`, etc.).
     - `DP` (`Jornada nacional`): según área (`AS, ED, PC, UTAC, FI, ULAB = 1,5`; `UEDA = 1,0`; `UTA, KI, TO = 0,5`; u `0`).
     - `DQ` (`Días Progresivos`): días progresivos ponderados por factor `0,7` (o `0,8` según parámetro en `DQ2`).
     - `DR` (`Días Complementario`): ponderados por factor `0,8` (`DR2`).
     - `DS` (`Ley 21530` Descanso Reparatorio): ponderados por factor `1,0` (`DS2`).
     - `DT` (`Vacaciones Legales Pendientes Periodos Anteriores`): factor `1,0` (`DT2`).
     - `DU` (`Jornadas Directiva Sindical`), `DV` (`Ajuste por diferencia en días contrato`), `DW` (`Lactancia Equiv_días/año`).
   - `DX` (`Meses 2do semestre`): cantidad de meses entre julio y diciembre (`BO..BT`) en que el profesional tiene semanas base $> 0$:
     $$DX = \text{COUNTIF}(\text{BO}:\text{BT}, ">0")$$
   - `DY` (`Semanas/mes a prorratear 2do semestre`): convierte la suma de días `DL..DW` a semanas ($\div 5$) y las prorratea en los $DX$ meses activos del 2° semestre:
     $$DY = \begin{cases} \dfrac{\sum (\text{DL}:\text{DW})}{5 \times DX} & \text{si } DX > 0 \\ 0 & \text{si } DX = 0 \end{cases}$$
4. **Descuentos en Meses Específicos (`DF` a `DK`)**:
   - `DF` (`Feriado legal zona extrema ene'26`): se descuenta en enero (`DF / 5`).
   - `DG` (`Feriado legal zona extrema feb'26`): se descuenta en febrero (`DG / 5`).
   - `DH` (`Mes cumpleaños`):
     $$DH = \begin{cases} 0 & \text{si } \text{Repetido (col L)} = \text{"R"} \\ \text{MONTH}(\text{Fecha Nacimiento}) & \text{en caso contrario} \end{cases}$$
   - `DI` (`Días de cumpleaños`): `1,0` o `0,5` días (o `0` si `Repetido = "R"`). Se descuenta `DI / 5` semanas en el mes $m = DH$.
   - `DJ` (`Vísperas Fiestas Patrias`): `0,5` días (`0,1` semanas). Se descuenta en **septiembre ($m = 9$)**.
   - `DK` (`Vísperas Navidad y Año Nuevo`): `1,0` día (`0,2` semanas). Se descuenta en **diciembre ($m = 12$)**.
5. **Fórmula Exacta de Semanas Hábiles Efectivas por Mes (`SemHab1..SemHab12`, cols `DZ..EK`)**:
   Para cada mes $m \in \{1..12\}$, si la semana calendario base $\text{Cal}_m = 0$, entonces $\text{SemHab}_m = 0$. Si $\text{Cal}_m > 0$:
   $$\text{Base}_m = \text{Cal}_m - \text{Lic}_m - \text{Vac}_m - \text{Perm}_m - \begin{cases} \text{DF}/5 & m=1 \\ \text{DG}/5 & m=2 \\ 0 & \text{otro} \end{cases} - \begin{cases} DY & m \ge 7 \\ 0 & m \le 6 \end{cases} - \begin{cases} \text{DI}/5 & m = DH \\ 0 & m \neq DH \end{cases} - \begin{cases} \text{DJ}/5 & m = 9 \\ \text{DK}/5 & m = 12 \\ 0 & \text{otro} \end{cases}$$
   - Para $m \in \{1, 2, 3, 4, 6, 8, 9, 10, 11, 12\}$:
     $$\text{SemHab}_m = \max(0,\; \text{Base}_m)$$
   - Para $m \in \{5, 7\}$ (mayo y julio, donde aplica redondeo hacia arriba al primer decimal en la planilla original):
     $$\text{SemHab}_m = \text{ROUNDUP}(\max(0,\; \text{Base}_m),\; 1)$$
   - **Total Semanas Hábiles en el Año (`col EL`)**:
     $$EL = \sum_{m=1}^{12} \text{SemHab}_m$$
   - **Total Días Hábiles Efectivos (`col EM`)**:
     $$EM = EL \times Z$$

---

### Paso 7: Proyección Mensual y Anual de Atenciones y Meta de Gestión (`EN` a `GV`)
1. **Atenciones Planificadas Directas por Mes (`PlanAteDir1..12`, cols `EN..EY`)**:
   - Para enero–abril ($m \in \{1..4\}$), usa la tasa semanal directa+grupal del período `ene - abr`:
     $$\text{PlanAteDir}_m = \text{ROUNDDOWN}\left(\text{SemHab}_m \times \text{BF}_{\text{ene-abr}},\; 0\right)$$
   - Para mayo–diciembre ($m \in \{5..12\}$), usa la tasa semanal directa+grupal del período `may - dic`:
     $$\text{PlanAteDir}_m = \text{ROUNDDOWN}\left(\text{SemHab}_m \times \text{BF}_{\text{may-dic}},\; 0\right)$$
   - **Total Plan Atenciones Directas Anual (`EZ` y `GN`)**:
     $$EZ = GN = \sum_{m=1}^{12} \text{PlanAteDir}_m$$
2. **Atenciones Planificadas Directas + Indirectas por Mes (`PlanAteDir+Ind1..12`, cols `FA..FL`)**:
   - Para $m \in \{1..4\}$:
     $$\text{PlanAteDir+Ind}_m = \text{ROUNDDOWN}\left(\text{SemHab}_m \times \text{BH}_{\text{ene-abr}},\; 0\right)$$
   - Para $m \in \{5..12\}$:
     $$\text{PlanAteDir+Ind}_m = \text{ROUNDDOWN}\left(\text{SemHab}_m \times \text{BH}_{\text{may-dic}},\; 0\right)$$
   - **Total Plan Atenciones Dir+Indir Anual (`FM` y `GP`)**:
     $$FM = GP = \sum_{m=1}^{12} \text{PlanAteDir+Ind}_m$$
   - **Plan Atención Anual Indirecta (`GO`)**:
     $$GO = GP - GN$$
3. **Meta de Gestión Mensual y Anual al 90% (`FN` a `GM` y `GT` a `GV`)**:
   Sea $\alpha_{\text{meta}} = 0{,}90$ (parámetro en celda `FZ3` / `Planificacion Total!L3`):
   - **Meta Mensual Directa (`MG_AteDir1..12`, cols `FN..FY`)**:
     $$\text{MG\_AteDir}_m = \text{ROUNDDOWN}\left(\text{PlanAteDir}_m \times \alpha_{\text{meta}},\; 0\right)$$
   - **Total MG Directa Anual (`FZ` y `GT`)**:
     $$FZ = GT = \sum_{m=1}^{12} \text{MG\_AteDir}_m$$
   - **Meta Mensual Directa + Indirecta (`MG_AteDir+Ind1..12`, cols `GA..GL`)**:
     $$\text{MG\_AteDir+Ind}_m = \text{ROUNDDOWN}\left(\text{PlanAteDir+Ind}_m \times \alpha_{\text{meta}},\; 0\right)$$
   - **Total MG Anual (`GM` y `GV`)**:
     $$GM = GV = \sum_{m=1}^{12} \text{MG\_AteDir+Ind}_m$$
   - **MG Atención Anual Indirecta (`GU`)**:
     $$GU = GV - GT$$
4. **Tiempo Anual en Atenciones (`GQ`, `GR`, `GS`)**:
   - `GS` (`Total Tiempo en Atenciones por Año`):
     $$GS = EL \times AX$$
   - `GQ` (`Tiempo en Atención Anual Directa`):
     $$GQ = GS \times \left(\frac{BF}{BH}\right)$$
   - `GR` (`Tiempo en Atención Anual Indirecta`):
     $$GR = GS \times \left(\frac{BG}{BH}\right)$$

---

## 4. Catálogos Oficiales Extraídos para Parametrización en Base de Datos

### 4.1 Catálogo de Institutos Teletón (`IT`)
| Código `IT` | Abreviatura (`Inst`) | Nombre Instituto | Reglas Regionales Especiales |
|:---:|:---:|---|---|
| `11` | `Ari` | Arica | Feriado regional: **07 de junio** (Asalto y Toma del Morro de Arica) |
| `12` | `Iqu` | Iquique | Feriado regional: **10 de agosto** (San Lorenzo de Tarapacá) |
| `21` | `Ant` | Antofagasta | Zona Norte |
| `22` | `Cal` | Calama | Zona Norte |
| `31` | `Cop` | Copiapó | Feriado regional: **10 de agosto** (San Lorenzo - Atacama) |
| `41` | `Coq` | Coquimbo | Feriado regional: **20 de septiembre** (Fiesta de la Pampilla) |
| `51` | `Val` | Valparaíso | Cuenta con Pabellón Quirúrgico (`OR`, `UR`) |
| `61` | `OHi` | O'Higgins | Rancagua |
| `71` | `Tal` | Talca | Maule |
| `81` | `Con` | Concepción | **Factor Grupal Mayor** (`3` y `6` pac/grupo) + Pabellón (`OR`, `UR`) |
| `91` | `Tem` | Temuco | Araucanía |
| `101` | `PMo` | Puerto Montt | Cuenta con Pabellón Quirúrgico (`OR`) |
| `102` | `Cas` | Castro | Chiloé |
| `111` | `Coy` | Coyhaique | **Zona Extrema Aysén**: 20 días hábiles vacaciones (`0,4` sem ene + `3,6` sem feb) + 5 días adic. |
| `131` | `San` | Santiago | **Factor Grupal Mayor** (`3` y `6` pac/grupo) + Alta Complejidad (`AC`) + Pabellón (`NE`, `OR`, `UR`) |
| `141` | `Vdv` | Valdivia | Los Ríos |
| `161` | `Nub` | Ñuble | Feriado regional: **20 de agosto** (Natalicio de Bernardo O'Higgins - Chillán) |

### 4.2 Catálogo de Asignaciones Especiales (`Distribucion_NAC` cols `R:T`)
| Tipo de Asignación Especial | Horas Semanales Estándar | Condición / Autorización |
|---|:---:|---|
| `Líder Equipo` | `2,0` | Semanales para actividades de rol de líder |
| `Investigación` | `2,0` | Semanales una vez aprobado por CEC |
| `Responsable técnico` | `1,0` | Hora semana por programa de voluntariado |
| `Casa de acogida Santiago` | `1,0` | Semanal |
| `Coordinación WeeFIM (Local) por IT` | `1,0` | Semanal |
| `GMFM (Local) Por IT` | `1,0` | Semanal |
| `Tutor docente` | `1,0` | Semanal |
| `Coordinador docente` | `1,0` | Semanal |
| `Coordinación Tetca/tetco/ Uta(Local) por IT` | `1,0` | Semanal |
| `Comité paritario` | `0,5` | Semanal |
| `Mesas Tecnicas SSYC` | `0,5` | Semanal |
| `Mesas Tecnicas SDM` | `0,5` | Semanal |
| `CEA` | `0,25` | Semanal |
| `Reunion Sindical (Dentro de contrato)` | `0,25` | Semanal (Asignación 4) |
| `Desarrollo técnico AM` / `Investigación AM` | `5,0` / `4,0` | Por contrato específico |
| `Capacitación` / `Acreditación` / `Docencia DMN` / `Coordinacion nacional WeeFIM` | *Parametrizable* | **Autorizado / Definido por DMN** |
| `Actividades de Coordinacion` / `Coordinacion ULAB` / `Dirigente Sindical` / `Atenciones Otro IT` | *Parametrizable* | **Definidos por GP / DMN** |

---

## 5. Condiciones de Evaluación (Rúbrica Oficial de Semifinales)

Para obtener el puntaje máximo (**4 · Destacado**) en las 9 dimensiones de la rúbrica de `Solicitud.md`:
1. **Problema y usuario**: cuantificar el problema real de Teletón con datos de los Excel (> 600 profesionales en histórico / 462 en base 2026, 17 institutos, > 200 columnas encadenadas, más de 1.042.866 atenciones anuales planificadas que impactan a más de 32.800 niños, niñas y jóvenes en rehabilitación).
2. **Pertinencia de la solución**: justificar claramente por qué **Blue Prism** es la pieza clave para la ingesta, validación, cruce de novedades BUK/SIC y orquestación sin intervención humana, conectado a un motor de reglas parametrizable.
3. **Innovación y propuesta de valor**: motor de reglas desacoplado del Excel, simulador "What-if" de cambios de jornada/parámetros en tiempo real, trazabilidad tipo control de versiones (Git-style diff de dotación) y alertas tempranas SIC vs Planificado.
4. **Funcionamiento del bot**: mostrar en vivo en la demo de 5 minutos cómo Blue Prism toma un archivo Excel de BUK con novedades (altas, bajas, cambios de jornada), valida errores, carga los cambios y actualiza instantáneamente las metas en el dashboard.
5. **Viabilidad técnica**: arquitectura limpia, local en la VM provista (Blue Prism + MSSQLLocalDB + Backend/API + Frontend Web interactivo), lista para integrarse en el futuro a las APIs de BUK y SIC.
6. **Accesibilidad e inclusión**: incluir y **demostrar en vivo** en el dashboard modo de alto contraste/lectura fácil, compatibilidad con lector de pantalla, navegación por teclado y síntesis de voz (lectura en voz alta de los KPIs nacionales/instituto).
7. **Impacto esperado**: mostrar métricas concretas: reducción del tiempo de procesamiento de días a segundos, 0% error de fórmulas rotas (`#REF!` presentes hoy en `Resumen distribucion NAS` y `Excepciones`), 100% trazabilidad de aprobaciones DMN y visibilidad inmediata para los 17 institutos.
8. **Claridad y manejo del tiempo**: estructura de Pitch de 3 minutos + Demo de 5 minutos cronometrada, incluyendo la **lámina obligatoria de arquitectura** (Blue Prism + Motor/App + Escalabilidad nacional).
9. **Dominio en preguntas**: respaldo total en las fórmulas exactas documentadas en este archivo.
