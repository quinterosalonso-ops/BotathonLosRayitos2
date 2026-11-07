# Plan del dashboard web

## Objetivo del MVP

Dar al equipo central una vista web del consolidado nacional y permitir consultar, ingresar y modificar información de planificación, con trazabilidad. Los reportes Excel de BUK/SIC se reciben desde la web y Blue Prism procesa las cargas mediante una carpeta de intercambio supervisada.

## Arquitectura

- **Web:** FastAPI con páginas HTML server-rendered y CSS responsive; no requiere frontend separado.
- **Datos:** MSSQLLocalDB en desarrollo local, con acceso mediante `pyodbc` y consultas parametrizadas.
- **Procesamiento:** Blue Prism sigue siendo responsable de las cargas y automatizaciones. La web no ejecuta SQL arbitrario ni reemplaza las reglas existentes.
- **Intercambio:** la web valida el archivo y lo deja en `inbox` con un identificador de ejecución y manifiesto. Blue Prism toma los trabajos pendientes, ejecuta el proceso con la ruta del archivo y escribe resultado/estado en SQL o en `outbox`; la web muestra el estado.
- **Identidad:** en una demo de un solo usuario, FastAPI y Blue Prism deben correr bajo una cuenta Windows que pueda acceder a la misma instancia LocalDB. Para uso compartido o desatendido, migrar a SQL Server Express/Developer y una identidad de servicio con permisos limitados.

## Pantallas del MVP

1. **Resumen nacional:** profesionales, horas, planificación/meta anual por instituto; filtro por instituto.
2. **Detalle por profesional:** área, cargo, horas, plan/meta anual y, al seleccionar mes, planificación mensual y ejecución SIC.
3. **Cargas:** subir Excel BUK/SIC, validar formato, tamaño y nombre de hoja, crear un trabajo pendiente e historial. La vista previa del contenido del libro queda pendiente.
4. **Parámetros:** editar parámetros globales numéricos permitidos y reglas de distribución por cargo; validar porcentajes y auditar cada cambio transaccionalmente.
5. **Auditoría y errores:** últimos cambios y errores de validación.

## Límites de edición

- No permitir editar directamente métricas calculadas ni ejecutar sentencias SQL desde el navegador.
- Separar datos fuente editables, parámetros y resultados calculados.
- Guardar cada cambio de negocio en `auditoria_cambios`, con usuario, justificación y periodo de vigencia.
- Validar jornada, instituto, cargo y fechas en el backend; recalcular en una transacción y mostrar el resultado antes/después.
- Conservar el archivo original y asociar cada carga con su nombre, hash, fecha, usuario y resultado.

## Contrato de intercambio con Blue Prism

1. FastAPI crea un `job_id` y guarda el archivo como `source.xlsx` en una subcarpeta única de `inbox`, con manifiesto y hash; evita reemplazar cargas anteriores.
2. Blue Prism toma solo trabajos `PENDING`, valida que el archivo y su manifiesto estén completos, y ejecuta el proceso correspondiente.
3. Blue Prism registra `RUNNING`, `SUCCEEDED` o `FAILED`, conteos y errores en SQL mediante el endpoint de callback de la web; el estado también se sincroniza con el manifiesto.
4. FastAPI consulta el estado y presenta el detalle. Una carga `PENDING` puede cancelarse y una carga `FAILED` puede reencolarse desde el dashboard.
5. Definir la carpeta de intercambio y permisos en configuración; no acoplarla a una ruta personal fija del equipo actual.

La carpeta por sí sola no inicia un proceso: Blue Prism debe tener un proceso programado/polling que recoja trabajos pendientes. La web no debe simular un clic de Control Room ni depender de una API de ejecución no confirmada.

El release actual todavía usa rutas fijas de archivos de demo y no consume esta cola web. El coordinador/polling y la adaptación de esos procesos son pendientes de integración; la carga web solo registra el trabajo y no debe presentarse como procesada hasta que el estado llegue a `SUCCEEDED`.

## Entidades y fuentes

- `vw_resumen_nacional_instituto`: indicadores agregados por instituto.
- `dotacion_asistencial` y `planificacion_mensual`: detalle planificado y mensual.
- `ejecucion_sic`: atenciones reales y cumplimiento SIC.
- `parametros_globales`, `distribucion_nac`, `calendario_2026`, `excepciones_multi_it`, `lop_instituto`: datos de reglas/fuentes.
- `errores_validacion` y `auditoria_cambios`: trazabilidad.
- Añadir una tabla de trabajos de carga, por ejemplo `web_import_jobs`, para `job_id`, archivo, tipo, hash, estado, usuario, timestamps y resultado.

La aplicación está implementada en `dashboard.py` y se inicia con `python iniciar_dashboard.py`. Los detalles de configuración y las limitaciones de esta primera versión están en [DASHBOARD_WEB.md](./DASHBOARD_WEB.md). La base de desarrollo puede tener solo esquema y no datos semilla.

## Fases y criterios de aceptación

### Fase 0 — Base de datos y ejecución

- Instalar el ODBC Driver 17/18 y completar la carga inicial.
- Confirmar qué identidad de Windows usa el servidor web y el runtime de Blue Prism.
- Confirmar filas en las tablas principales y conciliar totales con el Excel de referencia.

**Aceptación:** la web y Blue Prism leen/escriben la misma instancia/base; la carga inicial termina sin errores.

### Fase 1 — Lectura del dashboard (implementada)

- Crear la aplicación FastAPI, configuración local y acceso de solo lectura.
- Implementar resumen nacional, filtros y drill-down.
- Manejar errores de conexión explícitamente y no mostrar datos ficticios como si fueran reales.

**Aceptación:** filtros y totales corresponden a consultas SQL verificables y el sistema informa claramente cuando la base está vacía.

### Fase 2 — Cargas Excel y Blue Prism (parcial)

- Subida local, validación de extensión/tamaño/hojas y registro del trabajo en SQL.
- Crear trabajos de carga en carpeta `inbox` y exponer callback de estados.
- Pendiente: vista previa, coordinador/polling en Blue Prism, adaptación de los procesos para consumir rutas por trabajo y validación de negocio antes de aplicar cambios.

**Aceptación:** una carga válida se procesa una sola vez; una inválida muestra sus errores y no modifica datos maestros.

### Fase 3 — Edición y reglas (parcial)

- Formularios acotados para parámetros numéricos y distribución por cargo, con validación, justificación y auditoría.
- Pendiente: autenticación/roles y recálculo automático de métricas afectadas.

**Aceptación:** un cambio válido queda auditado y recalculado; un cambio inválido o fallido no deja escrituras parciales.

### Fase 4 — Endurecimiento

- Autenticación/autorización para más de un usuario, copias de seguridad, registros operativos y pruebas de concurrencia.
- Migrar de LocalDB si el servicio va a ser compartido o desatendido.

## Riesgos y decisiones pendientes

- **Referencia numérica no conciliada:** en la instancia local revisada, el archivo `Nac - Planificación Asistencial x mes.xlsx` y `dotacion_asistencial` contienen las mismas 471 filas, pero suman 16.165 H/S dic y 772.356 atenciones anuales (618.998 directas y 153.358 indirectas). La tabla de `Solicitud/Solicitud.md` espera 21.586 H/S y 1.042.866 atenciones (823.982 directas y 218.884 indirectas), además de dotación en Castro y Ñuble, que no aparece en esas 471 filas. Acordar cuál archivo/versión es la fuente de verdad antes de modificar fórmulas o presentar aceptación de cálculo exacto.
- LocalDB es por usuario y no es una buena base de servicio multiusuario; mantenerlo solo para desarrollo/demo.
- El runtime de Blue Prism debe consumir los trabajos de la carpeta de forma programada. Acordar si se modifican los tres procesos existentes o se crea un proceso coordinador.
- Definir roles (consulta, planificación, administrador) antes de habilitar edición.
- Identificar qué campos de negocio se pueden cambiar desde web y cuáles requieren aprobación.
- Instalar ODBC Driver 17/18 en el host del backend; actualmente no está disponible en este entorno.
