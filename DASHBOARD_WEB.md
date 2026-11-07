# Dashboard web local

## Requisitos

- Python 3.10 o superior.
- ODBC Driver 17 o 18 for SQL Server.
- MSSQLLocalDB con la base `TeletonPlanificacion` inicializada y permisos para la misma cuenta Windows que ejecuta Blue Prism.
- La carpeta de intercambio debe ser accesible para esa cuenta y para Blue Prism.

El driver ODBC no se instala con `pip`. Si falta o no hay conexión, la página muestra el error en vez de presentar datos inventados.

## Iniciar

Desde la raíz del proyecto:

```powershell
python -m pip install -r requirements.txt
python iniciar_dashboard.py
```

Abre <http://127.0.0.1:8000>. El proceso se detiene con `Ctrl+C`. El servidor escucha solo en la interfaz local de forma predeterminada.

Variables opcionales de configuración:

| Variable | Predeterminado | Uso |
|---|---|---|
| `TELETON_SQL_CONNECTION_STRING` | LocalDB `MSSQLLocalDB`, base `TeletonPlanificacion` | Reemplaza la cadena ODBC completa |
| `TELETON_SQL_SERVER` | `(LocalDB)\MSSQLLocalDB` | Servidor del inicializador `scripts/inicializar_bd_y_motor.py` (no lo usa el dashboard) |
| `TELETON_INBOX_DIR` | `blueprism_inbox` en la raíz del proyecto | Carpeta de intercambio compartida con Blue Prism |
| `TELETON_HOST` | `127.0.0.1` | Interfaz de escucha; no se recomienda exponer en una red |
| `TELETON_PORT` | `8000` | Puerto HTTP |
| `TELETON_JOB_TOKEN` | Sin token | Si se define, el callback de estado requiere `X-Job-Token` |

Ejemplo PowerShell:

```powershell
$env:TELETON_INBOX_DIR = 'C:\BluePrism\Teleton\inbox'
$env:TELETON_JOB_TOKEN = 'configura-un-secreto-compartido'
$env:TELETON_SQL_CONNECTION_STRING = 'DRIVER={ODBC Driver 18 for SQL Server};SERVER=localhost\SQLEXPRESS;DATABASE=TeletonPlanificacion;Trusted_Connection=yes;TrustServerCertificate=yes;'
python iniciar_dashboard.py
```

## Funciones disponibles

- **Resumen:** consolidado por instituto consultando `vw_resumen_nacional_instituto`, con horas directas/indirectas/NAS, atenciones anuales, meta al 90%, FOJA y total general del filtro.
- **Planificación:** detalle anual desde `dotacion_asistencial`; filtros combinables por instituto, área, cargo, PersonaId y período Ene-Abr/May-Dic. Al seleccionar un mes consulta `planificacion_mensual` y el contraste de `ejecucion_sic`.
- **Errores de filtros:** consultas SQL limitadas a 15 segundos. Si falla la conexión, la consulta o la validación del filtro, el dashboard muestra un aviso y mantiene los controles para corregir, reintentar o restablecer los filtros.
- **Cargas:** validación de extensión, límite de 20 MB, estructura XLSX y hoja requerida (`Dotacion_2026` para BUK o `Ejecucion_SIC` para SIC). Un archivo válido se conserva como `source.xlsx` con un `manifest.json` y se registra como `PENDING`.
- Desde el historial se puede **cancelar** una carga `PENDING` (queda `CANCELLED`, sin borrar el archivo) o **reencolar** una carga `FAILED`. No se puede cancelar una carga `RUNNING`.
- **Parámetros:** edición limitada de parámetros numéricos aprobados y porcentajes/minutos de `distribucion_nac`. Los tres porcentajes de cada tramo deben sumar 1.0; los cambios se guardan en transacción y se insertan en `auditoria_cambios`.
- **Auditoría:** consulta de hasta 200 cambios recientes y errores de validación.
- **Estado:** `/health` comprueba la conexión a la base.

El dashboard no edita dotación ni métricas calculadas, no ejecuta SQL del navegador y no procesa el Excel por sí solo.

## Contrato de carga para Blue Prism

Cada carga queda en:

```text
<TELETON_INBOX_DIR>\<job_id>\source.xlsx
<TELETON_INBOX_DIR>\<job_id>\manifest.json
```

El manifiesto contiene `job_id`, `type` (`BUK` o `SIC`), nombre original, hoja(s), tamaño, SHA-256, solicitante y estado. Blue Prism debe tener un proceso coordinador programado que:

1. Recoja solo trabajos `PENDING`, valide el manifiesto/hash y cambie el estado a `RUNNING`.
2. Ejecute el proceso de carga correspondiente usando `source.xlsx`.
3. Informe `SUCCEEDED` o `FAILED`, el número de filas y el error mediante `POST /api/jobs/{job_id}/status` con campos de formulario `status`, `rows_processed`, `error_message`.
4. Envíe el header `X-Job-Token` si está configurado `TELETON_JOB_TOKEN`.

El estado se refleja tanto en SQL como en `manifest.json`. Las transiciones principales son `PENDING → RUNNING → SUCCEEDED/FAILED`, `PENDING → CANCELLED` y `FAILED → PENDING` (reencolar desde el dashboard). El callback también admite `FAILED → RUNNING` para un reintento coordinado externamente. Blue Prism debe comprobar el estado antes de procesar y reclamar el trabajo cambiándolo a `RUNNING`; si ya está `RUNNING`, el dashboard no lo cancela. **El coordinador/polling aún debe configurarse en Blue Prism**: esta aplicación no inicia procesos de Blue Prism ni adapta automáticamente las rutas de los tres procesos existentes. La carpeta debe ser compartida y accesible por ambas aplicaciones.

La carga web actualmente verifica el formato XLSX y el nombre de hoja, pero no sustituye la validación de negocio de Blue Prism. El release actual de Blue Prism usa rutas fijas para los archivos de demo y no consume todavía el `job_id`, `source.xlsx` ni el manifiesto de esta cola. Por tanto, subir un archivo y ver `PENDING` no significa que se haya procesado; para cerrar el flujo falta implementar y programar el coordinador que lea la carpeta, adapte las rutas de los procesos y publique el callback.

## Conexión de Blue Prism a SQL Server

LocalDB es por usuario de Windows. La cadena `Integrated Security=True` conecta con la instancia `MSSQLLocalDB` de la identidad que ejecuta el proceso, no con una instancia global compartida. Si el dashboard se ejecuta desde una sesión interactiva y Blue Prism Runtime como servicio/u otra cuenta, pueden estar viendo instancias distintas o la cuenta del Runtime puede no tener una instancia LocalDB iniciada.

Para comprobarlo, ejecuta estos comandos bajo la identidad configurada en Blue Prism Runtime:

```powershell
whoami
sqllocaldb info MSSQLLocalDB
sqllocaldb start MSSQLLocalDB
```

Confirma también que `TeletonPlanificacion` existe en esa instancia y que esa cuenta puede abrirla. Si el Runtime es desatendido, la solución recomendada es **SQL Server Express/Developer con una instancia compartida**, concediendo permisos de base de datos a la misma cuenta de servicio que usa el dashboard y Blue Prism, en vez de depender de una LocalDB por usuario.

Ejemplo para una instancia local `SQLEXPRESS`; configura el mismo servidor/base en los Data Items `vConnectionString` de los tres procesos Blue Prism y en el dashboard:

```text
Blue Prism:
Data Source=localhost\SQLEXPRESS;Initial Catalog=TeletonPlanificacion;Integrated Security=True;Connect Timeout=30;TrustServerCertificate=True;

Dashboard (PowerShell):
$env:TELETON_SQL_CONNECTION_STRING = 'DRIVER={ODBC Driver 18 for SQL Server};SERVER=localhost\SQLEXPRESS;DATABASE=TeletonPlanificacion;Trusted_Connection=yes;TrustServerCertificate=yes;'
```

La identidad web y la identidad Runtime deben tener acceso a esa misma instancia y permisos limitados sobre `TeletonPlanificacion`. No basta con que las cadenas coincidan si se ejecutan bajo cuentas Windows diferentes.

Para inicializar una base vacía en `SQLEXPRESS`, configura además `TELETON_SQL_SERVER=localhost\SQLEXPRESS` antes de ejecutar el inicializador. **El inicializador elimina y reconstruye las tablas**, por lo que no se debe ejecutar sobre una base que contenga datos que se quieran conservar.

## Límites operativos

Esta versión es para desarrollo/demo local, sin inicio de sesión ni roles. No la expongas a otros equipos antes de implementar autenticación/autorización y revisar la configuración de red. LocalDB depende de la identidad Windows; web y Blue Prism deben usar una identidad con acceso a la misma instancia. Para ejecución multiusuario o desatendida, usa SQL Server Express/Developer con una identidad de servicio.

No ejecutar repetidamente `scripts/inicializar_bd_y_motor.py` sobre datos que deban conservarse: el inicializador reconstruye el esquema y carga datos.
