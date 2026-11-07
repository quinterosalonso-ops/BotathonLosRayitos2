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
| `TELETON_INBOX_DIR` | `blueprism_inbox` en la raíz del proyecto | Carpeta de intercambio compartida con Blue Prism |
| `TELETON_HOST` | `127.0.0.1` | Interfaz de escucha; no se recomienda exponer en una red |
| `TELETON_PORT` | `8000` | Puerto HTTP |
| `TELETON_JOB_TOKEN` | Sin token | Si se define, el callback de estado requiere `X-Job-Token` |

Ejemplo PowerShell:

```powershell
$env:TELETON_INBOX_DIR = 'C:\BluePrism\Teleton\inbox'
$env:TELETON_JOB_TOKEN = 'configura-un-secreto-compartido'
python iniciar_dashboard.py
```

## Funciones disponibles

- **Resumen:** consolidado por instituto consultando `vw_resumen_nacional_instituto`.
- **Planificación:** detalle anual desde `dotacion_asistencial`; al filtrar por mes consulta `planificacion_mensual` y el contraste de `ejecucion_sic`.
- **Cargas:** validación de extensión, límite de 20 MB, estructura XLSX y hoja requerida (`Dotacion_2026` para BUK o `Ejecucion_SIC` para SIC). Un archivo válido se conserva como `source.xlsx` con un `manifest.json` y se registra como `PENDING`.
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

El callback aplica estas transiciones: `PENDING → RUNNING → SUCCEEDED/FAILED` y `FAILED → RUNNING` para reintentos. **El coordinador/polling aún debe configurarse en Blue Prism**: esta aplicación no inicia procesos de Blue Prism ni adapta automáticamente las rutas de los tres procesos existentes. La carpeta debe ser compartida y accesible por ambas aplicaciones.

## Límites operativos

Esta versión es para desarrollo/demo local, sin inicio de sesión ni roles. No la expongas a otros equipos antes de implementar autenticación/autorización y revisar la configuración de red. LocalDB depende de la identidad Windows; web y Blue Prism deben usar una identidad con acceso a la misma instancia. Para ejecución multiusuario o desatendida, usa SQL Server Express/Developer con una identidad de servicio.

No ejecutar repetidamente `scripts/inicializar_bd_y_motor.py` sobre datos que deban conservarse: el inicializador reconstruye el esquema y carga datos.
