# Blue Prism + MSSQLLocalDB

## Conexión utilizada

- **Excel:** `MS Excel VBO - Extended` abre y lee las hojas de BUK/SIC.
- **Base de datos:** `Data - MSSQLLocalDB` ejecuta consultas y escribe colecciones en staging.
- **Cadena predeterminada:** `Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=TeletonPlanificacion;Integrated Security=True;Connect Timeout=30;TrustServerCertificate=True;`

El VBO de Excel no conecta la base de datos. El flujo lee el libro en una colección y el objeto `Data - MSSQLLocalDB` carga esa colección a SQL Server.

## Preparación

1. Instala SQL Server Express LocalDB, ODBC Driver 17 o 18 for SQL Server y Python 3.10+.
2. Desde la cuenta de Windows que ejecutará Blue Prism, confirma la instancia `MSSQLLocalDB`. Si no existe, crea e inicia la instancia LocalDB para esa cuenta.
3. Instala las dependencias Python con `python -m pip install -r requirements.txt` y ejecuta `python scripts/inicializar_bd_y_motor.py`. El inicializador crea la base, aplica `database/teleton_planificacion_mssql.sql`, carga los datos anonimizados desde `Solicitud/` y crea los reportes demo.
4. Si necesitas solo el esquema sin cargar las planillas, ejecuta `database/teleton_planificacion_mssql.sql` en SSMS. **Cada ejecución borra y recrea las tablas**; úsalo solo al inicializar o reseedear los datos de demostración.
5. Importa `blueprism/Botathon_Teleton_2026_Completo.bprelease` en Blue Prism. El release incluye los dos objetos y los tres procesos. Si importas los XML por separado, importa primero los objetos y después los procesos.
6. Revisa las rutas de archivos en los Data Items de cada proceso antes de ejecutarlos. La cadena de conexión se configura en el Data Item `vConnectionString`.
7. Para la demo, ejecuta `01_Carga_Base_Contratos_BUK`, `02_Carga_y_Calculo_Planificacion` y finalmente `03_Novedades_BUK_y_Contraste_SIC`.

El esquema por sí solo no contiene los catálogos ni la planificación base. El proceso 02 parte de esos datos ya cargados por el inicializador: Blue Prism conserva las lecturas en tablas de staging y actualiza indicadores/auditoría, pero no sustituye la carga inicial de las tablas maestras y de planificación. El proceso 03 incorpora las filas válidas del reporte SIC a `ejecucion_sic`, registra las inválidas y agrega el total ejecutado/cumplimiento SIC al consolidado exportado.

Los procesos comprueban `Success` en cada acción SQL que entrega ese resultado; `Set Connection Details` solo configura la cadena y no devuelve estado. La conexión real se abre en la primera lectura/escritura SQL y sus fallas detienen el flujo con el mensaje del objeto.

Si una versión anterior del proceso termina con `Data - MSSQLLocalDB failed:` y el detalle vacío, reemplázala por la versión actual: se había añadido una comprobación de `Success` después de `Set Connection Details`, acción que no tiene salida `Success`. Los procesos actuales pasan directamente a la primera operación SQL.

## Consideraciones de ejecución

LocalDB es una instancia local por usuario de Windows. La identidad que ejecuta el runtime de Blue Prism debe tener instalada/iniciada su propia instancia y acceso integrado; una instancia LocalDB de una sesión interactiva no necesariamente estará disponible para un runtime con otra cuenta. Para ejecución desatendida como servicio o varios usuarios, conviene SQL Server Express/Developer en lugar de LocalDB.

`Data - MSSQLLocalDB` usa `System.Data.SqlClient` de .NET Framework y no requiere configurar un DSN. El inicializador Python usa `pyodbc` y ODBC Driver 17/18.
