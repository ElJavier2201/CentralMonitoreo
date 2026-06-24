# Central de Monitoreo para Simulaciones de Circuitos y Prototipos

Bitácora técnica web (ASP Clásico + VBScript + Microsoft Access) para documentar el paso
a paso entre el diseño en simulador (Proteus / Tinkercad) y el ensamblaje físico del hardware
(ej. sensores de fuerza capacitivos con Arduino y placas de aluminio).

## 1. Requisitos

- Windows con **IIS** o **IIS Express**, con la característica **ASP clásico** habilitada
  (Panel de control → Activar o desactivar características de Windows → Internet Information
  Services → Servicios World Wide Web → Características de desarrollo de aplicaciones → ASP).
- **Microsoft Access Database Engine Redistributable** instalado (proveedor `Microsoft.ACE.OLEDB.12.0`).
  Debe coincidir en arquitectura (32/64 bits) con el Application Pool de IIS / IIS Express que
  ejecutará el sitio.
- Permisos de **escritura** para el usuario del sitio (`IUSR` / `IIS_IUSRS`) sobre la carpeta `/db`,
  necesarios para crear el archivo `.accdb` y para que Access pueda generar su archivo de bloqueo
  `.laccdb` durante las operaciones (sin esto, el archivo queda bloqueado en Windows o no se puede
  escribir).

## 2. Estructura del proyecto

```
CentralMonitoreo/
├── index.asp                 Dashboard general
├── conexion.asp              Conexión ADO + funciones de utilidad (Trim, escape SQL, fechas...)
├── web.config                Configura index.asp como documento por defecto
├── includes/
│   ├── header.asp            Encabezado y navegación
│   ├── footer.asp             Pie de página
│   └── estilos.css            Tema visual
├── db/
│   ├── CentralMonitoreo.sql   Esquema de referencia (documentación)
│   ├── instalar.asp           Crea el .accdb y las 3 tablas (ejecutar una sola vez)
│   ├── semilla.asp             Carga el proyecto de ejemplo (sensor capacitivo)
│   └── CentralMonitoreo.accdb  Se genera aquí al ejecutar instalar.asp
├── proyectos/
│   ├── listar.asp
│   ├── formulario.asp          Alta y edición (mismo formulario)
│   └── eliminar.asp
├── componentes/
│   ├── listar.asp              INNER JOIN con Proyectos
│   ├── formulario.asp
│   └── eliminar.asp
└── fallos/
    ├── listar.asp               INNER JOIN con Proyectos
    ├── formulario.asp
    └── eliminar.asp
```

## 3. Puesta en marcha

1. Copia la carpeta `CentralMonitoreo` completa y configúrala como **raíz del sitio** en
   IIS / IIS Express (importante: las rutas internas usan `Server.MapPath("/db/...")`,
   ancladas a la raíz del sitio web, así que el proyecto debe ser el sitio o la aplicación raíz,
   no una subcarpeta de otro sitio).
2. Navega a `http://localhost:PUERTO/db/instalar.asp`. Esto crea `CentralMonitoreo.accdb`
   y las tres tablas (`Proyectos`, `Componentes`, `Bitacora_Fallos`) si no existen.
3. (Opcional) Navega a `http://localhost:PUERTO/db/semilla.asp` para cargar el proyecto de
   ejemplo "Sensor de Fuerza Capacitivo" con sus componentes y fallos.
4. Abre `http://localhost:PUERTO/index.asp` para empezar a usar la bitácora.

## 4. Modelo de datos

- **Proyectos** — cabecera del experimento: nombre, plataforma de simulación
  (Proteus / Tinkercad) y microcontrolador (cerebro del circuito).
- **Componentes** — inventario y mapa de cableado por proyecto: tipo de pieza, valor
  calculado, pin de conexión y ubicación exacta en la protoboard. Relacionada con
  `Proyectos` por `ID_Proyecto`.
- **Bitacora_Fallos** — síntomas de error detectados en pruebas físicas y la solución
  electrónica o de código aplicada. Relacionada con `Proyectos` por `ID_Proyecto`.

## 5. Estándares de código aplicados

- **SQL nativo sobre ADO**: todo INSERT/UPDATE/DELETE se ejecuta con
  `objConn.Execute sql, , 129` (`adCmdText` + `adExecuteNoRecords`) directamente sobre el
  objeto `Connection`, sin abrir un `Recordset` intermedio, para no saturar la memoria de
  IIS Express.
- **INNER JOIN**: las pantallas de listado de `componentes/listar.asp` y `fallos/listar.asp`
  cruzan obligatoriamente con `Proyectos` mediante `INNER JOIN ... ON ... ID_Proyecto`.
- **Validación de formularios**: toda entrada pasa por `Limpiar()` (Trim + normalización de
  `Null`), se rechazan envíos con campos obligatorios vacíos y se reconstruye el formulario
  con los valores ya escritos para que el usuario no pierda lo capturado.
- **Cierre explícito de conexiones**: cada página llama a `CerrarConexion()` al finalizar,
  que hace `objConn.Close` y `Set objConn = Nothing` de forma explícita, evitando que el
  archivo `.accdb` quede bloqueado en Windows.
- **Anti-inyección básica**: `EscaparSQL()` duplica las comillas simples y `IDValido()`
  fuerza a que cualquier ID recibido por querystring/formulario sea estrictamente numérico
  antes de concatenarlo en una sentencia SQL.

## 6. Notas

- Si se requiere compatibilidad con archivos `.mdb` clásicos (Jet 4.0) en lugar de `.accdb`,
  basta con cambiar el `Provider` en `conexion.asp` a `Microsoft.Jet.OLEDB.4.0` y la extensión
  del archivo en `RUTA_BD` — el resto del código no cambia, ya que todo el SQL usado es
  compatible con ambos motores.
- El borrado de un proyecto elimina en cascada (de forma manual y explícita) sus
  componentes y registros de fallos asociados, ya que el proveedor OLEDB de Access no
  garantiza `ON DELETE CASCADE` de forma confiable.
