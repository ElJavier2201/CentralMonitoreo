
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Instalador de Base de Datos</title>
<link rel="stylesheet" href="../includes/estilos.css">
</head>
<body>
<main class="contenedor">
<div class="panel">
<h1>Instalador de Base de Datos</h1>
<p class="ayuda">
    Crea el archivo <strong>CentralMonitoreo.accdb</strong> y las tres tablas
    (Proyectos, Componentes, Bitacora_Fallos) si todavía no existen.
    Solo es necesario ejecutarlo una vez.
</p>
<pre class="consola">
<%
Dim objFSO, rutaBD, strConn, cat, objConn
Dim sqlProyectos, sqlComponentes, sqlFallos

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
rutaBD = Server.MapPath("CentralMonitoreo.accdb")
strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & rutaBD & ";Persist Security Info=False;"

Response.Write "Ruta del archivo: " & rutaBD & vbCrLf & vbCrLf

' --------------------------------------------------------------
' Paso 1: crear el archivo .accdb si no existe (vía ADOX.Catalog)
' --------------------------------------------------------------
If objFSO.FileExists(rutaBD) Then
    Response.Write "[OK] El archivo de base de datos ya existía. No se vuelve a crear." & vbCrLf
Else
    On Error Resume Next
    Set cat = Server.CreateObject("ADOX.Catalog")
    cat.Create strConn
    If Err.Number <> 0 Then
        Response.Write "[ERROR] No se pudo crear el archivo .accdb: " & Err.Description & vbCrLf
        Response.Write "        Verifique que el 'Microsoft Access Database Engine Redistributable'" & vbCrLf
        Response.Write "        (proveedor ACE OLEDB 12.0) esté instalado, y que la carpeta /db" & vbCrLf
        Response.Write "        tenga permiso de escritura para el usuario de IIS (IIS_IUSRS / IUSR)." & vbCrLf
        Err.Clear
    Else
        Response.Write "[OK] Archivo CentralMonitoreo.accdb creado correctamente." & vbCrLf
    End If
    Set cat = Nothing
    On Error Goto 0
End If

Response.Write vbCrLf

' --------------------------------------------------------------
' Paso 2: crear las tablas mediante SQL nativo (DDL) sobre ADO
' --------------------------------------------------------------
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open strConn

Sub EjecutarDDL(ByVal nombreTabla, ByVal sentenciaSQL, ByRef cn)
    On Error Resume Next
    cn.Execute sentenciaSQL, , 129   ' adCmdText + adExecuteNoRecords
    If Err.Number <> 0 Then
        ' La tabla ya existe -> no es un error real, se informa y se continúa
        Response.Write "[OK] La tabla '" & nombreTabla & "' ya existía (o no requirió cambios)." & vbCrLf
        Err.Clear
    Else
        Response.Write "[OK] Tabla '" & nombreTabla & "' creada correctamente." & vbCrLf
    End If
    On Error Goto 0
End Sub

sqlProyectos = "CREATE TABLE Proyectos (" & _
    "ID_Proyecto COUNTER PRIMARY KEY, " & _
    "Nombre_Proyecto VARCHAR(150) NOT NULL, " & _
    "Plataforma_Simulacion VARCHAR(80) NOT NULL, " & _
    "Microcontrolador VARCHAR(80) NOT NULL, " & _
    "Descripcion VARCHAR(255), " & _
    "Fecha_Creacion DATETIME)"

sqlComponentes = "CREATE TABLE Componentes (" & _
    "ID_Componente COUNTER PRIMARY KEY, " & _
    "ID_Proyecto LONG NOT NULL, " & _
    "Tipo_Componente VARCHAR(100) NOT NULL, " & _
    "Valor_Calculado VARCHAR(100), " & _
    "Pin_Conexion VARCHAR(50), " & _
    "Ubicacion_Protoboard VARCHAR(100), " & _
    "Notas VARCHAR(255), " & _
    "CONSTRAINT FK_Componentes_Proyectos FOREIGN KEY (ID_Proyecto) REFERENCES Proyectos (ID_Proyecto))"

sqlFallos = "CREATE TABLE Bitacora_Fallos (" & _
    "ID_Fallo COUNTER PRIMARY KEY, " & _
    "ID_Proyecto LONG NOT NULL, " & _
    "Fecha_Registro DATETIME, " & _
    "Sintoma_Error VARCHAR(255) NOT NULL, " & _
    "Solucion_Aplicada VARCHAR(255), " & _
    "Estado VARCHAR(20), " & _
    "CONSTRAINT FK_Fallos_Proyectos FOREIGN KEY (ID_Proyecto) REFERENCES Proyectos (ID_Proyecto))"

EjecutarDDL "Proyectos", sqlProyectos, objConn
EjecutarDDL "Componentes", sqlComponentes, objConn
EjecutarDDL "Bitacora_Fallos", sqlFallos, objConn

' Cierre y destrucción explícita de la conexión (evita bloqueo del .accdb)
If objConn.State = 1 Then objConn.Close
Set objConn = Nothing
Set objFSO = Nothing

Response.Write vbCrLf & "Instalación finalizada."
%>
</pre>
<div class="acciones">
    <a href="semilla.asp" class="boton boton-secundario">Cargar datos de ejemplo (sensor capacitivo)</a>
    <a href="../index.asp" class="boton boton-primario">Ir al Dashboard</a>
</div>
</div>
</main>
</body>
</html>
