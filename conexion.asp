
<%


'==========================================================================
' conexion.asp
' Modulo central de acceso a datos.
' Central de Monitoreo para Simulaciones de Circuitos y Prototipos
'
' Contiene:
'   - Apertura/cierre EXPLICITO de conexion ADO sobre Microsoft Access
'   - Funciones de limpieza (Trim) y escape para sentencias SQL nativas
'   - Validacion de IDs numericos (anti-inyeccion via querystring/form)
'   - Utilidades de fecha para concatenar literales SQL de Access (#fecha#)
'
' IMPORTANTE: la ruta usa "/" inicial porque Server.MapPath resuelve
' rutas relativas en funcion del script que se esta EJECUTANDO (no del
' archivo incluido). Con "/" se ancla siempre a la raiz del sitio en
' IIS Express, sin importar si el include se hizo desde la raiz o desde
' una subcarpeta (proyectos/, componentes/, fallos/).
'==========================================================================

Const RUTA_BD_FISICA = "C:\CentralMonitoreo\db\CentralMonitoreo.accdb"

' ----------------------------------------------------------------------
' Abre y devuelve una conexion ADO lista para usar.
' El llamador es responsable de cerrarla con CerrarConexion().
' ----------------------------------------------------------------------
Function AbrirConexion()
    Dim objConn
    Dim strConn
   

    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & RUTA_BD_FISICA & ";" & _
              "Persist Security Info=False;"

    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 15
    objConn.CommandTimeout = 30
    objConn.Open strConn

    Set AbrirConexion = objConn
End Function

' ----------------------------------------------------------------------
' Cierra y destruye EXPLICITAMENTE la conexion para liberar el bloqueo
' que Access (.accdb / .laccdb) mantiene sobre el archivo en Windows.
' Debe llamarse al finalizar CADA tarea, sin excepcion.
' ----------------------------------------------------------------------
Sub CerrarConexion(ByRef objConn)
    On Error Resume Next
    If IsObject(objConn) Then
        If Not (objConn Is Nothing) Then
            If objConn.State = 1 Then objConn.Close
        End If
    End If
    Set objConn = Nothing
    On Error Goto 0
End Sub

' ----------------------------------------------------------------------
' Limpia espacios sobrantes (Trim) y normaliza valores Null a cadena vacia.
' Debe usarse sobre TODO Request.Form / Request.QueryString antes de usarlo.
' ----------------------------------------------------------------------
Function Limpiar(ByVal valor)
    If IsNull(valor) Then
        Limpiar = ""
    ElseIf valor = "" Then
        Limpiar = ""
    Else
        Limpiar = Trim(CStr(valor))
    End If
End Function

' ----------------------------------------------------------------------
' Escapa comillas simples para que el dato no rompa la sentencia SQL
' concatenada (proteccion basica anti-inyeccion en SQL nativo).
' ----------------------------------------------------------------------
Function EscaparSQL(ByVal valor)
    EscaparSQL = Replace(Limpiar(valor), "'", "''")
End Function

' ----------------------------------------------------------------------
' Devuelve el valor de texto listo para concatenar en SQL: entre comillas
' simples, o NULL (literal SQL) si la cadena quedo vacia tras el Trim.
' ----------------------------------------------------------------------
Function ValorTextoSQL(ByVal valor)
    Dim v
    v = EscaparSQL(valor)
    If v = "" Then
        ValorTextoSQL = "NULL"
    Else
        ValorTextoSQL = "'" & v & "'"
    End If
End Function

' ----------------------------------------------------------------------
' Valida que el ID recibido (QueryString o Form) sea numerico entero.
' Si no lo es, devuelve 0 para que el llamador rechace la operacion.
' ----------------------------------------------------------------------
Function IDValido(ByVal valor)
    Dim v
    v = Limpiar(valor)
    If IsNumeric(v) And v <> "" Then
        IDValido = CLng(v)
    Else
        IDValido = 0
    End If
End Function

' ----------------------------------------------------------------------
' Formatea una fecha como literal SQL de Access: #m/d/yyyy#
' ----------------------------------------------------------------------
Function FechaSQL(ByVal valorFecha)
    If IsDate(valorFecha) Then
        FechaSQL = "#" & Month(valorFecha) & "/" & Day(valorFecha) & "/" & Year(valorFecha) & "#"
    Else
        FechaSQL = "#" & Month(Now) & "/" & Day(Now) & "/" & Year(Now) & "#"
    End If
End Function

' ----------------------------------------------------------------------
' Formatea una fecha para un <input type="date"> (yyyy-mm-dd)
' ----------------------------------------------------------------------
Function FormatearFechaInput(ByVal valorFecha)
    If IsDate(valorFecha) Then
        FormatearFechaInput = Year(valorFecha) & "-" & Right("0" & Month(valorFecha), 2) & "-" & Right("0" & Day(valorFecha), 2)
    Else
        FormatearFechaInput = ""
    End If
End Function

' ----------------------------------------------------------------------
' Formatea una fecha para mostrarla en pantalla (dd/mm/yyyy)
' ----------------------------------------------------------------------
Function FormatearFechaVisible(ByVal valorFecha)
    If IsDate(valorFecha) Then
        FormatearFechaVisible = Right("0" & Day(valorFecha), 2) & "/" & Right("0" & Month(valorFecha), 2) & "/" & Year(valorFecha)
    Else
        FormatearFechaVisible = "-"
    End If
End Function

' ----------------------------------------------------------------------
' Devuelve "activo" si la seccion actual coincide con la indicada.
' (VBScript NO tiene IIf nativo; esta es la alternativa correcta)
' ----------------------------------------------------------------------
Function ClaseSiActivo(ByVal actual, ByVal seccion)
    If actual = seccion Then
        ClaseSiActivo = "activo"
    Else
        ClaseSiActivo = ""
    End If
End Function
%>
