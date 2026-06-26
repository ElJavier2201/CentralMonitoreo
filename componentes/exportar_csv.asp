<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True
Response.CharSet = "UTF-8"

' 1. Configurar cabeceras HTTP para forzar la descarga del archivo
Response.ContentType = "text/csv"
Response.AddHeader "Content-Disposition", "attachment; filename=inventario_componentes.csv"

' 2. Enviar el BOM UTF-8 para que Excel abra el archivo con la codificación correcta
Response.BinaryWrite ChrB(&hEF) & ChrB(&hBB) & ChrB(&hBF)
%>
<%
Dim objConn, objRS, sql
Dim filtroProyecto, terminoBusqueda, clausulaWhere

filtroProyecto = IDValido(Request.QueryString("proyecto"))
terminoBusqueda = Limpiar(Request.QueryString("q"))
clausulaWhere = ""

Set objConn = AbrirConexion()

sql = "SELECT c.Tipo_Componente, c.Valor_Calculado, c.Pin_Conexion, c.Ubicacion_Protoboard, c.Notas, p.Nombre_Proyecto " & _
      "FROM Componentes c INNER JOIN Proyectos p ON c.ID_Proyecto = p.ID_Proyecto "

If filtroProyecto > 0 Then
    clausulaWhere = "WHERE p.ID_Proyecto = " & filtroProyecto & " "
End If

If terminoBusqueda <> "" Then
    If clausulaWhere = "" Then clausulaWhere = "WHERE " Else clausulaWhere = clausulaWhere & "AND "
    clausulaWhere = clausulaWhere & "(c.Tipo_Componente LIKE '%" & EscaparSQL(terminoBusqueda) & "%' OR c.Valor_Calculado LIKE '%" & EscaparSQL(terminoBusqueda) & "%') "
End If

sql = sql & clausulaWhere & "ORDER BY p.Nombre_Proyecto, c.Tipo_Componente"
Set objRS = objConn.Execute(sql)

' 3. Imprimir la línea de encabezados (Usando punto y coma como separador regional)
Response.Write "Proyecto;Componente;Valor Calculado;Pin;Ubicación Protoboard;Notas" & vbCrLf

' 4. Recorrer los registros y escribir las filas limpiando saltos de línea internos
Do While Not objRS.EOF
    Dim notaLimpia
    notaLimpia = Replace(Limpiar(objRS("Notas")), vbCrLf, " ")
    notaLimpia = Replace(notaLimpia, ";", ",") ' Evita que el punto y coma rompa la columna

    Response.Write """" & objRS("Nombre_Proyecto") & """;" & _
                   """" & objRS("Tipo_Componente") & """;" & _
                   """" & objRS("Valor_Calculado") & """;" & _
                   """" & objRS("Pin_Conexion") & """;" & _
                   """" & objRS("Ubicacion_Protoboard") & """;" & _
                   """" & notaLimpia & """" & vbCrLf
    objRS.MoveNext
Loop

objRS.Close
Set objRS = Nothing
CerrarConexion objConn
Response.End
%>