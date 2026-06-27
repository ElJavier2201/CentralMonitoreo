<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True
Response.CharSet = "UTF-8"

Dim rutaBase
rutaBase = "../"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Response.ContentType = "text/csv"
Response.AddHeader "Content-Disposition", "attachment; filename=inventario_componentes.csv"
Response.BinaryWrite ChrB(&hEF) & ChrB(&hBB) & ChrB(&hBF)

Dim objConn, objRS, objCmd, sql
Dim filtroProyecto, terminoBusqueda, clausulaWhere

filtroProyecto = IDValido(Request.QueryString("proyecto"))
terminoBusqueda = Limpiar(Request.QueryString("q"))
clausulaWhere = ""

Set objConn = AbrirConexion()
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = objConn
objCmd.CommandType = 1

sql = "SELECT c.Tipo_Componente, c.Valor_Calculado, c.Pin_Conexion, c.Ubicacion_Protoboard, c.Notas, p.Nombre_Proyecto " & _
      "FROM Componentes c INNER JOIN Proyectos p ON c.ID_Proyecto = p.ID_Proyecto "

If filtroProyecto > 0 Then
    clausulaWhere = "WHERE p.ID_Proyecto = ? "
    objCmd.Parameters.Append objCmd.CreateParameter("@proyecto", 3, 1, , filtroProyecto)
End If

If terminoBusqueda <> "" Then
    If clausulaWhere = "" Then
        clausulaWhere = "WHERE "
    Else
        clausulaWhere = clausulaWhere & "AND "
    End If
    clausulaWhere = clausulaWhere & "(c.Tipo_Componente LIKE ? OR c.Valor_Calculado LIKE ?) "
    objCmd.Parameters.Append objCmd.CreateParameter("@q1", 200, 1, 150, "%" & terminoBusqueda & "%")
    objCmd.Parameters.Append objCmd.CreateParameter("@q2", 200, 1, 150, "%" & terminoBusqueda & "%")
End If

objCmd.CommandText = sql & clausulaWhere & "ORDER BY p.Nombre_Proyecto, c.Tipo_Componente"
Set objRS = objCmd.Execute

Response.Write "Proyecto;Componente;Valor Calculado;Pin;Ubicación Protoboard;Notas" & vbCrLf

Do While Not objRS.EOF
    Response.Write CSV(objRS("Nombre_Proyecto")) & ";" & _
                   CSV(objRS("Tipo_Componente")) & ";" & _
                   CSV(objRS("Valor_Calculado")) & ";" & _
                   CSV(objRS("Pin_Conexion")) & ";" & _
                   CSV(objRS("Ubicacion_Protoboard")) & ";" & _
                   CSV(objRS("Notas")) & vbCrLf
    objRS.MoveNext
Loop

objRS.Close
Set objRS = Nothing
Set objCmd = Nothing
CerrarConexion objConn
Response.End

Function CSV(ByVal valor)
    Dim v
    v = Limpiar(valor)
    v = Replace(v, """", """""")
    v = Replace(v, vbCrLf, " ")
    v = Replace(v, vbCr, " ")
    v = Replace(v, vbLf, " ")
    CSV = """" & v & """"
End Function
%>
