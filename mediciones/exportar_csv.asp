<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True
Response.CharSet = "UTF-8"
Dim rutaBase: rutaBase = "../"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Response.ContentType = "text/csv"
Response.AddHeader "Content-Disposition", "attachment; filename=mediciones.csv"
Response.BinaryWrite ChrB(&hEF) & ChrB(&hBB) & ChrB(&hBF)
Function CSV(ByVal valor)
    Dim v
    v = Limpiar(valor)
    v = Replace(v, """", """""")
    v = Replace(v, vbCrLf, " ")
    v = Replace(v, vbLf, " ")
    CSV = """" & v & """"
End Function
Dim objConn, objRS, sql, filtroProyecto
filtroProyecto = IDValido(Request.QueryString("proyecto"))
Set objConn = AbrirConexion()
sql = "SELECT p.Nombre_Proyecto, m.Fecha_Medicion, m.Variable_Medida, m.Valor_Simulado, m.Valor_Fisico, m.Diferencia, m.Observaciones " & _
      "FROM Mediciones m INNER JOIN Proyectos p ON m.ID_Proyecto = p.ID_Proyecto "
If filtroProyecto > 0 Then sql = sql & "WHERE p.ID_Proyecto = " & filtroProyecto & " "
sql = sql & "ORDER BY p.Nombre_Proyecto, m.Fecha_Medicion DESC"
Set objRS = objConn.Execute(sql)
Response.Write "Proyecto;Fecha;Variable;Valor Simulado;Valor Fisico;Diferencia;Observaciones" & vbCrLf
Do While Not objRS.EOF
    Response.Write CSV(objRS("Nombre_Proyecto")) & ";" & CSV(FormatearFechaVisible(objRS("Fecha_Medicion"))) & ";" & CSV(objRS("Variable_Medida")) & ";" & CSV(objRS("Valor_Simulado")) & ";" & CSV(objRS("Valor_Fisico")) & ";" & CSV(objRS("Diferencia")) & ";" & CSV(objRS("Observaciones")) & vbCrLf
    objRS.MoveNext
Loop
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
Response.End
%>
