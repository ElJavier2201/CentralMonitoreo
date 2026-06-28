<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase: rutaBase = "../"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim idItem, objConn, cmd
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then idItem = IDValido(Request.Form("id_item")) Else idItem = 0
If idItem > 0 Then
    Set objConn = AbrirConexion()
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = objConn
    cmd.CommandType = 1
    cmd.CommandText = "UPDATE Checklist_Proyecto SET Completado = IIF(Completado=True, False, True), Fecha_Actualizacion = ? WHERE ID_Item = ?"
    cmd.Parameters.Append cmd.CreateParameter("@fecha", 135, 1, , Now())
    cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idItem)
    cmd.Execute , , 128
    Set cmd = Nothing
    CerrarConexion objConn
End If
Response.Redirect "listar.asp?ok=1"
%>
