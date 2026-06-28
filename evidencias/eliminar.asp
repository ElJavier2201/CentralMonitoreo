<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase: rutaBase = "../"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim idEvidencia, objConn, cmd
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then idEvidencia = IDValido(Request.Form("id_evidencia")) Else idEvidencia = 0
If idEvidencia > 0 Then
    Set objConn = AbrirConexion()
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = objConn
    cmd.CommandType = 1
    cmd.CommandText = "DELETE FROM Evidencias WHERE ID_Evidencia = ?"
    cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idEvidencia)
    cmd.Execute , , 128
    Set cmd = Nothing
    CerrarConexion objConn
End If
Response.Redirect "listar.asp?eliminado=1"
%>
