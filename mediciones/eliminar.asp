<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase: rutaBase = "../"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim idMedicion, objConn, cmd
idMedicion = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then idMedicion = IDValido(Request.Form("id_medicion"))
If idMedicion > 0 Then
    Set objConn = AbrirConexion()
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = objConn
    cmd.CommandType = 1
    cmd.CommandText = "DELETE FROM Mediciones WHERE ID_Medicion = ?"
    cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idMedicion)
    cmd.Execute , , 128
    Set cmd = Nothing
    CerrarConexion objConn
End If
Response.Redirect "listar.asp?eliminado=1"
%>
