<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase
rutaBase = "../"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim idCalculo, objConn, cmd
idCalculo = IDValido(Request.Form("id"))

If Request.ServerVariables("REQUEST_METHOD") = "POST" And idCalculo > 0 Then
    Set objConn = AbrirConexion()
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = objConn
    cmd.CommandType = 1
    cmd.CommandText = "DELETE FROM Calculos_Electricos WHERE ID_Calculo = ?"
    cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idCalculo)
    cmd.Execute , , 128
    Set cmd = Nothing
    CerrarConexion objConn
End If

Response.Redirect "listar.asp?eliminado=1"
Response.End
%>
