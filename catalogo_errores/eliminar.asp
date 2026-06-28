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
Dim idError, objConn, cmd
idError = IDValido(Request.Form("id"))

If Request.ServerVariables("REQUEST_METHOD") = "POST" And idError > 0 Then
    Set objConn = AbrirConexion()
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = objConn
    cmd.CommandType = 1
    cmd.CommandText = "DELETE FROM Catalogo_Errores WHERE ID_Error = ?"
    cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idError)
    cmd.Execute , , 128
    Set cmd = Nothing
    CerrarConexion objConn
End If

Response.Redirect "listar.asp?eliminado=1"
Response.End
%>
