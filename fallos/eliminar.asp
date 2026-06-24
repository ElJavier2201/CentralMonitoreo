<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
%>
<!--#include virtual="/conexion.asp"-->
<%
Dim idFallo, objConn, sql

idFallo = IDValido(Request.QueryString("id"))

If idFallo > 0 Then
    Set objConn = AbrirConexion()
    sql = "DELETE FROM Bitacora_Fallos WHERE ID_Fallo = " & idFallo
    objConn.Execute sql, , 129
    CerrarConexion objConn
End If

Response.Redirect "listar.asp?eliminado=1"
%>
