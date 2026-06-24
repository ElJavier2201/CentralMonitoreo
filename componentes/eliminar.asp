<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
%>
<!--#include virtual="/conexion.asp"-->
<%
Dim idComponente, objConn, sql

idComponente = IDValido(Request.QueryString("id"))

If idComponente > 0 Then
    Set objConn = AbrirConexion()
    sql = "DELETE FROM Componentes WHERE ID_Componente = " & idComponente
    objConn.Execute sql, , 129
    CerrarConexion objConn
End If

Response.Redirect "listar.asp?eliminado=1"
%>
