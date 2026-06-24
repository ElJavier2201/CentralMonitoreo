<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
%>
<!--#include virtual="/conexion.asp"-->
<%
Dim idProyecto, objConn, sql

idProyecto = IDValido(Request.QueryString("id"))

If idProyecto > 0 Then
    Set objConn = AbrirConexion()

    ' Access/Jet no garantiza ON DELETE CASCADE vía OLEDB de forma fiable,
    ' por lo que el borrado en cascada se hace explícito: primero los
    ' registros hijos (Bitacora_Fallos, Componentes) y al final la cabecera.
    sql = "DELETE FROM Bitacora_Fallos WHERE ID_Proyecto = " & idProyecto
    objConn.Execute sql, , 129

    sql = "DELETE FROM Componentes WHERE ID_Proyecto = " & idProyecto
    objConn.Execute sql, , 129

    sql = "DELETE FROM Proyectos WHERE ID_Proyecto = " & idProyecto
    objConn.Execute sql, , 129

    CerrarConexion objConn
End If

Response.Redirect "listar.asp?eliminado=1"
%>
