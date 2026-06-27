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
Dim idProyecto, objConn, objCmd, exitoBorrado
exitoBorrado = False

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Redirect "listar.asp"
    Response.End
End If

idProyecto = IDValido(Request.Form("id_proyecto"))

If idProyecto > 0 Then
    Set objConn = AbrirConexion()

    On Error Resume Next
    objConn.BeginTrans

    Set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = objConn
    objCmd.CommandType = 1

    objCmd.CommandText = "DELETE FROM Bitacora_Fallos WHERE ID_Proyecto = ?"
    objCmd.Parameters.Append objCmd.CreateParameter("@id", 3, 1, , idProyecto)
    objCmd.Execute , , 128

    objCmd.Parameters.Delete 0
    objCmd.CommandText = "DELETE FROM Componentes WHERE ID_Proyecto = ?"
    objCmd.Parameters.Append objCmd.CreateParameter("@id", 3, 1, , idProyecto)
    objCmd.Execute , , 128

    objCmd.Parameters.Delete 0
    objCmd.CommandText = "DELETE FROM Proyectos WHERE ID_Proyecto = ?"
    objCmd.Parameters.Append objCmd.CreateParameter("@id", 3, 1, , idProyecto)
    objCmd.Execute , , 128

    If Err.Number <> 0 Then
        objConn.RollbackTrans
        Err.Clear
    Else
        objConn.CommitTrans
        exitoBorrado = True
    End If

    On Error Goto 0
    Set objCmd = Nothing
    CerrarConexion objConn
End If

If exitoBorrado Then
    Response.Redirect "listar.asp?eliminado=1"
ElseIf idProyecto > 0 Then
    Response.Redirect "listar.asp?error=borrado_fallido"
Else
    Response.Redirect "listar.asp"
End If
%>
