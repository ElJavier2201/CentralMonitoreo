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
Dim idComponente, objConn, objCmd, exitoBorrado
exitoBorrado = False

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Redirect "listar.asp"
    Response.End
End If

idComponente = IDValido(Request.Form("id_componente"))

If idComponente > 0 Then
    Set objConn = AbrirConexion()

    On Error Resume Next
    Set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = objConn
    objCmd.CommandType = 1
    objCmd.CommandText = "DELETE FROM Componentes WHERE ID_Componente = ?"
    objCmd.Parameters.Append objCmd.CreateParameter("@id", 3, 1, , idComponente)
    objCmd.Execute , , 128

    If Err.Number = 0 Then
        exitoBorrado = True
    Else
        Err.Clear
    End If
    On Error Goto 0

    Set objCmd = Nothing
    CerrarConexion objConn
End If

If exitoBorrado Then
    Response.Redirect "listar.asp?eliminado=1"
ElseIf idComponente > 0 Then
    Response.Redirect "listar.asp?error=borrado_fallido"
Else
    Response.Redirect "listar.asp"
End If
%>
