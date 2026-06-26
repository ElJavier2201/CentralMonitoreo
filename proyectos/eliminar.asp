<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
%>
<%
Dim idProyecto, objConn, sql
Dim exitoBorrado

idProyecto = IDValido(Request.QueryString("id"))
exitoBorrado = False

If idProyecto > 0 Then
    Set objConn = AbrirConexion()

    ' 1. Desactivamos la interrupción automática por errores para manejarlos manualmente
    On Error Resume Next
    
    ' 2. Iniciamos la transacción
    objConn.BeginTrans

    ' 3. Ejecutamos los borrados en cascada explícitos
    sql = "DELETE FROM Bitacora_Fallos WHERE ID_Proyecto = " & idProyecto
    objConn.Execute sql, , 129

    sql = "DELETE FROM Componentes WHERE ID_Proyecto = " & idProyecto
    objConn.Execute sql, , 129

    sql = "DELETE FROM Proyectos WHERE ID_Proyecto = " & idProyecto
    objConn.Execute sql, , 129

    ' 4. Evaluamos si ocurrió algún error en cualquiera de los pasos anteriores
    If Err.Number <> 0 Then
        ' Ocurrió un error: cancelamos la transacción y restauramos los datos
        objConn.RollbackTrans
        Err.Clear ' Limpiamos el objeto de error
    Else
        ' No hubo errores: confirmamos los cambios de forma definitiva
        objConn.CommitTrans
        exitoBorrado = True
    End If
    
    ' 5. Restauramos el comportamiento normal de errores de VBScript
    On Error Goto 0

    CerrarConexion objConn
End If

' 6. Redirigimos según el resultado
If exitoBorrado Then
    Response.Redirect "listar.asp?eliminado=1"
ElseIf idProyecto > 0 Then
    ' Si se intentó borrar pero la transacción falló, enviamos una bandera de error
    Response.Redirect "listar.asp?error=borrado_fallido"
Else
    ' Si entró sin un ID válido
    Response.Redirect "listar.asp"
End If
%>