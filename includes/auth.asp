<%
'==========================================================================
' includes/auth.asp
' Valida que el usuario tenga una sesion autenticada antes de ejecutar
' consultas, descargas o modificaciones.
'
' IMPORTANTE: cada pagina debe definir rutaBase ANTES de incluir este archivo.
'   rutaBase = ""    para paginas en la raiz
'   rutaBase = "../" para paginas dentro de subcarpetas
'==========================================================================

Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

If Session("Autenticado") <> True Then
    Response.Redirect rutaBase & "login.asp"
    Response.End
End If
%>