

<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Const MAX_INTENTOS = 5
Const MINUTOS_BLOQUEO = 10

Dim errorMsg, pass, passwordConfigurada, bloqueado, intentos, bloqueoHasta
errorMsg = ""
bloqueado = False
passwordConfigurada = ObtenerPasswordConfigurada()

If IsNumeric(Session("IntentosLogin")) Then
    intentos = CInt(Session("IntentosLogin"))
Else
    intentos = 0
End If

If IsDate(Session("BloqueoHasta")) Then
    bloqueoHasta = CDate(Session("BloqueoHasta"))
    If bloqueoHasta > Now() Then
        bloqueado = True
        errorMsg = "Demasiados intentos fallidos. Intenta de nuevo después de " & bloqueoHasta & "."
    Else
        Session.Contents.Remove("BloqueoHasta")
        Session("IntentosLogin") = 0
        intentos = 0
    End If
End If

If passwordConfigurada = "" Then
    errorMsg = "La contraseña del sistema no está configurada. Define la variable de entorno CENTRAL_MONITOREO_PASSWORD y reinicia IIS."
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" And Not bloqueado And passwordConfigurada <> "" Then
    pass = Trim(Request.Form("password"))

    If pass = passwordConfigurada Then
        Session.Contents.RemoveAll
        Session("Autenticado") = True
        Session("LoginHora") = Now()
        Session.Timeout = 20

        Dim fsoLogin
        Set fsoLogin = Server.CreateObject("Scripting.FileSystemObject")
        If Not fsoLogin.FileExists(Server.MapPath("db/CentralMonitoreo.accdb")) Then
            Set fsoLogin = Nothing
            Response.Redirect "db/instalar.asp"
        Else
            Set fsoLogin = Nothing
            Response.Redirect "index.asp"
        End If
        Response.End
    Else
        intentos = intentos + 1
        Session("IntentosLogin") = intentos

        If intentos >= MAX_INTENTOS Then
            Session("BloqueoHasta") = DateAdd("n", MINUTOS_BLOQUEO, Now())
            errorMsg = "Demasiados intentos fallidos. Acceso bloqueado por " & MINUTOS_BLOQUEO & " minutos."
        Else
            errorMsg = "Contraseña incorrecta. Intento " & intentos & " de " & MAX_INTENTOS & "."
        End If
    End If
End If

Function ObtenerPasswordConfigurada()
    Dim shell, valor
    valor = ""
    On Error Resume Next
    Set shell = Server.CreateObject("WScript.Shell")
    If Err.Number = 0 Then
        valor = shell.Environment("PROCESS")("CENTRAL_MONITOREO_PASSWORD")
        If valor = "" Then valor = shell.Environment("SYSTEM")("CENTRAL_MONITOREO_PASSWORD")
        If valor = "" Then valor = shell.Environment("USER")("CENTRAL_MONITOREO_PASSWORD")
    End If
    Set shell = Nothing
    Err.Clear
    On Error Goto 0
    ObtenerPasswordConfigurada = valor
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Acceso - Central de Monitoreo</title>
    <link rel="stylesheet" href="includes/estilos.css">
</head>
<body>
    <main class="contenedor" style="max-width: 420px; margin-top: 10vh;">
        <div class="panel">
            <div style="text-align: center; margin-bottom: 20px;">
                <span class="led" style="display: inline-block; margin-bottom: 10px;"></span>
                <h1 style="font-size: 20px;">Central de Monitoreo</h1>
                <p class="ayuda">Simulaciones &amp; Prototipos</p>
            </div>

            <% If errorMsg <> "" Then %>
                <div class="alerta alerta-error"><%= Server.HTMLEncode(errorMsg) %></div>
            <% End If %>

            <form method="post" action="login.asp" autocomplete="off">
                <label for="password">Contraseña de acceso</label>
                <input type="password" id="password" name="password" required autofocus <% If bloqueado Or passwordConfigurada = "" Then %>disabled<% End If %>>

                <div class="acciones" style="margin-top: 20px; justify-content: center;">
                    <button type="submit" class="boton boton-primario" style="width: 100%;" <% If bloqueado Or passwordConfigurada = "" Then %>disabled<% End If %>>Ingresar al sistema</button>
                </div>
            </form>
        </div>
    </main>
</body>
</html>
