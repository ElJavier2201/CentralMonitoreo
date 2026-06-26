<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim errorMsg
errorMsg = ""

' Si el usuario envió el formulario
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim pass
    pass = Trim(Request.Form("password"))
    
    ' Aquí defines tu contraseña maestra
    If pass = "admin123" Then
        ' Contraseña correcta: Creamos la sesión y redirigimos al dashboard
        Session("Autenticado") = True
        Response.Redirect "index.asp"
        Response.End
    Else
        errorMsg = "Contraseña incorrecta. Inténtalo de nuevo."
    End If
End If
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
                <p class="ayuda">Simulaciones & Prototipos</p>
            </div>
            
            <% If errorMsg <> "" Then %>
                <div class="alerta alerta-error"><%= errorMsg %></div>
            <% End If %>
            
            <form method="post" action="login.asp">
                <label for="password">Contraseña de acceso</label>
                <input type="password" id="password" name="password" required autofocus>
                
                <div class="acciones" style="margin-top: 20px; justify-content: center;">
                    <button type="submit" class="boton boton-primario" style="width: 100%;">Ingresar al sistema</button>
                </div>
            </form>
        </div>
    </main>
</body>
</html>