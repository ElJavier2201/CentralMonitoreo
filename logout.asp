<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' Destruye todas las variables de sesión del usuario actual
Session.Abandon()

' Redirige de vuelta a la pantalla de login
Response.Redirect "login.asp"
%>