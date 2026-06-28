
<%
Dim titulo, seccion
If tituloPagina = "" Then
    titulo = "Central de Monitoreo de Circuitos"
Else
    titulo = tituloPagina
End If
If seccionActiva = "" Then
    seccion = ""
Else
    seccion = seccionActiva
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title><%= titulo %></title>
<link rel="stylesheet" href="<%= rutaBase %>includes/estilos.css">
</head>
<body>
<header class="topbar">
    <div class="marca">
        <span class="led"></span>
        <div class="marca-textos">
            <span class="marca-texto">CENTRAL DE MONITOREO</span>
            <span class="marca-sub">Simulaciones &amp; Prototipos</span>
        </div>
    </div>
    <nav class="menu">
        <a href="<%= rutaBase %>index.asp" class="<%= ClaseSiActivo(seccion, "dashboard") %>">Dashboard</a>
        <a href="<%= rutaBase %>proyectos/listar.asp" class="<%= ClaseSiActivo(seccion, "proyectos") %>">Proyectos</a>
        <a href="<%= rutaBase %>componentes/listar.asp" class="<%= ClaseSiActivo(seccion, "componentes") %>">Componentes</a>
        <a href="<%= rutaBase %>fallos/listar.asp" class="<%= ClaseSiActivo(seccion, "fallos") %>">Bitácora de Fallos</a>
        <a href="<%= rutaBase %>mediciones/listar.asp" class="<%= ClaseSiActivo(seccion, "mediciones") %>">Mediciones</a>
        <a href="<%= rutaBase %>checklist/listar.asp" class="<%= ClaseSiActivo(seccion, "checklist") %>">Checklist</a>
        <a href="<%= rutaBase %>evidencias/listar.asp" class="<%= ClaseSiActivo(seccion, "evidencias") %>">Evidencias</a>
        <a href="<%= rutaBase %>busqueda/index.asp" class="<%= ClaseSiActivo(seccion, "busqueda") %>">Buscar</a>
        <a href="<%= rutaBase %>herramientas/formulas.asp" class="<%= ClaseSiActivo(seccion, "herramientas") %>">Herramientas</a>
        <a href="<%= rutaBase %>reportes/index.asp" class="<%= ClaseSiActivo(seccion, "reportes") %>">Reportes</a>
        <a href="<%= rutaBase %>calculos/listar.asp" class="<%= ClaseSiActivo(seccion, "calculos") %>">Cálculos</a>
        <a href="<%= rutaBase %>catalogo_errores/listar.asp" class="<%= ClaseSiActivo(seccion, "catalogo_errores") %>">Errores frecuentes</a>
        <a href="<%= rutaBase %>ayuda/index.asp" class="<%= ClaseSiActivo(seccion, "ayuda") %>">Ayuda</a>
        <a href="<%= rutaBase %>logout.asp" style="color: var(--alerta); margin-left: 15px;">Salir</a>
    </nav>
</header>
<main class="contenedor">
