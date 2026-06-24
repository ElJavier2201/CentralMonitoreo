<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
%>
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim totalProyectos, totalComponentes, totalFallosPendientes
Dim rutaBase, tituloPagina, seccionActiva

rutaBase = ""
seccionActiva = "dashboard"
tituloPagina = "Dashboard - Central de Monitoreo"

Set objConn = AbrirConexion()

sql = "SELECT COUNT(*) AS Total FROM Proyectos"
Set objRS = objConn.Execute(sql)
totalProyectos = objRS("Total")
objRS.Close

sql = "SELECT COUNT(*) AS Total FROM Componentes"
Set objRS = objConn.Execute(sql)
totalComponentes = objRS("Total")
objRS.Close

sql = "SELECT COUNT(*) AS Total FROM Bitacora_Fallos WHERE Estado = 'Pendiente'"
Set objRS = objConn.Execute(sql)
totalFallosPendientes = objRS("Total")
objRS.Close

sql = "SELECT TOP 5 ID_Proyecto, Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Fecha_Creacion " & _
      "FROM Proyectos ORDER BY ID_Proyecto DESC"
Set objRS = objConn.Execute(sql)
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <h1>Panel de Monitoreo</h1>
    <p class="ayuda">Bitácora técnica del paso a paso entre el diseño en simulador y el ensamblaje físico del hardware.</p>

    <div class="tarjetas">
        <div class="tarjeta">
            <span class="tarjeta-numero"><%= totalProyectos %></span>
            <span class="tarjeta-etiqueta">Proyectos activos</span>
        </div>
        <div class="tarjeta">
            <span class="tarjeta-numero"><%= totalComponentes %></span>
            <span class="tarjeta-etiqueta">Componentes mapeados</span>
        </div>
        <div class="tarjeta tarjeta-alerta">
            <span class="tarjeta-numero"><%= totalFallosPendientes %></span>
            <span class="tarjeta-etiqueta">Fallos pendientes</span>
        </div>
    </div>

    <h2>Últimos proyectos</h2>
    <table class="tabla-datos">
        <thead>
        <tr><th>Proyecto</th><th>Plataforma</th><th>Microcontrolador</th><th>Creado</th><th></th></tr>
        </thead>
        <tbody>
        <%
        If objRS.EOF Then
        %>
        <tr><td colspan="5" class="vacio">No hay proyectos registrados todavía. <a href="proyectos/formulario.asp">Crear el primero</a>.</td></tr>
        <%
        Else
            Do While Not objRS.EOF
        %>
        <tr>
            <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
            <td><span class="etiqueta"><%= Server.HTMLEncode(objRS("Plataforma_Simulacion")) %></span></td>
            <td><%= Server.HTMLEncode(objRS("Microcontrolador")) %></td>
            <td><%= FormatearFechaVisible(objRS("Fecha_Creacion")) %></td>
            <td><a href="proyectos/formulario.asp?id=<%= objRS("ID_Proyecto") %>">Ver / editar</a></td>
        </tr>
        <%
                objRS.MoveNext
            Loop
        End If
        %>
        </tbody>
    </table>

    <% If totalProyectos = 0 Then %>
    <p class="ayuda" style="margin-top:18px;">
        ¿Primera vez aquí? Ejecuta primero <a href="db/instalar.asp">db/instalar.asp</a> para crear la base de datos,
        y opcionalmente <a href="db/semilla.asp">db/semilla.asp</a> para cargar un proyecto de ejemplo.
    </p>
    <% End If %>
</div>

<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>
