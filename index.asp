
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = ""
seccionActiva = "dashboard"
tituloPagina = "Dashboard - Central de Monitoreo"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim totalProyectos, totalComponentes, totalFallosPendientes, totalFallosResueltos
Dim totalMediciones, totalEvidencias, checklistPendiente
Dim pPlaneado, pSim, pArmado, pPruebas, pFin, pCancelado

pPlaneado = 0
pSim = 0
pArmado = 0
pPruebas = 0
pFin = 0
pCancelado = 0

Set objConn = AbrirConexion()

Set objRS = objConn.Execute("SELECT COUNT(*) AS Total FROM Proyectos")
totalProyectos = objRS("Total")
objRS.Close

Set objRS = objConn.Execute("SELECT COUNT(*) AS Total FROM Componentes")
totalComponentes = objRS("Total")
objRS.Close

Set objRS = objConn.Execute("SELECT COUNT(*) AS Total FROM Bitacora_Fallos WHERE Estado = 'Pendiente'")
totalFallosPendientes = objRS("Total")
objRS.Close

Set objRS = objConn.Execute("SELECT COUNT(*) AS Total FROM Bitacora_Fallos WHERE Estado = 'Resuelto'")
totalFallosResueltos = objRS("Total")
objRS.Close

Set objRS = objConn.Execute("SELECT COUNT(*) AS Total FROM Mediciones")
totalMediciones = objRS("Total")
objRS.Close

Set objRS = objConn.Execute("SELECT COUNT(*) AS Total FROM Evidencias")
totalEvidencias = objRS("Total")
objRS.Close

Set objRS = objConn.Execute("SELECT COUNT(*) AS Total FROM Checklist_Proyecto WHERE Completado = False")
checklistPendiente = objRS("Total")
objRS.Close

Set objRS = objConn.Execute("SELECT Estado, COUNT(*) AS Total FROM Proyectos GROUP BY Estado")
Do While Not objRS.EOF
    Select Case Limpiar(objRS("Estado"))
        Case "Planeado"
            pPlaneado = objRS("Total")
        Case "En simulacion"
            pSim = objRS("Total")
        Case "En armado fisico"
            pArmado = objRS("Total")
        Case "En pruebas"
            pPruebas = objRS("Total")
        Case "Finalizado"
            pFin = objRS("Total")
        Case "Cancelado"
            pCancelado = objRS("Total")
    End Select
    objRS.MoveNext
Loop
objRS.Close

sql = "SELECT TOP 5 ID_Proyecto, Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Estado, Fecha_Creacion " & _
      "FROM Proyectos ORDER BY ID_Proyecto DESC"
Set objRS = objConn.Execute(sql)
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <h1>Panel de Monitoreo</h1>
    <p class="ayuda">Bitácora técnica del paso a paso entre el diseño en simulador, el armado físico y la validación del circuito.</p>

    <div class="tarjetas">
        <div class="tarjeta">
            <span class="tarjeta-numero"><%= totalProyectos %></span>
            <span class="tarjeta-etiqueta">Proyectos</span>
        </div>
        <div class="tarjeta">
            <span class="tarjeta-numero"><%= totalComponentes %></span>
            <span class="tarjeta-etiqueta">Componentes</span>
        </div>
        <div class="tarjeta tarjeta-alerta">
            <span class="tarjeta-numero"><%= totalFallosPendientes %></span>
            <span class="tarjeta-etiqueta">Fallos pendientes</span>
        </div>
        <div class="tarjeta">
            <span class="tarjeta-numero"><%= totalMediciones %></span>
            <span class="tarjeta-etiqueta">Mediciones</span>
        </div>
        <div class="tarjeta">
            <span class="tarjeta-numero"><%= checklistPendiente %></span>
            <span class="tarjeta-etiqueta">Checklist pendiente</span>
        </div>
        <div class="tarjeta">
            <span class="tarjeta-numero"><%= totalEvidencias %></span>
            <span class="tarjeta-etiqueta">Evidencias</span>
        </div>
    </div>

    <h2>Resumen de fallos</h2>
    <table class="tabla-datos">
        <thead>
        <tr>
            <th>Estado</th>
            <th>Total</th>
        </tr>
        </thead>
        <tbody>
        <tr>
            <td>Pendientes</td>
            <td><%= totalFallosPendientes %></td>
        </tr>
        <tr>
            <td>Resueltos</td>
            <td><%= totalFallosResueltos %></td>
        </tr>
        </tbody>
    </table>

    <h2>Proyectos por estado</h2>
    <table class="tabla-datos">
        <thead>
        <tr>
            <th>Estado</th>
            <th>Total</th>
        </tr>
        </thead>
        <tbody>
        <tr><td>Planeado</td><td><%= pPlaneado %></td></tr>
        <tr><td>En simulación</td><td><%= pSim %></td></tr>
        <tr><td>En armado físico</td><td><%= pArmado %></td></tr>
        <tr><td>En pruebas</td><td><%= pPruebas %></td></tr>
        <tr><td>Finalizado</td><td><%= pFin %></td></tr>
        <tr><td>Cancelado</td><td><%= pCancelado %></td></tr>
        </tbody>
    </table>

    <h2>Últimos proyectos</h2>
    <table class="tabla-datos">
        <thead>
        <tr>
            <th>Proyecto</th>
            <th>Plataforma</th>
            <th>Microcontrolador</th>
            <th>Estado</th>
            <th>Creado</th>
            <th></th>
        </tr>
        </thead>
        <tbody>
        <%
        If objRS.EOF Then
        %>
        <tr>
            <td colspan="6" class="vacio">No hay proyectos registrados todavía. <a href="proyectos/formulario.asp">Crear el primero</a>.</td>
        </tr>
        <%
        Else
            Do While Not objRS.EOF
        %>
        <tr>
            <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
            <td><span class="etiqueta"><%= Server.HTMLEncode(objRS("Plataforma_Simulacion")) %></span></td>
            <td><%= Server.HTMLEncode(objRS("Microcontrolador")) %></td>
            <td><span class="estado-proyecto"><%= Server.HTMLEncode(objRS("Estado")) %></span></td>
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
    <p class="ayuda" style="margin-top:18px;">¿Primera vez aquí? Ejecuta <a href="db/instalar.asp">db/instalar.asp</a> y después <a href="db/semilla.asp">db/semilla.asp</a>.</p>
    <% End If %>
</div>

<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>