<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "proyectos"
tituloPagina = "Proyectos - Central de Monitoreo"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql

Set objConn = AbrirConexion()

sql = "SELECT p.ID_Proyecto, p.Nombre_Proyecto, p.Plataforma_Simulacion, p.Microcontrolador, " & _
      "p.Fecha_Creacion, " & _
      "(SELECT COUNT(*) FROM Componentes c WHERE c.ID_Proyecto = p.ID_Proyecto) AS TotalComponentes, " & _
      "(SELECT COUNT(*) FROM Bitacora_Fallos f WHERE f.ID_Proyecto = p.ID_Proyecto) AS TotalFallos " & _
      "FROM Proyectos p ORDER BY p.Nombre_Proyecto"

Set objRS = objConn.Execute(sql)
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <div class="panel-cabecera">
        <h1>Proyectos</h1>
        <a href="formulario.asp" class="boton boton-primario">+ Nuevo proyecto</a>
    </div>
    <p class="ayuda">Cabecera de cada experimento: nombre, plataforma de simulación y cerebro del circuito.</p>

    <% If Request.QueryString("ok") = "1" Then %>
        <div class="alerta alerta-ok">Proyecto guardado correctamente.</div>
    <% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %>
        <div class="alerta alerta-ok">Proyecto eliminado junto con sus componentes y fallos asociados.</div>
    <% End If %>
    <% If Request.QueryString("error") = "borrado_fallido" Then %>
        <div class="alerta alerta-error">No se pudo eliminar el proyecto.</div>
    <% End If %>

    <table class="tabla-datos">
        <thead>
        <tr>
            <th>Proyecto</th>
            <th>Plataforma</th>
            <th>Microcontrolador</th>
            <th>Creado</th>
            <th>Componentes</th>
            <th>Fallos</th>
            <th>Acciones</th>
        </tr>
        </thead>
        <tbody>
        <% If objRS.EOF Then %>
        <tr><td colspan="7" class="vacio">Aún no hay proyectos registrados.</td></tr>
        <% Else
            Do While Not objRS.EOF
        %>
        <tr>
            <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
            <td><span class="etiqueta"><%= Server.HTMLEncode(objRS("Plataforma_Simulacion")) %></span></td>
            <td><%= Server.HTMLEncode(objRS("Microcontrolador")) %></td>
            <td><%= FormatearFechaVisible(objRS("Fecha_Creacion")) %></td>
            <td class="centrado">
                <a href="../componentes/listar.asp?proyecto=<%= objRS("ID_Proyecto") %>"><%= objRS("TotalComponentes") %></a>
            </td>
            <td class="centrado">
                <a href="../fallos/listar.asp?proyecto=<%= objRS("ID_Proyecto") %>"><%= objRS("TotalFallos") %></a>
            </td>
            <td class="acciones-fila">
                <a href="imprimir_plano.asp?id=<%= objRS("ID_Proyecto") %>" target="_blank">Imprimir</a>
                <a href="formulario.asp?id=<%= objRS("ID_Proyecto") %>">Editar</a>
                <form method="post" action="eliminar.asp" class="form-eliminar" onsubmit="return confirm('¿Eliminar este proyecto y todos sus componentes y fallos asociados?');">
                    <input type="hidden" name="id_proyecto" value="<%= objRS("ID_Proyecto") %>">
                    <button type="submit" class="boton-enlace enlace-peligro">Eliminar</button>
                </form>
            </td>
        </tr>
        <%
                objRS.MoveNext
            Loop
        End If
        %>
        </tbody>
    </table>
</div>

<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>
