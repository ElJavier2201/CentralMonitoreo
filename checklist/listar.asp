<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "checklist"
tituloPagina = "Checklist de armado - Central de Monitoreo"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql, filtroProyecto, nombreProyectoFiltro
filtroProyecto = IDValido(Request.QueryString("proyecto"))
nombreProyectoFiltro = ""
Set objConn = AbrirConexion()

sql = "SELECT ch.ID_Item, ch.Descripcion, ch.Completado, ch.Fecha_Actualizacion, p.ID_Proyecto, p.Nombre_Proyecto " & _
      "FROM Checklist_Proyecto ch INNER JOIN Proyectos p ON ch.ID_Proyecto = p.ID_Proyecto "
If filtroProyecto > 0 Then sql = sql & "WHERE p.ID_Proyecto = " & filtroProyecto & " "
sql = sql & "ORDER BY p.Nombre_Proyecto, ch.Completado, ch.ID_Item"
Set objRS = objConn.Execute(sql)

If filtroProyecto > 0 Then
    Dim objRSNombre
    Set objRSNombre = objConn.Execute("SELECT Nombre_Proyecto FROM Proyectos WHERE ID_Proyecto = " & filtroProyecto)
    If Not objRSNombre.EOF Then nombreProyectoFiltro = objRSNombre("Nombre_Proyecto")
    objRSNombre.Close
    Set objRSNombre = Nothing
End If
%>
<!--#include virtual="/includes/header.asp"-->
<div class="panel">
    <div class="panel-cabecera">
        <h1>Checklist de armado</h1>
        <a href="formulario.asp<% If filtroProyecto > 0 Then %>?proyecto=<%= filtroProyecto %><% End If %>" class="boton boton-primario">+ Nuevo punto</a>
    </div>
    <p class="ayuda">Lista de verificación para pasar de simulación a protoboard/hardware real.
        <% If filtroProyecto > 0 Then %> Mostrando: <strong><%= Server.HTMLEncode(nombreProyectoFiltro) %></strong> &mdash; <a href="listar.asp">ver todos</a><% End If %>
    </p>
    <% If Request.QueryString("ok") = "1" Then %><div class="alerta alerta-ok">Checklist actualizado correctamente.</div><% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %><div class="alerta alerta-ok">Punto eliminado correctamente.</div><% End If %>
    <table class="tabla-datos">
        <thead><tr><th>Proyecto</th><th>Estado</th><th>Punto de revisión</th><th>Última actualización</th><th>Acciones</th></tr></thead>
        <tbody>
        <% If objRS.EOF Then %>
            <tr><td colspan="5" class="vacio">No hay puntos de checklist registrados.</td></tr>
        <% Else
            Do While Not objRS.EOF
        %>
            <tr>
                <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
                <td>
                    <% If CBool(objRS("Completado")) Then %>
                        <span class="estado-resuelto">Completado</span>
                    <% Else %>
                        <span class="estado-pendiente">Pendiente</span>
                    <% End If %>
                </td>
                <td><%= Server.HTMLEncode(objRS("Descripcion")) %></td>
                <td><%= FormatearFechaVisible(objRS("Fecha_Actualizacion")) %></td>
                <td class="acciones-fila">
                    <form method="post" action="toggle.asp" class="form-eliminar">
                        <input type="hidden" name="id_item" value="<%= objRS("ID_Item") %>">
                        <button type="submit" class="boton-enlace"><% If CBool(objRS("Completado")) Then %>Marcar pendiente<% Else %>Completar<% End If %></button>
                    </form>
                    <a href="formulario.asp?id=<%= objRS("ID_Item") %>">Editar</a>
                    <form method="post" action="eliminar.asp" class="form-eliminar" onsubmit="return confirm('¿Eliminar este punto del checklist?');">
                        <input type="hidden" name="id_item" value="<%= objRS("ID_Item") %>">
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
