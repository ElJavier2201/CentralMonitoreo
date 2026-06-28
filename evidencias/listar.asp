<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "evidencias"
tituloPagina = "Evidencias - Central de Monitoreo"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql, filtroProyecto, nombreProyectoFiltro
filtroProyecto = IDValido(Request.QueryString("proyecto"))
nombreProyectoFiltro = ""
Set objConn = AbrirConexion()
sql = "SELECT e.ID_Evidencia, e.Tipo_Evidencia, e.Ruta_Archivo, e.Descripcion, e.Fecha_Subida, p.ID_Proyecto, p.Nombre_Proyecto " & _
      "FROM Evidencias e INNER JOIN Proyectos p ON e.ID_Proyecto = p.ID_Proyecto "
If filtroProyecto > 0 Then sql = sql & "WHERE p.ID_Proyecto = " & filtroProyecto & " "
sql = sql & "ORDER BY e.Fecha_Subida DESC, e.ID_Evidencia DESC"
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
        <h1>Evidencias</h1>
        <a href="formulario.asp<% If filtroProyecto > 0 Then %>?proyecto=<%= filtroProyecto %><% End If %>" class="boton boton-primario">+ Nueva evidencia</a>
    </div>
    <p class="ayuda">Registra rutas o enlaces a capturas de Proteus/Tinkercad, fotos del protoboard, diagramas o mediciones.
        <% If filtroProyecto > 0 Then %> Mostrando: <strong><%= Server.HTMLEncode(nombreProyectoFiltro) %></strong> &mdash; <a href="listar.asp">ver todos</a><% End If %>
    </p>
    <% If Request.QueryString("ok") = "1" Then %><div class="alerta alerta-ok">Evidencia guardada correctamente.</div><% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %><div class="alerta alerta-ok">Evidencia eliminada correctamente.</div><% End If %>
    <table class="tabla-datos">
        <thead><tr><th>Proyecto</th><th>Tipo</th><th>Ruta / enlace</th><th>Descripción</th><th>Fecha</th><th>Acciones</th></tr></thead>
        <tbody>
        <% If objRS.EOF Then %>
            <tr><td colspan="6" class="vacio">No hay evidencias registradas.</td></tr>
        <% Else
            Do While Not objRS.EOF
        %>
            <tr>
                <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
                <td><span class="etiqueta"><%= Server.HTMLEncode(objRS("Tipo_Evidencia")) %></span></td>
                <td><a href="<%= Server.HTMLEncode(objRS("Ruta_Archivo")) %>" target="_blank"><%= Server.HTMLEncode(objRS("Ruta_Archivo")) %></a></td>
                <td><%= Server.HTMLEncode(objRS("Descripcion")) %></td>
                <td><%= FormatearFechaVisible(objRS("Fecha_Subida")) %></td>
                <td class="acciones-fila">
                    <a href="formulario.asp?id=<%= objRS("ID_Evidencia") %>">Editar</a>
                    <form method="post" action="eliminar.asp" class="form-eliminar" onsubmit="return confirm('¿Eliminar esta evidencia?');">
                        <input type="hidden" name="id_evidencia" value="<%= objRS("ID_Evidencia") %>">
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
