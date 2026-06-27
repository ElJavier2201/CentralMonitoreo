<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "fallos"
tituloPagina = "Bitacora de Fallos - Central de Monitoreo"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, objCmd, sql
Dim filtroProyecto, nombreProyectoFiltro, claseEstado

filtroProyecto = IDValido(Request.QueryString("proyecto"))
nombreProyectoFiltro = ""

Set objConn = AbrirConexion()
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = objConn
objCmd.CommandType = 1

sql = "SELECT f.ID_Fallo, f.Sintoma_Error, f.Solucion_Aplicada, f.Estado, f.Fecha_Registro, " & _
      "p.ID_Proyecto, p.Nombre_Proyecto " & _
      "FROM Bitacora_Fallos f INNER JOIN Proyectos p ON f.ID_Proyecto = p.ID_Proyecto "

If filtroProyecto > 0 Then
    sql = sql & "WHERE p.ID_Proyecto = ? "
    objCmd.Parameters.Append objCmd.CreateParameter("@proyecto", 3, 1, , filtroProyecto)
End If

objCmd.CommandText = sql & "ORDER BY f.Fecha_Registro DESC, f.ID_Fallo DESC"
Set objRS = objCmd.Execute

If filtroProyecto > 0 Then
    Dim objRSNombre, cmdNombre
    Set cmdNombre = Server.CreateObject("ADODB.Command")
    cmdNombre.ActiveConnection = objConn
    cmdNombre.CommandType = 1
    cmdNombre.CommandText = "SELECT Nombre_Proyecto FROM Proyectos WHERE ID_Proyecto = ?"
    cmdNombre.Parameters.Append cmdNombre.CreateParameter("@id", 3, 1, , filtroProyecto)
    Set objRSNombre = cmdNombre.Execute
    If Not objRSNombre.EOF Then nombreProyectoFiltro = objRSNombre("Nombre_Proyecto")
    objRSNombre.Close
    Set objRSNombre = Nothing
    Set cmdNombre = Nothing
End If
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <div class="panel-cabecera">
        <h1>Bitácora de Fallos</h1>
        <a href="formulario.asp<% If filtroProyecto > 0 Then %>?proyecto=<%= filtroProyecto %><% End If %>" class="boton boton-primario">+ Nuevo registro</a>
    </div>
    <p class="ayuda">
        Síntomas detectados durante las pruebas físicas y la solución electrónica o de código aplicada.
        <% If filtroProyecto > 0 Then %>
            Mostrando solo: <strong><%= Server.HTMLEncode(nombreProyectoFiltro) %></strong>
            &mdash; <a href="listar.asp">ver todos</a>
        <% End If %>
    </p>

    <% If Request.QueryString("ok") = "1" Then %>
        <div class="alerta alerta-ok">Registro de fallo guardado correctamente.</div>
    <% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %>
        <div class="alerta alerta-ok">Registro eliminado correctamente.</div>
    <% End If %>
    <% If Request.QueryString("error") = "borrado_fallido" Then %>
        <div class="alerta alerta-error">No se pudo eliminar el registro.</div>
    <% End If %>

    <table class="tabla-datos">
        <thead>
        <tr>
            <th>Proyecto</th>
            <th>Fecha</th>
            <th>Síntoma</th>
            <th>Solución aplicada</th>
            <th>Estado</th>
            <th>Acciones</th>
        </tr>
        </thead>
        <tbody>
        <% If objRS.EOF Then %>
        <tr><td colspan="6" class="vacio">No hay fallos registrados todavía.</td></tr>
        <% Else
            Do While Not objRS.EOF
                If Limpiar(objRS("Estado")) = "Resuelto" Then
                    claseEstado = "estado-resuelto"
                Else
                    claseEstado = "estado-pendiente"
                End If
        %>
        <tr>
            <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
            <td><%= FormatearFechaVisible(objRS("Fecha_Registro")) %></td>
            <td><%= Server.HTMLEncode(objRS("Sintoma_Error")) %></td>
            <td><%= Server.HTMLEncode(objRS("Solucion_Aplicada")) %></td>
            <td class="<%= claseEstado %>"><%= Server.HTMLEncode(objRS("Estado")) %></td>
            <td class="acciones-fila">
                <a href="formulario.asp?id=<%= objRS("ID_Fallo") %>">Editar</a>
                <form method="post" action="eliminar.asp" class="form-eliminar" onsubmit="return confirm('¿Eliminar este registro de la bitácora?');">
                    <input type="hidden" name="id_fallo" value="<%= objRS("ID_Fallo") %>">
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
Set objCmd = Nothing
CerrarConexion objConn
%>
