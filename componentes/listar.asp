

<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "componentes"
tituloPagina = "Componentes - Central de Monitoreo"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, objCmd, sql
Dim filtroProyecto, nombreProyectoFiltro, terminoBusqueda, clausulaWhere
Dim tamanoPagina, paginaActual, totalPaginas, registrosMostrados

filtroProyecto = IDValido(Request.QueryString("proyecto"))
terminoBusqueda = Limpiar(Request.QueryString("q"))
nombreProyectoFiltro = ""
clausulaWhere = ""

tamanoPagina = 20
paginaActual = IDValido(Request.QueryString("p"))
If paginaActual < 1 Then paginaActual = 1

Set objConn = AbrirConexion()
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = objConn
objCmd.CommandType = 1

sql = "SELECT c.ID_Componente, c.Tipo_Componente, c.Valor_Calculado, c.Pin_Conexion, " & _
      "c.Ubicacion_Protoboard, p.ID_Proyecto, p.Nombre_Proyecto " & _
      "FROM Componentes c INNER JOIN Proyectos p ON c.ID_Proyecto = p.ID_Proyecto "

If filtroProyecto > 0 Then
    clausulaWhere = "WHERE p.ID_Proyecto = ? "
    objCmd.Parameters.Append objCmd.CreateParameter("@proyecto", 3, 1, , filtroProyecto)
End If

If terminoBusqueda <> "" Then
    If clausulaWhere = "" Then
        clausulaWhere = "WHERE "
    Else
        clausulaWhere = clausulaWhere & "AND "
    End If

    clausulaWhere = clausulaWhere & "(c.Tipo_Componente LIKE ? OR c.Valor_Calculado LIKE ?) "
    objCmd.Parameters.Append objCmd.CreateParameter("@q1", 200, 1, 150, "%" & terminoBusqueda & "%")
    objCmd.Parameters.Append objCmd.CreateParameter("@q2", 200, 1, 150, "%" & terminoBusqueda & "%")
End If

objCmd.CommandText = sql & clausulaWhere & "ORDER BY p.Nombre_Proyecto, c.Tipo_Componente"

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorLocation = 3
objRS.Open objCmd, , 3, 1

If Not objRS.EOF Then
    objRS.PageSize = tamanoPagina
    totalPaginas = objRS.PageCount
    If paginaActual > totalPaginas Then paginaActual = totalPaginas
    objRS.AbsolutePage = paginaActual
Else
    totalPaginas = 0
End If

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
        <h1>Componentes &amp; mapa de cableado</h1>
        <div class="acciones" style="margin-top:0;">
            <a href="formulario.asp<% If filtroProyecto > 0 Then %>?proyecto=<%= filtroProyecto %><% End If %>" class="boton boton-primario">+ Nuevo componente</a>
            <a href="exportar_csv.asp?proyecto=<%= filtroProyecto %>&q=<%= Server.URLEncode(terminoBusqueda) %>" class="boton boton-secundario">Exportar a Excel (CSV)</a>
        </div>
    </div>

    <form method="get" action="listar.asp" class="form-busqueda">
        <% If filtroProyecto > 0 Then %><input type="hidden" name="proyecto" value="<%= filtroProyecto %>"><% End If %>
        <input type="text" name="q" placeholder="Buscar componente o valor (Ej: Resistencia, 10k)..." value="<%= Server.HTMLEncode(terminoBusqueda) %>" style="flex:1;">
        <button type="submit" class="boton boton-primario">Buscar</button>
        <% If terminoBusqueda <> "" Then %>
            <a href="listar.asp<% If filtroProyecto > 0 Then %>?proyecto=<%= filtroProyecto %><% End If %>" class="boton boton-secundario">Limpiar filtro</a>
        <% End If %>
    </form>

    <p class="ayuda">
        Inventario por proyecto: tipo de pieza, valor calculado y ubicación exacta en la protoboard.
        <% If filtroProyecto > 0 Then %>
            Mostrando solo: <strong><%= Server.HTMLEncode(nombreProyectoFiltro) %></strong>
            &mdash; <a href="listar.asp">ver todos</a>
        <% End If %>
    </p>

    <% If Request.QueryString("ok") = "1" Then %>
        <div class="alerta alerta-ok">Componente guardado correctamente.</div>
    <% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %>
        <div class="alerta alerta-ok">Componente eliminado correctamente.</div>
    <% End If %>
    <% If Request.QueryString("error") = "borrado_fallido" Then %>
        <div class="alerta alerta-error">No se pudo eliminar el componente.</div>
    <% End If %>

    <table class="tabla-datos">
        <thead>
            <tr>
                <th>Proyecto</th>
                <th>Componente</th>
                <th>Valor calculado</th>
                <th>Pin</th>
                <th>Ubicación en protoboard</th>
                <th>Acciones</th>
            </tr>
        </thead>
        <tbody>
            <% If objRS.EOF Then %>
                <tr><td colspan="6" class="vacio">No hay componentes registrados.</td></tr>
            <% Else
                registrosMostrados = 0
                Do While Not objRS.EOF And registrosMostrados < tamanoPagina
            %>
                    <tr>
                        <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
                        <td><%= Server.HTMLEncode(objRS("Tipo_Componente")) %></td>
                        <td><%= Server.HTMLEncode(objRS("Valor_Calculado")) %></td>
                        <td><span class="codigo"><%= Server.HTMLEncode(objRS("Pin_Conexion")) %></span></td>
                        <td><%= Server.HTMLEncode(objRS("Ubicacion_Protoboard")) %></td>
                        <td class="acciones-fila">
                            <a href="formulario.asp?id=<%= objRS("ID_Componente") %>">Editar</a>
                            <form method="post" action="eliminar.asp" class="form-eliminar" onsubmit="return confirm('¿Eliminar este componente?');">
                                <input type="hidden" name="id_componente" value="<%= objRS("ID_Componente") %>">
                                <button type="submit" class="boton-enlace enlace-peligro">Eliminar</button>
                            </form>
                        </td>
                    </tr>
            <%
                    registrosMostrados = registrosMostrados + 1
                    objRS.MoveNext
                Loop
            End If
            %>
        </tbody>
    </table>

    <% If totalPaginas > 1 Then %>
        <div style="margin-top: 20px; display: flex; justify-content: center; align-items: center; gap: 15px;">
            <% If paginaActual > 1 Then %>
                <a href="listar.asp?p=<%= paginaActual - 1 %><% If filtroProyecto > 0 Then %>&proyecto=<%= filtroProyecto %><% End If %><% If terminoBusqueda <> "" Then %>&q=<%= Server.URLEncode(terminoBusqueda) %><% End If %>" class="boton boton-secundario">&laquo; Anterior</a>
            <% End If %>

            <span style="color: var(--texto-tenue); font-size: 14px;">Página <%= paginaActual %> de <%= totalPaginas %></span>

            <% If paginaActual < totalPaginas Then %>
                <a href="listar.asp?p=<%= paginaActual + 1 %><% If filtroProyecto > 0 Then %>&proyecto=<%= filtroProyecto %><% End If %><% If terminoBusqueda <> "" Then %>&q=<%= Server.URLEncode(terminoBusqueda) %><% End If %>" class="boton boton-secundario">Siguiente &raquo;</a>
            <% End If %>
        </div>
    <% End If %>
</div>

<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
Set objCmd = Nothing
CerrarConexion objConn
%>
