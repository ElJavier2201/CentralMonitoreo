<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Cálculos eléctricos - Central de Monitoreo"
seccionActiva = "calculos"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim filtroProyecto, nombreProyectoFiltro
Dim q, whereSql

filtroProyecto = IDValido(Request.QueryString("proyecto"))
q = Limpiar(Request.QueryString("q"))
nombreProyectoFiltro = ""
whereSql = ""

Set objConn = AbrirConexion()

sql = "SELECT c.ID_Calculo, c.ID_Proyecto, c.Tipo_Calculo, c.Entrada, c.Resultado, c.Fecha_Calculo, " & _
      "p.Nombre_Proyecto FROM Calculos_Electricos c INNER JOIN Proyectos p ON c.ID_Proyecto = p.ID_Proyecto "

If filtroProyecto > 0 Then
    whereSql = "WHERE c.ID_Proyecto = " & filtroProyecto & " "
End If

If q <> "" Then
    If whereSql = "" Then
        whereSql = "WHERE "
    Else
        whereSql = whereSql & "AND "
    End If
    whereSql = whereSql & "(c.Tipo_Calculo LIKE '%" & EscaparSQL(q) & "%' OR c.Entrada LIKE '%" & EscaparSQL(q) & "%' OR c.Resultado LIKE '%" & EscaparSQL(q) & "%' OR p.Nombre_Proyecto LIKE '%" & EscaparSQL(q) & "%') "
End If

sql = sql & whereSql & "ORDER BY c.Fecha_Calculo DESC, c.ID_Calculo DESC"
Set objRS = objConn.Execute(sql)

If filtroProyecto > 0 Then
    Dim rsProyecto
    Set rsProyecto = objConn.Execute("SELECT Nombre_Proyecto FROM Proyectos WHERE ID_Proyecto = " & filtroProyecto)
    If Not rsProyecto.EOF Then nombreProyectoFiltro = rsProyecto("Nombre_Proyecto")
    rsProyecto.Close
    Set rsProyecto = Nothing
End If
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <div class="panel-cabecera">
        <h1>Cálculos eléctricos guardados</h1>
        <a href="formulario.asp<% If filtroProyecto > 0 Then %>?proyecto=<%= filtroProyecto %><% End If %>" class="boton boton-primario">+ Nuevo cálculo</a>
    </div>

    <p class="ayuda">
        Guarda cálculos de resistencias, Ley de Ohm, divisor de voltaje y error porcentual asociados a cada proyecto.
        <% If filtroProyecto > 0 Then %>
            Mostrando solo: <strong><%= Server.HTMLEncode(nombreProyectoFiltro) %></strong> &mdash; <a href="listar.asp">ver todos</a>
        <% End If %>
    </p>

    <% If Request.QueryString("ok") = "1" Then %>
        <div class="alerta alerta-ok">Cálculo guardado correctamente.</div>
    <% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %>
        <div class="alerta alerta-ok">Cálculo eliminado correctamente.</div>
    <% End If %>

    <form method="get" action="listar.asp" class="formulario-linea">
        <% If filtroProyecto > 0 Then %><input type="hidden" name="proyecto" value="<%= filtroProyecto %>"><% End If %>
        <input type="text" name="q" value="<%= Server.HTMLEncode(q) %>" placeholder="Buscar cálculo, resultado o proyecto...">
        <button type="submit" class="boton boton-primario">Buscar</button>
        <% If q <> "" Then %><a href="listar.asp" class="boton boton-secundario">Limpiar</a><% End If %>
    </form>

    <table class="tabla-datos">
        <thead>
            <tr>
                <th>Proyecto</th>
                <th>Tipo</th>
                <th>Entrada</th>
                <th>Resultado</th>
                <th>Fecha</th>
                <th>Acciones</th>
            </tr>
        </thead>
        <tbody>
        <% If objRS.EOF Then %>
            <tr><td colspan="6" class="vacio">No hay cálculos guardados.</td></tr>
        <% Else %>
            <% Do While Not objRS.EOF %>
            <tr>
                <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
                <td><span class="etiqueta"><%= Server.HTMLEncode(objRS("Tipo_Calculo")) %></span></td>
                <td><%= Server.HTMLEncode(objRS("Entrada")) %></td>
                <td><strong><%= Server.HTMLEncode(objRS("Resultado")) %></strong></td>
                <td><%= FormatearFechaVisible(objRS("Fecha_Calculo")) %></td>
                <td class="acciones-fila">
                    <a href="formulario.asp?id=<%= objRS("ID_Calculo") %>">Editar</a>
                    <form method="post" action="eliminar.asp" class="form-inline" onsubmit="return confirm('¿Eliminar este cálculo?');">
                        <input type="hidden" name="id" value="<%= objRS("ID_Calculo") %>">
                        <button type="submit" class="boton-link peligro">Eliminar</button>
                    </form>
                </td>
            </tr>
            <% objRS.MoveNext : Loop %>
        <% End If %>
        </tbody>
    </table>
</div>

<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>
