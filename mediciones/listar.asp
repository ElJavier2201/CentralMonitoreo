<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "mediciones"
tituloPagina = "Mediciones - Central de Monitoreo"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql, filtroProyecto, nombreProyectoFiltro
filtroProyecto = IDValido(Request.QueryString("proyecto"))
nombreProyectoFiltro = ""

Set objConn = AbrirConexion()

sql = "SELECT m.ID_Medicion, m.Fecha_Medicion, m.Variable_Medida, m.Valor_Simulado, m.Valor_Fisico, " & _
      "m.Diferencia, m.Observaciones, p.ID_Proyecto, p.Nombre_Proyecto " & _
      "FROM Mediciones m INNER JOIN Proyectos p ON m.ID_Proyecto = p.ID_Proyecto "
If filtroProyecto > 0 Then sql = sql & "WHERE p.ID_Proyecto = " & filtroProyecto & " "
sql = sql & "ORDER BY m.Fecha_Medicion DESC, m.ID_Medicion DESC"
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
        <h1>Mediciones simulación vs físico</h1>
        <div class="acciones">
            <a href="formulario.asp<% If filtroProyecto > 0 Then %>?proyecto=<%= filtroProyecto %><% End If %>" class="boton boton-primario">+ Nueva medición</a>
            <a href="exportar_csv.asp?proyecto=<%= filtroProyecto %>" class="boton boton-secundario">Exportar CSV</a>
        </div>
    </div>
    <p class="ayuda">Registra valores simulados y físicos para comparar resultados reales contra Proteus/Tinkercad.
        <% If filtroProyecto > 0 Then %> Mostrando: <strong><%= Server.HTMLEncode(nombreProyectoFiltro) %></strong> &mdash; <a href="listar.asp">ver todos</a><% End If %>
    </p>

    <% If Request.QueryString("ok") = "1" Then %><div class="alerta alerta-ok">Medición guardada correctamente.</div><% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %><div class="alerta alerta-ok">Medición eliminada correctamente.</div><% End If %>

    <table class="tabla-datos">
        <thead><tr><th>Proyecto</th><th>Fecha</th><th>Variable</th><th>Simulado</th><th>Físico</th><th>Diferencia</th><th>Observaciones</th><th>Acciones</th></tr></thead>
        <tbody>
        <% If objRS.EOF Then %>
            <tr><td colspan="8" class="vacio">No hay mediciones registradas.</td></tr>
        <% Else
            Do While Not objRS.EOF
        %>
            <tr>
                <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
                <td><%= FormatearFechaVisible(objRS("Fecha_Medicion")) %></td>
                <td><strong><%= Server.HTMLEncode(objRS("Variable_Medida")) %></strong></td>
                <td><span class="codigo"><%= Server.HTMLEncode(objRS("Valor_Simulado")) %></span></td>
                <td><span class="codigo"><%= Server.HTMLEncode(objRS("Valor_Fisico")) %></span></td>
                <td><%= Server.HTMLEncode(objRS("Diferencia")) %></td>
                <td><%= Server.HTMLEncode(objRS("Observaciones")) %></td>
                <td class="acciones-fila">
                    <a href="formulario.asp?id=<%= objRS("ID_Medicion") %>">Editar</a>
                    <form method="post" action="eliminar.asp" class="form-eliminar" onsubmit="return confirm('¿Eliminar esta medición?');">
                        <input type="hidden" name="id_medicion" value="<%= objRS("ID_Medicion") %>">
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
