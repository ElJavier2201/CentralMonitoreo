<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Reportes - Central de Monitoreo"
seccionActiva = "reportes"
%>
<!--#include virtual="/conexion.asp"-->
<!--#include virtual="/includes/auth.asp"-->
<%
Dim objConn, objRS, sql
Set objConn = AbrirConexion()
sql = "SELECT ID_Proyecto, Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Fecha_Creacion FROM Proyectos ORDER BY Nombre_Proyecto"
Set objRS = objConn.Execute(sql)
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <div class="panel-cabecera">
        <h1>Reportes técnicos</h1>
    </div>
    <p class="ayuda">Genera documentación imprimible por proyecto con componentes, fallos, mediciones, checklist y evidencias.</p>

    <table class="tabla-datos">
        <thead>
            <tr>
                <th>Proyecto</th>
                <th>Plataforma</th>
                <th>Microcontrolador</th>
                <th>Creado</th>
                <th>Reportes</th>
            </tr>
        </thead>
        <tbody>
        <% If objRS.EOF Then %>
            <tr><td colspan="5" class="vacio">No hay proyectos registrados.</td></tr>
        <% Else %>
            <% Do While Not objRS.EOF %>
                <tr>
                    <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
                    <td><%= Server.HTMLEncode(objRS("Plataforma_Simulacion")) %></td>
                    <td><%= Server.HTMLEncode(objRS("Microcontrolador")) %></td>
                    <td><%= FormatearFechaVisible(objRS("Fecha_Creacion")) %></td>
                    <td class="acciones-fila">
                        <a href="proyecto_completo.asp?id=<%= objRS("ID_Proyecto") %>" target="_blank">Reporte completo</a>
                        <a href="../proyectos/imprimir_plano.asp?id=<%= objRS("ID_Proyecto") %>" target="_blank">Plano de armado</a>
                    </td>
                </tr>
            <%
                objRS.MoveNext
            Loop
            %>
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
