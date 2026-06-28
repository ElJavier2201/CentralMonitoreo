<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Catálogo de errores - Central de Monitoreo"
seccionActiva = "catalogo_errores"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql, q, whereSql
q = Limpiar(Request.QueryString("q"))
whereSql = ""

Set objConn = AbrirConexion()

sql = "SELECT ID_Error, Nombre_Error, Categoria, Causa_Probable, Solucion_Recomendada, Nivel_Riesgo FROM Catalogo_Errores "
If q <> "" Then
    whereSql = "WHERE Nombre_Error LIKE '%" & EscaparSQL(q) & "%' OR Categoria LIKE '%" & EscaparSQL(q) & "%' OR Causa_Probable LIKE '%" & EscaparSQL(q) & "%' OR Solucion_Recomendada LIKE '%" & EscaparSQL(q) & "%' "
End If
sql = sql & whereSql & "ORDER BY Categoria, Nombre_Error"
Set objRS = objConn.Execute(sql)
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <div class="panel-cabecera">
        <h1>Catálogo de errores frecuentes</h1>
        <a href="formulario.asp" class="boton boton-primario">+ Nuevo error</a>
    </div>
    <p class="ayuda">Biblioteca de problemas comunes de circuitos, causas probables y soluciones recomendadas.</p>

    <% If Request.QueryString("ok") = "1" Then %><div class="alerta alerta-ok">Error guardado correctamente.</div><% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %><div class="alerta alerta-ok">Error eliminado correctamente.</div><% End If %>

    <form method="get" action="listar.asp" class="formulario-linea">
        <input type="text" name="q" value="<%= Server.HTMLEncode(q) %>" placeholder="Buscar por error, causa, solución o categoría...">
        <button type="submit" class="boton boton-primario">Buscar</button>
        <% If q <> "" Then %><a href="listar.asp" class="boton boton-secundario">Limpiar</a><% End If %>
    </form>

    <table class="tabla-datos">
        <thead>
            <tr>
                <th>Error</th>
                <th>Categoría</th>
                <th>Causa probable</th>
                <th>Solución recomendada</th>
                <th>Riesgo</th>
                <th>Acciones</th>
            </tr>
        </thead>
        <tbody>
        <% If objRS.EOF Then %>
            <tr><td colspan="6" class="vacio">No hay errores registrados en el catálogo.</td></tr>
        <% Else %>
            <% Do While Not objRS.EOF %>
            <tr>
                <td><strong><%= Server.HTMLEncode(objRS("Nombre_Error")) %></strong></td>
                <td><span class="etiqueta"><%= Server.HTMLEncode(objRS("Categoria")) %></span></td>
                <td><%= Server.HTMLEncode(objRS("Causa_Probable")) %></td>
                <td><%= Server.HTMLEncode(objRS("Solucion_Recomendada")) %></td>
                <td><%= Server.HTMLEncode(objRS("Nivel_Riesgo")) %></td>
                <td class="acciones-fila">
                    <a href="formulario.asp?id=<%= objRS("ID_Error") %>">Editar</a>
                    <form method="post" action="eliminar.asp" class="form-inline" onsubmit="return confirm('¿Eliminar este error del catálogo?');">
                        <input type="hidden" name="id" value="<%= objRS("ID_Error") %>">
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
