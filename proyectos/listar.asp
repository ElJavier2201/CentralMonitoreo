
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
Dim objConn, objRS, sql, filtroEstado, whereEstado
filtroEstado = Limpiar(Request.QueryString("estado"))
whereEstado = ""
If filtroEstado <> "" Then
    whereEstado = "WHERE p.Estado = '" & EscaparSQL(filtroEstado) & "' "
End If

Set objConn = AbrirConexion()

sql = "SELECT p.ID_Proyecto, p.Nombre_Proyecto, p.Plataforma_Simulacion, p.Microcontrolador, p.Estado, " & _
      "p.Fecha_Creacion, " & _
      "(SELECT COUNT(*) FROM Componentes c WHERE c.ID_Proyecto = p.ID_Proyecto) AS TotalComponentes, " & _
      "(SELECT COUNT(*) FROM Bitacora_Fallos f WHERE f.ID_Proyecto = p.ID_Proyecto) AS TotalFallos, " & _
      "(SELECT COUNT(*) FROM Mediciones m WHERE m.ID_Proyecto = p.ID_Proyecto) AS TotalMediciones, " & _
      "(SELECT COUNT(*) FROM Checklist_Proyecto ch WHERE ch.ID_Proyecto = p.ID_Proyecto AND ch.Completado = False) AS ChecklistPendiente " & _
      "FROM Proyectos p " & whereEstado & "ORDER BY p.Nombre_Proyecto"

Set objRS = objConn.Execute(sql)
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <div class="panel-cabecera">
        <h1>Proyectos</h1>
        <a href="formulario.asp" class="boton boton-primario">+ Nuevo proyecto</a>
    </div>
    <p class="ayuda">Cabecera de cada experimento y acceso rápido a componentes, fallos, mediciones, checklist y reporte técnico.</p>

    <form method="get" action="listar.asp" class="barra-filtros">
        <select name="estado">
            <option value="">Todos los estados</option>
            <option value="Planeado" <% If filtroEstado = "Planeado" Then %>selected<% End If %>>Planeado</option>
            <option value="En simulacion" <% If filtroEstado = "En simulacion" Then %>selected<% End If %>>En simulación</option>
            <option value="En armado fisico" <% If filtroEstado = "En armado fisico" Then %>selected<% End If %>>En armado físico</option>
            <option value="En pruebas" <% If filtroEstado = "En pruebas" Then %>selected<% End If %>>En pruebas</option>
            <option value="Finalizado" <% If filtroEstado = "Finalizado" Then %>selected<% End If %>>Finalizado</option>
            <option value="Cancelado" <% If filtroEstado = "Cancelado" Then %>selected<% End If %>>Cancelado</option>
        </select>
        <button type="submit" class="boton boton-secundario">Filtrar</button>
        <% If filtroEstado <> "" Then %><a href="listar.asp" class="boton boton-secundario">Limpiar</a><% End If %>
    </form>

    <% If Request.QueryString("ok") = "1" Then %><div class="alerta alerta-ok">Proyecto guardado correctamente.</div><% End If %>
    <% If Request.QueryString("eliminado") = "1" Then %><div class="alerta alerta-ok">Proyecto eliminado junto con sus registros asociados.</div><% End If %>
    <% If Request.QueryString("error") = "borrado_fallido" Then %><div class="alerta alerta-error">No se pudo eliminar el proyecto.</div><% End If %>

    <table class="tabla-datos">
        <thead>
        <tr>
            <th>Proyecto</th><th>Plataforma</th><th>Microcontrolador</th><th>Estado</th><th>Creado</th>
            <th>Comp.</th><th>Fallos</th><th>Med.</th><th>Checklist</th><th>Acciones</th>
        </tr>
        </thead>
        <tbody>
        <% If objRS.EOF Then %>
        <tr><td colspan="10" class="vacio">Aún no hay proyectos registrados.</td></tr>
        <% Else
            Do While Not objRS.EOF
        %>
        <tr>
            <td><strong><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></strong></td>
            <td><span class="etiqueta"><%= Server.HTMLEncode(objRS("Plataforma_Simulacion")) %></span></td>
            <td><%= Server.HTMLEncode(objRS("Microcontrolador")) %></td>
            <td><span class="estado-proyecto"><%= Server.HTMLEncode(objRS("Estado")) %></span></td>
            <td><%= FormatearFechaVisible(objRS("Fecha_Creacion")) %></td>
            <td class="centrado"><a href="../componentes/listar.asp?proyecto=<%= objRS("ID_Proyecto") %>"><%= objRS("TotalComponentes") %></a></td>
            <td class="centrado"><a href="../fallos/listar.asp?proyecto=<%= objRS("ID_Proyecto") %>"><%= objRS("TotalFallos") %></a></td>
            <td class="centrado"><a href="../mediciones/listar.asp?proyecto=<%= objRS("ID_Proyecto") %>"><%= objRS("TotalMediciones") %></a></td>
            <td class="centrado"><a href="../checklist/listar.asp?proyecto=<%= objRS("ID_Proyecto") %>"><%= objRS("ChecklistPendiente") %> pend.</a></td>
            <td class="acciones-fila acciones-apiladas">
                <a href="../reportes/proyecto_completo.asp?id=<%= objRS("ID_Proyecto") %>" target="_blank">Reporte completo</a>
                <a href="formulario.asp?id=<%= objRS("ID_Proyecto") %>">Editar</a>
                <a href="imprimir_plano.asp?id=<%= objRS("ID_Proyecto") %>" target="_blank">Plano</a>
                <form method="post" action="eliminar.asp" class="form-eliminar" onsubmit="return confirm('¿Eliminar este proyecto y todos sus componentes, fallos, mediciones, checklist y evidencias?');">
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