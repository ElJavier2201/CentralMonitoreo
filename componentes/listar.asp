<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% 
Option Explicit

Dim objConn, objRS, sql
Dim rutaBase, tituloPagina, seccionActiva 
Dim filtroProyecto, nombreProyectoFiltro
Dim terminoBusqueda, clausulaWhere

rutaBase = "../"
seccionActiva = "componentes"
tituloPagina = "Componentes - Central de Monitoreo" 
filtroProyecto = IDValido(Request.QueryString("proyecto"))
terminoBusqueda = Limpiar(Request.QueryString("q")) ' Capturamos lo que el usuario escribe
nombreProyectoFiltro = ""

Set objConn = AbrirConexion()

' --- Cruce de tablas obligatorio: Componentes INNER JOIN Proyectos ---
sql = "SELECT c.ID_Componente, c.Tipo_Componente, c.Valor_Calculado, c.Pin_Conexion, " & _
      "c.Ubicacion_Protoboard, p.ID_Proyecto, p.Nombre_Proyecto " & _
      "FROM Componentes c INNER JOIN Proyectos p ON c.ID_Proyecto = p.ID_Proyecto " 

clausulaWhere = "" 

If filtroProyecto > 0 Then
    clausulaWhere = "WHERE p.ID_Proyecto = " & filtroProyecto & " "
End If

If terminoBusqueda <> "" Then
    If clausulaWhere = "" Then
        clausulaWhere = "WHERE "
    Else
        clausulaWhere = clausulaWhere & "AND "
    End If

    ' Buscamos coincidencias en el tipo o en el valor, protegiendo con EscaparSQL
    clausulaWhere = clausulaWhere & "(c.Tipo_Componente LIKE '%" & EscaparSQL(terminoBusqueda) & "%' " & _
                                    "OR c.Valor_Calculado LIKE '%" & EscaparSQL(terminoBusqueda) & "%') "
End If

' Unimos todo (Se eliminó la línea de ORDER BY que estaba repetida)
sql = sql & clausulaWhere & "ORDER BY p.Nombre_Proyecto, c.Tipo_Componente"

Set objRS = objConn.Execute(sql)

If filtroProyecto > 0 Then
    Dim objRSNombre
    Set objRSNombre = objConn.Execute("SELECT Nombre_Proyecto FROM Proyectos WHERE ID_Proyecto = " & filtroProyecto)
    If Not objRSNombre.EOF Then nombreProyectoFiltro = objRSNombre("Nombre_Proyecto")
    objRSNombre.Close
    Set objRSNombre = Nothing
End If
%>
<div class="panel">
    <div class="panel-cabecera">
        <h1>Componentes &amp; mapa de cableado</h1>
        <a href="formulario.asp<% If filtroProyecto > 0 Then %>?proyecto=<%= filtroProyecto %><% End If %>" class="boton boton-primario">+ Nuevo componente</a>
        
        <form method="get" action="listar.asp" style="display:flex; gap:10px; margin-bottom:20px;">
            <input type="text" name="q" placeholder="Buscar componente o valor (Ej: Resistencia, 10k)..." value="<%= Server.HTMLEncode(terminoBusqueda) %>" style="flex:1;">
            <button type="submit" class="boton boton-primario">Buscar en Inventario</button>
            <% If terminoBusqueda <> "" Then %>
                <a href="listar.asp" class="boton boton-secundario">Limpiar filtro</a>
            <% End If %>
        </form>
    </div>
    
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
                <tr>
                    <td colspan="6" class="vacio">No hay componentes registrados.</td>
                </tr>
            <% Else %>
                <% Do While Not objRS.EOF %>
                    <tr>
                        <td><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></td>
                        <td><%= Server.HTMLEncode(objRS("Tipo_Componente")) %></td>
                        <td><%= Server.HTMLEncode(objRS("Valor_Calculado")) %></td>
                        <td><span class="codigo"><%= Server.HTMLEncode(objRS("Pin_Conexion")) %></span></td>
                        <td><%= Server.HTMLEncode(objRS("Ubicacion_Protoboard")) %></td>
                        <td class="acciones-fila">
                            <a href="formulario.asp?id=<%= objRS("ID_Componente") %>">Editar</a>
                            <a href="eliminar.asp?id=<%= objRS("ID_Componente") %>" class="enlace-peligro" onclick="return confirm('¿Eliminar este componente?');">Eliminar</a>
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

<% 
objRS.Close 
Set objRS = Nothing 
CerrarConexion objConn 
%>