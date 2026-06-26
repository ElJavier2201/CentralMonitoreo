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

' --- VARIABLES PARA PAGINACIÓN ---
Dim tamanoPagina, paginaActual, totalPaginas, registrosMostrados
tamanoPagina = 20 ' Cantidad de componentes a mostrar por página
paginaActual = IDValido(Request.QueryString("p"))
If paginaActual < 1 Then paginaActual = 1

' --- APERTURA DEL RECORDSET PREPARADO PARA PAGINACIÓN ---
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorLocation = 3 ' adUseClient (Necesario para contar páginas)
' Abrimos: (Consulta, Conexión, adOpenStatic = 3, adLockReadOnly = 1)
objRS.Open sql, objConn, 3, 1 

' --- CÁLCULO DE PÁGINAS ---
If Not objRS.EOF Then
    objRS.PageSize = tamanoPagina
    totalPaginas = objRS.PageCount
    
    ' Seguridad: Si piden una página mayor a la existente, ir a la última
    If paginaActual > totalPaginas Then paginaActual = totalPaginas
    
    ' Posicionamos el cursor en la página solicitada
    objRS.AbsolutePage = paginaActual
Else
    totalPaginas = 0
End If

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
        <a href="exportar_csv.asp?proyecto=<%= filtroProyecto %>&q=<%= Server.URLEncode(terminoBusqueda) %>" class="boton boton-secundario">Exportar a Excel (CSV)</a>
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
                <% 
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
                            <a href="eliminar.asp?id=<%= objRS("ID_Componente") %>" class="enlace-peligro" onclick="return confirm('¿Eliminar este componente?');">Eliminar</a>
                        </td>
                    </tr>
                <% 
                    registrosMostrados = registrosMostrados + 1
                    objRS.MoveNext 
                Loop 
                %>
            <% End If %>
        </tbody>
    </table>

    <% If totalPaginas > 1 Then %>
        <div style="margin-top: 20px; display: flex; justify-content: center; align-items: center; gap: 15px;">
            <% If paginaActual > 1 Then %>
                <a href="listar.asp?p=<%= paginaActual - 1 %><% If filtroProyecto > 0 Then %>&proyecto=<%= filtroProyecto %><% End If %><% If terminoBusqueda <> "" Then %>&q=<%= Server.URLEncode(terminoBusqueda) %><% End If %>" class="boton boton-secundario">&laquo; Anterior</a>
            <% End If %>
            
            <span style="color: var(--texto-tenue); font-size: 14px;">
                Página <%= paginaActual %> de <%= totalPaginas %>
            </span>
            
            <% If paginaActual < totalPaginas Then %>
                <a href="listar.asp?p=<%= paginaActual + 1 %><% If filtroProyecto > 0 Then %>&proyecto=<%= filtroProyecto %><% End If %><% If terminoBusqueda <> "" Then %>&q=<%= Server.URLEncode(terminoBusqueda) %><% End If %>" class="boton boton-secundario">Siguiente &raquo;</a>
            <% End If %>
        </div>
    <% End If %>

</div>    


<% 
objRS.Close 
Set objRS = Nothing 
CerrarConexion objConn 
%>