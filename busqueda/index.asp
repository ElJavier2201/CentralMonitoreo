<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Buscador global - Central de Monitoreo"
seccionActiva = "busqueda"
%>
<!--#include virtual="/conexion.asp"-->
<!--#include virtual="/includes/auth.asp"-->
<%
Dim objConn, q, qSQL, totalResultados
q = Limpiar(Request.QueryString("q"))
qSQL = EscaparSQL(q)
totalResultados = 0
Set objConn = AbrirConexion()

Sub PintarFila(ByVal modulo, ByVal titulo, ByVal detalle, ByVal enlace)
    totalResultados = totalResultados + 1
%>
    <tr>
        <td><span class="etiqueta"><%= Server.HTMLEncode(modulo) %></span></td>
        <td><%= Server.HTMLEncode(titulo) %></td>
        <td><%= Server.HTMLEncode(detalle) %></td>
        <td><a href="<%= enlace %>">Abrir</a></td>
    </tr>
<%
End Sub

Sub BuscarSQL(ByVal modulo, ByVal sql, ByVal campoTitulo, ByVal campoDetalle, ByVal campoId, ByVal urlBase)
    Dim rs, titulo, detalle, enlace
    On Error Resume Next
    Set rs = objConn.Execute(sql)
    If Err.Number <> 0 Then
        Err.Clear
        On Error Goto 0
        Exit Sub
    End If
    On Error Goto 0

    Do While Not rs.EOF
        titulo = ""
        detalle = ""
        If Not IsNull(rs(campoTitulo)) Then titulo = CStr(rs(campoTitulo))
        If Not IsNull(rs(campoDetalle)) Then detalle = CStr(rs(campoDetalle))
        enlace = urlBase & rs(campoId)
        PintarFila modulo, titulo, detalle, enlace
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <div class="panel-cabecera">
        <h1>Buscador global</h1>
    </div>
    <p class="ayuda">Busca en proyectos, componentes, fallos, mediciones, checklist y evidencias.</p>

    <form method="get" action="index.asp" class="formulario-linea">
        <input type="text" name="q" maxlength="100" value="<%= Server.HTMLEncode(q) %>" placeholder="Ej: resistencia, Arduino, A0, lectura inestable...">
        <button type="submit" class="boton boton-primario">Buscar</button>
        <% If q <> "" Then %>
            <a href="index.asp" class="boton boton-secundario">Limpiar</a>
        <% End If %>
    </form>

    <% If q = "" Then %>
        <div class="alerta alerta-ok" style="margin-top:18px;">Escribe una palabra clave para buscar en todo el sistema.</div>
    <% ElseIf Len(q) < 2 Then %>
        <div class="alerta alerta-error" style="margin-top:18px;">Escribe al menos 2 caracteres para buscar.</div>
    <% Else %>
        <table class="tabla-datos">
            <thead>
                <tr>
                    <th>Módulo</th>
                    <th>Resultado</th>
                    <th>Detalle</th>
                    <th>Acción</th>
                </tr>
            </thead>
            <tbody>
            <%
            BuscarSQL "Proyecto", _
                "SELECT ID_Proyecto, Nombre_Proyecto, Plataforma_Simulacion & ' / ' & Microcontrolador AS Detalle FROM Proyectos " & _
                "WHERE Nombre_Proyecto LIKE '%" & qSQL & "%' OR Plataforma_Simulacion LIKE '%" & qSQL & "%' OR Microcontrolador LIKE '%" & qSQL & "%' OR Descripcion LIKE '%" & qSQL & "%'", _
                "Nombre_Proyecto", "Detalle", "ID_Proyecto", "../proyectos/formulario.asp?id="

            BuscarSQL "Componente", _
                "SELECT c.ID_Componente, c.Tipo_Componente, p.Nombre_Proyecto & ' / ' & c.Valor_Calculado & ' / ' & c.Pin_Conexion AS Detalle " & _
                "FROM Componentes c INNER JOIN Proyectos p ON c.ID_Proyecto = p.ID_Proyecto " & _
                "WHERE c.Tipo_Componente LIKE '%" & qSQL & "%' OR c.Valor_Calculado LIKE '%" & qSQL & "%' OR c.Pin_Conexion LIKE '%" & qSQL & "%' OR c.Ubicacion_Protoboard LIKE '%" & qSQL & "%' OR c.Notas LIKE '%" & qSQL & "%'", _
                "Tipo_Componente", "Detalle", "ID_Componente", "../componentes/formulario.asp?id="

            BuscarSQL "Fallo", _
                "SELECT f.ID_Fallo, f.Sintoma_Error, p.Nombre_Proyecto & ' / ' & f.Estado AS Detalle " & _
                "FROM Bitacora_Fallos f INNER JOIN Proyectos p ON f.ID_Proyecto = p.ID_Proyecto " & _
                "WHERE f.Sintoma_Error LIKE '%" & qSQL & "%' OR f.Solucion_Aplicada LIKE '%" & qSQL & "%' OR f.Estado LIKE '%" & qSQL & "%'", _
                "Sintoma_Error", "Detalle", "ID_Fallo", "../fallos/formulario.asp?id="

            BuscarSQL "Medición", _
                "SELECT m.ID_Medicion, m.Variable_Medida, p.Nombre_Proyecto & ' / Simulado: ' & m.Valor_Simulado & ' / Físico: ' & m.Valor_Fisico AS Detalle " & _
                "FROM Mediciones m INNER JOIN Proyectos p ON m.ID_Proyecto = p.ID_Proyecto " & _
                "WHERE m.Variable_Medida LIKE '%" & qSQL & "%' OR m.Valor_Simulado LIKE '%" & qSQL & "%' OR m.Valor_Fisico LIKE '%" & qSQL & "%' OR m.Diferencia LIKE '%" & qSQL & "%' OR m.Observaciones LIKE '%" & qSQL & "%'", _
                "Variable_Medida", "Detalle", "ID_Medicion", "../mediciones/formulario.asp?id="

            BuscarSQL "Checklist", _
                "SELECT ch.ID_Item, ch.Descripcion, p.Nombre_Proyecto AS Detalle " & _
                "FROM Checklist_Proyecto ch INNER JOIN Proyectos p ON ch.ID_Proyecto = p.ID_Proyecto " & _
                "WHERE ch.Descripcion LIKE '%" & qSQL & "%'", _
                "Descripcion", "Detalle", "ID_Item", "../checklist/formulario.asp?id="

            BuscarSQL "Evidencia", _
                "SELECT e.ID_Evidencia, e.Tipo_Evidencia, p.Nombre_Proyecto & ' / ' & e.Ruta_Archivo AS Detalle " & _
                "FROM Evidencias e INNER JOIN Proyectos p ON e.ID_Proyecto = p.ID_Proyecto " & _
                "WHERE e.Tipo_Evidencia LIKE '%" & qSQL & "%' OR e.Ruta_Archivo LIKE '%" & qSQL & "%' OR e.Descripcion LIKE '%" & qSQL & "%'", _
                "Tipo_Evidencia", "Detalle", "ID_Evidencia", "../evidencias/formulario.asp?id="
            %>
            <% If totalResultados = 0 Then %>
                <tr><td colspan="4" class="vacio">No se encontraron resultados para: <strong><%= Server.HTMLEncode(q) %></strong></td></tr>
            <% End If %>
            </tbody>
        </table>
        <p class="ayuda" style="margin-top:14px;">Resultados encontrados: <strong><%= totalResultados %></strong></p>
    <% End If %>
</div>

<!--#include virtual="/includes/footer.asp"-->
<%
CerrarConexion objConn
%>
