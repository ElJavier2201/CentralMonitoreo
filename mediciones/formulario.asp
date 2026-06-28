<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "mediciones"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, idMedicion, modoEdicion, idProyectoSel
Dim fechaMedicion, variableMedida, valorSimulado, valorFisico, diferencia, observaciones, errores

idMedicion = IDValido(Request.QueryString("id"))
modoEdicion = (idMedicion > 0)
idProyectoSel = IDValido(Request.QueryString("proyecto"))
fechaMedicion = Date()
variableMedida = ""
valorSimulado = ""
valorFisico = ""
diferencia = ""
observaciones = ""
errores = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    idMedicion = IDValido(Request.Form("id_medicion"))
    modoEdicion = (idMedicion > 0)
    idProyectoSel = IDValido(Request.Form("id_proyecto"))
    fechaMedicion = Limpiar(Request.Form("fecha_medicion"))
    variableMedida = Limpiar(Request.Form("variable_medida"))
    valorSimulado = Limpiar(Request.Form("valor_simulado"))
    valorFisico = Limpiar(Request.Form("valor_fisico"))
    diferencia = Limpiar(Request.Form("diferencia"))
    observaciones = Limpiar(Request.Form("observaciones"))

    If Not IsDate(fechaMedicion) Then fechaMedicion = Date()
    If idProyectoSel = 0 Then errores = errores & "<li>Debe seleccionar un proyecto.</li>"
    If variableMedida = "" Then errores = errores & "<li>Debe indicar la variable medida.</li>"

    If errores = "" Then
        Set objConn = AbrirConexion()
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = objConn
        cmd.CommandType = 1
        If modoEdicion Then
            cmd.CommandText = "UPDATE Mediciones SET ID_Proyecto = ?, Fecha_Medicion = ?, Variable_Medida = ?, Valor_Simulado = ?, Valor_Fisico = ?, Diferencia = ?, Observaciones = ? WHERE ID_Medicion = ?"
        Else
            cmd.CommandText = "INSERT INTO Mediciones (ID_Proyecto, Fecha_Medicion, Variable_Medida, Valor_Simulado, Valor_Fisico, Diferencia, Observaciones) VALUES (?, ?, ?, ?, ?, ?, ?)"
        End If
        cmd.Parameters.Append cmd.CreateParameter("@idp", 3, 1, , idProyectoSel)
        cmd.Parameters.Append cmd.CreateParameter("@fecha", 135, 1, , CDate(fechaMedicion))
        cmd.Parameters.Append cmd.CreateParameter("@var", 200, 1, 100, variableMedida)
        cmd.Parameters.Append cmd.CreateParameter("@sim", 200, 1, 80, valorSimulado)
        cmd.Parameters.Append cmd.CreateParameter("@fis", 200, 1, 80, valorFisico)
        cmd.Parameters.Append cmd.CreateParameter("@dif", 200, 1, 80, diferencia)
        If observaciones = "" Then
            cmd.Parameters.Append cmd.CreateParameter("@obs", 200, 1, 255, Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@obs", 200, 1, 255, observaciones)
        End If
        If modoEdicion Then cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idMedicion)
        cmd.Execute , , 128
        Set cmd = Nothing
        CerrarConexion objConn
        Response.Redirect "listar.asp?ok=1"
        Response.End
    End If
End If

If modoEdicion And Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Set objConn = AbrirConexion()
    Dim cmdGet
    Set cmdGet = Server.CreateObject("ADODB.Command")
    cmdGet.ActiveConnection = objConn
    cmdGet.CommandType = 1
    cmdGet.CommandText = "SELECT * FROM Mediciones WHERE ID_Medicion = ?"
    cmdGet.Parameters.Append cmdGet.CreateParameter("@id", 3, 1, , idMedicion)
    Set objRS = cmdGet.Execute
    If Not objRS.EOF Then
        idProyectoSel = objRS("ID_Proyecto")
        fechaMedicion = objRS("Fecha_Medicion")
        variableMedida = Limpiar(objRS("Variable_Medida"))
        valorSimulado = Limpiar(objRS("Valor_Simulado"))
        valorFisico = Limpiar(objRS("Valor_Fisico"))
        diferencia = Limpiar(objRS("Diferencia"))
        observaciones = Limpiar(objRS("Observaciones"))
    Else
        modoEdicion = False
        idMedicion = 0
    End If
    objRS.Close
    Set objRS = Nothing
    Set cmdGet = Nothing
    CerrarConexion objConn
End If

If modoEdicion Then tituloPagina = "Editar medición" Else tituloPagina = "Nueva medición"
Set objConn = AbrirConexion()
Set objRS = objConn.Execute("SELECT ID_Proyecto, Nombre_Proyecto FROM Proyectos ORDER BY Nombre_Proyecto")
%>
<!--#include virtual="/includes/header.asp"-->
<div class="panel">
    <h1><% If modoEdicion Then %>Editar medición<% Else %>Nueva medición<% End If %></h1>
    <p class="ayuda">Compara una variable del circuito entre simulación y medición física.</p>
    <% If errores <> "" Then %><div class="alerta alerta-error"><strong>Revisa:</strong><ul><%= errores %></ul></div><% End If %>
    <% If objRS.EOF Then %>
        <div class="alerta alerta-error">Primero crea un proyecto.</div>
    <% Else %>
    <form method="post" action="formulario.asp<% If modoEdicion Then %>?id=<%= idMedicion %><% End If %>">
        <input type="hidden" name="id_medicion" value="<%= idMedicion %>">
        <label for="id_proyecto">Proyecto *</label>
        <select id="id_proyecto" name="id_proyecto">
            <option value="">-- Seleccione --</option>
            <% Do While Not objRS.EOF %>
            <option value="<%= objRS("ID_Proyecto") %>" <% If CLng(objRS("ID_Proyecto")) = CLng(idProyectoSel) Then %>selected<% End If %>><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></option>
            <% objRS.MoveNext: Loop %>
        </select>
        <label for="fecha_medicion">Fecha</label>
        <input type="date" id="fecha_medicion" name="fecha_medicion" value="<%= FormatearFechaInput(fechaMedicion) %>">
        <label for="variable_medida">Variable medida *</label>
        <input type="text" id="variable_medida" name="variable_medida" maxlength="100" value="<%= Server.HTMLEncode(variableMedida) %>" placeholder="Ej: Voltaje en A0, resistencia equivalente, corriente del LED">
        <label for="valor_simulado">Valor simulado</label>
        <input type="text" id="valor_simulado" name="valor_simulado" maxlength="80" value="<%= Server.HTMLEncode(valorSimulado) %>" placeholder="Ej: 2.50 V">
        <label for="valor_fisico">Valor físico</label>
        <input type="text" id="valor_fisico" name="valor_fisico" maxlength="80" value="<%= Server.HTMLEncode(valorFisico) %>" placeholder="Ej: 2.38 V">
        <label for="diferencia">Diferencia</label>
        <input type="text" id="diferencia" name="diferencia" maxlength="80" value="<%= Server.HTMLEncode(diferencia) %>" placeholder="Ej: 0.12 V">
        <label for="observaciones">Observaciones</label>
        <textarea id="observaciones" name="observaciones" rows="3"><%= Server.HTMLEncode(observaciones) %></textarea>
        <div class="acciones"><button type="submit" class="boton boton-primario">Guardar medición</button><a href="listar.asp" class="boton boton-secundario">Cancelar</a></div>
    </form>
    <% End If %>
</div>
<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>
