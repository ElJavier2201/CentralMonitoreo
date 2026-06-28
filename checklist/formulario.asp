<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "checklist"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, idItem, modoEdicion, idProyectoSel, descripcion, completado, errores
idItem = IDValido(Request.QueryString("id"))
modoEdicion = (idItem > 0)
idProyectoSel = IDValido(Request.QueryString("proyecto"))
descripcion = ""
completado = False
errores = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    idItem = IDValido(Request.Form("id_item"))
    modoEdicion = (idItem > 0)
    idProyectoSel = IDValido(Request.Form("id_proyecto"))
    descripcion = Limpiar(Request.Form("descripcion"))
    completado = (Limpiar(Request.Form("completado")) = "1")
    If idProyectoSel = 0 Then errores = errores & "<li>Debe seleccionar un proyecto.</li>"
    If descripcion = "" Then errores = errores & "<li>Debe escribir el punto de revisión.</li>"
    If errores = "" Then
        Set objConn = AbrirConexion()
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = objConn
        cmd.CommandType = 1
        If modoEdicion Then
            cmd.CommandText = "UPDATE Checklist_Proyecto SET ID_Proyecto = ?, Descripcion = ?, Completado = ?, Fecha_Actualizacion = ? WHERE ID_Item = ?"
        Else
            cmd.CommandText = "INSERT INTO Checklist_Proyecto (ID_Proyecto, Descripcion, Completado, Fecha_Actualizacion) VALUES (?, ?, ?, ?)"
        End If
        cmd.Parameters.Append cmd.CreateParameter("@idp", 3, 1, , idProyectoSel)
        cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 255, descripcion)
        cmd.Parameters.Append cmd.CreateParameter("@comp", 11, 1, , completado)
        cmd.Parameters.Append cmd.CreateParameter("@fecha", 135, 1, , Now())
        If modoEdicion Then cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idItem)
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
    cmdGet.CommandText = "SELECT * FROM Checklist_Proyecto WHERE ID_Item = ?"
    cmdGet.Parameters.Append cmdGet.CreateParameter("@id", 3, 1, , idItem)
    Set objRS = cmdGet.Execute
    If Not objRS.EOF Then
        idProyectoSel = objRS("ID_Proyecto")
        descripcion = Limpiar(objRS("Descripcion"))
        completado = CBool(objRS("Completado"))
    Else
        modoEdicion = False
        idItem = 0
    End If
    objRS.Close
    Set objRS = Nothing
    Set cmdGet = Nothing
    CerrarConexion objConn
End If
If modoEdicion Then tituloPagina = "Editar checklist" Else tituloPagina = "Nuevo checklist"
Set objConn = AbrirConexion()
Set objRS = objConn.Execute("SELECT ID_Proyecto, Nombre_Proyecto FROM Proyectos ORDER BY Nombre_Proyecto")
%>
<!--#include virtual="/includes/header.asp"-->
<div class="panel">
    <h1><% If modoEdicion Then %>Editar punto de checklist<% Else %>Nuevo punto de checklist<% End If %></h1>
    <p class="ayuda">Agrega pasos de verificación como alimentación, continuidad, polaridad, conexiones y pruebas.</p>
    <% If errores <> "" Then %><div class="alerta alerta-error"><strong>Revisa:</strong><ul><%= errores %></ul></div><% End If %>
    <% If objRS.EOF Then %>
        <div class="alerta alerta-error">Primero crea un proyecto.</div>
    <% Else %>
    <form method="post" action="formulario.asp<% If modoEdicion Then %>?id=<%= idItem %><% End If %>">
        <input type="hidden" name="id_item" value="<%= idItem %>">
        <label for="id_proyecto">Proyecto *</label>
        <select id="id_proyecto" name="id_proyecto">
            <option value="">-- Seleccione --</option>
            <% Do While Not objRS.EOF %>
            <option value="<%= objRS("ID_Proyecto") %>" <% If CLng(objRS("ID_Proyecto")) = CLng(idProyectoSel) Then %>selected<% End If %>><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></option>
            <% objRS.MoveNext: Loop %>
        </select>
        <label for="descripcion">Punto de revisión *</label>
        <textarea id="descripcion" name="descripcion" rows="3" maxlength="255" placeholder="Ej: Verificar continuidad entre GND del Arduino y riel negativo de la protoboard"><%= Server.HTMLEncode(descripcion) %></textarea>
        <label class="checkline"><input type="checkbox" name="completado" value="1" <% If completado Then %>checked<% End If %>> Marcar como completado</label>
        <div class="acciones"><button type="submit" class="boton boton-primario">Guardar</button><a href="listar.asp" class="boton boton-secundario">Cancelar</a></div>
    </form>
    <% End If %>
</div>
<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>
