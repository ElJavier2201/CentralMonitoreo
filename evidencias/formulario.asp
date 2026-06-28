<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "evidencias"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, idEvidencia, modoEdicion, idProyectoSel, tipoEvidencia, rutaArchivo, descripcion, fechaSubida, errores
idEvidencia = IDValido(Request.QueryString("id"))
modoEdicion = (idEvidencia > 0)
idProyectoSel = IDValido(Request.QueryString("proyecto"))
tipoEvidencia = "Foto protoboard"
rutaArchivo = ""
descripcion = ""
fechaSubida = Date()
errores = ""

Function TipoPermitido(ByVal tipo)
    tipo = Limpiar(tipo)
    Select Case tipo
        Case "Captura simulacion", "Foto protoboard", "Diagrama", "Medicion", "Codigo", "Otro"
            TipoPermitido = tipo
        Case Else
            TipoPermitido = "Otro"
    End Select
End Function

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    idEvidencia = IDValido(Request.Form("id_evidencia"))
    modoEdicion = (idEvidencia > 0)
    idProyectoSel = IDValido(Request.Form("id_proyecto"))
    tipoEvidencia = TipoPermitido(Request.Form("tipo_evidencia"))
    rutaArchivo = Limpiar(Request.Form("ruta_archivo"))
    descripcion = Limpiar(Request.Form("descripcion"))
    fechaSubida = Limpiar(Request.Form("fecha_subida"))
    If Not IsDate(fechaSubida) Then fechaSubida = Date()
    If idProyectoSel = 0 Then errores = errores & "<li>Debe seleccionar un proyecto.</li>"
    If rutaArchivo = "" Then errores = errores & "<li>Debe escribir la ruta o enlace de la evidencia.</li>"
    If errores = "" Then
        Set objConn = AbrirConexion()
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = objConn
        cmd.CommandType = 1
        If modoEdicion Then
            cmd.CommandText = "UPDATE Evidencias SET ID_Proyecto = ?, Tipo_Evidencia = ?, Ruta_Archivo = ?, Descripcion = ?, Fecha_Subida = ? WHERE ID_Evidencia = ?"
        Else
            cmd.CommandText = "INSERT INTO Evidencias (ID_Proyecto, Tipo_Evidencia, Ruta_Archivo, Descripcion, Fecha_Subida) VALUES (?, ?, ?, ?, ?)"
        End If
        cmd.Parameters.Append cmd.CreateParameter("@idp", 3, 1, , idProyectoSel)
        cmd.Parameters.Append cmd.CreateParameter("@tipo", 200, 1, 60, tipoEvidencia)
        cmd.Parameters.Append cmd.CreateParameter("@ruta", 200, 1, 255, rutaArchivo)
        If descripcion = "" Then
            cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 255, Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 255, descripcion)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@fecha", 135, 1, , CDate(fechaSubida))
        If modoEdicion Then cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idEvidencia)
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
    cmdGet.CommandText = "SELECT * FROM Evidencias WHERE ID_Evidencia = ?"
    cmdGet.Parameters.Append cmdGet.CreateParameter("@id", 3, 1, , idEvidencia)
    Set objRS = cmdGet.Execute
    If Not objRS.EOF Then
        idProyectoSel = objRS("ID_Proyecto")
        tipoEvidencia = TipoPermitido(objRS("Tipo_Evidencia"))
        rutaArchivo = Limpiar(objRS("Ruta_Archivo"))
        descripcion = Limpiar(objRS("Descripcion"))
        fechaSubida = objRS("Fecha_Subida")
    Else
        modoEdicion = False
        idEvidencia = 0
    End If
    objRS.Close
    Set objRS = Nothing
    Set cmdGet = Nothing
    CerrarConexion objConn
End If
If modoEdicion Then tituloPagina = "Editar evidencia" Else tituloPagina = "Nueva evidencia"
Set objConn = AbrirConexion()
Set objRS = objConn.Execute("SELECT ID_Proyecto, Nombre_Proyecto FROM Proyectos ORDER BY Nombre_Proyecto")
%>
<!--#include virtual="/includes/header.asp"-->
<div class="panel">
    <h1><% If modoEdicion Then %>Editar evidencia<% Else %>Nueva evidencia<% End If %></h1>
    <p class="ayuda">Por compatibilidad simple con ASP clásico, este módulo guarda la ruta o enlace del archivo. Puedes poner una URL, una ruta local o una carpeta compartida.</p>
    <% If errores <> "" Then %><div class="alerta alerta-error"><strong>Revisa:</strong><ul><%= errores %></ul></div><% End If %>
    <% If objRS.EOF Then %>
        <div class="alerta alerta-error">Primero crea un proyecto.</div>
    <% Else %>
    <form method="post" action="formulario.asp<% If modoEdicion Then %>?id=<%= idEvidencia %><% End If %>">
        <input type="hidden" name="id_evidencia" value="<%= idEvidencia %>">
        <label for="id_proyecto">Proyecto *</label>
        <select id="id_proyecto" name="id_proyecto"><option value="">-- Seleccione --</option>
        <% Do While Not objRS.EOF %><option value="<%= objRS("ID_Proyecto") %>" <% If CLng(objRS("ID_Proyecto")) = CLng(idProyectoSel) Then %>selected<% End If %>><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></option><% objRS.MoveNext: Loop %>
        </select>
        <label for="tipo_evidencia">Tipo de evidencia</label>
        <select id="tipo_evidencia" name="tipo_evidencia">
            <option value="Captura simulacion" <% If tipoEvidencia = "Captura simulacion" Then %>selected<% End If %>>Captura de simulación</option>
            <option value="Foto protoboard" <% If tipoEvidencia = "Foto protoboard" Then %>selected<% End If %>>Foto de protoboard</option>
            <option value="Diagrama" <% If tipoEvidencia = "Diagrama" Then %>selected<% End If %>>Diagrama</option>
            <option value="Medicion" <% If tipoEvidencia = "Medicion" Then %>selected<% End If %>>Medición</option>
            <option value="Codigo" <% If tipoEvidencia = "Codigo" Then %>selected<% End If %>>Código</option>
            <option value="Otro" <% If tipoEvidencia = "Otro" Then %>selected<% End If %>>Otro</option>
        </select>
        <label for="ruta_archivo">Ruta o enlace *</label>
        <input type="text" id="ruta_archivo" name="ruta_archivo" maxlength="255" value="<%= Server.HTMLEncode(rutaArchivo) %>" placeholder="Ej: evidencias/sensor/foto1.jpg o https://...">
        <label for="fecha_subida">Fecha</label>
        <input type="date" id="fecha_subida" name="fecha_subida" value="<%= FormatearFechaInput(fechaSubida) %>">
        <label for="descripcion">Descripción</label>
        <textarea id="descripcion" name="descripcion" rows="3"><%= Server.HTMLEncode(descripcion) %></textarea>
        <div class="acciones"><button type="submit" class="boton boton-primario">Guardar evidencia</button><a href="listar.asp" class="boton boton-secundario">Cancelar</a></div>
    </form>
    <% End If %>
</div>
<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>
