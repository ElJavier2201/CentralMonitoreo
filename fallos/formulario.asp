<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "fallos"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim idFallo, modoEdicion
Dim idProyectoSel, sintomaError, solucionAplicada, estadoFallo, fechaRegistro
Dim errores

errores = ""
idFallo = IDValido(Request.QueryString("id"))
modoEdicion = (idFallo > 0)

idProyectoSel    = IDValido(Request.QueryString("proyecto"))
sintomaError     = ""
solucionAplicada = ""
estadoFallo      = "Pendiente"
fechaRegistro    = Date()

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    idFallo          = IDValido(Request.Form("id_fallo"))
    modoEdicion      = (idFallo > 0)
    idProyectoSel    = IDValido(Request.Form("id_proyecto"))
    sintomaError     = Limpiar(Request.Form("sintoma_error"))
    solucionAplicada = Limpiar(Request.Form("solucion_aplicada"))
    estadoFallo      = Limpiar(Request.Form("estado"))
    fechaRegistro    = Limpiar(Request.Form("fecha_registro"))

    If Not IsDate(fechaRegistro) Then fechaRegistro = Date()
    If estadoFallo <> "Pendiente" And estadoFallo <> "Resuelto" Then estadoFallo = "Pendiente"

    If idProyectoSel = 0 Then errores = errores & "<li>Debe seleccionar el proyecto donde ocurrió el fallo.</li>"
    If sintomaError = "" Then errores = errores & "<li>Debe describir el síntoma del error.</li>"

    If errores = "" Then
        Set objConn = AbrirConexion()

        Dim objCmd
        Set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = objConn
        objCmd.CommandType = 1

        If modoEdicion Then
            objCmd.CommandText = "UPDATE Bitacora_Fallos SET ID_Proyecto = ?, Fecha_Registro = ?, Sintoma_Error = ?, Solucion_Aplicada = ?, Estado = ? WHERE ID_Fallo = ?"
            objCmd.Parameters.Append objCmd.CreateParameter("@id_proyecto", 3, 1, , idProyectoSel)
            objCmd.Parameters.Append objCmd.CreateParameter("@fecha", 135, 1, , CDate(fechaRegistro))
            objCmd.Parameters.Append objCmd.CreateParameter("@sintoma", 200, 1, 255, sintomaError)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@solucion", solucionAplicada, 255)
            objCmd.Parameters.Append objCmd.CreateParameter("@estado", 200, 1, 20, estadoFallo)
            objCmd.Parameters.Append objCmd.CreateParameter("@id", 3, 1, , idFallo)
        Else
            objCmd.CommandText = "INSERT INTO Bitacora_Fallos (ID_Proyecto, Fecha_Registro, Sintoma_Error, Solucion_Aplicada, Estado) VALUES (?, ?, ?, ?, ?)"
            objCmd.Parameters.Append objCmd.CreateParameter("@id_proyecto", 3, 1, , idProyectoSel)
            objCmd.Parameters.Append objCmd.CreateParameter("@fecha", 135, 1, , CDate(fechaRegistro))
            objCmd.Parameters.Append objCmd.CreateParameter("@sintoma", 200, 1, 255, sintomaError)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@solucion", solucionAplicada, 255)
            objCmd.Parameters.Append objCmd.CreateParameter("@estado", 200, 1, 20, estadoFallo)
        End If

        objCmd.Execute , , 128
        Set objCmd = Nothing
        CerrarConexion objConn

        Response.Redirect "listar.asp?ok=1"
        Response.End
    End If
End If

If modoEdicion And Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Set objConn = AbrirConexion()

    Dim cmdFallo
    Set cmdFallo = Server.CreateObject("ADODB.Command")
    cmdFallo.ActiveConnection = objConn
    cmdFallo.CommandType = 1
    cmdFallo.CommandText = "SELECT ID_Fallo, ID_Proyecto, Fecha_Registro, Sintoma_Error, Solucion_Aplicada, Estado FROM Bitacora_Fallos WHERE ID_Fallo = ?"
    cmdFallo.Parameters.Append cmdFallo.CreateParameter("@id", 3, 1, , idFallo)
    Set objRS = cmdFallo.Execute

    If Not objRS.EOF Then
        idProyectoSel    = objRS("ID_Proyecto")
        fechaRegistro    = objRS("Fecha_Registro")
        sintomaError     = Limpiar(objRS("Sintoma_Error"))
        solucionAplicada = Limpiar(objRS("Solucion_Aplicada"))
        estadoFallo      = Limpiar(objRS("Estado"))
    Else
        modoEdicion = False
        idFallo = 0
    End If

    objRS.Close
    Set objRS = Nothing
    Set cmdFallo = Nothing
    CerrarConexion objConn
End If

If modoEdicion Then
    tituloPagina = "Editar fallo - Central de Monitoreo"
Else
    tituloPagina = "Nuevo fallo - Central de Monitoreo"
End If

Set objConn = AbrirConexion()
sql = "SELECT ID_Proyecto, Nombre_Proyecto FROM Proyectos ORDER BY Nombre_Proyecto"
Set objRS = objConn.Execute(sql)

Function CrearTextoNull(ByRef cmd, ByVal nombre, ByVal valor, ByVal tamano)
    If Limpiar(valor) = "" Then
        Set CrearTextoNull = cmd.CreateParameter(nombre, 200, 1, tamano, Null)
    Else
        Set CrearTextoNull = cmd.CreateParameter(nombre, 200, 1, tamano, valor)
    End If
End Function
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <h1><% If modoEdicion Then %>Editar registro de fallo<% Else %>Nuevo registro de fallo<% End If %></h1>
    <p class="ayuda">Síntoma detectado durante la prueba física y la solución electrónica o de código aplicada.</p>

    <% If errores <> "" Then %>
    <div class="alerta alerta-error">
        <strong>Revisa los siguientes campos:</strong>
        <ul><%= errores %></ul>
    </div>
    <% End If %>

    <% If objRS.EOF Then %>
    <div class="alerta alerta-error">
        Aún no hay proyectos creados. <a href="../proyectos/formulario.asp">Crea primero un proyecto</a> antes de registrar fallos.
    </div>
    <% Else %>
    <form method="post" action="formulario.asp<% If modoEdicion Then %>?id=<%= idFallo %><% End If %>">
        <input type="hidden" name="id_fallo" value="<%= idFallo %>">

        <label for="id_proyecto">Proyecto *</label>
        <select id="id_proyecto" name="id_proyecto">
            <option value="">-- Seleccione un proyecto --</option>
            <% Do While Not objRS.EOF %>
            <option value="<%= objRS("ID_Proyecto") %>" <% If CLng(objRS("ID_Proyecto")) = CLng(idProyectoSel) Then %>selected<% End If %>>
                <%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %>
            </option>
            <%
                objRS.MoveNext
            Loop
            %>
        </select>

        <label for="fecha_registro">Fecha del registro</label>
        <input type="date" id="fecha_registro" name="fecha_registro" value="<%= FormatearFechaInput(fechaRegistro) %>">

        <label for="sintoma_error">Síntoma del error *</label>
        <textarea id="sintoma_error" name="sintoma_error" rows="3"
                  placeholder="Ej: Lecturas inestables en el monitor serie al aplicar presión"><%= Server.HTMLEncode(sintomaError) %></textarea>

        <label for="solucion_aplicada">Solución aplicada (electrónica o de código)</label>
        <textarea id="solucion_aplicada" name="solucion_aplicada" rows="3"
                  placeholder="Ej: Se agregó resistencia pull-down y se filtró con promedio móvil"><%= Server.HTMLEncode(solucionAplicada) %></textarea>

        <label for="estado">Estado</label>
        <select id="estado" name="estado">
            <option value="Pendiente" <% If estadoFallo = "Pendiente" Then %>selected<% End If %>>Pendiente</option>
            <option value="Resuelto" <% If estadoFallo = "Resuelto" Then %>selected<% End If %>>Resuelto</option>
        </select>

        <div class="acciones">
            <button type="submit" class="boton boton-primario">Guardar registro</button>
            <a href="listar.asp" class="boton boton-secundario">Cancelar</a>
        </div>
    </form>
    <% End If %>
</div>

<!--#include virtual="/includes/footer.asp"-->
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>
