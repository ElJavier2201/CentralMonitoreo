<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
%>
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim idFallo, modoEdicion
Dim idProyectoSel, sintomaError, solucionAplicada, estadoFallo, fechaRegistro
Dim errores
Dim rutaBase, tituloPagina, seccionActiva

rutaBase = "../"
seccionActiva = "fallos"
errores = ""

idFallo = IDValido(Request.QueryString("id"))
modoEdicion = (idFallo > 0)

idProyectoSel    = IDValido(Request.QueryString("proyecto"))
sintomaError     = ""
solucionAplicada = ""
estadoFallo      = "Pendiente"
fechaRegistro    = Date()

' ===================================================================
' PROCESAMIENTO DEL FORMULARIO (POST) - validación + SQL nativo
' ===================================================================
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

        If modoEdicion Then
            sql = "UPDATE Bitacora_Fallos SET " & _
                  "ID_Proyecto = " & idProyectoSel & ", " & _
                  "Fecha_Registro = " & FechaSQL(fechaRegistro) & ", " & _
                  "Sintoma_Error = " & ValorTextoSQL(sintomaError) & ", " & _
                  "Solucion_Aplicada = " & ValorTextoSQL(solucionAplicada) & ", " & _
                  "Estado = " & ValorTextoSQL(estadoFallo) & " " & _
                  "WHERE ID_Fallo = " & idFallo
        Else
            sql = "INSERT INTO Bitacora_Fallos " & _
                  "(ID_Proyecto, Fecha_Registro, Sintoma_Error, Solucion_Aplicada, Estado) " & _
                  "VALUES (" & idProyectoSel & ", " & _
                  FechaSQL(fechaRegistro) & ", " & _
                  ValorTextoSQL(sintomaError) & ", " & _
                  ValorTextoSQL(solucionAplicada) & ", " & _
                  ValorTextoSQL(estadoFallo) & ")"
        End If

        objConn.Execute sql, , 129
        CerrarConexion objConn

        Response.Redirect "listar.asp?ok=1"
        Response.End
    End If
End If

' ===================================================================
' CARGA DE DATOS EXISTENTES (modo edición, primer GET)
' ===================================================================
If modoEdicion And Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Set objConn = AbrirConexion()
    sql = "SELECT ID_Fallo, ID_Proyecto, Fecha_Registro, Sintoma_Error, Solucion_Aplicada, Estado " & _
          "FROM Bitacora_Fallos WHERE ID_Fallo = " & idFallo
    Set objRS = objConn.Execute(sql)

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
    CerrarConexion objConn
End If

If modoEdicion Then
    tituloPagina = "Editar fallo - Central de Monitoreo"
Else
    tituloPagina = "Nuevo fallo - Central de Monitoreo"
End If

' Recordset auxiliar para llenar el combo de proyectos disponibles
Set objConn = AbrirConexion()
sql = "SELECT ID_Proyecto, Nombre_Proyecto FROM Proyectos ORDER BY Nombre_Proyecto"
Set objRS = objConn.Execute(sql)
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
            <%
            Do While Not objRS.EOF
            %>
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
