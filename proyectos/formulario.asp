<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
%>
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim idProyecto, modoEdicion
Dim nombreProyecto, plataforma, microcontrolador, descripcion, fechaCreacion
Dim errores
Dim rutaBase, tituloPagina, seccionActiva

rutaBase = "../"
seccionActiva = "proyectos"
errores = ""

idProyecto = IDValido(Request.QueryString("id"))
modoEdicion = (idProyecto > 0)

' Valores por defecto del formulario
nombreProyecto   = ""
plataforma       = ""
microcontrolador = ""
descripcion      = ""
fechaCreacion    = Date()

' ===================================================================
' PROCESAMIENTO DEL FORMULARIO (POST) - validación + SQL nativo
' ===================================================================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    nombreProyecto   = Limpiar(Request.Form("nombre_proyecto"))
    plataforma       = Limpiar(Request.Form("plataforma_simulacion"))
    microcontrolador = Limpiar(Request.Form("microcontrolador"))
    descripcion      = Limpiar(Request.Form("descripcion"))
    fechaCreacion    = Limpiar(Request.Form("fecha_creacion"))
    idProyecto       = IDValido(Request.Form("id_proyecto"))
    modoEdicion      = (idProyecto > 0)

    If Not IsDate(fechaCreacion) Then fechaCreacion = Date()

    ' --- Validación: ningún campo obligatorio puede llegar vacío ---
    If nombreProyecto = "" Then errores = errores & "<li>El nombre del proyecto es obligatorio.</li>"
    If plataforma = "" Then errores = errores & "<li>Debe indicar la plataforma de simulación (Proteus, Tinkercad, etc.).</li>"
    If microcontrolador = "" Then errores = errores & "<li>Debe indicar el microcontrolador / cerebro del circuito.</li>"

    If errores = "" Then

        Set objConn = AbrirConexion()

        If modoEdicion Then
            sql = "UPDATE Proyectos SET " & _
                  "Nombre_Proyecto = " & ValorTextoSQL(nombreProyecto) & ", " & _
                  "Plataforma_Simulacion = " & ValorTextoSQL(plataforma) & ", " & _
                  "Microcontrolador = " & ValorTextoSQL(microcontrolador) & ", " & _
                  "Descripcion = " & ValorTextoSQL(descripcion) & ", " & _
                  "Fecha_Creacion = " & FechaSQL(fechaCreacion) & " " & _
                  "WHERE ID_Proyecto = " & idProyecto
        Else
            sql = "INSERT INTO Proyectos " & _
                  "(Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Descripcion, Fecha_Creacion) " & _
                  "VALUES (" & _
                  ValorTextoSQL(nombreProyecto) & ", " & _
                  ValorTextoSQL(plataforma) & ", " & _
                  ValorTextoSQL(microcontrolador) & ", " & _
                  ValorTextoSQL(descripcion) & ", " & _
                  FechaSQL(fechaCreacion) & ")"
        End If

        ' adCmdText (1) + adExecuteNoRecords (128) = 129: ejecuta el DML
        ' directamente sobre la conexión sin generar un Recordset en memoria
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
    sql = "SELECT ID_Proyecto, Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Descripcion, Fecha_Creacion " & _
          "FROM Proyectos WHERE ID_Proyecto = " & idProyecto

    Set objRS = objConn.Execute(sql)

    If Not objRS.EOF Then
        nombreProyecto   = Limpiar(objRS("Nombre_Proyecto"))
        plataforma       = Limpiar(objRS("Plataforma_Simulacion"))
        microcontrolador = Limpiar(objRS("Microcontrolador"))
        descripcion      = Limpiar(objRS("Descripcion"))
        fechaCreacion    = objRS("Fecha_Creacion")
    Else
        modoEdicion = False
        idProyecto = 0
    End If

    objRS.Close
    Set objRS = Nothing
    CerrarConexion objConn
End If

If modoEdicion Then
    tituloPagina = "Editar proyecto - Central de Monitoreo"
Else
    tituloPagina = "Nuevo proyecto - Central de Monitoreo"
End If
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <h1><% If modoEdicion Then %>Editar proyecto<% Else %>Nuevo proyecto<% End If %></h1>
    <p class="ayuda">Cabecera del experimento: nombre, plataforma de simulación y cerebro del circuito.</p>

    <% If errores <> "" Then %>
    <div class="alerta alerta-error">
        <strong>Revisa los siguientes campos:</strong>
        <ul><%= errores %></ul>
    </div>
    <% End If %>

    <form method="post" action="formulario.asp<% If modoEdicion Then %>?id=<%= idProyecto %><% End If %>">
        <input type="hidden" name="id_proyecto" value="<%= idProyecto %>">

        <label for="nombre_proyecto">Nombre del proyecto *</label>
        <input type="text" id="nombre_proyecto" name="nombre_proyecto" maxlength="150"
               value="<%= Server.HTMLEncode(nombreProyecto) %>" placeholder="Ej: Sensor de Fuerza Capacitivo">

        <label for="plataforma_simulacion">Plataforma de simulación *</label>
        <input type="text" id="plataforma_simulacion" name="plataforma_simulacion" maxlength="80" list="lista_plataformas"
               value="<%= Server.HTMLEncode(plataforma) %>" placeholder="Ej: Proteus, Tinkercad">
        <datalist id="lista_plataformas">
            <option value="Proteus">
            <option value="Tinkercad">
        </datalist>

        <label for="microcontrolador">Microcontrolador (cerebro del circuito) *</label>
        <input type="text" id="microcontrolador" name="microcontrolador" maxlength="80"
               value="<%= Server.HTMLEncode(microcontrolador) %>" placeholder="Ej: Arduino Uno, ESP32">

        <label for="fecha_creacion">Fecha de creación</label>
        <input type="date" id="fecha_creacion" name="fecha_creacion" value="<%= FormatearFechaInput(fechaCreacion) %>">

        <label for="descripcion">Descripción</label>
        <textarea id="descripcion" name="descripcion" rows="4" placeholder="Objetivo del experimento..."><%= Server.HTMLEncode(descripcion) %></textarea>

        <div class="acciones">
            <button type="submit" class="boton boton-primario">Guardar proyecto</button>
            <a href="listar.asp" class="boton boton-secundario">Cancelar</a>
        </div>
    </form>
</div>

<!--#include virtual="/includes/footer.asp"-->
