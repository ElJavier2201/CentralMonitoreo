

<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "proyectos"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS
Dim idProyecto, modoEdicion
Dim nombreProyecto, plataforma, microcontrolador, estadoProyecto, descripcion, fechaCreacion
Dim errores

errores = ""
idProyecto = IDValido(Request.QueryString("id"))
modoEdicion = (idProyecto > 0)

nombreProyecto   = ""
plataforma       = ""
microcontrolador = ""
estadoProyecto   = "Planeado"
descripcion      = ""
fechaCreacion    = Date()

Function EstadoPermitido(ByVal estado)
    estado = Limpiar(estado)
    Select Case estado
        Case "Planeado", "En simulacion", "En armado fisico", "En pruebas", "Finalizado", "Cancelado"
            EstadoPermitido = estado
        Case Else
            EstadoPermitido = "Planeado"
    End Select
End Function

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    nombreProyecto   = Limpiar(Request.Form("nombre_proyecto"))
    plataforma       = Limpiar(Request.Form("plataforma_simulacion"))
    microcontrolador = Limpiar(Request.Form("microcontrolador"))
    estadoProyecto   = EstadoPermitido(Request.Form("estado"))
    descripcion      = Limpiar(Request.Form("descripcion"))
    fechaCreacion    = Limpiar(Request.Form("fecha_creacion"))
    idProyecto       = IDValido(Request.Form("id_proyecto"))
    modoEdicion      = (idProyecto > 0)

    If Not IsDate(fechaCreacion) Then fechaCreacion = Date()

    If nombreProyecto = "" Then errores = errores & "<li>El nombre del proyecto es obligatorio.</li>"
    If plataforma = "" Then errores = errores & "<li>Debe indicar la plataforma de simulación.</li>"
    If microcontrolador = "" Then errores = errores & "<li>Debe indicar el microcontrolador / cerebro del circuito.</li>"

    If errores = "" Then
        Set objConn = AbrirConexion()

        Dim objCmd
        Set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = objConn
        objCmd.CommandType = 1

        If modoEdicion Then
            objCmd.CommandText = "UPDATE Proyectos SET Nombre_Proyecto = ?, Plataforma_Simulacion = ?, Microcontrolador = ?, Estado = ?, Descripcion = ?, Fecha_Creacion = ? WHERE ID_Proyecto = ?"
            objCmd.Parameters.Append objCmd.CreateParameter("@nombre", 200, 1, 150, nombreProyecto)
            objCmd.Parameters.Append objCmd.CreateParameter("@plataforma", 200, 1, 80, plataforma)
            objCmd.Parameters.Append objCmd.CreateParameter("@micro", 200, 1, 80, microcontrolador)
            objCmd.Parameters.Append objCmd.CreateParameter("@estado", 200, 1, 30, estadoProyecto)
            If descripcion = "" Then
                objCmd.Parameters.Append objCmd.CreateParameter("@desc", 200, 1, 255, Null)
            Else
                objCmd.Parameters.Append objCmd.CreateParameter("@desc", 200, 1, 255, descripcion)
            End If
            objCmd.Parameters.Append objCmd.CreateParameter("@fecha", 135, 1, , CDate(fechaCreacion))
            objCmd.Parameters.Append objCmd.CreateParameter("@id", 3, 1, , idProyecto)
        Else
            objCmd.CommandText = "INSERT INTO Proyectos (Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Estado, Descripcion, Fecha_Creacion) VALUES (?, ?, ?, ?, ?, ?)"
            objCmd.Parameters.Append objCmd.CreateParameter("@nombre", 200, 1, 150, nombreProyecto)
            objCmd.Parameters.Append objCmd.CreateParameter("@plataforma", 200, 1, 80, plataforma)
            objCmd.Parameters.Append objCmd.CreateParameter("@micro", 200, 1, 80, microcontrolador)
            objCmd.Parameters.Append objCmd.CreateParameter("@estado", 200, 1, 30, estadoProyecto)
            If descripcion = "" Then
                objCmd.Parameters.Append objCmd.CreateParameter("@desc", 200, 1, 255, Null)
            Else
                objCmd.Parameters.Append objCmd.CreateParameter("@desc", 200, 1, 255, descripcion)
            End If
            objCmd.Parameters.Append objCmd.CreateParameter("@fecha", 135, 1, , CDate(fechaCreacion))
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

    Dim cmdProyecto
    Set cmdProyecto = Server.CreateObject("ADODB.Command")
    cmdProyecto.ActiveConnection = objConn
    cmdProyecto.CommandType = 1
    cmdProyecto.CommandText = "SELECT ID_Proyecto, Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Estado, Descripcion, Fecha_Creacion FROM Proyectos WHERE ID_Proyecto = ?"
    cmdProyecto.Parameters.Append cmdProyecto.CreateParameter("@id", 3, 1, , idProyecto)
    Set objRS = cmdProyecto.Execute

    If Not objRS.EOF Then
        nombreProyecto   = Limpiar(objRS("Nombre_Proyecto"))
        plataforma       = Limpiar(objRS("Plataforma_Simulacion"))
        microcontrolador = Limpiar(objRS("Microcontrolador"))
        estadoProyecto   = EstadoPermitido(objRS("Estado"))
        descripcion      = Limpiar(objRS("Descripcion"))
        fechaCreacion    = objRS("Fecha_Creacion")
    Else
        modoEdicion = False
        idProyecto = 0
    End If

    objRS.Close
    Set objRS = Nothing
    Set cmdProyecto = Nothing
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
    <p class="ayuda">Cabecera del experimento: plataforma, microcontrolador, estado y descripción.</p>

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
            <option value="Multisim">
            <option value="LTspice">
        </datalist>

        <label for="microcontrolador">Microcontrolador / cerebro del circuito *</label>
        <input type="text" id="microcontrolador" name="microcontrolador" maxlength="80"
               value="<%= Server.HTMLEncode(microcontrolador) %>" placeholder="Ej: Arduino Uno, ESP32">

        <label for="estado">Estado del proyecto</label>
        <select id="estado" name="estado">
            <option value="Planeado" <% If estadoProyecto = "Planeado" Then %>selected<% End If %>>Planeado</option>
            <option value="En simulacion" <% If estadoProyecto = "En simulacion" Then %>selected<% End If %>>En simulación</option>
            <option value="En armado fisico" <% If estadoProyecto = "En armado fisico" Then %>selected<% End If %>>En armado físico</option>
            <option value="En pruebas" <% If estadoProyecto = "En pruebas" Then %>selected<% End If %>>En pruebas</option>
            <option value="Finalizado" <% If estadoProyecto = "Finalizado" Then %>selected<% End If %>>Finalizado</option>
            <option value="Cancelado" <% If estadoProyecto = "Cancelado" Then %>selected<% End If %>>Cancelado</option>
        </select>

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
