
<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
seccionActiva = "componentes"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim idComponente, modoEdicion
Dim idProyectoSel, tipoComponente, valorCalculado, pinConexion, ubicacionProtoboard, notas
Dim errores

errores = ""
idComponente = IDValido(Request.QueryString("id"))
modoEdicion = (idComponente > 0)

idProyectoSel        = IDValido(Request.QueryString("proyecto"))
tipoComponente       = ""
valorCalculado       = ""
pinConexion          = ""
ubicacionProtoboard  = ""
notas                = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    idComponente         = IDValido(Request.Form("id_componente"))
    modoEdicion          = (idComponente > 0)
    idProyectoSel        = IDValido(Request.Form("id_proyecto"))
    tipoComponente       = Limpiar(Request.Form("tipo_componente"))
    valorCalculado       = Limpiar(Request.Form("valor_calculado"))
    pinConexion          = Limpiar(Request.Form("pin_conexion"))
    ubicacionProtoboard  = Limpiar(Request.Form("ubicacion_protoboard"))
    notas                = Limpiar(Request.Form("notas"))

    If idProyectoSel = 0 Then errores = errores & "<li>Debe seleccionar el proyecto al que pertenece el componente.</li>"
    If tipoComponente = "" Then errores = errores & "<li>El tipo de componente es obligatorio.</li>"

    If errores = "" Then
        Set objConn = AbrirConexion()

        Dim objCmd
        Set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = objConn
        objCmd.CommandType = 1

        If modoEdicion Then
            objCmd.CommandText = "UPDATE Componentes SET ID_Proyecto = ?, Tipo_Componente = ?, Valor_Calculado = ?, Pin_Conexion = ?, Ubicacion_Protoboard = ?, Notas = ? WHERE ID_Componente = ?"
            objCmd.Parameters.Append objCmd.CreateParameter("@id_proyecto", 3, 1, , idProyectoSel)
            objCmd.Parameters.Append objCmd.CreateParameter("@tipo", 200, 1, 100, tipoComponente)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@valor", valorCalculado, 100)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@pin", pinConexion, 50)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@ubicacion", ubicacionProtoboard, 100)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@notas", notas, 255)
            objCmd.Parameters.Append objCmd.CreateParameter("@id", 3, 1, , idComponente)
        Else
            objCmd.CommandText = "INSERT INTO Componentes (ID_Proyecto, Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas) VALUES (?, ?, ?, ?, ?, ?)"
            objCmd.Parameters.Append objCmd.CreateParameter("@id_proyecto", 3, 1, , idProyectoSel)
            objCmd.Parameters.Append objCmd.CreateParameter("@tipo", 200, 1, 100, tipoComponente)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@valor", valorCalculado, 100)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@pin", pinConexion, 50)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@ubicacion", ubicacionProtoboard, 100)
            objCmd.Parameters.Append CrearTextoNull(objCmd, "@notas", notas, 255)
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

    Dim cmdComponente
    Set cmdComponente = Server.CreateObject("ADODB.Command")
    cmdComponente.ActiveConnection = objConn
    cmdComponente.CommandType = 1
    cmdComponente.CommandText = "SELECT ID_Componente, ID_Proyecto, Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas FROM Componentes WHERE ID_Componente = ?"
    cmdComponente.Parameters.Append cmdComponente.CreateParameter("@id", 3, 1, , idComponente)
    Set objRS = cmdComponente.Execute

    If Not objRS.EOF Then
        idProyectoSel       = objRS("ID_Proyecto")
        tipoComponente      = Limpiar(objRS("Tipo_Componente"))
        valorCalculado      = Limpiar(objRS("Valor_Calculado"))
        pinConexion         = Limpiar(objRS("Pin_Conexion"))
        ubicacionProtoboard = Limpiar(objRS("Ubicacion_Protoboard"))
        notas               = Limpiar(objRS("Notas"))
    Else
        modoEdicion = False
        idComponente = 0
    End If

    objRS.Close
    Set objRS = Nothing
    Set cmdComponente = Nothing
    CerrarConexion objConn
End If

If modoEdicion Then
    tituloPagina = "Editar componente - Central de Monitoreo"
Else
    tituloPagina = "Nuevo componente - Central de Monitoreo"
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
    <h1><% If modoEdicion Then %>Editar componente<% Else %>Nuevo componente<% End If %></h1>
    <p class="ayuda">Inventario y mapa de cableado: tipo de pieza, valor calculado y ubicación exacta en la protoboard.</p>

    <% If errores <> "" Then %>
    <div class="alerta alerta-error">
        <strong>Revisa los siguientes campos:</strong>
        <ul><%= errores %></ul>
    </div>
    <% End If %>

    <% If objRS.EOF Then %>
    <div class="alerta alerta-error">
        Aún no hay proyectos creados. <a href="../proyectos/formulario.asp">Crea primero un proyecto</a> antes de agregar componentes.
    </div>
    <% Else %>
    <form method="post" action="formulario.asp<% If modoEdicion Then %>?id=<%= idComponente %><% End If %>">
        <input type="hidden" name="id_componente" value="<%= idComponente %>">

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

        <label for="tipo_componente">Tipo de componente *</label>
        <input type="text" id="tipo_componente" name="tipo_componente" maxlength="100"
               value="<%= Server.HTMLEncode(tipoComponente) %>" placeholder="Ej: Placa de aluminio, Resistencia, Sensor">

        <label for="valor_calculado">Valor calculado</label>
        <input type="text" id="valor_calculado" name="valor_calculado" maxlength="100"
               value="<%= Server.HTMLEncode(valorCalculado) %>" placeholder="Ej: 4.7 kOhm, 100 nF">

        <label for="pin_conexion">Pin de conexión</label>
        <input type="text" id="pin_conexion" name="pin_conexion" maxlength="50"
               value="<%= Server.HTMLEncode(pinConexion) %>" placeholder="Ej: A0, D2">

        <label for="ubicacion_protoboard">Ubicación en protoboard</label>
        <input type="text" id="ubicacion_protoboard" name="ubicacion_protoboard" maxlength="100"
               value="<%= Server.HTMLEncode(ubicacionProtoboard) %>" placeholder="Ej: Fila 12, columna E">

        <label for="notas">Notas</label>
        <textarea id="notas" name="notas" rows="3"><%= Server.HTMLEncode(notas) %></textarea>

        <div class="acciones">
            <button type="submit" class="boton boton-primario">Guardar componente</button>
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
