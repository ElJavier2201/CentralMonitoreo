<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
%>
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim idComponente, modoEdicion
Dim idProyectoSel, tipoComponente, valorCalculado, pinConexion, ubicacionProtoboard, notas
Dim errores
Dim rutaBase, tituloPagina, seccionActiva

rutaBase = "../"
seccionActiva = "componentes"
errores = ""

idComponente = IDValido(Request.QueryString("id"))
modoEdicion = (idComponente > 0)

idProyectoSel        = IDValido(Request.QueryString("proyecto"))
tipoComponente       = ""
valorCalculado       = ""
pinConexion          = ""
ubicacionProtoboard  = ""
notas                = ""

' ===================================================================
' PROCESAMIENTO DEL FORMULARIO (POST) - validación + SQL nativo
' ===================================================================
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

        If modoEdicion Then
            sql = "UPDATE Componentes SET " & _
                  "ID_Proyecto = " & idProyectoSel & ", " & _
                  "Tipo_Componente = " & ValorTextoSQL(tipoComponente) & ", " & _
                  "Valor_Calculado = " & ValorTextoSQL(valorCalculado) & ", " & _
                  "Pin_Conexion = " & ValorTextoSQL(pinConexion) & ", " & _
                  "Ubicacion_Protoboard = " & ValorTextoSQL(ubicacionProtoboard) & ", " & _
                  "Notas = " & ValorTextoSQL(notas) & " " & _
                  "WHERE ID_Componente = " & idComponente
        Else
            sql = "INSERT INTO Componentes " & _
                  "(ID_Proyecto, Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas) " & _
                  "VALUES (" & idProyectoSel & ", " & _
                  ValorTextoSQL(tipoComponente) & ", " & _
                  ValorTextoSQL(valorCalculado) & ", " & _
                  ValorTextoSQL(pinConexion) & ", " & _
                  ValorTextoSQL(ubicacionProtoboard) & ", " & _
                  ValorTextoSQL(notas) & ")"
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
    sql = "SELECT ID_Componente, ID_Proyecto, Tipo_Componente, Valor_Calculado, Pin_Conexion, " & _
          "Ubicacion_Protoboard, Notas FROM Componentes WHERE ID_Componente = " & idComponente
    Set objRS = objConn.Execute(sql)

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
    CerrarConexion objConn
End If

If modoEdicion Then
    tituloPagina = "Editar componente - Central de Monitoreo"
Else
    tituloPagina = "Nuevo componente - Central de Monitoreo"
End If

' Recordset auxiliar para llenar el combo de proyectos disponibles
Set objConn = AbrirConexion()
sql = "SELECT ID_Proyecto, Nombre_Proyecto FROM Proyectos ORDER BY Nombre_Proyecto"
Set objRS = objConn.Execute(sql)
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
