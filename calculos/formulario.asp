<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Cálculo eléctrico - Central de Monitoreo"
seccionActiva = "calculos"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Function Numero(ByVal valor)
    Dim v
    v = Trim(CStr(valor))
    v = Replace(v, ",", ".")
    If IsNumeric(v) Then
        Numero = CDbl(v)
    Else
        Numero = Null
    End If
End Function

Function TieneNumero(ByVal valor)
    Dim v
    v = Trim(CStr(valor))
    v = Replace(v, ",", ".")
    TieneNumero = (v <> "" And IsNumeric(v))
End Function

Function FormatoOhms(ByVal valor)
    FormatoOhms = Replace(FormatNumber(valor, 3), ",", "") & " Ω"
End Function

Function FormatoVolts(ByVal valor)
    FormatoVolts = Replace(FormatNumber(valor, 3), ",", "") & " V"
End Function

Function FormatoAmps(ByVal valor)
    FormatoAmps = Replace(FormatNumber(valor, 6), ",", "") & " A"
End Function

Function CalcularResultado(ByVal tipo, ByVal a, ByVal b, ByVal c, ByRef entrada, ByRef resultado, ByRef mensajeError)
    Dim n1, n2, n3, suma, inv
    entrada = ""
    resultado = ""
    mensajeError = ""

    If tipo = "Serie" Then
        suma = 0
        If TieneNumero(a) Then suma = suma + Numero(a)
        If TieneNumero(b) Then suma = suma + Numero(b)
        If TieneNumero(c) Then suma = suma + Numero(c)
        If suma <= 0 Then
            mensajeError = "Ingresa al menos una resistencia mayor a 0."
        Else
            entrada = "R1=" & a & " Ω; R2=" & b & " Ω; R3=" & c & " Ω"
            resultado = "Req serie = " & FormatoOhms(suma)
        End If
    ElseIf tipo = "Paralelo" Then
        inv = 0
        If TieneNumero(a) Then
            n1 = Numero(a)
            If n1 > 0 Then inv = inv + (1 / n1)
        End If
        If TieneNumero(b) Then
            n2 = Numero(b)
            If n2 > 0 Then inv = inv + (1 / n2)
        End If
        If TieneNumero(c) Then
            n3 = Numero(c)
            If n3 > 0 Then inv = inv + (1 / n3)
        End If
        If inv <= 0 Then
            mensajeError = "Ingresa al menos una resistencia mayor a 0."
        Else
            entrada = "R1=" & a & " Ω; R2=" & b & " Ω; R3=" & c & " Ω"
            resultado = "Req paralelo = " & FormatoOhms(1 / inv)
        End If
    ElseIf tipo = "Ley de Ohm - Voltaje" Then
        If Not TieneNumero(a) Or Not TieneNumero(b) Then
            mensajeError = "Ingresa corriente y resistencia."
        Else
            entrada = "I=" & a & " A; R=" & b & " Ω"
            resultado = "V = " & FormatoVolts(Numero(a) * Numero(b))
        End If
    ElseIf tipo = "Ley de Ohm - Corriente" Then
        If Not TieneNumero(a) Or Not TieneNumero(b) Or Numero(b) = 0 Then
            mensajeError = "Ingresa voltaje y resistencia mayor a 0."
        Else
            entrada = "V=" & a & " V; R=" & b & " Ω"
            resultado = "I = " & FormatoAmps(Numero(a) / Numero(b))
        End If
    ElseIf tipo = "Ley de Ohm - Resistencia" Then
        If Not TieneNumero(a) Or Not TieneNumero(b) Or Numero(b) = 0 Then
            mensajeError = "Ingresa voltaje y corriente mayor a 0."
        Else
            entrada = "V=" & a & " V; I=" & b & " A"
            resultado = "R = " & FormatoOhms(Numero(a) / Numero(b))
        End If
    ElseIf tipo = "Divisor de voltaje" Then
        If Not TieneNumero(a) Or Not TieneNumero(b) Or Not TieneNumero(c) Or (Numero(b) + Numero(c)) = 0 Then
            mensajeError = "Ingresa Vin, R1 y R2. R1 + R2 debe ser mayor a 0."
        Else
            entrada = "Vin=" & a & " V; R1=" & b & " Ω; R2=" & c & " Ω"
            resultado = "Vout = " & FormatoVolts(Numero(a) * (Numero(c) / (Numero(b) + Numero(c))))
        End If
    ElseIf tipo = "Error porcentual" Then
        If Not TieneNumero(a) Or Not TieneNumero(b) Or Numero(a) = 0 Then
            mensajeError = "Ingresa valor teórico mayor a 0 y valor físico."
        Else
            entrada = "Teórico=" & a & "; Físico=" & b
            resultado = "Error = " & Replace(FormatNumber(Abs((Numero(b) - Numero(a)) / Numero(a)) * 100, 3), ",", "") & " %"
        End If
    Else
        mensajeError = "Selecciona un tipo de cálculo válido."
    End If
End Function

Dim objConn, objRS, cmd, sql
Dim idCalculo, modoEdicion, idProyectoSel, tipoCalculo, entrada, resultado, fechaCalculo
Dim v1, v2, v3, errores, resultadoCalculado, entradaCalculada, errorCalc

idCalculo = IDValido(Request.QueryString("id"))
modoEdicion = (idCalculo > 0)
idProyectoSel = IDValido(Request.QueryString("proyecto"))
tipoCalculo = "Serie"
entrada = ""
resultado = ""
fechaCalculo = Date()
v1 = ""
v2 = ""
v3 = ""
errores = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    idCalculo = IDValido(Request.Form("id_calculo"))
    modoEdicion = (idCalculo > 0)
    idProyectoSel = IDValido(Request.Form("id_proyecto"))
    tipoCalculo = Limpiar(Request.Form("tipo_calculo"))
    fechaCalculo = Limpiar(Request.Form("fecha_calculo"))

    If Not IsDate(fechaCalculo) Then fechaCalculo = Date()
    If idProyectoSel = 0 Then errores = errores & "<li>Selecciona un proyecto.</li>"

    If modoEdicion Then
        entrada = Limpiar(Request.Form("entrada"))
        resultado = Limpiar(Request.Form("resultado"))
        If tipoCalculo = "" Then errores = errores & "<li>El tipo de cálculo es obligatorio.</li>"
        If entrada = "" Then errores = errores & "<li>La entrada del cálculo es obligatoria.</li>"
        If resultado = "" Then errores = errores & "<li>El resultado es obligatorio.</li>"
    Else
        v1 = Limpiar(Request.Form("valor1"))
        v2 = Limpiar(Request.Form("valor2"))
        v3 = Limpiar(Request.Form("valor3"))
        Call CalcularResultado(tipoCalculo, v1, v2, v3, entradaCalculada, resultadoCalculado, errorCalc)
        If errorCalc <> "" Then errores = errores & "<li>" & errorCalc & "</li>"
        entrada = entradaCalculada
        resultado = resultadoCalculado
    End If

    If errores = "" Then
        Set objConn = AbrirConexion()
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = objConn
        cmd.CommandType = 1

        If modoEdicion Then
            cmd.CommandText = "UPDATE Calculos_Electricos SET ID_Proyecto=?, Tipo_Calculo=?, Entrada=?, Resultado=?, Fecha_Calculo=? WHERE ID_Calculo=?"
            cmd.Parameters.Append cmd.CreateParameter("@idp", 3, 1, , idProyectoSel)
            cmd.Parameters.Append cmd.CreateParameter("@tipo", 200, 1, 50, tipoCalculo)
            cmd.Parameters.Append cmd.CreateParameter("@entrada", 200, 1, 255, entrada)
            cmd.Parameters.Append cmd.CreateParameter("@resultado", 200, 1, 255, resultado)
            cmd.Parameters.Append cmd.CreateParameter("@fecha", 135, 1, , fechaCalculo)
            cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idCalculo)
        Else
            cmd.CommandText = "INSERT INTO Calculos_Electricos (ID_Proyecto, Tipo_Calculo, Entrada, Resultado, Fecha_Calculo) VALUES (?, ?, ?, ?, ?)"
            cmd.Parameters.Append cmd.CreateParameter("@idp", 3, 1, , idProyectoSel)
            cmd.Parameters.Append cmd.CreateParameter("@tipo", 200, 1, 50, tipoCalculo)
            cmd.Parameters.Append cmd.CreateParameter("@entrada", 200, 1, 255, entrada)
            cmd.Parameters.Append cmd.CreateParameter("@resultado", 200, 1, 255, resultado)
            cmd.Parameters.Append cmd.CreateParameter("@fecha", 135, 1, , fechaCalculo)
        End If

        cmd.Execute , , 128
        Set cmd = Nothing
        CerrarConexion objConn
        Response.Redirect "listar.asp?ok=1"
        Response.End
    End If
End If

If modoEdicion And Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Set objConn = AbrirConexion()
    sql = "SELECT ID_Calculo, ID_Proyecto, Tipo_Calculo, Entrada, Resultado, Fecha_Calculo FROM Calculos_Electricos WHERE ID_Calculo = " & idCalculo
    Set objRS = objConn.Execute(sql)
    If Not objRS.EOF Then
        idProyectoSel = objRS("ID_Proyecto")
        tipoCalculo = Limpiar(objRS("Tipo_Calculo"))
        entrada = Limpiar(objRS("Entrada"))
        resultado = Limpiar(objRS("Resultado"))
        fechaCalculo = objRS("Fecha_Calculo")
    Else
        modoEdicion = False
        idCalculo = 0
    End If
    objRS.Close
    Set objRS = Nothing
    CerrarConexion objConn
End If

Set objConn = AbrirConexion()
Set objRS = objConn.Execute("SELECT ID_Proyecto, Nombre_Proyecto FROM Proyectos ORDER BY Nombre_Proyecto")
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <h1><% If modoEdicion Then %>Editar cálculo<% Else %>Nuevo cálculo eléctrico<% End If %></h1>
    <p class="ayuda">Calcula y guarda resultados técnicos asociados al proyecto.</p>

    <% If errores <> "" Then %>
        <div class="alerta alerta-error"><strong>Revisa:</strong><ul><%= errores %></ul></div>
    <% End If %>

    <% If objRS.EOF Then %>
        <div class="alerta alerta-error">Primero crea un proyecto antes de guardar cálculos.</div>
    <% Else %>
    <form method="post" action="formulario.asp<% If modoEdicion Then %>?id=<%= idCalculo %><% End If %>">
        <input type="hidden" name="id_calculo" value="<%= idCalculo %>">

        <label for="id_proyecto">Proyecto *</label>
        <select id="id_proyecto" name="id_proyecto">
            <option value="">-- Seleccione --</option>
            <% Do While Not objRS.EOF %>
                <option value="<%= objRS("ID_Proyecto") %>" <% If CLng(objRS("ID_Proyecto")) = CLng(idProyectoSel) Then %>selected<% End If %>><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></option>
            <% objRS.MoveNext : Loop %>
        </select>

        <label for="tipo_calculo">Tipo de cálculo *</label>
        <select id="tipo_calculo" name="tipo_calculo">
            <option value="Serie" <% If tipoCalculo = "Serie" Then %>selected<% End If %>>Resistencias en serie</option>
            <option value="Paralelo" <% If tipoCalculo = "Paralelo" Then %>selected<% End If %>>Resistencias en paralelo</option>
            <option value="Ley de Ohm - Voltaje" <% If tipoCalculo = "Ley de Ohm - Voltaje" Then %>selected<% End If %>>Ley de Ohm: calcular voltaje</option>
            <option value="Ley de Ohm - Corriente" <% If tipoCalculo = "Ley de Ohm - Corriente" Then %>selected<% End If %>>Ley de Ohm: calcular corriente</option>
            <option value="Ley de Ohm - Resistencia" <% If tipoCalculo = "Ley de Ohm - Resistencia" Then %>selected<% End If %>>Ley de Ohm: calcular resistencia</option>
            <option value="Divisor de voltaje" <% If tipoCalculo = "Divisor de voltaje" Then %>selected<% End If %>>Divisor de voltaje</option>
            <option value="Error porcentual" <% If tipoCalculo = "Error porcentual" Then %>selected<% End If %>>Error porcentual</option>
        </select>

        <label for="fecha_calculo">Fecha</label>
        <input type="date" id="fecha_calculo" name="fecha_calculo" value="<%= FormatearFechaInput(fechaCalculo) %>">

        <% If modoEdicion Then %>
            <label for="entrada">Entrada *</label>
            <textarea id="entrada" name="entrada" rows="3"><%= Server.HTMLEncode(entrada) %></textarea>
            <label for="resultado">Resultado *</label>
            <textarea id="resultado" name="resultado" rows="3"><%= Server.HTMLEncode(resultado) %></textarea>
        <% Else %>
            <div class="caja-ayuda">
                <strong>Uso de campos:</strong><br>
                Serie/Paralelo: Valor 1, Valor 2 y Valor 3 son resistencias en Ω.<br>
                Ohm voltaje: Valor 1=I(A), Valor 2=R(Ω). Ohm corriente: Valor 1=V, Valor 2=R. Ohm resistencia: Valor 1=V, Valor 2=I.<br>
                Divisor: Valor 1=Vin, Valor 2=R1, Valor 3=R2. Error: Valor 1=Teórico, Valor 2=Físico.
            </div>

            <label for="valor1">Valor 1</label>
            <input type="text" id="valor1" name="valor1" value="<%= Server.HTMLEncode(v1) %>" placeholder="Ej: 1000">

            <label for="valor2">Valor 2</label>
            <input type="text" id="valor2" name="valor2" value="<%= Server.HTMLEncode(v2) %>" placeholder="Ej: 2200">

            <label for="valor3">Valor 3</label>
            <input type="text" id="valor3" name="valor3" value="<%= Server.HTMLEncode(v3) %>" placeholder="Ej: 3300">
        <% End If %>

        <div class="acciones">
            <button type="submit" class="boton boton-primario">Guardar cálculo</button>
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
