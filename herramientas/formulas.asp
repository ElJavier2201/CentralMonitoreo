<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Fórmulas y calculadora - Central de Monitoreo"
seccionActiva = "herramientas"
%>
<!--#include virtual="/conexion.asp"-->
<!--#include virtual="/includes/auth.asp"-->
<%
Function DecimalSistema()
    DecimalSistema = Mid(CStr(1.1), 2, 1)
End Function

Function Num(ByVal valor)
    Dim v, sep
    v = Trim(CStr(valor))
    v = Replace(v, "Ω", "")
    v = Replace(v, "ohm", "")
    v = Replace(v, "Ohm", "")
    v = Replace(v, " ", "")
    sep = DecimalSistema()
    If sep = "," Then
        v = Replace(v, ".", ",")
    Else
        v = Replace(v, ",", ".")
    End If
    If v = "" Or Not IsNumeric(v) Then
        Num = Null
    Else
        Num = CDbl(v)
    End If
End Function

Function Fmt(ByVal valor)
    If IsNull(valor) Then
        Fmt = "-"
    Else
        Fmt = Replace(FormatNumber(valor, 4, -1, 0, 0), ",0000", "")
    End If
End Function

Function UnidadResistencia(ByVal r)
    If IsNull(r) Then
        UnidadResistencia = "-"
    ElseIf Abs(r) >= 1000000 Then
        UnidadResistencia = Fmt(r / 1000000) & " MΩ"
    ElseIf Abs(r) >= 1000 Then
        UnidadResistencia = Fmt(r / 1000) & " kΩ"
    Else
        UnidadResistencia = Fmt(r) & " Ω"
    End If
End Function

Function Hay(ByVal v)
    Hay = Not IsNull(v)
End Function

Dim resistenciasTxt, partes, i, r, serie, invParalelo, cuenta, paralelo, errorRes
Dim vOhm, iOhm, rOhm, resOhmTexto
Dim vin, r1, r2, vout
Dim teorico, fisico, errorPorc

vOhm = Null
iOhm = Null
rOhm = Null
vin = Null
r1 = Null
r2 = Null
vout = Null
teorico = Null
fisico = Null
errorPorc = Null
resOhmTexto = ""

resistenciasTxt = Limpiar(Request.Form("resistencias"))
serie = 0
invParalelo = 0
cuenta = 0
paralelo = Null
errorRes = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    partes = Split(Replace(Replace(resistenciasTxt, vbCrLf, ";"), vbLf, ";"), ";")
    For i = 0 To UBound(partes)
        r = Num(partes(i))
        If Hay(r) Then
            If r > 0 Then
                serie = serie + r
                invParalelo = invParalelo + (1 / r)
                cuenta = cuenta + 1
            End If
        End If
    Next
    If cuenta > 0 And invParalelo > 0 Then paralelo = 1 / invParalelo

    vOhm = Num(Request.Form("voltaje"))
    iOhm = Num(Request.Form("corriente"))
    rOhm = Num(Request.Form("resistencia_ohm"))
    resOhmTexto = ""

    If Not Hay(vOhm) And Hay(iOhm) And Hay(rOhm) Then
        resOhmTexto = "Voltaje calculado: " & Fmt(iOhm * rOhm) & " V"
    ElseIf Hay(vOhm) And Not Hay(iOhm) And Hay(rOhm) And rOhm <> 0 Then
        resOhmTexto = "Corriente calculada: " & Fmt(vOhm / rOhm) & " A"
    ElseIf Hay(vOhm) And Hay(iOhm) And Not Hay(rOhm) And iOhm <> 0 Then
        resOhmTexto = "Resistencia calculada: " & UnidadResistencia(vOhm / iOhm)
    ElseIf Hay(vOhm) And Hay(iOhm) And Hay(rOhm) Then
        resOhmTexto = "Comprobación: V esperado = " & Fmt(iOhm * rOhm) & " V"
    End If

    vin = Num(Request.Form("vin"))
    r1 = Num(Request.Form("r1"))
    r2 = Num(Request.Form("r2"))
    vout = Null
    If Hay(vin) And Hay(r1) And Hay(r2) And (r1 + r2) <> 0 Then
        vout = vin * r2 / (r1 + r2)
    End If

    teorico = Num(Request.Form("teorico"))
    fisico = Num(Request.Form("fisico"))
    errorPorc = Null
    If Hay(teorico) And Hay(fisico) And teorico <> 0 Then
        errorPorc = Abs(fisico - teorico) / Abs(teorico) * 100
    End If
End If
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <div class="panel-cabecera">
        <h1>Fórmulas y calculadora de circuitos</h1>
    </div>
    <p class="ayuda">Herramientas básicas para circuitos resistivos, Ley de Ohm, divisor de voltaje y error porcentual.</p>

    <form method="post" action="formulas.asp">
        <h2>1. Resistencia equivalente</h2>
        <p class="ayuda">Escribe resistencias en ohms separadas por punto y coma. Ejemplo: <span class="codigo">1000; 2200; 4700</span></p>
        <label for="resistencias">Resistencias</label>
        <textarea id="resistencias" name="resistencias" rows="3"><%= Server.HTMLEncode(resistenciasTxt) %></textarea>

        <% If Request.ServerVariables("REQUEST_METHOD") = "POST" Then %>
            <div class="tarjetas">
                <div class="tarjeta">
                    <span class="tarjeta-numero"><%= cuenta %></span>
                    <span class="tarjeta-etiqueta">Resistencias válidas</span>
                </div>
                <div class="tarjeta">
                    <span class="tarjeta-numero"><%= UnidadResistencia(serie) %></span>
                    <span class="tarjeta-etiqueta">Equivalente en serie</span>
                </div>
                <div class="tarjeta">
                    <span class="tarjeta-numero"><%= UnidadResistencia(paralelo) %></span>
                    <span class="tarjeta-etiqueta">Equivalente en paralelo</span>
                </div>
            </div>
        <% End If %>

        <h2>2. Ley de Ohm</h2>
        <p class="ayuda">Llena dos campos y deja uno vacío para calcularlo. Fórmula: <strong>V = I × R</strong>.</p>
        <div class="grid-formulas">
            <div>
                <label for="voltaje">Voltaje V</label>
                <input type="text" id="voltaje" name="voltaje" value="<%= Server.HTMLEncode(Request.Form("voltaje")) %>" placeholder="Ej: 5">
            </div>
            <div>
                <label for="corriente">Corriente I en amperes</label>
                <input type="text" id="corriente" name="corriente" value="<%= Server.HTMLEncode(Request.Form("corriente")) %>" placeholder="Ej: 0.02">
            </div>
            <div>
                <label for="resistencia_ohm">Resistencia R en Ω</label>
                <input type="text" id="resistencia_ohm" name="resistencia_ohm" value="<%= Server.HTMLEncode(Request.Form("resistencia_ohm")) %>" placeholder="Ej: 220">
            </div>
        </div>
        <% If resOhmTexto <> "" Then %>
            <div class="alerta alerta-ok"><%= Server.HTMLEncode(resOhmTexto) %></div>
        <% End If %>

        <h2>3. Divisor de voltaje</h2>
        <p class="ayuda">Fórmula: <strong>Vout = Vin × R2 / (R1 + R2)</strong>.</p>
        <div class="grid-formulas">
            <div>
                <label for="vin">Vin</label>
                <input type="text" id="vin" name="vin" value="<%= Server.HTMLEncode(Request.Form("vin")) %>" placeholder="Ej: 5">
            </div>
            <div>
                <label for="r1">R1 en Ω</label>
                <input type="text" id="r1" name="r1" value="<%= Server.HTMLEncode(Request.Form("r1")) %>" placeholder="Ej: 10000">
            </div>
            <div>
                <label for="r2">R2 en Ω</label>
                <input type="text" id="r2" name="r2" value="<%= Server.HTMLEncode(Request.Form("r2")) %>" placeholder="Ej: 10000">
            </div>
        </div>
        <% If Hay(vout) Then %>
            <div class="alerta alerta-ok">Vout calculado: <strong><%= Fmt(vout) %> V</strong></div>
        <% End If %>

        <h2>4. Error porcentual</h2>
        <p class="ayuda">Útil para comparar valor teórico/simulado contra medición física.</p>
        <div class="grid-formulas">
            <div>
                <label for="teorico">Valor teórico o simulado</label>
                <input type="text" id="teorico" name="teorico" value="<%= Server.HTMLEncode(Request.Form("teorico")) %>" placeholder="Ej: 5">
            </div>
            <div>
                <label for="fisico">Valor físico medido</label>
                <input type="text" id="fisico" name="fisico" value="<%= Server.HTMLEncode(Request.Form("fisico")) %>" placeholder="Ej: 4.82">
            </div>
        </div>
        <% If Hay(errorPorc) Then %>
            <div class="alerta alerta-ok">Error porcentual: <strong><%= Fmt(errorPorc) %>%</strong></div>
        <% End If %>

        <div class="acciones">
            <button type="submit" class="boton boton-primario">Calcular</button>
            <a href="formulas.asp" class="boton boton-secundario">Limpiar</a>
        </div>
    </form>
</div>

<div class="panel" style="margin-top:18px;">
    <h1>Formulario rápido de fórmulas</h1>
    <table class="tabla-datos">
        <thead><tr><th>Concepto</th><th>Fórmula</th><th>Uso</th></tr></thead>
        <tbody>
            <tr><td>Ley de Ohm</td><td><span class="codigo">V = I × R</span></td><td>Voltaje, corriente y resistencia.</td></tr>
            <tr><td>Serie</td><td><span class="codigo">Req = R1 + R2 + ...</span></td><td>Resistencias una tras otra.</td></tr>
            <tr><td>Paralelo</td><td><span class="codigo">1/Req = 1/R1 + 1/R2 + ...</span></td><td>Resistencias con extremos comunes.</td></tr>
            <tr><td>Divisor de voltaje</td><td><span class="codigo">Vout = Vin × R2 / (R1 + R2)</span></td><td>Lecturas analógicas y sensores.</td></tr>
            <tr><td>Error porcentual</td><td><span class="codigo">|medido - teórico| / |teórico| × 100</span></td><td>Comparar simulación contra físico.</td></tr>
        </tbody>
    </table>
</div>

<!--#include virtual="/includes/footer.asp"-->
