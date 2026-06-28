<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Ayuda técnica - Central de Monitoreo"
seccionActiva = "ayuda"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <h1>Ayuda técnica de circuitos y prototipos</h1>
    <p class="ayuda">Guía rápida para documentación, armado en protoboard, fórmulas y revisión de errores comunes.</p>

    <h2>1. Fórmulas básicas</h2>
    <table class="tabla-datos">
        <thead><tr><th>Tema</th><th>Fórmula</th><th>Uso</th></tr></thead>
        <tbody>
            <tr><td>Ley de Ohm</td><td><span class="codigo">V = I × R</span></td><td>Calcular voltaje, corriente o resistencia.</td></tr>
            <tr><td>Resistencia en serie</td><td><span class="codigo">Req = R1 + R2 + R3 + ...</span></td><td>Cuando las resistencias están una después de otra.</td></tr>
            <tr><td>Resistencia en paralelo</td><td><span class="codigo">1/Req = 1/R1 + 1/R2 + ...</span></td><td>Cuando las resistencias comparten los mismos dos nodos.</td></tr>
            <tr><td>Divisor de voltaje</td><td><span class="codigo">Vout = Vin × R2 / (R1 + R2)</span></td><td>Obtener una tensión menor a partir de dos resistencias.</td></tr>
            <tr><td>Error porcentual</td><td><span class="codigo">Error = |Físico - Teórico| / Teórico × 100</span></td><td>Comparar simulación o teoría contra medición real.</td></tr>
        </tbody>
    </table>

    <h2>2. Checklist básico antes de energizar</h2>
    <ul class="lista-ayuda">
        <li>Verifica continuidad en líneas de alimentación y tierra.</li>
        <li>Confirma que no exista corto entre <strong>5V</strong> y <strong>GND</strong>.</li>
        <li>Revisa polaridad de LEDs, capacitores electrolíticos y diodos.</li>
        <li>Confirma valores de resistencias con código de colores y multímetro.</li>
        <li>Verifica que cada pin del microcontrolador coincida con el plano registrado.</li>
        <li>Haz primero una prueba con bajo consumo antes de conectar sensores o cargas.</li>
    </ul>

    <h2>3. Errores comunes en protoboard</h2>
    <table class="tabla-datos">
        <thead><tr><th>Error</th><th>Causa probable</th><th>Solución</th></tr></thead>
        <tbody>
            <tr><td>Lectura analógica inestable</td><td>Entrada flotante o ruido eléctrico.</td><td>Agregar resistencia pull-down/pull-up y filtrar lecturas.</td></tr>
            <tr><td>El circuito no enciende</td><td>Alimentación mal conectada o rieles partidos.</td><td>Revisar rieles positivo/negativo y continuidad.</td></tr>
            <tr><td>LED no prende</td><td>Polaridad invertida o falta resistencia.</td><td>Revisar ánodo/cátodo y colocar resistencia limitadora.</td></tr>
            <tr><td>Valor físico no coincide con simulación</td><td>Tolerancia de componentes o conexiones reales.</td><td>Medir componentes reales y recalibrar.</td></tr>
        </tbody>
    </table>

    <h2>4. Código de colores de resistencias</h2>
    <table class="tabla-datos">
        <thead><tr><th>Color</th><th>Dígito</th><th>Multiplicador</th></tr></thead>
        <tbody>
            <tr><td>Negro</td><td>0</td><td>×1</td></tr>
            <tr><td>Café</td><td>1</td><td>×10</td></tr>
            <tr><td>Rojo</td><td>2</td><td>×100</td></tr>
            <tr><td>Naranja</td><td>3</td><td>×1,000</td></tr>
            <tr><td>Amarillo</td><td>4</td><td>×10,000</td></tr>
            <tr><td>Verde</td><td>5</td><td>×100,000</td></tr>
            <tr><td>Azul</td><td>6</td><td>×1,000,000</td></tr>
            <tr><td>Violeta</td><td>7</td><td>×10,000,000</td></tr>
            <tr><td>Gris</td><td>8</td><td>×100,000,000</td></tr>
            <tr><td>Blanco</td><td>9</td><td>×1,000,000,000</td></tr>
        </tbody>
    </table>

    <h2>5. Flujo recomendado de documentación</h2>
    <ol class="lista-ayuda">
        <li>Crea el proyecto.</li>
        <li>Registra componentes y ubicación en protoboard.</li>
        <li>Guarda cálculos eléctricos relevantes.</li>
        <li>Registra mediciones simuladas y físicas.</li>
        <li>Sube evidencias: captura del simulador, foto del protoboard y monitor serial.</li>
        <li>Documenta fallos encontrados y su solución.</li>
        <li>Genera el reporte técnico completo.</li>
    </ol>
</div>

<!--#include virtual="/includes/footer.asp"-->
