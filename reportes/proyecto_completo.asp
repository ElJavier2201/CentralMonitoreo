<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase
rutaBase = "../"
%>
<!--#include virtual="/conexion.asp"-->
<!--#include virtual="/includes/auth.asp"-->
<%
Dim objConn, objRS, sql, idProyecto
Dim nombreProyecto, plataforma, micro, descripcion, fechaCreacion, estadoProyecto

idProyecto = IDValido(Request.QueryString("id"))
If idProyecto = 0 Then
    Response.Write "Proyecto no especificado."
    Response.End
End If

Set objConn = AbrirConexion()

On Error Resume Next
sql = "SELECT ID_Proyecto, Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Descripcion, Fecha_Creacion, Estado FROM Proyectos WHERE ID_Proyecto = " & idProyecto
Set objRS = objConn.Execute(sql)
If Err.Number <> 0 Then
    Err.Clear
    sql = "SELECT ID_Proyecto, Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Descripcion, Fecha_Creacion FROM Proyectos WHERE ID_Proyecto = " & idProyecto
    Set objRS = objConn.Execute(sql)
    estadoProyecto = "-"
End If
On Error Goto 0

If objRS.EOF Then
    Response.Write "Proyecto no encontrado."
    Response.End
End If

nombreProyecto = Limpiar(objRS("Nombre_Proyecto"))
plataforma = Limpiar(objRS("Plataforma_Simulacion"))
micro = Limpiar(objRS("Microcontrolador"))
descripcion = Limpiar(objRS("Descripcion"))
fechaCreacion = objRS("Fecha_Creacion")
On Error Resume Next
If estadoProyecto = "" Then estadoProyecto = Limpiar(objRS("Estado"))
If Err.Number <> 0 Then estadoProyecto = "-" : Err.Clear
On Error Goto 0
objRS.Close
Set objRS = Nothing

Sub SeccionTabla(ByVal titulo, ByVal sqlConsulta, ByVal encabezados, ByVal campos)
    Dim rs, i
%>
    <h2><%= titulo %></h2>
<%
    On Error Resume Next
    Set rs = objConn.Execute(sqlConsulta)
    If Err.Number <> 0 Then
        Response.Write "<p class='nota'>Sección no disponible. Verifica que la tabla correspondiente exista.</p>"
        Err.Clear
        On Error Goto 0
        Exit Sub
    End If
    On Error Goto 0
%>
    <table>
        <thead>
            <tr>
                <% For i = 0 To UBound(encabezados) %>
                    <th><%= encabezados(i) %></th>
                <% Next %>
            </tr>
        </thead>
        <tbody>
        <% If rs.EOF Then %>
            <tr><td colspan="<%= UBound(encabezados) + 1 %>" class="vacio">Sin registros.</td></tr>
        <% Else %>
            <% Do While Not rs.EOF %>
                <tr>
                    <% For i = 0 To UBound(campos) %>
                        <td><%= Server.HTMLEncode(Limpiar(rs(campos(i)))) %></td>
                    <% Next %>
                </tr>
            <%
                rs.MoveNext
            Loop
            %>
        <% End If %>
        </tbody>
    </table>
<%
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Reporte completo - <%= Server.HTMLEncode(nombreProyecto) %></title>
<style>
    body { font-family: "Segoe UI", Arial, sans-serif; color:#111; background:#fff; padding:24px; }
    .barra { margin-bottom:20px; }
    .boton { display:inline-block; padding:10px 14px; border:1px solid #333; background:#f2f2f2; color:#111; text-decoration:none; border-radius:5px; cursor:pointer; }
    h1 { font-size:26px; margin:0 0 6px; border-bottom:3px solid #111; padding-bottom:8px; }
    h2 { font-size:18px; margin-top:28px; border-bottom:1px solid #999; padding-bottom:4px; }
    .meta { display:grid; grid-template-columns: 180px 1fr; gap:6px 12px; margin-top:16px; font-size:14px; }
    .meta strong { background:#eee; padding:6px; border:1px solid #ccc; }
    .meta span { padding:6px; border:1px solid #ddd; }
    table { width:100%; border-collapse:collapse; margin-top:10px; font-size:13px; }
    th, td { border:1px solid #444; padding:8px; text-align:left; vertical-align:top; }
    th { background:#eee; text-transform:uppercase; font-size:12px; }
    .vacio, .nota { color:#555; font-style:italic; }
    .firma { margin-top:36px; display:flex; gap:60px; }
    .firma div { width:260px; border-top:1px solid #111; text-align:center; padding-top:8px; }
    @media print { .barra { display:none; } body { padding:0; } }
</style>
</head>
<body>
<div class="barra">
    <button onclick="window.print()" class="boton">Imprimir reporte</button>
    <a href="index.asp" class="boton">Volver a reportes</a>
</div>

<h1>Reporte técnico completo</h1>
<p>Central de Monitoreo para Simulaciones de Circuitos y Prototipos</p>

<div class="meta">
    <strong>Proyecto</strong><span><%= Server.HTMLEncode(nombreProyecto) %></span>
    <strong>Plataforma</strong><span><%= Server.HTMLEncode(plataforma) %></span>
    <strong>Microcontrolador</strong><span><%= Server.HTMLEncode(micro) %></span>
    <strong>Estado</strong><span><%= Server.HTMLEncode(estadoProyecto) %></span>
    <strong>Fecha de creación</strong><span><%= FormatearFechaVisible(fechaCreacion) %></span>
    <strong>Descripción</strong><span><%= Server.HTMLEncode(descripcion) %></span>
</div>

<%
SeccionTabla "1. Componentes y mapa de cableado", _
    "SELECT Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas FROM Componentes WHERE ID_Proyecto = " & idProyecto & " ORDER BY Tipo_Componente", _
    Array("Componente", "Valor", "Pin", "Ubicación", "Notas"), _
    Array("Tipo_Componente", "Valor_Calculado", "Pin_Conexion", "Ubicacion_Protoboard", "Notas")

SeccionTabla "2. Bitácora de fallos", _
    "SELECT Fecha_Registro, Sintoma_Error, Solucion_Aplicada, Estado FROM Bitacora_Fallos WHERE ID_Proyecto = " & idProyecto & " ORDER BY Fecha_Registro DESC, ID_Fallo DESC", _
    Array("Fecha", "Síntoma", "Solución", "Estado"), _
    Array("Fecha_Registro", "Sintoma_Error", "Solucion_Aplicada", "Estado")

SeccionTabla "3. Mediciones simulación vs físico", _
    "SELECT Fecha_Medicion, Variable_Medida, Valor_Simulado, Valor_Fisico, Diferencia, Observaciones FROM Mediciones WHERE ID_Proyecto = " & idProyecto & " ORDER BY Fecha_Medicion DESC, ID_Medicion DESC", _
    Array("Fecha", "Variable", "Simulado", "Físico", "Diferencia", "Observaciones"), _
    Array("Fecha_Medicion", "Variable_Medida", "Valor_Simulado", "Valor_Fisico", "Diferencia", "Observaciones")

SeccionTabla "4. Checklist de armado", _
    "SELECT Descripcion, Completado, Fecha_Actualizacion FROM Checklist_Proyecto WHERE ID_Proyecto = " & idProyecto & " ORDER BY ID_Item", _
    Array("Actividad", "Completado", "Actualización"), _
    Array("Descripcion", "Completado", "Fecha_Actualizacion")

SeccionTabla "5. Evidencias", _
    "SELECT Fecha_Subida, Tipo_Evidencia, Ruta_Archivo, Descripcion FROM Evidencias WHERE ID_Proyecto = " & idProyecto & " ORDER BY Fecha_Subida DESC, ID_Evidencia DESC", _
    Array("Fecha", "Tipo", "Ruta / enlace", "Descripción"), _
    Array("Fecha_Subida", "Tipo_Evidencia", "Ruta_Archivo", "Descripcion")
%>

<h2>6. Observaciones finales</h2>
<p style="min-height:70px; border:1px solid #ccc; padding:10px;">&nbsp;</p>

<div class="firma">
    <div>Responsable del proyecto</div>
    <div>Fecha de revisión</div>
</div>

<p style="margin-top:30px; font-size:12px; text-align:center; color:#555;">
    Generado el <%= Date() %> a las <%= Time() %> desde la Central de Monitoreo.
</p>

</body>
</html>
<%
CerrarConexion objConn
%>
