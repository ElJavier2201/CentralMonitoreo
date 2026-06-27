<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"

Dim rutaBase
rutaBase = "../"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, idProyecto, nombreProyecto, micro
Dim cmdProyecto, cmdComponentes

idProyecto = IDValido(Request.QueryString("id"))

If idProyecto = 0 Then
    Response.Write "Error: Proyecto no especificado."
    Response.End
End If

Set objConn = AbrirConexion()

Set cmdProyecto = Server.CreateObject("ADODB.Command")
cmdProyecto.ActiveConnection = objConn
cmdProyecto.CommandType = 1
cmdProyecto.CommandText = "SELECT Nombre_Proyecto, Microcontrolador FROM Proyectos WHERE ID_Proyecto = ?"
cmdProyecto.Parameters.Append cmdProyecto.CreateParameter("@id", 3, 1, , idProyecto)
Set objRS = cmdProyecto.Execute

If objRS.EOF Then
    objRS.Close
    Set objRS = Nothing
    Set cmdProyecto = Nothing
    CerrarConexion objConn
    Response.Write "Proyecto no encontrado."
    Response.End
End If

nombreProyecto = objRS("Nombre_Proyecto")
micro = objRS("Microcontrolador")
objRS.Close
Set objRS = Nothing
Set cmdProyecto = Nothing

Set cmdComponentes = Server.CreateObject("ADODB.Command")
cmdComponentes.ActiveConnection = objConn
cmdComponentes.CommandType = 1
cmdComponentes.CommandText = "SELECT Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas FROM Componentes WHERE ID_Proyecto = ? ORDER BY Tipo_Componente"
cmdComponentes.Parameters.Append cmdComponentes.CreateParameter("@id", 3, 1, , idProyecto)
Set objRS = cmdComponentes.Execute
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Plano de Armado - <%= Server.HTMLEncode(nombreProyecto) %></title>
    <style>
        body { font-family: "Segoe UI", Arial, sans-serif; color: #000; background: #fff; padding: 20px; }
        h1 { font-size: 24px; border-bottom: 2px solid #000; padding-bottom: 5px; margin-bottom: 5px;}
        h3 { font-size: 16px; color: #333; margin-top: 0; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px; }
        th, td { border: 1px solid #000; padding: 10px; text-align: left; }
        th { background: #eee; text-transform: uppercase; }
        .pin { font-family: "Consolas", monospace; font-weight: bold; }
        @media print { .no-imprimir { display: none !important; } }
    </style>
</head>
<body>
    <div class="no-imprimir" style="margin-bottom: 20px;">
        <button onclick="window.print()" style="padding: 10px 15px; font-weight: bold; cursor:pointer;">🖨️ Imprimir Plano</button>
        <button onclick="window.close()" style="padding: 10px 15px; cursor:pointer;">Cerrar</button>
    </div>

    <h1>Guía de Ensamblaje: <%= Server.HTMLEncode(nombreProyecto) %></h1>
    <h3>Microcontrolador Principal: <%= Server.HTMLEncode(micro) %></h3>

    <table>
        <thead>
            <tr>
                <th>Componente</th>
                <th>Valor / Especificación</th>
                <th>Pin de Conexión</th>
                <th>Ubicación en Protoboard</th>
            </tr>
        </thead>
        <tbody>
            <% If objRS.EOF Then %>
            <tr><td colspan="4">No hay componentes registrados.</td></tr>
            <% Else
                Do While Not objRS.EOF
            %>
            <tr>
                <td><%= Server.HTMLEncode(objRS("Tipo_Componente")) %></td>
                <td><%= Server.HTMLEncode(objRS("Valor_Calculado")) %></td>
                <td class="pin"><%= Server.HTMLEncode(objRS("Pin_Conexion")) %></td>
                <td><%= Server.HTMLEncode(objRS("Ubicacion_Protoboard")) %></td>
            </tr>
            <%
                    objRS.MoveNext
                Loop
            End If
            %>
        </tbody>
    </table>

    <p style="margin-top: 30px; font-size: 12px; font-style: italic; text-align: center;">
        Generado desde la Central de Monitoreo el <%= Date() %> a las <%= Time() %>
    </p>

    <script>
        window.onload = function() { window.print(); }
    </script>
</body>
</html>
<%
objRS.Close
Set objRS = Nothing
Set cmdComponentes = Nothing
CerrarConexion objConn
%>
