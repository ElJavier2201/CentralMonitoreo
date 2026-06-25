<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit 
Response.CharSet = "UTF-8"
%>
<%
Dim objConn, objRS, sql, idProyecto, nombreProyecto, micro

idProyecto = IDValido(Request.QueryString("id"))

If idProyecto = 0 Then
    Response.Write "Error: Proyecto no especificado."
    Response.End
End If

Set objConn = AbrirConexion()

' Extraemos primero la cabecera del proyecto
sql = "SELECT Nombre_Proyecto, Microcontrolador FROM Proyectos WHERE ID_Proyecto = " & idProyecto
Set objRS = objConn.Execute(sql)
If objRS.EOF Then
    Response.Write "Proyecto no encontrado."
    Response.End
End If
nombreProyecto = objRS("Nombre_Proyecto")
micro = objRS("Microcontrolador")
objRS.Close

' Ahora extraemos el listado de componentes
sql = "SELECT Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas " & _
      "FROM Componentes WHERE ID_Proyecto = " & idProyecto & " ORDER BY Tipo_Componente"
Set objRS = objConn.Execute(sql)
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Plano de Armado - <%= Server.HTMLEncode(nombreProyecto) %></title>
    <style>
        /* CSS Minimalista pensado exclusivamente para hojas en blanco y negro */
        body { font-family: "Segoe UI", Arial, sans-serif; color: #000; background: #fff; padding: 20px; }
        h1 { font-size: 24px; border-bottom: 2px solid #000; padding-bottom: 5px; margin-bottom: 5px;}
        h3 { font-size: 16px; color: #333; margin-top: 0; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px; }
        th, td { border: 1px solid #000; padding: 10px; text-align: left; }
        th { background: #eee; text-transform: uppercase; }
        .pin { font-family: "Consolas", monospace; font-weight: bold; }
        
        /* Ocultar botones al momento de imprimir físicamente */
        @media print {
            .no-imprimir { display: none !important; }
        }
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
            <% Do While Not objRS.EOF %>
            <tr>
                <td><%= Server.HTMLEncode(objRS("Tipo_Componente")) %></td>
                <td><%= Server.HTMLEncode(objRS("Valor_Calculado")) %></td>
                <td class="pin"><%= Server.HTMLEncode(objRS("Pin_Conexion")) %></td>
                <td><%= Server.HTMLEncode(objRS("Ubicacion_Protoboard")) %></td>
            </tr>
            <% 
                objRS.MoveNext
            Loop 
            %>
        </tbody>
    </table>

    <p style="margin-top: 30px; font-size: 12px; font-style: italic; text-align: center;">
        Generado desde la Central de Monitoreo el <%= Date() %> a las <%= Time() %>
    </p>

    <script>
        window.onload = function() {
            window.print();
        }
    </script>
</body>
</html>
<%
objRS.Close
Set objRS = Nothing
CerrarConexion objConn
%>