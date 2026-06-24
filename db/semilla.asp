
%>
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, sql
Dim idNuevoProyecto

Set objConn = AbrirConexion()

' Evita duplicar la semilla si ya se cargó antes
sql = "SELECT COUNT(*) AS Total FROM Proyectos WHERE Nombre_Proyecto = 'Sensor de Fuerza Capacitivo'"
Set objRS = objConn.Execute(sql)

If objRS("Total") = 0 Then

    objRS.Close
    Set objRS = Nothing

    sql = "INSERT INTO Proyectos (Nombre_Proyecto, Plataforma_Simulacion, Microcontrolador, Descripcion, Fecha_Creacion) VALUES (" & _
          "'Sensor de Fuerza Capacitivo', 'Proteus', 'Arduino Uno', " & _
          "'Sensor capacitivo armado con dos placas de aluminio para medir fuerza aplicada.', " & _
          FechaSQL(Date()) & ")"
    objConn.Execute sql, , 129

    sql = "SELECT MAX(ID_Proyecto) AS UltimoID FROM Proyectos"
    Set objRS = objConn.Execute(sql)
    idNuevoProyecto = objRS("UltimoID")
    objRS.Close
    Set objRS = Nothing

    sql = "INSERT INTO Componentes (ID_Proyecto, Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas) VALUES (" & _
          idNuevoProyecto & ", 'Placa de aluminio (electrodo)', '2 placas de 5x5 cm, separadas 3 mm', 'A0', " & _
          "'Fuera de protoboard, cableado directo', 'Forma el capacitor variable; la separación cambia con la fuerza aplicada')"
    objConn.Execute sql, , 129

    sql = "INSERT INTO Componentes (ID_Proyecto, Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas) VALUES (" & _
          idNuevoProyecto & ", 'Resistencia pull-down', '1 MOhm', 'A0', " & _
          "'Fila 12, columna E', 'Necesaria para estabilizar la lectura analógica del electrodo')"
    objConn.Execute sql, , 129

    sql = "INSERT INTO Componentes (ID_Proyecto, Tipo_Componente, Valor_Calculado, Pin_Conexion, Ubicacion_Protoboard, Notas) VALUES (" & _
          idNuevoProyecto & ", 'Microcontrolador', 'Arduino Uno R3', 'A0 / 5V / GND', " & _
          "'N/A (fuera de protoboard)', 'Cerebro del circuito; lee el valor analógico equivalente a la capacitancia')"
    objConn.Execute sql, , 129

    sql = "INSERT INTO Bitacora_Fallos (ID_Proyecto, Fecha_Registro, Sintoma_Error, Solucion_Aplicada, Estado) VALUES (" & _
          idNuevoProyecto & ", " & FechaSQL(Date()) & ", " & _
          "'Lecturas inestables y saturadas en el monitor serie al aplicar presión leve sobre las placas.', " & _
          "'Se agregó una resistencia pull-down de 1 MOhm en A0 y se promediaron 10 lecturas por ciclo en el código para filtrar el ruido.', " & _
          "'Resuelto')"
    objConn.Execute sql, , 129

    sql = "INSERT INTO Bitacora_Fallos (ID_Proyecto, Fecha_Registro, Sintoma_Error, Solucion_Aplicada, Estado) VALUES (" & _
          idNuevoProyecto & ", " & FechaSQL(Date()) & ", " & _
          "'El valor simulado en Proteus no coincide con el valor físico medido en el Arduino real.', " & _
          "'Pendiente de recalibrar la curva de conversión analógico-fuerza con datos físicos reales.', " & _
          "'Pendiente')"
    objConn.Execute sql, , 129

Else
    objRS.Close
    Set objRS = Nothing
End If

CerrarConexion objConn
%>
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><title>Datos de ejemplo cargados</title>
<link rel="stylesheet" href="../includes/estilos.css"></head>
<body>
<main class="contenedor">
<div class="panel">
<h1>Datos de ejemplo</h1>
<p>Se cargó el proyecto de ejemplo <strong>"Sensor de Fuerza Capacitivo"</strong> con sus componentes
y fallos asociados (si no existía previamente).</p>
<div class="acciones">
    <a href="../index.asp" class="boton boton-primario">Ir al Dashboard</a>
</div>
</div>
</main>
</body>
</html>
