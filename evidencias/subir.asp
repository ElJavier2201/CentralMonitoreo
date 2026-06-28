<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Subir evidencia - Central de Monitoreo"
seccionActiva = "evidencias"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Function BinaryToText(ByVal binData)
    Dim stm
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write binData
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "ISO-8859-1"
    BinaryToText = stm.ReadText
    stm.Close
    Set stm = Nothing
End Function

Sub SaveTextAsBinaryFile(ByVal textData, ByVal filePath)
    Dim stm
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "ISO-8859-1"
    stm.Open
    stm.WriteText textData
    stm.Position = 0
    stm.Type = 1
    stm.SaveToFile filePath, 2
    stm.Close
    Set stm = Nothing
End Sub

Function ObtenerEntre(ByVal texto, ByVal inicio, ByVal fin)
    Dim p1, p2
    p1 = InStr(1, texto, inicio, vbTextCompare)
    If p1 = 0 Then
        ObtenerEntre = ""
        Exit Function
    End If
    p1 = p1 + Len(inicio)
    p2 = InStr(p1, texto, fin, vbTextCompare)
    If p2 = 0 Then
        ObtenerEntre = ""
    Else
        ObtenerEntre = Mid(texto, p1, p2 - p1)
    End If
End Function

Function NombreArchivoSeguro(ByVal nombre)
    Dim n, i, ch, limpio
    n = Replace(nombre, "\", "/")
    If InStrRev(n, "/") > 0 Then n = Mid(n, InStrRev(n, "/") + 1)
    limpio = ""
    For i = 1 To Len(n)
        ch = Mid(n, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "." Or ch = "_" Or ch = "-" Then
            limpio = limpio & ch
        ElseIf ch = " " Then
            limpio = limpio & "_"
        End If
    Next
    If limpio = "" Then limpio = "evidencia"
    NombreArchivoSeguro = limpio
End Function

Function ExtensionPermitida(ByVal nombre)
    Dim ext
    If InStrRev(nombre, ".") = 0 Then
        ExtensionPermitida = False
        Exit Function
    End If
    ext = LCase(Mid(nombre, InStrRev(nombre, ".") + 1))
    ExtensionPermitida = (ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "gif" Or ext = "webp")
End Function

Sub AsegurarCarpeta(ByVal ruta)
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(ruta) Then fso.CreateFolder(ruta)
    Set fso = Nothing
End Sub

Sub ParseMultipart(ByVal cuerpo, ByVal boundary, ByRef campos, ByRef nombreArchivo, ByRef contenidoArchivo)
    Dim partes, i, parte, posSep, headers, contenido, nombreCampo, archivoOriginal
    partes = Split(cuerpo, boundary)
    For i = 0 To UBound(partes)
        parte = partes(i)
        posSep = InStr(1, parte, vbCrLf & vbCrLf, vbBinaryCompare)
        If posSep > 0 Then
            headers = Left(parte, posSep - 1)
            contenido = Mid(parte, posSep + 4)
            If Right(contenido, 2) = vbCrLf Then contenido = Left(contenido, Len(contenido) - 2)

            nombreCampo = ObtenerEntre(headers, "name=""", """)
            archivoOriginal = ObtenerEntre(headers, "filename=""", """)

            If nombreCampo <> "" Then
                If archivoOriginal <> "" Then
                    nombreArchivo = NombreArchivoSeguro(archivoOriginal)
                    contenidoArchivo = contenido
                Else
                    campos(nombreCampo) = contenido
                End If
            End If
        End If
    Next
End Sub

Function ValorCampo(ByVal campos, ByVal nombre)
    If campos.Exists(nombre) Then
        ValorCampo = Limpiar(campos(nombre))
    Else
        ValorCampo = ""
    End If
End Function

Dim objConn, objRS, cmd, sql
Dim idProyectoSel, tipoEvidencia, descripcion, errores, okMsg
Dim contentType, boundary, cuerpo, campos, bytes, nombreArchivo, contenidoArchivo
Dim nombreFinal, rutaRelativa, rutaFisica, carpetaUploads

errores = ""
okMsg = ""
idProyectoSel = 0
tipoEvidencia = "Foto del protoboard"
descripcion = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Request.TotalBytes > (5 * 1024 * 1024) Then
        errores = errores & "<li>El archivo no debe superar 5 MB.</li>"
    Else
        contentType = Request.ServerVariables("CONTENT_TYPE")
        If InStr(1, contentType, "multipart/form-data", vbTextCompare) = 0 Then
            errores = errores & "<li>El formulario no se recibió como multipart/form-data.</li>"
        Else
            boundary = "--" & Mid(contentType, InStr(contentType, "boundary=") + 9)
            bytes = Request.BinaryRead(Request.TotalBytes)
            cuerpo = BinaryToText(bytes)
            Set campos = Server.CreateObject("Scripting.Dictionary")
            nombreArchivo = ""
            contenidoArchivo = ""
            Call ParseMultipart(cuerpo, boundary, campos, nombreArchivo, contenidoArchivo)

            idProyectoSel = IDValido(ValorCampo(campos, "id_proyecto"))
            tipoEvidencia = ValorCampo(campos, "tipo_evidencia")
            descripcion = ValorCampo(campos, "descripcion")

            If idProyectoSel = 0 Then errores = errores & "<li>Selecciona un proyecto.</li>"
            If tipoEvidencia = "" Then tipoEvidencia = "Evidencia"
            If nombreArchivo = "" Then errores = errores & "<li>Selecciona una imagen.</li>"
            If nombreArchivo <> "" And Not ExtensionPermitida(nombreArchivo) Then errores = errores & "<li>Solo se permiten imágenes JPG, JPEG, PNG, GIF o WEBP.</li>"

            If errores = "" Then
                Randomize
                nombreFinal = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & _
                              Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & "_" & _
                              CStr(Int(Rnd() * 100000)) & "_" & nombreArchivo
                carpetaUploads = Server.MapPath("/uploads/evidencias")
                Call AsegurarCarpeta(carpetaUploads)
                rutaFisica = Server.MapPath("/uploads/evidencias/" & nombreFinal)
                rutaRelativa = "uploads/evidencias/" & nombreFinal
                Call SaveTextAsBinaryFile(contenidoArchivo, rutaFisica)

                Set objConn = AbrirConexion()
                Set cmd = Server.CreateObject("ADODB.Command")
                cmd.ActiveConnection = objConn
                cmd.CommandType = 1
                cmd.CommandText = "INSERT INTO Evidencias (ID_Proyecto, Tipo_Evidencia, Ruta_Archivo, Descripcion, Fecha_Subida) VALUES (?, ?, ?, ?, ?)"
                cmd.Parameters.Append cmd.CreateParameter("@idp", 3, 1, , idProyectoSel)
                cmd.Parameters.Append cmd.CreateParameter("@tipo", 200, 1, 60, tipoEvidencia)
                cmd.Parameters.Append cmd.CreateParameter("@ruta", 200, 1, 255, rutaRelativa)
                If descripcion = "" Then
                    cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 255, Null)
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@desc", 200, 1, 255, descripcion)
                End If
                cmd.Parameters.Append cmd.CreateParameter("@fecha", 135, 1, , Date())
                cmd.Execute , , 128
                Set cmd = Nothing
                CerrarConexion objConn
                okMsg = "Imagen subida correctamente. Ruta guardada: " & rutaRelativa
            End If
        End If
    End If
End If

Set objConn = AbrirConexion()
Set objRS = objConn.Execute("SELECT ID_Proyecto, Nombre_Proyecto FROM Proyectos ORDER BY Nombre_Proyecto")
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <h1>Subir imagen de evidencia</h1>
    <p class="ayuda">Sube una foto del protoboard, captura de Proteus/Tinkercad, monitor serial o resultado físico.</p>

    <% If errores <> "" Then %>
        <div class="alerta alerta-error"><strong>Revisa:</strong><ul><%= errores %></ul></div>
    <% End If %>
    <% If okMsg <> "" Then %>
        <div class="alerta alerta-ok"><%= Server.HTMLEncode(okMsg) %></div>
    <% End If %>

    <div class="caja-ayuda">
        <strong>Importante:</strong> la carpeta <span class="codigo">uploads/evidencias</span> debe tener permiso de escritura para IIS/IUSR. Tamaño máximo: 5 MB.
    </div>

    <% If objRS.EOF Then %>
        <div class="alerta alerta-error">Primero crea un proyecto antes de subir evidencias.</div>
    <% Else %>
    <form method="post" action="subir.asp" enctype="multipart/form-data">
        <label for="id_proyecto">Proyecto *</label>
        <select id="id_proyecto" name="id_proyecto">
            <option value="">-- Seleccione --</option>
            <% Do While Not objRS.EOF %>
                <option value="<%= objRS("ID_Proyecto") %>" <% If CLng(objRS("ID_Proyecto")) = CLng(idProyectoSel) Then %>selected<% End If %>><%= Server.HTMLEncode(objRS("Nombre_Proyecto")) %></option>
            <% objRS.MoveNext : Loop %>
        </select>

        <label for="tipo_evidencia">Tipo de evidencia</label>
        <select id="tipo_evidencia" name="tipo_evidencia">
            <option value="Foto del protoboard" <% If tipoEvidencia = "Foto del protoboard" Then %>selected<% End If %>>Foto del protoboard</option>
            <option value="Captura de simulación" <% If tipoEvidencia = "Captura de simulación" Then %>selected<% End If %>>Captura de simulación</option>
            <option value="Monitor serial" <% If tipoEvidencia = "Monitor serial" Then %>selected<% End If %>>Monitor serial</option>
            <option value="Medición con multímetro" <% If tipoEvidencia = "Medición con multímetro" Then %>selected<% End If %>>Medición con multímetro</option>
            <option value="Resultado físico" <% If tipoEvidencia = "Resultado físico" Then %>selected<% End If %>>Resultado físico</option>
        </select>

        <label for="archivo">Imagen *</label>
        <input type="file" id="archivo" name="archivo" accept="image/*">

        <label for="descripcion">Descripción</label>
        <textarea id="descripcion" name="descripcion" rows="3"><%= Server.HTMLEncode(descripcion) %></textarea>

        <div class="acciones">
            <button type="submit" class="boton boton-primario">Subir evidencia</button>
            <a href="listar.asp" class="boton boton-secundario">Volver a evidencias</a>
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
