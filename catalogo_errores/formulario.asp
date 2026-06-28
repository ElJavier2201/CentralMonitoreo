<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CharSet = "UTF-8"
Dim rutaBase, tituloPagina, seccionActiva
rutaBase = "../"
tituloPagina = "Error frecuente - Central de Monitoreo"
seccionActiva = "catalogo_errores"
%>
<!--#include virtual="/includes/auth.asp"-->
<!--#include virtual="/conexion.asp"-->
<%
Dim objConn, objRS, cmd, sql
Dim idError, modoEdicion, nombreError, categoria, causa, solucion, riesgo, errores

idError = IDValido(Request.QueryString("id"))
modoEdicion = (idError > 0)
nombreError = ""
categoria = "Circuito"
causa = ""
solucion = ""
riesgo = "Medio"
errores = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    idError = IDValido(Request.Form("id_error"))
    modoEdicion = (idError > 0)
    nombreError = Limpiar(Request.Form("nombre_error"))
    categoria = Limpiar(Request.Form("categoria"))
    causa = Limpiar(Request.Form("causa_probable"))
    solucion = Limpiar(Request.Form("solucion_recomendada"))
    riesgo = Limpiar(Request.Form("nivel_riesgo"))

    If nombreError = "" Then errores = errores & "<li>El nombre del error es obligatorio.</li>"
    If categoria = "" Then categoria = "General"
    If riesgo <> "Bajo" And riesgo <> "Medio" And riesgo <> "Alto" Then riesgo = "Medio"

    If errores = "" Then
        Set objConn = AbrirConexion()
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = objConn
        cmd.CommandType = 1

        If modoEdicion Then
            cmd.CommandText = "UPDATE Catalogo_Errores SET Nombre_Error=?, Categoria=?, Causa_Probable=?, Solucion_Recomendada=?, Nivel_Riesgo=? WHERE ID_Error=?"
            cmd.Parameters.Append cmd.CreateParameter("@nombre", 200, 1, 120, nombreError)
            cmd.Parameters.Append cmd.CreateParameter("@cat", 200, 1, 80, categoria)
            cmd.Parameters.Append cmd.CreateParameter("@causa", 200, 1, 255, causa)
            cmd.Parameters.Append cmd.CreateParameter("@sol", 200, 1, 255, solucion)
            cmd.Parameters.Append cmd.CreateParameter("@riesgo", 200, 1, 30, riesgo)
            cmd.Parameters.Append cmd.CreateParameter("@id", 3, 1, , idError)
        Else
            cmd.CommandText = "INSERT INTO Catalogo_Errores (Nombre_Error, Categoria, Causa_Probable, Solucion_Recomendada, Nivel_Riesgo) VALUES (?, ?, ?, ?, ?)"
            cmd.Parameters.Append cmd.CreateParameter("@nombre", 200, 1, 120, nombreError)
            cmd.Parameters.Append cmd.CreateParameter("@cat", 200, 1, 80, categoria)
            cmd.Parameters.Append cmd.CreateParameter("@causa", 200, 1, 255, causa)
            cmd.Parameters.Append cmd.CreateParameter("@sol", 200, 1, 255, solucion)
            cmd.Parameters.Append cmd.CreateParameter("@riesgo", 200, 1, 30, riesgo)
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
    sql = "SELECT ID_Error, Nombre_Error, Categoria, Causa_Probable, Solucion_Recomendada, Nivel_Riesgo FROM Catalogo_Errores WHERE ID_Error = " & idError
    Set objRS = objConn.Execute(sql)
    If Not objRS.EOF Then
        nombreError = Limpiar(objRS("Nombre_Error"))
        categoria = Limpiar(objRS("Categoria"))
        causa = Limpiar(objRS("Causa_Probable"))
        solucion = Limpiar(objRS("Solucion_Recomendada"))
        riesgo = Limpiar(objRS("Nivel_Riesgo"))
    Else
        modoEdicion = False
        idError = 0
    End If
    objRS.Close
    Set objRS = Nothing
    CerrarConexion objConn
End If
%>
<!--#include virtual="/includes/header.asp"-->

<div class="panel">
    <h1><% If modoEdicion Then %>Editar error frecuente<% Else %>Nuevo error frecuente<% End If %></h1>
    <p class="ayuda">Registra problemas comunes del armado físico y cómo resolverlos.</p>

    <% If errores <> "" Then %><div class="alerta alerta-error"><strong>Revisa:</strong><ul><%= errores %></ul></div><% End If %>

    <form method="post" action="formulario.asp<% If modoEdicion Then %>?id=<%= idError %><% End If %>">
        <input type="hidden" name="id_error" value="<%= idError %>">

        <label for="nombre_error">Nombre del error *</label>
        <input type="text" id="nombre_error" name="nombre_error" maxlength="120" value="<%= Server.HTMLEncode(nombreError) %>" placeholder="Ej: Lectura analógica inestable">

        <label for="categoria">Categoría</label>
        <input type="text" id="categoria" name="categoria" maxlength="80" value="<%= Server.HTMLEncode(categoria) %>" placeholder="Ej: Arduino, alimentación, sensor, protoboard">

        <label for="causa_probable">Causa probable</label>
        <textarea id="causa_probable" name="causa_probable" rows="3" maxlength="255"><%= Server.HTMLEncode(causa) %></textarea>

        <label for="solucion_recomendada">Solución recomendada</label>
        <textarea id="solucion_recomendada" name="solucion_recomendada" rows="3" maxlength="255"><%= Server.HTMLEncode(solucion) %></textarea>

        <label for="nivel_riesgo">Nivel de riesgo</label>
        <select id="nivel_riesgo" name="nivel_riesgo">
            <option value="Bajo" <% If riesgo = "Bajo" Then %>selected<% End If %>>Bajo</option>
            <option value="Medio" <% If riesgo = "Medio" Then %>selected<% End If %>>Medio</option>
            <option value="Alto" <% If riesgo = "Alto" Then %>selected<% End If %>>Alto</option>
        </select>

        <div class="acciones">
            <button type="submit" class="boton boton-primario">Guardar error</button>
            <a href="listar.asp" class="boton boton-secundario">Cancelar</a>
        </div>
    </form>
</div>

<!--#include virtual="/includes/footer.asp"-->
