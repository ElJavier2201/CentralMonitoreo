<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Write "Server.MapPath(""/"") = " & Server.MapPath("/") & "<br>"
Response.Write "Server.MapPath(""/db/CentralMonitoreo.accdb"") = " & Server.MapPath("/db/CentralMonitoreo.accdb") & "<br>"
Response.Write "APPL_PHYSICAL_PATH = " & Request.ServerVariables("APPL_PHYSICAL_PATH") & "<br>"
%>