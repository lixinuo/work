<!-- #Include File="inc/config.asp" -->
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<%
dim conn,connstr,dbimages
	dbimages = DbAccessConnection
	Set conn = Server.CreateObject("ADODB.connection")
	connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&dbimages&"")
	 if err then 
      err.clear
   else
        conn.open connstr
        if err then 
           err.clear
        end if
   end if
set rs=server.CreateObject("adodb.recordset")
rs.Open "select W_name from S_Main ",conn,1,1
webname=rs(0)
rs.close:set rs=nothing
%>
