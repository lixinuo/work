<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
set rs=server.createobject("adodb.recordset")
conn.execute "delete from o_gbook where id="&trim(request.querystring("id"))
response.redirect "Guestbook_manage.asp"
%>
