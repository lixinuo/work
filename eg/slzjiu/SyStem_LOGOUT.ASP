<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="connection.asp" -->
<%'Session.Abandon()
Response.Cookies(Cookies_name)=""
			
if Request.QueryString("target")="index" then
Response.Redirect("System_first.asp")
else
Response.Clear()
Server.Transfer("SyStem_login.asp")
end if
%>