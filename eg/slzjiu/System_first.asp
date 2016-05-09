<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<%
if request.Cookies(Cookies_name)("Grade")="" then 
	response.redirect "error.asp"
	response.End()
else
%>
<!--#INCLUDE FILE="data.asp" -->
<%
Sub SetEmpty()
request.Cookies(Cookies_name)("Name")=""
request.Cookies(Cookies_name)("Grade")=""
request.Cookies(Cookies_name)("Pwd")=""
End Sub

s_type=Trim(request.QueryString("type"))

Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select top 1 * from S_main order by id desc"
rs.open strSql,Conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站后台管理系统</title>
<link rel="stylesheet" type="text/css" href="images/cssyullhao.css">
</head>

<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" style="overflow-x: hidden; overflow-y: hidden; width: 98%;" bgcolor="#F6FBFF">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

<table width="98%"  border="0">
  <tr>
    <td><img src="images/welcome.gif" width="215" height="30"></td>
  </tr>
  <tr>
    <td align="center" width="90%">
<IFRAME ID="Welcome" src="inc/aspcheckObj.asp" frameborder="0" scrolling="no" width="98%" height="280">浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</IFRAME>
	</td>
  </tr>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
call closeconn
end if
%>