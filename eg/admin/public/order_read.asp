<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%
dim id
id=trim(request.QueryString("id"))
if id="" or id=null then
response.Write("<script language=javascript>alert('没有该定单！');window.close;</script>")
response.End()
else
id=cint(id)
end if

sql="select * from o_s_orders where id="&cint(id)
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Conn,1,1
%>
<html>
<head>
<title>信息查看</title>
<style type="text/css">
<!--
.STYLE1 {
	font-size: 14pt;
	font-weight: bold;
}
.STYLE2 {font-size: 12px;}
.STYLE3 {font-size: 13px;}
-->
</style>
<link href="css/css.css" rel="stylesheet" type="text/css" />

<meta http-equiv="Content-Type" content="text/html; charset=utf-8"><style type="text/css">
<!--
body {
	background-color: #DDEEFF;
}
-->
</style></head>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30"><span class="STYLE1">信息唯一ID:<%=rs("id")%></span></td>
  </tr>
  <tr>
    <td><span class="STYLE3"><%=trim(rs("s_order"))%></span></td>
  </tr>
  <tr>
    <td align="right"><span class="STYLE2">[<a href=# onClick="javascript:history.back();">返回</a>]</span></td>
  </tr>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
set conn=nothing
%>

