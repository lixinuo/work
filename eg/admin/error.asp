<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<html>
<head>
<title>网站管理系统——登陆错误页面</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style>
body {font-size:12px}
table {font-size:12px}
A:visited{TEXT-DECORATION: none;color: #000000}
A:active{TEXT-DECORATION: none}
A:hover{TEXT-DECORATION: underline;CURSOR: font-size: 12px; font-family: "宋体"; 
position: relative; left: 1px; top: 1px; clip: rect( );crosshair;Color:#000000;}
A:link{text-decoration: none; color: #000000}
.style1 {color: #666666}
.style2 {color: #FF9900}
</style>
<Script language="JavaScript">
if (top==self)
{
alert("对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。");
self.location.href="../index.asp";
}
</Script>
</head>
<body bgcolor="<%=Color_0%>" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" 
style="overflow-x: hidden; overflow-y: hidden; width: 100%;"
onLoad="javascript:window.err.focus();" 
onKeyDown="javascript:window.err.focus();">
<table width="72%" border="0" cellspacing="0" cellpadding="0" height="270" align="center" 
onMouseOver="javascript:window.err.focus();">
  <tr> 
    <td height="96" colspan="3">&nbsp;</td>
  </tr>
  <tr bgcolor="#999999"> 
    <td height="17" colspan="3"> 
      <div align="center"></div>
    </td>
  </tr>
  <tr> 
    <td height="57" width="12" bgcolor="#999999"> 
      <div align="center"></div>
    </td>
    <td height="57" bgcolor="<%=Color_0%>"><div align="center"><span class="style2">
	用户名/或密码错误/或你发呆的时间过长(超过了<%= CStr(Session.Timeout)%>分钟) 
	</span></div></td>
    <td height="57" width="12" bgcolor="#999999">&nbsp;</td>
  </tr>
  <tr bgcolor="#999999"> 
    <td height="14" colspan="3">&nbsp;</td>
  </tr>
  <tr> 
   <td height="57" width="12"> 
      <div align="center"></div>
    </td>
    <td align="center"> 
      <font size="2">
	  <a name="err" id="err" href="SyStem_login.asp"><span class="style1">重新登陆&gt;&gt;&gt;</span></a>
	  </font><br>
<!--
<input type="button" name="view" value="查看源码" onClick='window.location.href="view-source:" +window.location.href;' 
style="font-size:9pt;">
-->
    </td>
	<td height="57" width="12">&nbsp;</td>
  </tr>
</table>
</body>
</html>