<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
set rs=server.createobject("adodb.recordset")
rs.open "select * from o_gbook where id="&request("id"),conn,3,2
if request("action")<>"save" then
%>
<html>
<head>
<title>回复</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="images/cssyullhao.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	background-color: #DDEEFF;
}
body,td,th {
	font-size: 12px;
}
-->
</style></head>
<body>
<form action="guestbook_reply.asp?action=save&id=<%=request("id")%>" method="post" name="form" id="form">
  <table width="88%" border="1" align="center" cellpadding="0" cellspacing="2" bordercolor="#CCCCCC" class=log_table>
    <tr> 
      <td align="center"> <table  width="100%" border="0" class=log_titlewidth="100%">
          <tr> 
            <td>&nbsp;&nbsp;&nbsp;留言内容：</td>
          </tr>
          <tr>
            <td> 
              <%
Function unHtml(content)
	ON ERROR RESUME NEXT
	unHtml=content
	IF content <> "" Then
		unHtml=Server.HTMLEncode(unHtml)
		unHtml=Replace(unHtml,vbcrlf,"<br>")
		unHtml=Replace(unHtml,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
		unHtml=Replace(unHtml," ","&nbsp;")
	End IF
	IF Err.Number <>0 Then
		unHtml= "HTML转换中出错请联系管理员<br>"
		Err.Clear
	End IF
End Function
%>
              &nbsp;&nbsp;&nbsp;<%=unHtml(rs("s_content"))%></td>
          </tr>
          <tr> 
            <td><strong>&nbsp;&nbsp;&nbsp;管 理 员 回 复 </strong></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="65%">&nbsp;&nbsp;&nbsp;回复<font color="#990033"><%=rs("s_name")%></font>的留言</td>
            <td rowspan="2"><div align="right"></div></td>
          </tr>
          <tr>
            <td>&nbsp;&nbsp;&nbsp;不超过255个字符！</td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td valign="middle">&nbsp;&nbsp;&nbsp;回复内容： <br>
        &nbsp;&nbsp;&nbsp;
        <textarea name="reply" cols="58" rows="8" class="wenbenkuang" id="reply"></textarea>
      </td>
    </tr>
    <tr> 
      <td height="30"> 
       &nbsp;&nbsp;&nbsp; <input type="submit" class="go-wenbenkuang" value="回复" name="button"> 
        &nbsp;&nbsp;<input name="Submit2" type="reset" class="go-wenbenkuang" id="Submit2" value="重置"> 
      </td>
    </tr>
  </table>
</form>
<%
else
if request.form("reply")="" then
response.write"<SCRIPT language=JavaScript>alert('对不起，公告内容不能为空！');"
response.write"javascript:history.go(-1)</SCRIPT>"
else
rs("s_reply")=request.form("reply")
rs.update
rs.close
set rs=nothing
response.redirect "Guestbook_manage.asp"
end if
end if
%>
</body>
</html>
