<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
s_lan=request("s_lan")
action=request("action")
if s_lan="" then w("请选择语言") else s_lan=cint(s_lan)
if request("action")="edit" then
  id=request("id")
	m_class=trim(request("m_class"))
	m_url=trim(request("m_url"))
	m_order=trim(request("m_order"))
	conn.execute("update menu set m_class='"&m_class&"',m_url='"&m_url&"',m_order="&m_order&" Where id=" & id & "")
end if
if request("action")="add" then
  s_lan=trim(request("s_lan"))
	m_class=trim(request("m_class"))
	m_url=trim(request("m_url"))
	m_order=trim(request("m_order"))
	conn.execute("Insert into menu(s_lan,m_class,m_url,m_order) values ("&s_lan&",'"&m_class&"','"&m_url&"',"&m_order&")")
end if
if request("action")="del" then
	id=request("id")
	conn.execute("Delete from menu where id="&id)
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../images/cssyullhao.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
.STYLE1 {
	color: #000000;
	font-weight: bold;
}
-->
</style>
</head>
<body text="#000000" >
<table width="100%"   border="0" cellpadding="5" cellspacing="0">
 <tr>
  <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
   <li></li>
   导航管理</span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
  <td align="center">
   <table width="100%">
    <tr >
     <td height="30" colspan="4" align="center" bgcolor="<%=Color_0%>" class="STYLE1">导航列表</td>
    </tr>
    <tr >
     <td width="30%" align="center" bgcolor="<%=Color_0%>">栏目导航</td>
     <td width="30%" align="center" bgcolor="<%=Color_0%>">导航链接</td>
     <td width="20%" align="center" bgcolor="<%=Color_0%>">显示排序</td>
     <td width="20%" align="center" bgcolor="<%=Color_0%>">确定操作</td>
    </tr>
    <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from menu where s_lan="&s_lan&" order by m_order asc,id desc",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>还没有栏目导航</font></center>"
		  paixu=0
		  else
		  do while not rs.EOF
		  %>
    <form name="form1" method="post" action="menu.asp?action=edit&id=<%=int(rs("id"))%>&s_lan=<%=s_lan%>">
     <tr  align="center">
      <td bgcolor="<%=Color_0%>">
       <input name="m_class" type="text" id="m_class" size="30" value="<%=trim(rs("m_class"))%>">
      </td>
      <td bgcolor="<%=Color_0%>">
       <input name="m_url" type="text" id="m_url" size="30" value="<%=trim(rs("m_url"))%>">
      </td>
      <td bgcolor="<%=Color_0%>">
       <input name="m_order" type="text" id="m_order" size="4" value="<%=int(rs("m_order"))%>">
      </td>
      <td bgcolor="<%=Color_0%>">
       <input type="submit" name="Submit" value="修 改">
       &nbsp;
       <%
		 if rs("Isshow") then
		    response.Write("<a title=取消显示 href=?s_lan="&s_lan&"&isshow=0&id="&rs("id")&" onclick=""return confirm('取消显示');""><font color=blue>显示</font></a>")
		 else
		    response.Write("<a title=设置为隐藏' href=?s_lan="&s_lan&"&isshow=1&id="&rs("id")&" onclick=""return confirm('是否设置为隐藏');""><font color=red>隐藏</font></a>")
		 end if
%>
       <a href="menu.asp?s_lan=<%=s_lan%>&id=<%=int(rs("id"))%>&action=del" onClick="return confirm('您确定要删除该分类吗？')"><font color=red>
       <!--删除-->
       </font></a> </td>
     </tr>
    </form>
<%rs.MoveNext
loop
paixu=rs.RecordCount
end if%>
   </table>
  </td>
 </tr>
 <tr>
  <td>
   <table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
     <td height="30" colspan="4" align="center" background="../images/admin_bg_1.gif" bgcolor="<%=Color_0%>"><span class="STYLE1">栏目添加</span></td>
    </tr>
    <tr  align="center">
     <td width="30%" align="center" bgcolor="<%=Color_0%>"> 分类名称</td>
     <td width="30%" align="center" bgcolor="<%=Color_0%>">导航链接</td>
     <td width="20%" align="center" bgcolor="<%=Color_0%>"> 分类排序</td>
     <td width="20%" align="center" bgcolor="<%=Color_0%>"> 确定操作</td>
    </tr>
    <form name="form1" method="post" action="menu.asp?action=add">
    <input type="hidden" name="s_lan" value="<%=s_lan%>">
     <tr  align="center">
      <td bgcolor="<%=Color_0%>">
       <input name="m_class" type="text" id="m_class" size="30">
      </td>
      <td bgcolor="<%=Color_0%>">
       <input name="m_url" type="text" id="m_url" size="30">
      </td>
      <td bgcolor="<%=Color_0%>">
       <input name="m_order" type="text" id="m_order" size="4" value="<%=paixu+1%>">
      </td>
      <td bgcolor="<%=Color_0%>">
       <input type="submit" name="Submit3" value="添 加">
      </td>
     </tr>
    </form>
   </table>
  </td>
 </tr>
</table>
<%rs.close
set rs=nothing
dim id,Isshow,Pa
id=Request.QueryString("id")
Isshow=Request.QueryString("Isshow")
if Isshow<>empty then
  conn.execute("update menu set isshow=0 where isshow=1 and id="&id&"")
  conn.execute("update menu set Isshow="&Isshow&" where id="&id&"")
  response.redirect "?"
end if
%>
</body>
</html>
