<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
table_name="s_menu_class"
id=trim(request("id"))
action=request("action")

if action="s_order" then
	strSQL="update "&table_name&" set s_order="&request("s_order")&" Where id=" &id& ""
	conn.execute strSQL
end if
if action="edit" then
	strSQL="update "&table_name&" set s_name='"&request("s_name")&"' Where id=" &id& ""
	conn.execute strSQL
end if
if action="s_ok" then
	strSQL="update "&table_name&" set s_ok="&request("s_ok")&" Where id=" &id& ""
	conn.execute strSQL
end if
if action="add" then
	strSQL="insert into "&table_name&" (s_name,s_order) values ('"&request("s_name")&"',"&request("s_order")&")"
	conn.execute strSQL
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body >
<table width="100%"   border="0" cellpadding="5" cellspacing="0">
 
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
  <td align="center"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
    <tr>
      <td align="center">栏目列表</td>
    </tr>
  </table>
    <table width="98%" border="0" cellpadding="0" cellspacing="0" class="lmlb">
    
    <tr >
     <td width="40%" height="30" align="center">分类名称</td>
     <td width="20%" align="center">分类排序</td>
     <td width="20%" align="center">确定操作</td>
    </tr>
    <%
		  set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from s_menu_class where parent_id=0 order by S_order asc,id desc",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>还没有分类</font></center>"
		  paixu=0
		  else
		  do while not rs.EOF
		  %>
    <form name="form1" method="post" action="?action=edit&id=<%=int(rs(0))%>">
     <tr align="center" bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
      <td height="30"><input name="s_name" type="text" size="30" value="<%=trim(rs("s_name"))%>">      </td>
      <td><input name="s_order<%=rs(0)%>" type="text"  size="4" value="<%=int(rs("s_order"))%>" onChange="location='?id=<%=trim(rs("id"))%>&action=s_order&s_order=' + this.value">      </td>
      <td><input type="submit" name="Submit" class="inputkkys" value="修 改"> | <%
		 if rs("S_ok") then
		    response.Write("<a title=取消显示 href=?action=s_ok&s_ok=0&id="&rs(0)&" onclick=""return confirm('取消显示');""><font color=blue>显示</font></a>")
		 else
		    response.Write("<a title=取消隐藏' href=?action=s_ok&s_ok=1&id="&rs(0)&" onclick=""return confirm('取消隐藏');""><font color=red>隐藏</font></a>")
		 end if
%>
       <a href=".asp?id=<%=int(rs(0))%>&action=del" onClick="return confirm('您确定要删除该分类吗？')"><font color=red>
       <!--删除-->
       </font></a> </td>
     </tr>
    </form>
    <%
		rs.MoveNext
		loop
		paixu=rs.RecordCount
		end if
		%>
   </table></td>
 </tr>
 <tr>
  <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
    <tr>
      <td align="center">栏目添加</td>
    </tr>
  </table>
    <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="lmlb">
    
    <tr  align="center">
     <td width="40%" height="30" align="center"> 分类名称</td>
     <td width="20%" align="center"> 分类排序</td>
     <td width="20%" align="center"> 确定操作</td>
    </tr>
    <form name="form1" method="post" action="?action=add">
     <tr  align="center">
      <td height="30" bgcolor="<%=Color_0%>"><input name="s_name" type="text" id="s_name" size="30">      </td>
      <td bgcolor="<%=Color_0%>"><input name="S_order" type="text" id="S_order" size="4" value="<%=paixu+1%>">      </td>
      <td bgcolor="<%=Color_0%>"><input type="submit" name="Submit3" class="inputkkys" value="添 加">      </td>
     </tr>
    </form>
   </table></td>
 </tr>
</table>
<%rs.close
set rs=nothing
closeconn%>
</body>
</html>
