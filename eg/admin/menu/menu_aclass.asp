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
if action="del" then
	strSQL="delete from "&table_name&" where id="&id&" "
	conn.execute strSQL
end if
%>
<html>
<head>
<title>栏目列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body >
<div class="aclass">
	<ul>
    	<li class="bold">栏目列表</li>
    	<li>
        	<span class="width_50">分类名称</span>
            <span class="width_25">分类排序</span>
            <span class="width_25">确定操作</span>
        </li>
		<%
        set rs=server.CreateObject("adodb.recordset")
        rs.Open "select * from s_menu_class where parent_id=0 order by S_order asc,id desc",conn,1,1
        if rs.EOF and rs.BOF then
			response.Write "<div align=center><font color=red>还没有分类</font></center>"
			paixu=0
        else
			do while not rs.EOF
        %>
        <li>
        	<span class="width_50"><input name="s_name" type="text" size="30" value="<%=trim(rs("s_name"))%>"></span>
            <span class="width_25"><input name="s_order<%=rs(0)%>" type="text"  size="4" value="<%=int(rs("s_order"))%>" onChange="location='?id=<%=trim(rs("id"))%>&action=s_order&s_order=' + this.value"></span>
            <span class="width_25"><input type="submit" name="Submit" class="inputkkys" value="修 改"> | <%
			 if rs("S_ok") then
				response.Write("<a title=取消显示 href=?action=s_ok&s_ok=0&id="&rs(0)&" onclick=""return confirm('取消显示');""><font color=blue>显示</font></a>")
			 else
				response.Write("<a title=取消隐藏' href=?action=s_ok&s_ok=1&id="&rs(0)&" onclick=""return confirm('取消隐藏');""><font color=red>隐藏</font></a>")
			 end if
			%>
            	<a href="?id=<%=int(rs(0))%>&action=del" onClick="return confirm('您确定要删除该分类吗？')"><font color=red>删除</font></a>
       		</span>
        </li>
        <%
			rs.MoveNext
			loop
			paixu=rs.RecordCount
		end if
		rs.close
		set rs=nothing
		%>
        <li></li>
        <li class="bold">栏目添加</li>
        <li>
        <form name="form1" method="post" action="?action=add">
        <span class="width_50"><input name="s_name" type="text" id="s_name" size="30"></span>
        <span class="width_25"><input name="S_order" type="text" id="S_order" size="4" value="999"></span>
        <span class="width_25"><input type="submit" name="Submit3" class="inputkkys" value="添 加"></span> 
        </form>
        </li>
    </ul>
</div>
</body>
</html>
