<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%
table_name="s_menu_class"
id=trim(request("id"))
o_id=isid(request.QueryString("o_id"),0)
action=request("action")

id_a=db_a("select id from "&table_name&" where parent_id="&o_id&" order by s_order asc,id desc")
  pre_id=get_content(get_positon(id)-1)
  next_id=get_content(get_positon(id)+1)

if action="s_order" then
	strSQL="update "&table_name&" set s_order="&request("s_order")&" Where id=" &id& ""
	conn.execute strSQL
end if

if action="s_order_up" then

	if pre_id<>"" then pre_id_order=db_f(table_name,"s_order",pre_id)
	if next_id<>"" then next_id_order=db_f(table_name,"s_order",next_id)
	now_id_order=db_f(table_name,"s_order",id)

	strSQL="update "&table_name&" set s_order="&now_id_order&" Where id=" &pre_id& ""
	conn.execute strSQL
	strSQL="update "&table_name&" set s_order="&pre_id_order&" Where id=" &id& ""
	conn.execute strSQL
end if

if action="s_order_down" then

	if pre_id<>"" then pre_id_order=db_f(table_name,"s_order",pre_id)
	if next_id<>"" then next_id_order=db_f(table_name,"s_order",next_id)
	now_id_order=db_f(table_name,"s_order",id)

	strSQL="update "&table_name&" set s_order="&now_id_order&" Where id=" &next_id& ""
	conn.execute strSQL
	strSQL="update "&table_name&" set s_order="&next_id_order&" Where id=" &id& ""
	conn.execute strSQL
end if


if action="edit" then
	strSQL="update "&table_name&" set s_name='"&request("s_name")&"',s_url='"&request("s_url")&"' Where id=" &id& ""
	conn.execute strSQL
end if
if action="s_ok" then
	strSQL="update "&table_name&" set s_ok="&request("s_ok")&" Where id=" &id& ""
	conn.execute strSQL
end if
if action="add" then
	strSQL="insert into "&table_name&" (s_name,s_url,s_order,parent_id) values ('"&request("s_name")&"','"&request("s_url")&"',"&request("s_order")&","&o_id&")"
	conn.execute strSQL
end if
if action="move" then
conn.execute "update "&table_name&" set parent_id="&o_id&" where id="&id
end if
%>
<html>
<head>
<title>菜单管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body>
<div class="aclass">
	<ul>
    	<li class="bold">菜单管理</li>
    	<li style="text-align:left;">
        	<select name="select" class="chose_item" onChange="location='?o_id='+this.options[this.selectedIndex].value" >
                <option >选择栏目</option>
				<%call my_optionid(0,o_id,table_name)%>
            </select>
        </li>
        <li>
        	<span class="width_15">菜单名称</span>
            <span class="width_50">链接地址</span>
            <span class="width_15">菜单排序</span>
            <span class="width_20">确定操作</span>
        </li>
		<%
        if o_id=0 then
        else
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from "&table_name&" where parent_id="&o_id&" order by s_order asc,id desc",conn,1,1
			if rs.EOF and rs.BOF then
				response.Write "<div align=center><font color=red>还没有菜单</font></center>"
				paixu=0
			else
				formi=0
				do while not rs.EOF
        %>
        <form name="formlist<%=formi%>" method="post" action="?action=edit&id=<%=rs(0)%>&o_id=<%= o_id %>">
        <li>
        	<span class="width_15">
            	<input name="s_name" type="text" id="s_name" size="12" value="<%=trim(rs("s_name"))%>">
            </span>
            <span class="width_50">
            	<input name="s_url" type="text" id="s_url" size="45" value="<%=trim(rs("s_url"))%>">
            </span>
            <span class="width_15">
            	<%if Isid(db_F(table_name,"top 1 s_order","s_order<"&Rs("s_order")&" and s_ok=1 and parent_id<>0  order by s_order desc"),0)<>0 then%><a href="?action=s_order_up&id=<%=trim(rs("id"))%>&o_id=<%= o_id %>">上</a><%else%>顶<%end if%>
                <input name="s_order<%=rs(0)%>" type="text"  size="2" value="<%=int(rs("s_order"))%>" onChange="location='?action=s_order&id=<%=trim(rs("id"))%>&o_id=<%= o_id %>&s_order=' + this.value">
				<%if Isid(db_F(table_name,"top 1 s_order","s_order>"&Rs("s_order")&" and s_ok=1 and parent_id<>0  order by s_order asc"),0)<>0 then%><a href="?action=s_order_down&id=<%=trim(rs("id"))%>&o_id=<%= o_id %>">下</a><%else%>底<%end if%>
            </span>
            <span class="width_20">
          <input type="submit" class="inputkkys" name="Submit" value="修改">&nbsp;|&nbsp;<%
		 if rs("s_ok") then
		    response.Write("<a title=取消显示 href=?action=s_ok&s_ok=0&id="&rs(0)&"&O_id="&O_id&" ><font color=blue>显示</font></a>")
		 else
		    response.Write("<a title=取消隐藏' href=?action=s_ok&s_ok=1&id="&rs(0)&"&O_id="&O_id&" ><font color=red>隐藏</font></a>")
		 end if
%>&nbsp;|&nbsp;<a href="?id=<%=int(rs("id"))%>&action=del&o_id=<%=o_id%>" onClick="return confirm('您确定进行删除操作吗？')">删除</a>
            </span>
        </li>
        </form>
        <%
				rs.movenext:formi=formi+1
				loop
				paixu=rs.RecordCount
				rs.close
				set rs=nothing
			end if
        end if
		%>
        <li></li>
        <li class="bold">菜单添加</li>
        <li>
        <form name="form2" method="post" action="?action=add&o_id=<%= o_id %>">
        <span class="width_15"><input name="s_name" type="text" id="s_name" size="12"></span>
        <span class="width_50"><input name="s_url" type="text" id="s_url" size="45"></span>
        <span class="width_15"><input name="s_order" type="text" id="s_order" size="4" value="<%=paixu+1%>"></span> 
        <span class="width_20"><input type="submit" class="inputkkys" name="Submit2" value="添加菜单"></span>
        </form>
        </li>
    </ul>
</div>
</body>
</html>