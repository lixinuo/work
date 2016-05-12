<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%dim anclassid,anclass,paixu
s_typename=request.QueryString("s_typename")
action=request("action")
show_del=request("show_del")
show_add=request("show_add")
show_edit=request("show_edit")
if not isnumeric(show_del) then show_del=0
if not isnumeric(show_add) then show_add=0
if not isnumeric(show_edit) then show_edit=1
base_url="?s_typename="&s_typename&"&show_edit="&show_edit&"&show_del="&show_del&"&show_add="&show_add&"&edit_name="&edit_name

'//////////////根据action来选择操作
select case action
'//增加数据
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from add_contact",conn,1,3
rs.AddNew
	rs("s_type")=trim(request("s_type"))
	rs("s_hm")=trim(request("s_hm"))
	rs("s_name")=trim(request("s_name"))
	rs("s_order")=trim(request("s_order"))
rs.Update
rs.Close
set rs=nothing

'//修改数据
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from add_contact where id="&request.QueryString("id"),conn,1,3
	rs("s_type")=trim(request("s_type"))
	rs("s_name")=trim(request("s_name"))
	rs("s_hm")=trim(request("s_hm"))
	rs("s_order")=trim(request("s_order"))
rs.update
rs.close
set rs=nothing

'//删除数据
case "del"
conn.execute ("delete from add_contact where id="&request.QueryString("id"))
end select
%>
<html>
<head>
<title>添加联系方式</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body>
<div class="aclass">
	<ul>
    	<li class="bold">浮动客服</li>
        <li>
            <span class="width_10">类型</span>
            <span class="width_30">名称</span>
            <span class="width_30">账号</span>
            <span class="width_10">排序</span>
            <span class="width_10">确定操作</span>
        </li>
		<%
		set rs=server.CreateObject("adodb.recordset")
		rs.Open "select * from add_contact order by s_order asc,id desc",conn,1,1
		if rs.EOF and rs.BOF then
			response.Write "<li align=center><font color=red>还没有"&s_typename&"</font></li>"
			paixu=0
		else
			i=1
			do while not rs.EOF
				s_type=rs("s_type")
        %>  	
    	<form name="form<%=rs("id")%>" method="post" action="<%=base_url%>&action=edit&id=<%=rs("id")%>">
        <li>
            <span class="width_10">
            	<select name="s_type">
                    <option value="QQ"<%if s_type="QQ" then response.Write" selected"%>>QQ</option>
                    <option value="MSN"<%if s_type="MSN" then response.Write" selected"%>>MSN</option>
                    <option value="Skype"<%if s_type="Skype" then response.Write" selected"%>>Skype</option>
                    <option value="阿里旺旺"<%if s_type="阿里旺旺" then response.Write" selected"%>>阿里旺旺</option>
                    <option value="淘宝旺旺"<%if s_type="淘宝旺旺" then response.Write" selected"%>>淘宝旺旺</option>
                    <option value="雅虎通"<%if s_type="雅虎通" then response.Write" selected"%>>雅虎通</option>
                </select>
            </span>
            <span class="width_30"><input name="s_name" type="text" id="s_name" value="<%=trim(rs("s_name"))%>"></span>
            <span class="width_30"><input name="s_hm" type="text" id="s_hm" value="<%=trim(rs("s_hm"))%>"></span>
            <span class="width_10"><input name="s_order" type="text" id="s_order" style="width:30px;" value="<%=trim(rs("s_order"))%>"></span>
            <span class="width_10">
            	<input type="submit" name="Submit" class="inputkkys" value="修改">
			<%if show_del=1 then%>
                <input type="button" name="delme" value="删除" class="inputkkys" onClick="location='<%=base_url%>&action=del&id=<%=rs("id")%>'">
			<%end if%>
            </span>
        </li>
        </form>
		<%
        rs.movenext
        i=i+1
        loop
        paixu=rs.RecordCount
        rs.close
        set rs=nothing
        end if
        %>
        <li></li>
        <li class="bold">信息添加</li>
        <%if show_add=1 then%>
        <li>
        <form name="my_form" method="post" action="<%=base_url%>&action=add">
			<span class="width_10">
            	<select name="s_type" id="s_type">
                    <option value="">请选择</option>
                    <option value="QQ">QQ</option>
                    <option value="MSN">MSN</option>
                    <option value="Skype">Skype</option>
                    <option value="阿里旺旺">阿里旺旺</option>
                    <option value="淘宝旺旺">淘宝旺旺</option>
                    <option value="雅虎通">雅虎通</option>
                </select>
            </span>
            <span class="width_30"><input name="s_name" type="text" id="s_name"></span>
            <span class="width_30"><input name="s_hm" type="text" id="s_hm"></span>
            <span class="width_10"><input name="s_order" type="text" id="s_order" style="width:30px;" value="<%=paixu+1%>"></span>
            <span class="width_10"><input type="submit" class="inputkkys" name="Submit2" value="添加"></span>
        </form>
        </li>
        <%end if%>
    </ul>
</div>
</body>
</html>
