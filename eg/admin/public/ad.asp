<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%dim anclassid,anclass,paixu
s_type=request.QueryString("s_type")
s_typename=request.QueryString("s_typename")
action=request("action")
show_del=request("show_del")
show_add=request("show_add")
show_edit=request("show_edit")
if not isnumeric(show_del) then show_del=0
if not isnumeric(show_add) then show_add=0
if not isnumeric(show_edit) then show_edit=1
base_url="?s_type="&s_type&"&s_typename="&s_typename&"&show_edit="&show_edit&"&show_del="&show_del&"&show_add="&show_add&"&edit_name="&edit_name

'//////////////根据action来选择操作
select case action

'//增加数据
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from add_pic",conn,1,3
rs.AddNew
rs("s_type")=int(request("s_type"))
rs("s_img")=trim(replace(request("s_img"&request.QueryString("id")),"../",""))
'rs("s_img1")=trim(replace(request("s_img1"&request.QueryString("id")),"../",""))
rs("s_url")=trim(request("s_url"))
rs("s_name")=trim(request("s_name"))
rs("s_order")=trim(request("s_order"))
rs.Update
rs.Close
set rs=nothing

'//修改数据
case "edit"
d_imgurl=db_img("add_pic",request.QueryString("id"))
''如果图片路径被更改，说明新上传啦图片，则先删除原来的图片文件
if d_imgurl<>trim(replace(request("s_img"&request.QueryString("id")),"../","")) then
	d_imgurl="../../"&d_imgurl
	deletefile d_imgurl
end if

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from add_pic where id="&request.QueryString("id"),conn,1,3
rs("s_img")=trim(replace(request("s_img"&request.QueryString("id")),"../",""))
'rs("s_img1")=trim(replace(request("s_img1"&request.QueryString("id")),"../",""))
rs("s_url")=trim(request("s_url"))
rs("s_name")=trim(request("s_name"))
rs("s_order")=trim(request("s_order"))
rs.update
rs.close
set rs=nothing

'//删除数据
case "del"
'删除图片文件
set rspic=db("select S_img,S_img1 from add_pic where id="&request.QueryString("id"),2)
if not rspic.eof then
	delpic1="../../"&rspic(0)
	deletefile delpic1
end if
conn.execute ("delete from add_pic where id="&request.QueryString("id"))
end select

'文件操作函数
function deletefile(filedir)
'on error resume next
dim fso
set fso = Server.CreateObject("Scripting.FileSystemObject")
if (fso.fileexists(SM(filedir))) then fso.deletefile(SM(filedir))
set fso = Nothing
if err.number<>0 then err.clear
end Function
%>
<html>
<head>
<title>添加图片</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body>
<%=Upload_Init()%>
<div class="aclass">
	<ul>
    	<li class="bold">首页图片</li>
        <li>
            <span class="width_10">说明</span>
            <span class="width_30">图片路径（630*430）</span>
            <span class="width_30">链接url</span>
            <span class="width_10">排序</span>
            <span class="width_10">确定操作</span>
        </li>
		<%
        set rs=server.CreateObject("adodb.recordset")
        rs.Open "select * from add_pic where s_type="&s_type&" order by s_order asc,id desc",conn,1,1
        if rs.EOF and rs.BOF then
			response.Write "<li align=center><font color=red>还没有"&s_typename&"</font></li>"
			paixu=0
        else
			i=1
			do while not rs.EOF
        %>  	
    	<form name="form<%=rs("id")%>" method="post" action="<%=base_url%>&action=edit&id=<%=rs("id")%>">
        <li>
            <span class="width_10"><input name="s_name" type="text" id="s_name" style="width:100px;" value="<%=trim(rs("s_name"))%>"></span>
            <span class="width_30"><%=Upload_Input("S_img"&rs("id"),trim(rs("s_img")))%></span>
            <span class="width_30"><input name="s_url" type="text" id="s_url" value="<%=trim(rs("s_url"))%>"></span>
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
			<span class="width_10"><input name="s_name" type="text" id="s_name" style="width:100px;" value="<%=s_typename&paixu+1%>"></span>
            <span class="width_30"><%=Upload_Input("S_img","")%></span>
            <span class="width_30"><input name="s_url" type="text" id="s_url"></span>
            <span class="width_10"><input name="s_order" type="text" id="s_order" style="width:30px;" value="<%=paixu+1%>"></span>
            <span class="width_10"><input type="submit" class="inputkkys" name="Submit2" value="添加"></span>
        </form>
        </li>
        <%end if%>
    </ul>
</div>
</body>
</html>
