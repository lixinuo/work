<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%dim anclassid,anclass,paixu
s_type=isid(request.QueryString("s_type"),0)
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
rs.open "select * from P_color",conn,1,3
rs.AddNew
rs("s_name")=trim(request("s_name"))
rs("s_order")=trim(request("s_order"))
rs.Update
rs.Close
set rs=nothing
'//修改数据
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from P_color where id="&request.QueryString("id"),conn,1,3
rs("s_name")=trim(request("s_name"))
rs("s_order")=trim(request("s_order"))
rs.update
rs.close
set rs=nothing
'//删除数据
case "del"
conn.execute ("delete from P_color where id="&request.QueryString("id"))
end select
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
 background-color: <%=Color_0%>;
}
.style1 {
	color: #000000;
	font-weight: bold;
	font-size: 14px;
}
.style55 {
	color: #666666
}
body, td, th {
	color: #666666;
}
.style56 {
	color: #FF0000
}
-->
</style>
<link href="../images/cssyullhao.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE57 {
	color: #000000
}
-->
</style>
</head>
<body>
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#DDEEFF">
 <tr>
  <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
   <li><%=s_typename%></li>
   </span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
  <td width="100%" bgcolor="#FFFFFF"  valign="top"><table border="0" align="center" cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
     <td align="center" bgcolor="#DDEEFF">ID</td>
     <td align="center" bgcolor="#DDEEFF">名称</td>
     <td width="100" align="center" bgcolor="#DDEEFF">排序</td>
     <td width="100" align="center" bgcolor="#DDEEFF">确定操作</td>
    </tr>
     <%
		if s_type="" Then
		response.Write "<div align=center><font color=red>暂时没有选择分类</font></div>"
		else
		set rs=server.CreateObject("adodb.recordset")
		rs.Open "select * from P_color where parent_id="&s_type&" order by s_order asc,id desc",conn,1,1
		if rs.EOF and rs.BOF then
		response.Write "<div align=center><font color=red>还没有"&s_typename&"</font></center>"
		paixu=0
		Else
		do while not rs.EOF
     %>
    <form name="form<%=rs("id")%>" method="post" action="<%=base_url%>&action=edit&id=<%=rs("id")%>">
     <tr>
      <td align="center" bgcolor="#DDEEFF"><%=RS(0)%></td>
      <td align="center" bgcolor="#DDEEFF"><input name="s_name" type="text" id="s_name" size="40" value="<%=trim(rs("s_name"))%>" <%if show_edit=0 then w(" style='border:1px #ccc solid; background-color:#EEE; color:gray' readonly")%>></td>
      <td align="center" bgcolor="#DDEEFF"><input name="s_order" type="text" id="s_order" size="4" value="<%=trim(rs("s_order"))%>"<%if show_edit=0 then w(" style='border:1px #ccc solid; background-color:#EEE; color:gray' readonly")%>></td>
      <td align="center" bgcolor="#DDEEFF"><input type="submit" name="Submit" value="修改">
			<%if show_del=1 then%>
      <input type="button" name="delme" value="删除" onClick="location='<%=base_url%>&action=del&id=<%=rs("id")%>'">
			<%end if%>			</td>
     </tr>
    </form>
    <%rs.movenext
        loop
        paixu=rs.RecordCount
        rs.close
        set rs=nothing
        end if
        end if
		%>
   </table></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
</table>
<%if show_add=1 then%>
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#DDEEFF">
 <tr>
  <td width="100%" bgcolor="<%=Color_1%>"><table border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
     <td align="center" bgcolor="#DDEEFF">名称</td>
     <td width="100" align="center" bgcolor="#DDEEFF">排序</td>
     <td width="100" align="center" bgcolor="#DDEEFF">确定操作
      </div></td>
    </tr>
    <form name="my_form" method="post" action="<%=base_url%>&action=add">
     <tr>
      <td align="center" bgcolor="#DDEEFF"><input name="s_name" type="text" id="s_name" size="40" value="<%=s_typename&paixu+1%>"></td>
      <td align="center" bgcolor="#DDEEFF"><input name="s_order" type="text" id="s_order" size="4" value="<%=paixu+1%>"></td>
      <td align="center" bgcolor="#DDEEFF">&nbsp;
       <input type="submit" name="Submit2" value="添加"></td></tr>
    </form>
   </table></td>
 </tr>
</table>
<%end if%>
</body>
</html>
