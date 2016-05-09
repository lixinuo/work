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
rs.open "select * from o_kf",conn,1,3
rs.AddNew
rs("s_type")=int(request("s_type"))
rs("s_hm")=trim(request("s_hm"))
rs("s_name")=trim(request("s_name"))
rs("s_order")=trim(request("s_order"))
rs.Update
rs.Close
set rs=nothing
'//修改数据
case "edit"


set rs=server.CreateObject("adodb.recordset")
rs.open "select * from o_kf where id="&request.QueryString("id"),conn,1,3
rs("s_type")=int(request("s_type"))
rs("s_name")=trim(request("s_name"))
rs("s_hm")=trim(request("s_hm"))
rs("s_order")=trim(request("s_order"))
rs.update
rs.close
set rs=nothing
'//删除数据
case "del"

conn.execute ("delete from o_kf where id="&request.QueryString("id"))
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
<!--#include file="../inc/g_links.asp"-->
<style type="text/css">
<!--
.STYLE57 {
	color: #000000
}
-->
</style>
</head>
<body>
<%=Upload_Init()%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
  <tr>
    <td align="center"><%=s_typename%></td>
  </tr>
</table>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
  <tr>
    <td width="3%" height="30" align="center">ID</td>
    <td width="12%" align="center">类型</td>
    <td width="39%" align="center">名称</td>
    <td width="25%" align="center">帐号</td>
    <td width="4%" align="center">排序</td>
    <td width="17%" align="center">确定操作</td>
  </tr>
  <%

set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from o_kf order by s_order asc,id desc",conn,1,1
if rs.EOF and rs.BOF then
response.Write "<div align=center><font color=red>还没有"&s_typename&"</font></center>"
paixu=0
else
i=1
do while not rs.EOF
s_type=rs("s_type")
%>
  <form name="form<%=rs("id")%>" method="post" action="<%=base_url%>&action=edit&id=<%=rs("id")%>">
    <tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
      <td height="30" align="center"><%=rs("id")%></td>
      <td align="center"><select name="s_type">
			<!--<option value="">请选择</option>-->
			<option value="1"<%if s_type=1 then response.Write" selected"%>>QQ</option>
<!--			<option value="2"<%if s_type=2 then response.Write" selected"%>>MSN</option>
			<option value="3"<%if s_type=3 then response.Write" selected"%>>Skype</option>
			<option value="4"<%if s_type=4 then response.Write" selected"%>>阿里旺旺</option>
			<option value="5"<%if s_type=5 then response.Write" selected"%>>淘宝旺旺</option>
			<option value="6"<%if s_type=6 then response.Write" selected"%>>雅虎通</option>-->
			</select></td>
      <td align="center"><input name="s_name" type="text" id="s_name" size="25" value="<%=trim(rs("s_name"))%>"></td>
      <td align="center"><input name="s_hm" type="text" id="s_hm" size="25" value="<%=trim(rs("s_hm"))%>"></td>
      <td align="center"><input name="s_order" type="text" id="s_order2" style="width:30px;" value="<%=trim(rs("s_order"))%>"<%if show_edit=0 then w(" style='border:1px #ccc solid; background-color:#EEE; color:gray' readonly")%>></td>
      <td align="center"><input type="submit" name="Submit" class="inputkkys" value="修改">
          <%if show_del=1 then%>
          <input type="button" name="delme" value="删除" class="inputkkys" onClick="location='<%=base_url%>&action=del&id=<%=rs("id")%>'">
          <%end if%>
      </td>
    </tr>
  </form>
  <%rs.movenext
	i=i+1
        loop
        paixu=rs.RecordCount
        rs.close
        set rs=nothing
        end if
		%>
</table>
<%if show_add=1 then%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
  <tr>
    <td align="center">信息添加</td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
  <tr>
    <td width="14%" align="center">类型</td>
    <td width="40%" align="center">名称</td>
    <td width="27%" align="center">帐号</td>
    <td width="4%" align="center">排序</td>
    <td width="15%" align="center">确定操作
      </div></td>
  </tr>
  <form name="my_form" method="post" action="<%=base_url%>&action=add">
    <tr bgcolor="#FFFFFF">
      <td align="center"><select name="s_type">
			<!--<option value="">请选择</option>-->
			<option value="1">QQ</option>
<!--			<option value="2">MSN</option>
			<option value="3">Skype</option>
			<option value="4">阿里旺旺</option>
			<option value="5">淘宝旺旺</option>
			<option value="6">雅虎通</option>-->
			</select></td>
      <td align="center"><input name="s_name" type="text" id="s_name" size="28"></td>
      <td align="center"><input name="s_hm" type="text" id="s_hm" size="28"></td>
      <td align="center"><input name="s_order" type="text" id="s_order" style="width:30px;" value="<%=paixu+1%>"></td>
      <td align="center">&nbsp;
          <input type="submit" class="inputkkys" name="Submit2" value="添加"></td>
    </tr>
  </form>
</table>
<%end if%>
</body>
</html>
