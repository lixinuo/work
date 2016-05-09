<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%dim anclassid,anclass,paixu
s_type=isid(request("s_type"),0)
if s_type=0 then Response.Write "<script Language=Javascript>alert('ERRO!');history.go(-1);</script>"

s_typename=db_name("p_main",s_type) '产品名称
action=request("action")
show_del=1
show_add=1
show_edit=1

base_url="?s_type="&s_type&"&s_typename="&s_typename&"&show_edit="&show_edit&"&show_del="&show_del&"&show_add="&show_add&"&edit_name="&edit_name

'//////////////根据action来选择操作
select case action
'//增加数据
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from p_img",conn,1,3
rs.AddNew
rs("s_type")=int(request("s_type"))
rs("s_img")=trim(replace(request("s_img"&request.QueryString("id")),"../",""))
rs("s_order")=0
rs.Update
rs.Close
set rs=nothing
'//修改数据
case "edit"
d_imgurl=db_img("p_img",request.QueryString("id"))
''如果图片路径被更改，说明新上传啦图片，则先删除原来的图片文件
if d_imgurl<>trim(replace(request("s_img"&request.QueryString("id")),"../","")) then
d_imgurl="../../"&d_imgurl
deletefile d_imgurl
end if


set rs=server.CreateObject("adodb.recordset")
rs.open "select * from p_img where id="&request.QueryString("id"),conn,1,3
rs("s_img")=trim(replace(request("s_img"&request.QueryString("id")),"../",""))
rs("s_order")=trim(request("s_order"))
rs.update
rs.close
set rs=nothing
'//删除数据
case "del"

'删除图片文件
set rspic=db("select S_img from p_img where id="&request.QueryString("id"),2)
if not rspic.eof then
 delpic1="../../"&rspic(0)
 deletefile delpic1
end if
conn.execute ("delete from p_img where id="&request.QueryString("id"))
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
    <td align="center"><%=s_typename%>图片管理</td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
  <tr>
    <td width="3%" align="center">ID</td>
    <td width="17%" height="30" align="center">图片</td>
    <td width="52%" align="center">图片路径</td>
    <td width="5%" align="center">排序</td>
    <td width="23%" align="center">确定操作</td>
  </tr>
  <%
if s_type="" then
response.Write "<div align=center><font color=red>暂时没有选择分类</font></div>"
else
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from p_img where s_type="&s_type&" order by s_order asc,id desc",conn,1,1
if rs.EOF and rs.BOF then
response.Write "<div align=center><font color=red>还没有"&s_typename&"图片</font></center>"
paixu=0
else
i=1
do while not rs.EOF
%>
  <form name="form<%=rs("id")%>" method="post" action="<%=base_url%>&action=edit&id=<%=rs("id")%>">
    <tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
      <td height="110" align="center"><%=rs("id")%></td>
      <td align="center"><a href="../../<%=Rs("s_img")%>" target="_blank"><img src="../../<%=Rs("s_img")%>" width="100" height="100" border="0" /></a> </td>
      <td align="center"><%=Upload_Input("S_img"&rs("id"),trim(rs("s_img")))%></td>
      <td align="center"><input name="s_order" type="text" id="s_order" style="width:30px;" value="<%=trim(rs("s_order"))%>"<%if show_edit=0 then w(" style='border:1px #ccc solid; background-color:#EEE; color:gray' readonly")%>></td>
      <td align="center"><input type="submit" name="Submit" value="修改">
          <%if show_del=1 then%>
          <input type="button" name="delme" value="删除" onClick="location='<%=base_url%>&action=del&id=<%=rs("id")%>'">
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
        end if
		%>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
  <tr>
    <td align="center">图片添加</td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
  
  <form name="my_form" method="post" action="<%=base_url%>&action=add">
    <tr>
      <td width="13%" height="30" align="center">上传图片</td>
      <td width="62%" align="left"><%=Upload_Input("S_img","")%></td>
      <td width="25%" align="center">&nbsp;
          <input type="submit" name="Submit2" value="添加"></td>
    </tr>
  </form>
</table>
</body>
</html>
