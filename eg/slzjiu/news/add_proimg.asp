<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%dim anclassid,anclass,paixu
s_type=isid(request("s_type"),0)
if s_type=0 then Response.Write "<script Language=Javascript>alert('ERRO!');history.go(-1);</script>"

s_typename=db_name("a_main",s_type) '产品名称
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
rs("s_name")=request("s_name")
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
<%=Upload_Init()%>
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#DDEEFF">
 <tr>
  <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
   <li><%=s_typename%>图片管理</li>
   </span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
  <td width="100%" bgcolor="<%=Color_0%>"  valign="top">
  <table width="882" border="0" align="center" cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
     <td width="20" align="center" bgcolor="#3399CC">ID</td>
     <td width="164" align="center" bgcolor="#3399CC">图片</td>
     <td width="308" align="center" bgcolor="#3399CC">图片说明</td>
     <td width="149" align="center" bgcolor="#3399CC">图片路径</td>
     <td width="94" align="center" bgcolor="#3399CC">排序</td>
     <td width="121" align="center" bgcolor="#3399CC">确定操作</td>
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
     <tr>
      <td align="center" bgcolor="<%=mm_color%>"><%=rs("id")%></td>
      <td align="center" bgcolor="<%=mm_color%>">
      
      <a href="../../<%=Rs("s_img")%>" target="_blank"><img src="../../<%=Rs("s_img")%>" width="100" height="100" border="0" /></a>      </td>
      <td align="center" bgcolor="<%=mm_color%>"><input name="s_name" type="text" id="s_name" value="<%=rs("s_name")%>" size="50"></td>
      <td align="center" bgcolor="<%=mm_color%>"><%=Upload_Input("S_img"&rs("id"),trim(rs("s_img")))%></td>
      <td align="center" bgcolor="<%=mm_color%>"><input name="s_order" type="text" id="s_order" style="width:30px;" value="<%=trim(rs("s_order"))%>"<%if show_edit=0 then w(" style='border:1px #ccc solid; background-color:#EEE; color:gray' readonly")%>></td>
      <td align="center" bgcolor="<%=mm_color%>">
       <input type="submit" name="Submit" value="修改">
       <%if show_del=1 then%>
       <input type="button" name="delme" value="删除" onClick="location='<%=base_url%>&action=del&id=<%=rs("id")%>'">
       <%end if%>      </td>
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
   </table></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
</table>

<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#DDEEFF">
 <tr>
  <td width="100%" height="80" bgcolor="<%=Color_1%>"><table width="882" height="74" border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
     <td colspan="3" align="center" bgcolor="#CC9999">说明/图片添加</div></td>
     </tr>
    <form name="my_form" method="post" action="<%=base_url%>&action=add">
     <tr>
      <td width="91" align="center" bgcolor="#DDEEFF">上传图片</td>
      <td width="429" align="left" bgcolor="#DDEEFF"><input name="s_name" type="text" id="s_name" size="50">
        <br>

      <%=Upload_Input("S_img","")%></td>
      <td width="163" align="center" bgcolor="#DDEEFF">&nbsp;<input type="submit" name="Submit2" value="添加"></td>
      </tr>
    </form>
   </table></td>
 </tr>
</table>

</body>
</html>
