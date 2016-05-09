<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#include file="../inc/data.asp"-->
<!--#INCLUDE FILE="../inc/gFunc_Page.asp" -->
<html><head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body>
<%'开始分页
skey=kill_sql(request("skey")):s_type=kill_sql(request("s_type"))
sqlstr="select * from [U_Main] where id<>0"
if s_type<>"" then sqlstr=sqlstr&" and "&s_type&" like '%"&skey&"%'"
sqlstr=sqlstr&" order by S_order asc,id desc"
set rs=server.CreateObject("adodb.recordset")
rs.open sqlstr,conn,1,1
if rs.eof And rs.bof then
Response.Write "<p align='center' class='contents'> 数据库中无数据！</p><br>"
else
int_RPP=20 '设置每页显示数目
int_showNumberLink_=6 '数字导航显示数目
showMorePageGo_Type_ = 0 '是下拉菜单还是输入值跳转，当多次调用时只能选1
str_nonLinkColor_="#000000" '非热链接颜色
toF_="首页"   			'首页 
toP10_=""			'上十 
toP1_=" 上一页"			'上一
toN1_=" 下一页"			'下一
toN10_=""			'下十
toL_="尾页"				'尾页
rs.PageSize=int_RPP
cPageNo=Request.QueryString("Page")
If cPageNo="" or not isnumeric(cPageNo) Then cPageNo = 1
cPageNo = Clng(cPageNo)
If cPageNo<1 Then cPageNo=1
If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
rs.AbsolutePage=cPageNo
 
count=0 
i=1
%>
			
<table width="100%"   border="0" cellpadding="5" cellspacing="0" bgcolor="<%=Color_0%>">
  <tr>
    <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
    <li>用户列表
    </span></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#004A80"></td>
  </tr>
  <tr>
    <td valign="top"> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="tableBorder">
<tr> 

<td height="100" valign="top">

<table width="100%" border="0" cellpadding="3" cellspacing="1" >
<tr  align="left">
  <td colspan="8" align="left" bgcolor="<%=Color_2%>">
	<form action="?action=search" method="post" name="search" class="style1">
	 <strong>
   请填写关键字
   <input name="skey" type="text" value="关键字">
 请选择搜索方式
 <select name="s_type" id="s_type">
  <option value="S_name" <%if stype="S_name" then w("selected")%>>客户编码</option>
  <option value="S_Quesion" <%if stype="S_Quesion" then w("selected")%>>联系人</option>
  <option value="S_tel" <%if stype="S_tel" then w("selected")%>>联系电话</option>
 </select>
 <input name="gosearch" type="submit" value=" 搜索 ">
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <input name="adduser" type="button" value="添加用户" onClick="location='user_do.asp';">
   </strong>
	</form>    </td>
  </tr>
<tr  align="center" bgcolor="#3399CC">
<td width="12%"> 客户编号 </td>
<td width="15%">联系人</td>
<td width="13%">联系电话</td>
<td width="19%"> 注册时间</td>
<td width="17%">最后登录</td>
<td width="9%"> 登陆次数</td>
<td width="7%">修改资料</td>
<td width="8%"> 选 择</td>
</tr>
<form name="form1" method="post" action="user_save.asp?action=del">
<%do while not (rs.eof or rs.bof) and count<rs.PageSize

if i mod 2=0 then 
mm_color="#C2DBFF"
else
mm_color="#F1F5FA"
end if

%>
<tr align="center">
<td bgcolor="<%=mm_color%>" style="PADDING-LEFT: 10px"><a href=user_do.asp?id=<%=rs(0)%>><%=trim(rs("s_name"))%></a></td>
<td bgcolor="<%=mm_color%>" style="PADDING-LEFT: 10px"><%=trim(rs("S_realname"))%></td>
<td bgcolor="<%=mm_color%>" style="PADDING-LEFT: 10px"><%=trim(rs("s_tel"))%></td>
<td bgcolor="<%=mm_color%>" style="PADDING-LEFT: 10px"><%=rs("S_addtime")%></td>
<td bgcolor="<%=mm_color%>" style="PADDING-LEFT: 10px"><%=rs("S_lasttime")%></td>
<td align="center" bgcolor="<%=mm_color%>"><%=rs("S_logins")%> 次</td>
<td align="center" bgcolor="<%=mm_color%>"><a href=user_do.asp?id=<%=rs(0)%>><img src="../images/edit.gif" border="0"></a></td>
<td align="center" bgcolor="<%=mm_color%>">
<input name="userid" type="checkbox" id="userid" value="<%=rs(0)%>"></td>
</tr>
<%
rs.movenext
i=i+1
count=count+1
loop
%>
<tr align="center" bgcolor="#3399CC">
  <td colspan="8" style="PADDING-LEFT: 10px">
 <%response.Write( fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
          <input type="checkbox" name="checkbox" value="Check All" onClick="mm()">
          删除全选
          <input type="submit" name="Submit" value="删 除" onClick="return test();">
          &nbsp;</td>
  </tr></form>
</table>

</td>

</tr>
</table>
<%end if%>
</td></tr>
  <tr>
    <td height="2" bgcolor="#004A80"></td>
  </tr>
<tr><td>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF" class="tableBorder">
</table>
</td></tr></table>
</body>
</html>
<script>
function test()
{
  if(!confirm('确认删除吗？')) return false;
}
</script>
<script language=javascript>
function mm()
{
   var a = document.getElementsByTagName("input");
   for (var i=1; i<a.length; i++)
			a[i].checked = a[a.length-2].checked;
}
</script>
