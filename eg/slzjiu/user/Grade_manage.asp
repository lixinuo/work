<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/gFunc_Page.asp"-->
<%
page=request("page")
if request("action")="edit" then
  strIDs=request("strIDs")

	typename=trim(request("Gradename"))
	conn.execute("update s_user_Grade set Gradename='"&Gradename&"' Where id=" & strIDs & "")
end if
if request("action")="del" then
  Gradeid=request("Gradeid")
    if Gradeid="" then msg "没有选择删除项！","Grade_manage.asp"
	conn.execute("delete from s_user_Grade Where id in(" & Gradeid & ")")
end if
if request("action")="add" then
	addname=request("addname")
	conn.execute("Insert into s_user_Grade(Gradename) values ('"&addname&"')")
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
  <tr>
    <td align="center">权限管理</td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
<form name="form2" method="post" action="?action=add">
      <tr>
        <td class="lmtjdt">增加权限&nbsp;<input name="addname" type="text" id="addname" class="inputkkys" size="20">&nbsp;<input type="submit" name="button" class="inputkkys" id="button" value="提交"></td>
      </tr>
  </form>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj" >
  <%
sqlstr="select id,GradeName from [s_user_Grade] where id<>0"
sqlstr=sqlstr&" order by id desc"
set rs=server.CreateObject("adodb.recordset")

rs.open sqlstr,conn,1,1
if rs.eof then
w("<tr><td colspan=4>没有权限！</td></tr>")
else
%>
  <tr  align="center">
    <td width="25%" height="30">权限ID</td>
    <td width="15%" height="30" >权限名称</td>
    <td width="15%" height="30" >权限人数</td>
    <td width="15%" height="30" >修改权限</td>
    <td width="10%" height="30" > 选 择</td>
  </tr>
  <form name="form1" method="post" action="?action=del">
    <%
    int_RPP=20 '设置每页显示数目
		int_showNumberLink_=5 '数字导航显示数目
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
do while not (rs.eof or rs.bof) and count<rs.PageSize

if i mod 2=0 then 
mm_color="#C2DBFF"
else
mm_color="#F1F5FA"
end if
%>
    <tr align="center" bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
      <td height="30"  ><%=rs(0)%></td>
      <td height="30"  ><%=rs(1)%></td>
      <td height="30" align="center" style="color:#FF0000" ><%=db_s("select count(id) from s_user where Gradeid="&rs("id"))%></td>
      <td height="30" align="center" ><a href="grade_edit.asp?id=<%=rs(0)%>"><span class="ff_links"><img src="../images/detail_off.gif" border=0></span></a> </td>
      <td height="30" align="center" ><%
		 
		 
		 r_shum=db_s("select count(id) from s_user where Gradeid="&rs("id"))
		 
		 if r_shum>0 then
		 w "不能删除"
		 else
		 %>
          <input name="Gradeid" type="checkbox" id="Gradeid" value="<%=rs("id")%>">
          <%end if%>
      </td>
    </tr>
    <%
rs.movenext
i=i+1
count=count+1
loop
end if
%>
    <tr align="">
      <td height="30" colspan="5" align="center"  style="PADDING-LEFT: 10px"><%response.Write( fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="submit" name="Submit" class="inputkkys" value="删除" onClick="javascript:this.form.submit();">
        <input type="checkbox" name="checkbox2" value="Check All" onClick="mm()">
      全选 </td>
    </tr>
  </form>
</table>
</body>
</html>
<script language=javascript>
function mm()
{
   var a = document.getElementsByTagName("input");
   for (var i=0; i<a.length; i++)
			a[i].checked = a[a.length-1].checked;
}
</script>
