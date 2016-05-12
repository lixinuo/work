<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<!--#INCLUDE FILE="../inc/gFunc_Page.asp" -->
<%

id=Isid(request("id"),0)
if request("action")="orderby" then
  strIDs=request("strIDs")
	orderby=request("orderby")
	strSQL="update P_class set s_order='"&orderby&"' Where id=" & strIDs & ""
	conn.execute strSQL
end if
if request("action")="del" then
  selectbookid=request("selectbookid")
	if selectbookid<>"" then
	strSQL="Delete From P_class Where id in (" & selectbookid & ")"
	conn.execute strSQL
	end if
end if
show_depth=request("show_depth"):show_add=request("show_add"):show_next=request("show_next"):show_del=request("show_del")
if show_depth="" then
if session("show_depth")="" then session("show_depth")=0
else
session("show_depth")=show_depth
end if

if show_add="" then
if session("show_add")="" then session("show_add")=1
else
session("show_add")=show_add
end if

if show_next="" then
if session("show_next")="" then session("show_next")=1
else
session("show_next")=show_next
end if

if show_del="" then
if session("show_del")="" then session("show_del")=1
else
session("show_del")=show_del
end if

s_pai=trim(request("s_pai"))
if s_pai="" then
s_pai=0
else
s_pai=cint(s_pai)
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<!--#include file="../inc/g_links.asp"-->
</head><body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
      <tr>
        <td class="lmtjdt">分类列表 <%if id=0 then%><%else%>*当前父类信息：<strong style="color:#F00;"><%=db_name("p_class",id)%><%end if%></strong></td>
      </tr>
</table>
<table width="98%"   border="0" align="center" cellpadding="5" cellspacing="0">
 
 <tr>
  <td valign="top" bgcolor="<%=Color_0%>"><%
page_url="class_list.asp"


strSQL ="select * from P_class where s_pai="&s_pai
if id<>0 then
strSQL =strSQL&" and parent_id="&id
else
strSQL =strSQL&" and parent_id=0"
end if
strSQL =strSQL&" order by s_order asc,id desc"


 Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,1,1
if rs.eof then
response.Redirect("class_do.asp?parent_id="&id&"&s_pai="&s_pai)
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
   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="lmtj">
    <tr valign="middle">
     <td height="30" colspan="6"><table width="90%" border="0" align="right" cellpadding="0" cellspacing="0">
       <tr>
        <td><%if session("show_add")=1 then%>
         <input type="button" name="addnew" class="inputkkys" onClick="location='class_do.asp?parent_id=<%=id%>&s_pai=<%=s_pai%>';" value="添加本级分类">
         <%end if%>
         <input type="button" name="backlist" class="inputkkys" onClick="location='class_list.asp?s_pai=<%=s_pai%>&show_add=<%=show_add%>&show_next=<%=show_next%>&show_del=<%=show_del%>';" value="返回">        </td>
        <td></td>
       </tr>
      </table></td>
    </tr>
    <form name="form1" method="post" action="?action=del&id=<%=id%>&s_pai=<%=s_pai%>&show_add=<%=show_add%>&show_next=<%=show_next%>&show_del=<%=show_del%>">
    <tr>
     <td width="14%" height="30" align="left">&nbsp;分类ID号</td>
     <td width="36%" height="30" align="left">分类标题</td>
     <%if session("show_next")=1 and rs("class_depth")<session("show_depth")+0  then%>
     <td width="18%" height="30" align="center">添加下级分类</td>
     <%end if%>
     <td width="10%" height="30" align="center">修改</td>
     <td width="10%" height="30" align="center">排序</td>
     <%if session("show_del")=1 then%>
     <td width="14%" height="30" align="center">删除</td>
     <%end if%>
    </tr>
    <%
do while not (rs.eof or rs.bof) and count<int_RPP 


if i mod 2=0 then 
mm_color="#C2DBFF"
else
mm_color="#F1F5FA"
end if
%>
    <tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
     <td height="30">&nbsp;<font color=red><b><%= rs(0)%></b></font></td>
     <td height="30" ><%=rs("S_name")%></td>
     <%	if session("show_next")=1 and rs("class_depth")<session("show_depth")+0 then%>
     <td width="18%" height="30"><%response.Write("<a href='class_list.asp?id="&rs("id")&"&s_pai="&s_pai&"'>进入下级分类("&db_s("select count(id) from P_class  where parent_id="&rs(0))&")</a>")%>     </td>
     <%end if%>
     <td height="30" ><div align="center"> <A HREF="class_do.asp?id=<%=rs(0)%>&parent_id=<%=id%>&s_pai=<%=s_pai%>"><img src="../images/detail_off.gif" border=0></A></div></td>
     <td height="30" ><input name="orderby<%=trim(rs("id"))%>" type="text" id="orderby" size="4" maxlength="5" value="<%=trim(rs("s_order"))%>" onChange="location='?strIDs=<%=trim(rs("id"))%>&id=<%=id%>&action=orderby&s_pai=<%=s_pai%>&show_add=<%=show_add%>&show_next=<%=show_next%>&show_del=<%=show_del%>&orderby=' + this.value"></td>
     <%if session("show_del")=1 then%>
     <td height="30" align="center"><%
sttSQL="select * from P_main,P_class where (P_main.Classid="&rs("id")&" and P_main.s_pai="&s_pai&") or (P_class.parent_id="&rs("id")&" and P_class.s_pai="&s_pai&") "
 Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sttSQL,Conn,1,1
if rs1.recordcount=0 then
w "<input name=""selectbookid"" type=""checkbox"" id=""selectbookid"" value="""&rs("id")&""">"
else
response.Write("不能删除")
end if
rs1.close
set rs1=nothing
%>     </td>
     <%end if%>
    </tr>
    <%
rs.movenext 
count=count+1
i=i+1
loop
%>
    <tr>
     <td height="30" colspan="6" align="center"><%response.Write( fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
		 <% if session("show_del")=1 then%>
      <input type="checkbox" name="checkbox" value="Check All" onClick="mm()">
      全选
      <input type="submit" name="Submit" class="inputkkys" value="删除所选" onClick="return test();">
      <%
			 end if
			end if
			%>     </td>
    </tr>
</form>
   </table></td>
 </tr>
</table>
<%rs.close
set rs=nothing
closeconn%>
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