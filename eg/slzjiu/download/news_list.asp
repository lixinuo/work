<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<!--#INCLUDE FILE="../inc/gFunc_Page.asp" -->
<%
tablename="A_main"
s_pai=trim(request("s_pai"))
if s_pai="" then
s_pai=0
else
s_pai=cint(s_pai)
end if
show_img=trim(request("show_img"))
if show_img="" then
show_img=0
else
show_img=cint(show_img)
end if
show_hot=trim(request("show_hot"))
if show_hot="" then
show_hot=0
else
show_hot=cint(show_hot)
end if
action=trim(request("action"))
page=request("page")
url="news_list.asp?s_pai="&s_pai&"&page="&page&"&show_img="&show_img&"&show_hot="&show_hot

if action="del_all" then'//删除全部
	selectbookid=request("selectbookid")
	if selectbookid<>"" then
		conn.execute "delete * from "&tablename&" where id in ("&selectbookid&")"
		response.Redirect url
		response.End
	end if
end if
'//热点图片
S_ok=Request.QueryString("S_ok")
if S_ok<>"" then
   id=Request.QueryString("id")
  conn.execute("update "&tablename&" set S_ok="&S_ok&" where id="&id&"")
end if

S_hot=Request.QueryString("S_hot")
if S_hot<>"" then
   id=Request.QueryString("id")
  conn.execute("update "&tablename&" set S_hot="&S_hot&" where id="&id&"")
end if

if action="change_class" then'//删除全部
	selectbookid=request("selectbookid")
	change_classid=request("change_classid")
	if selectbookid<>"" and change_classid<>"" then
	  conn.execute "update "&tablename&" set classid="&change_classid&" where id in ("&selectbookid&")"
		response.Redirect url
		response.End
	end if
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<!--#include file="../inc/g_links.asp"-->
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
    <tr>
      <td align="center">信息列表</td>
    </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
<FORM action="<%=url%>" method="post" name="sele">
      <tr>
        <td class="lmtjdt">查询 信息标题：
        <input type="text" name="title" size="10" class=input>
       &nbsp; 
       文章类型：
       <select name="classid" size="1" class="input" style="width:120; ">
        <option value="">全部</option>
				<% call db_ChildID(0,classid,"A_class",s_pai) %>
       </select>
       &nbsp;
       <input type="submit" value="提交" name="B1" class=inputkkys>
			 <input type="button" value="添加信息" name="B2" class="inputkkys" onClick="location='news_do.asp?s_pai=<%=s_pai%>&show_img=<%=show_img%>'"> </td>
      </tr>
  </form>
</table>
<table width="98%"   border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
 
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
  <td valign="top" bgcolor="<%=Color_0%>">
   <%  

if request("action")="orderby" then
  strIDs=request("id")
	orderby=request("orderby")
	conn.execute("update "&tablename&" set s_order='"&orderby&"' Where id=" & strIDs & "")
end if

title=changechr(trim(request.Form("title")))
classid=request("classid")


PageSize = 12
strSQL ="select * from "&tablename&" where s_pai="&s_pai
if title<>"" then
	strSQL = strSQL & " and S_name like '%"&title&"%'"
end if
if classid<>"" then
	strSQL = strSQL & " and classid="&classid&""
end if
	strSQL = strSQL & " ORDER BY  s_order asc,id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,1,1

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
   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
    

    <form name="form1" method="post" action="<%=url%>&action=del_all">
    
    <tr>
     <td width="17%" height="30" align="center">
      文章类别</td>
     <td width="33%" height="30" align="center">
      信息标题</td>
     <td width="18%" height="30" align="center">
      创建日期</td>
     <td width="16%" height="30" align="center">
      修改</td>
     <td width="16%" height="30" align="center">
      选择|/排序</td>
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
     <td height="30" align="center"> <%=db_name("a_class",rs("classid"))%></td>
     <td height="30"><%if s_pai=1 then%><a href="add_proimg.asp?s_type=<%=Rs(0)%>"><%end if%><%=left(rs("S_name"),24)%><%if s_pai=1 then%></a><%end if%></td>
     <td height="30" align="center"><%=rs("s_time")%></td>
     <td width="16%" height="30" align="center">
      <%
		response.Write("<A title=修改内容 HREF='news_do.asp?id="&rs("id")&"&s_pai="&s_pai&"&show_img="&show_img&"'>编辑</A>&nbsp;")
		if show_hot=1 then
		 if rs("S_ok") then
		    response.Write("<a href="&url&"&action=s_ok&s_ok=0&id="&rs("id")&" onclick=""return confirm('是否取消推荐');""><font color=blue>推荐</font></a>&nbsp;")
		 else
		    response.Write("<a href="&url&"&action=s_ok&s_ok=1&id="&rs("id")&" onclick=""return confirm('是否设置为推荐');""><font color=black>推荐</font></a>&nbsp;")
		 end if
		end if
'		if s_pai=1 then
'			if rs("S_hot") then
'		    response.Write("<a href="&url&"&action=S_hot&S_hot=0&id="&rs("id")&" onclick=""return confirm('是否取消推荐');""><font color=blue>推荐</font></a>")
'		 else
'		    response.Write("<a href="&url&"&action=S_hot&S_hot=1&id="&rs("id")&" onclick=""return confirm('是否设置为推荐');""><font color=black>推荐</font></a>")
'		 end if
		
'		end if
		
		%>     </td>
     <td height="30" align="center" width="16%">
      <input name="selectbookid" type="checkbox" id="selectbookid" value="<%=rs("id")%>">
      <input name="orderby<%=trim(rs("id"))%>" type="text" id="orderby" size="4" maxlength="5" value="<%=trim(rs("s_order"))%>" onChange="location='<%=url%>&id=<%=rs(0)%>&action=orderby&orderby=' + this.value">     </td>
    </tr>
<%
rs.movenext 
count=count+1
i=i+1
loop
%>
    <tr>
     <td height="30" colspan="5" align="right">
      <%response.Write( fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
       <select name="change_classid" size="1" >
        <option value="">全部</option>
				<% call db_ChildID(0,classid,"A_class",s_pai) %>
       </select>
			<input type="submit" name="Submit" value="批量归类" class="inputkkys" onClick="this.form.action='<%=url%>&action=change_class';this.form.submit();">
			      <input type="submit" name="Submit" value="批量删除" class="inputkkys" onClick="return test();">
      <input type="checkbox" name="checkbox" value="Check All" onClick="mm()">全选/反选&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr></form>
   </table>
   <%end if%>  </td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
</table>
<%rs.close
set rs=nothing
closeconn
%>
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
			a[i].checked = a[a.length-1].checked;
}
</script>