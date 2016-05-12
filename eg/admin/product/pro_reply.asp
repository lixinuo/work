<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<!--#INCLUDE FILE="../inc/gFunc_Page.asp" -->
<%
tablename="p_reply"
page=request("Page"):s_content=request("s_content"):s_stateid=isid(request("s_stateid"),9)
selectid=request("selectid"):action=request("action")
parent_id=isid(request("parent_id"),0):id=isid(request("id"),0)
url="?page="&page&"&parent_id="&parent_id
url_=url&"&s_stateid="&s_stateid


state_a_str=""

if action="s_reply" then call post_reply()

sub post_reply()
conn.execute("update "&tablename&" set s_reply='"&trim(request("s_reply"))&"' where id="&id)
end sub
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<!--#include file="../inc/g_links.asp"-->
</head><body>
<table width="100%"   border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
 <tr>
  <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
   <li>评论列表 
   </span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
  <td valign="top" bgcolor="<%=Color_0%>"><%  
strSQL ="select * from "&tablename&" where parent_id="&parent_id
strSQL = strSQL & " ORDER BY id DESC"
'w strsql&"<br>s_stateid="&s_stateid
Set rs = db(strsql,2)
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
   <form name="form1" method="post" action="<%=url_%>&action=del">
    <%
do while not (rs.eof or rs.bof) and count<int_RPP 
%>
    <table width="800" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
     <tr bgcolor="<%=Color_0%>">
      <td width="100" height="30" align="left">姓名：<%=rs("s_name")%></td>
      <td width="150" align="left">时间：<%=rs("s_time")%></td>
      <td width="50" height="30" align="center"><a href="?action=del&id=<%=rs(0)%>">删除</a></td>
     </tr>
     <tr bgcolor="<%=Color_0%>">
      <td colspan="4" style="padding-left:10px;"><%=rs("s_content")%></td>
     </tr>
     <tr bgcolor="<%=Color_0%>">
      <td colspan="4" style="padding-left:10px;">回复:<textarea name="s_reply<%=rs(0)%>" cols="60" rows="4"><%if str_isnull(rs("s_reply")) then w "暂无回复" else w rs("s_reply")%></textarea>
      <input type="button" name="s_reply_submit<%=rs(0)%>" onClick="location='?action=s_reply&parent_id=<%=parent_id%>&id=<%=rs(0)%>&s_reply='+document.getElementById('s_reply<%=rs(0)%>').value" value="更新回复"></td>
     </tr>
    </table>
    <%
rs.movenext 
count=count+1
i=i+1
loop
%>
    <%end if%>
    <table align="center">
     <tr>
      <td bgcolor="<%=Color_0%>" colspan="6" align="center"><%response.Write( fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
      </td>
     </tr>
    </table>
   </form></td>
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
</html><script>
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
