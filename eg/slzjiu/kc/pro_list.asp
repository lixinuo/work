<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<!--#INCLUDE FILE="../inc/gFunc_Page.asp" -->
<%
page=trim(request("page"))
skey=trim(request("skey"))
stype=trim(request("stype"))
if s_pai="" then
s_pai=0
else
s_pai=cint(s_pai)
end if
url="Pro_list.asp?page="&page
%>
<html>
<head>
<title>商品管理</title>
<meta http-equiv="Content-Type" S_content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"--></head>
<body text="#000000" >
<%
if request("action")="orderby" then
  strIDs=request("id")
	orderby=request("orderby")
	strSQL="update S_kc set S_order="&orderby&" Where id=" & strIDs & ""
	conn.execute strSQL
end if


if request("action")="change_class" then'//删除全部
	selectid=request("selectid")
	change_classid=request("change_classid")
	if selectid<>"" and change_classid<>"" then
	  conn.execute "update P_main set classid="&change_classid&" where id in ("&selectid&")"
		response.Redirect url
		response.End
	end if
end if

selectid=request("selectid")
if selectid<>"" then
conn.execute "delete from S_kc where id in ("&selectid&")"
response.Redirect url
response.End
end if
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
    <tr>
      <td align="center">库存管理</td>
    </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
          <form name="search" method="post" action="<%=url%>">
            <tr>
              <td width="142" height="30" align="left" class="lmtjdt">按关键字查询：</td>
              <td width="204" class="lmtjdt"><input name="skey" type="text" id="skey" onFocus="this.value=''" value="请输入关键字">              </td>
              <td width="125" align="left" class="lmtjdt"><select name="stype" id="stype">
                  <option value="S_xh">按型号</option>
                  <option value="id">按资料序号</option>
                </select>              </td>
              <td width="223" height="30" align="left" class="lmtjdt"><input type="submit" class="inputkkys" name="Submit2" value="查询">              </td>
              <td width="292" align="left" class="lmtjdt"><input type="button" name="bb" class="inputkkys" onClick="location='pro_do.asp'" value="添加库存">              </td>
            </tr>
          </form>
        </table>
<table width="100%" border="0" cellpadding="5" cellspacing="0">
 <tr>
  <td bgcolor="<%=Color_0%>">
   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
     <td>
      <%'开始分页
sqlstr="select * from S_kc where 1=1"
if skey<>"" then
sqlstr=sqlstr&" and "&stype&" like '%"&skey&"%'"
end if
sqlstr=sqlstr&" order by S_order asc,id desc"		
			
set rs=server.CreateObject("adodb.recordset")
rs.open sqlstr,conn,1,1
if err.number<>0 then
response.write "NO data!"
end if

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
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="tableBorder">
       <form name="form1" method="post" action="<%=url%>">
	   <tr>
         <td colspan="6"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
         <tr>
         <td width="14%" height="30" align="center"><strong>ID</strong></td>
         <td width="14%" align="center"><strong>型号</strong></td>
         
         <td width="34%" align="center"><strong>数量</strong></td>
         <td width="15%" align="center"><strong>加入时间</strong></td>
         <td width="10%" align="center"><strong>修改</strong></td>
         <td width="13%" align="center"><strong>选 择/排序</strong></td>
        </tr>
        <%do while not (rs.eof or rs.bof) and count<rs.PageSize
		
if i mod 2=0 then 
mm_color="#C2DBFF"
else
mm_color="#F1F5FA"
end if
		
		
		%>
        <tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
         <td height="30" align="center"><%=i%></td>
         <td align="center"><%=rs("S_xh")%></td>
        
         <td align="center"><%=rs("s_sl")%></td>
         <td align="center"><%=rs("S_time")%></td>
         <td align="center">
          <%
		response.Write("<A title=修改内容 HREF='pro_do.asp?id="&rs("id")&"&s_pai="&s_pai&"'>编辑</A>&nbsp;")
		%>         </td>
         <td align="center">
          <input name="selectid" type="checkbox" id="selectid" value="<%=rs("id")%>">
          <input name="orderby<%=trim(rs("id"))%>" type="text" id="orderby" size="4" maxlength="5" value="<%=trim(rs("S_order"))%>" onChange="location='?id=<%=trim(rs("id"))%>&action=orderby&orderby=' + this.value">         </td>
        </tr>
        <%
rs.movenext
i=i+1
count=count+1
loop
%>
        <tr>
         <td height="30" colspan="8" align="center">
          <%response.Write( fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
	<input type="submit" name="Submit" value="批量删除" class="inputkkys" onClick="return test();">
      <input type="checkbox" name="checkbox" value="Check All" onClick="mm()">全选/反选&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
<%    
end if   
%>
         </table></td>
         </tr>
 
       </form>
      </table>
     </td>
    </tr>
   </table>
  </td>
 </tr>
</table>
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
