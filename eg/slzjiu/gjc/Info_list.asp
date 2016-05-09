<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<!--#include file="../inc/g_links.asp"-->
<link rel="stylesheet" href="../htmledit/themes/default/default.css" />
<script charset="utf-8" src="../htmledit/kindeditor.js"></script>
<script charset="utf-8" src="../htmledit/lang/zh_CN.js"></script>
<script>
KindEditor.ready(function(K) {
	K.create('textarea[name="s_content"]', {
		allowFileManager : true
	});
	K.create('textarea[name="s_content1"]', {
		allowFileManager : true
	});
	K.create('textarea[name="s_content2"]', {
		allowFileManager : true
	});
});
</script>
</head><body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
    <tr>
      <td align="center">单页关键字管理</td>
    </tr>
</table>
<table width="98%"   border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
 
 <tr>
  <td valign="top" bgcolor="<%=Color_0%>"><%
action=trim(request("action")):title=trim(request("title")):s_pai=trim(request("s_pai")):page=request("page")
show_del=trim(request("show_del")):if show_del="" then show_del=0
show_img=trim(request("show_img")):if show_img="" then show_img=0
page_url="Info_list.asp?s_pai="&s_pai&"&show_img="&show_img

select case action
case "orderby"
	strSQL="update A_info set s_order='"&request("orderby")&"' Where id=" & request("strIDs") & ""
	conn.execute strSQL:call list()
case "del"
	strSQL="DELETE FROM A_info where id="&Request("delid")
	conn.execute strSQL:call list()
case "save_edit"
    id=Request("id")
	S_name=changechr(trim(request("S_name")))
	s_name1=changechr(trim(Request("S_name1")))
	s_name2=changechr(trim(Request("S_name2")))	
	s_gjc=changechr(trim(request("s_gjc")))
	s_gjc1=changechr(trim(Request("s_gjc1")))
	s_gjc2=changechr(trim(Request("s_gjc2")))	
	s_ms=changechr(trim(request("s_ms")))
	s_ms1=changechr(trim(Request("s_ms1")))
	s_ms2=changechr(trim(Request("s_ms2")))
	s_time=Now()
	
	
	sql = "SELECT * FROM A_info where id="&id 
	Set rs = server.createobject("ADODB.recordset")
	rs.Open sql,conn,1,3
	If NOT Rs.Eof then
	Rs("s_name")=s_name
	Rs("s_name1")=s_name1
	Rs("s_name2")=s_name2
	Rs("s_bt")=s_name
	Rs("s_bt1")=s_name1
	Rs("s_bt2")=s_name2
	Rs("s_gjc")=s_gjc
	Rs("s_gjc1")=s_gjc1
	Rs("s_gjc2")=s_gjc2
	Rs("s_ms")=s_ms
	Rs("s_ms1")=s_ms1
	Rs("s_ms2")=s_ms2
	Rs.Update()
	End If
	Rs.close
	set Rs=NOthing
	Response.Write "<script Language=Javascript>alert('操作成功!');history.go(-1);</script>"
	'response.Redirect("?action=edit&id="&Request("id")&"&s_pai="&s_pai&"&show_del="&show_del&"&show_img="&show_img)
	'msg "修改成功","?action=edit&id="&Request("id")&"&s_pai="&s_pai&"&show_del="&show_del&"&show_img="&show_img
case "save_add"
    s_pai=Request("s_pai")
	S_name=changechr(trim(request("S_name")))
	s_name1=changechr(trim(Request("S_name1")))
	s_name2=changechr(trim(Request("S_name2")))	
	s_gjc=changechr(trim(request("s_gjc")))
	s_gjc1=changechr(trim(Request("s_gjc1")))
	s_gjc2=changechr(trim(Request("s_gjc2")))	
	s_ms=changechr(trim(request("s_ms")))
	s_ms1=changechr(trim(Request("s_ms1")))
	s_ms2=changechr(trim(Request("s_ms2")))	

	s_time=Now()
	
	
	sql = "SELECT * FROM A_info where id is null" 
	Set rs = server.createobject("ADODB.recordset")
	rs.Open sql,conn,1,3
	Rs.Addnew()
	Rs("s_pai")=s_pai
	Rs("s_name")=s_name
	Rs("s_name1")=s_name1
	Rs("s_name2")=s_name2
	Rs("s_bt")=s_name
	Rs("s_bt1")=s_name1
	Rs("s_bt2")=s_name2
	Rs("s_gjc")=s_gjc
	Rs("s_gjc1")=s_gjc1
	Rs("s_gjc2")=s_gjc2
	Rs("s_ms")=s_ms
	Rs("s_ms1")=s_ms1
	Rs("s_ms2")=s_ms2
	Rs.Update()
	Rs.close
	set Rs=NOthing

	Response.Write "<script Language=Javascript>alert('操作成功!');history.go(-1);</script>"


	'strSQL="insert into A_info (s_pai,s_img,S_name,s_name1,s_name2,s_content,s_content1,s_content2) values ("&trim(request("s_pai"))&",'"&Replace(trim(request("s_img")),"../../","") &"','"&trim(request("S_name")) &"','"&trim(request("s_name1")) &"','"&trim(request("s_name2")) &"','"&trim(request("s_content")) &"','"&trim(request("s_content1")) &"','"&trim(request("s_content2")) &"')"
'	conn.execute strSQL:call list()
case "add"
  call addedit()
case "edit"
  call addedit()
case else
  call list()
end select

sub list()
PageSize = 20

strSQL ="select * from A_info where id<>0"
if s_pai<>"" then
	strSQL = strSQL & " and s_pai="&s_pai
end if
strSQL = strSQL & " ORDER BY s_order asc,id desc"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,1,1


    rs.PageSize = PageSize
	totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if not isnumeric(page) then page=1
				if isnumeric(page) then:if cint(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

%>
   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <FORM action="<%= page_url%>&show_del=<%=show_del%>" method=post name=sele>
     <tr valign="middle" bgcolor="#f3f3f3">
      <td height="30" colspan="7">
	  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
      <tr>
        <td class="lmtjdt">[<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个]
          <%if page=1 then%>
          [首 页] [上一页]
          <% else %>
          [<a href="<%= page_url%>&page=1&type=<%=type1%>&show_del=<%=show_del%>">首 页</a>] 
          [<a href="<%= page_url%>&page=<%=page-1%>&type=<%=type1%>&show_del=<%=show_del%>">上一页</a>]
          <%end if%>
          <%if rs.pagecount-page<1 then%>
          [下一页] [尾 页]
          <%else%>
          [<a href="<%= page_url%>&page=<%=page+1%>&type=<%=type1%>&show_del=<%=show_del%>">下一页</a>]  [<a href="<%= page_url%>&page=<%=rs.pagecount%>&type=<%=type1%>&show_del=<%=show_del%>">尾 页</a>]</FONT>
          <%end if%>
          <input type='text' name='page' size=3 maxlength=10 value="<%=page%>" align="center" style="background-color:<%= Color_2%>; ">
          <input class="inputkkys" type='submit'  value=' Goto '   size=2>
          <%if show_del=1 then%>
          <input type="button" value=" 关键词添加" class="inputkkys"  name="buttonurl" onClick="location='?action=add&s_pai=<%=s_pai%>&show_del=<%=show_del%>&show_img=<%=show_img%>'">
          <%end if%>  </td>
      </tr>
    </table>	  </td>
     </tr>
    </form>
	<tr>
     <td height="30" colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0" class="lmtj">
     <tr>
     <td width="5%" align="center">ID</td>
     <td width="28%" height="30" align="left">信息标题</td>
     <td width="13%" height="30" align="center">创建日期</td>
     <td width="13%" align="center">排序</td>
     <td width="11%" height="30" align="center">修改</td>
    </tr>
    <%
if rs.eof then
response.write no_thing
else
count=0 
i=i+1
do while not (rs.eof or rs.bof) and count<rs.PageSize 

if i mod 2=0 then 
mm_color="#C2DBFF"
else
mm_color="#F1F5FA"
end if  
%>
    <tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
     <td align="center"><%=rs(0)%></td>
     <td height="30"><%=trim(rs("S_name"))%></font></td>
     <td height="30" align="center"><%=trim(rs("s_time"))%></td>
     <td align="center"><input name="orderby<%=trim(rs("id"))%>" type="text" id="orderby" size="4" maxlength="5" value="<%=trim(rs("s_order"))%>" onChange="location='?strIDs=<%=trim(rs("id"))%>&id=<%=id%>&page=<%=page%>&action=orderby&s_pai=<%=s_pai%>&show_add=<%=show_add%>&show_next=<%=show_next%>&show_del=<%=show_del%>&orderby=' + this.value"></td>
     <td width="11%" height="30" align="center"> <A HREF="?page=<%=page%>&action=edit&id=<%=trim(rs("id"))%>&s_pai=<%=s_pai%>&show_del=<%=show_del%>&show_img=<%=show_img%>"><img src="../images/detail_off.gif" border=0></A>&nbsp;&nbsp;
       <%if show_del=1 then%>
       <A HREF="?action=del&page=<%=page%>&delid=<%=trim(rs("id"))%>&s_pai=<%=s_pai%>&show_del=<%=show_del%>&show_img=<%=show_img%>"><img src="../images/del.gif" border=0></A>
       <%end if%>      </td>
    </tr>
	<%
rs.movenext 
count=count+1
i=i+1
loop 
end if
rs.close:set rs=nothing
end sub
sub addedit()
if action="edit" then
actstr="?action=save_edit" 
strSql="select * from A_info where id="&request("id")
set rs=db(strsql,2)
id=rs("id"):s_pai=rs("s_pai"):S_name=rs("S_name"):s_name1=rs("s_name1"):s_name2=rs("s_name2"):S_gjc=rs("S_gjc"):S_gjc1=rs("S_gjc1"):S_gjc2=rs("S_gjc2"):S_ms=rs("S_ms"):S_ms1=rs("S_ms1"):S_ms2=rs("S_ms2")
rs.close:set rs=nothing
else
actstr="?action=save_add" 
end if
%>
     </table></td>
     </tr>
   
    
    <tr>
     <td colspan="5"><form method="POST" action="<%=actstr%>&page=<%=page%>" name="myform">
       <input type="hidden" name="id" value="<%=trim(id)%>" >
       <input type="hidden" name="s_pai" value="<%=trim(s_pai)%>" >
       <input type="hidden" name="show_del" value="<%=trim(request("show_del"))%>" >
       <input type="hidden" name="show_img" value="<%=trim(request("show_img"))%>" >
	   <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="tjnrym">
  <tr>
    <td height="40" class="tjnrbt">更新信息</td>
  </tr>
  <tr>
		<td height="40"><table width="96%" border="0" cellpadding="0" cellspacing="0" class="tjnrnk">
		  <tr>
			<td><table width="100%" border="0" cellspacing="2" cellpadding="3" align="center" height="">
        <tr>
         <td width="13%" height="20" align="right" valign="middle" bgcolor="<%=Color_0%>">Title:</td>
         <td width="87%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_name" type="text" class=inputkkys style="width:99%" size="60" value="<%=S_name%>"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="13%" height="20" align="right" valign="middle" bgcolor="<%=Color_0%>">Title(en):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="s_name1" style="width:99%" type="text" class=inputkkys  size="60" value="<%=s_name1%>"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="13%" height="20" align="right" valign="middle">Title(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="s_name2" style="width:99%" type="text" class=inputkkys  size="60" value="<%=s_name2%>"></td>
        </tr>
        <%end if%>
		<tr>
         <td width="13%" height="30" align="right" valign="middle" bgcolor="<%=Color_0%>">Keyword:</td>
         <td width="87%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc" style="width:99%" type="text" class="inputkkys" value="<%=S_gjc%>" size="60"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="13%" height="30" align="right" valign="middle" bgcolor="<%=Color_0%>">Keyword(en):</td>
         <td valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc1" type="text" style="width:99%" class="inputkkys" value="<%=S_gjc1%>" size="60"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="13%" height="30" align="right" valign="middle">Keyword(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc2" style="width:99%" type="text" class="inputkkys" value="<%=S_gjc2%>" size="60"></td>
        </tr>
        <%end if%>
		<tr>
         <td width="13%" height="30" align="right" valign="middle" bgcolor="<%=Color_0%>">Description:</td>
         <td width="87%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms" style="width:99%" type="text" class="inputkkys" value="<%=S_ms%>" size="60"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="13%" height="30" align="right" valign="middle" bgcolor="<%=Color_0%>">Description(en):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms1" type="text" style="width:99%" class="inputkkys" value="<%=S_ms1%>" size="60"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="13%" height="30" align="right" valign="middle">Description(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms2" type="text" style="width:99%" class="inputkkys" value="<%=S_ms2%>" size="60"></td>
        </tr>
        <%end if%>
        <tr>
         <td height="30" colspan="2" valign="middle" bgcolor="<%=Color_0%>"><table width="100%"  border="0" cellspacing="0" cellpadding="3">
           <tr>
            <td width="18%">　</td>
            <td width="65%" valign="middle"><input type="submit" name="button" value=" 确 定 " class=inputkkys onClick="cmdForm()">
             &nbsp;&nbsp;&nbsp;
             <input name="button" type="button" class=inputkkys id="button" value=" 返 回 &gt;&gt;" onClick="history.go(-1)">            </td>
            <td width="17%" valign="middle"></td>
           </tr>
          </table></td>
        </tr>
       </table></td>
		  </tr>
		</table></td>
	  </tr>
	</table>
       
      </form></td>
    </tr>
    <%
				end sub
				%>
   </table></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
</table>
</body>
</html>
<%closeconn%>
<%
function edit_html1(str)
if str=null or str=empty or str="" then str="&nbsp;"
edit_html1=server.HTMLEncode(str)
end function
%>
