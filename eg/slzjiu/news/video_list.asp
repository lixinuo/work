<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<!--#include file="../inc/g_links.asp"-->
</head><body>
<table width="100%"   border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
 <tr>
  <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
   <li>信息管理 
   </span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
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
    s_img=Replace(trim(request("s_img")),"../../","")
	S_name=changechr(trim(request("S_name")))
	s_name1=changechr(trim(Request("S_name1")))
	s_name2=changechr(trim(Request("S_name2")))	
	s_content=changechr(trim(request("s_content")))
	s_content1=changechr(trim(request("s_content1")))
	s_content2=changechr(trim(request("s_content2")))
	s_time=Now()
	
	
	sql = "SELECT * FROM A_info where id="&id 
	Set rs = server.createobject("ADODB.recordset")
	rs.Open sql,conn,1,3
	If NOT Rs.Eof then
	Rs("s_img")=s_img
	Rs("s_name")=s_name
	Rs("s_name1")=s_name1
	Rs("s_name2")=s_name2
	Rs("s_content")=ReplaceStr(s_content,"'","&acute")
	Rs("s_content1")=ReplaceStr(s_content1,"'","&acute"	)
	Rs("s_content2")=ReplaceStr(s_content2,"'","&acute"	)
	Rs.Update()
	End If
	Rs.close
	set Rs=NOthing
	Response.Write "<script Language=Javascript>alert('操作成功!');history.go(-1);</script>"
	'response.Redirect("?action=edit&id="&Request("id")&"&s_pai="&s_pai&"&show_del="&show_del&"&show_img="&show_img)
	'msg "修改成功","?action=edit&id="&Request("id")&"&s_pai="&s_pai&"&show_del="&show_del&"&show_img="&show_img
case "save_add"
    s_pai=Request("s_pai")
    s_img=Replace(trim(request("s_img")),"../../","")
	S_name=changechr(trim(request("S_name")))
	s_name1=changechr(trim(Request("S_name1")))
	s_name2=changechr(trim(Request("S_name2")))	
	s_content=changechr(trim(request("s_content")))
	s_content1=changechr(trim(request("s_content1")))
	s_content2=changechr(trim(request("s_content2")))
	s_time=Now()
	
	
	sql = "SELECT * FROM A_info where id is null" 
	Set rs = server.createobject("ADODB.recordset")
	rs.Open sql,conn,1,3
	Rs.Addnew()
	Rs("s_pai")=s_pai
	Rs("s_img")=s_img
	Rs("s_name")=s_name
	Rs("s_name1")=s_name1
	Rs("s_name2")=s_name2
	Rs("s_content")=ReplaceStr(s_content,"'","&acute")
	Rs("s_content1")=ReplaceStr(s_content1,"'","&acute"	)
	Rs("s_content2")=ReplaceStr(s_content2,"'","&acute"	)
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
   <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
    <FORM action="<%= page_url%>&show_del=<%=show_del%>" method=post name=sele>
     <tr valign="middle" bgcolor="#f3f3f3">
      <td height="30" colspan="7" bgcolor="<%= Color_2%>"><table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
         <td>[<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个]
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
          <input style="border:1px solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal; background-color:<%= Color_2%>;" type='submit'  value=' Goto '   size=2>
          <%if show_del=1 then%>
          <input type="button" value=" 添加文章 "  name="buttonurl" onClick="location='?action=add&s_pai=<%=s_pai%>&show_del=<%=show_del%>&show_img=<%=show_img%>'">
          <%end if%>
         </td>
         <td></td>
        </tr>
       </table></td>
     </tr>
    </form>
    <tr bgcolor="#3399CC">
     <td width="5%"><div align="center">ID</div></td>
     <td width="28%" height="30"><div align="center">信息标题</div></td>
     <td width="13%" height="30"><div align="center">创建日期</div></td>
     <td width="13%"><div align="center">排序</div></td>
     <td width="11%" height="30"><div align="center">修改</div></td>
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
    <tr bgcolor="<%=mm_color%>">
     <td style="padding-left:10px;"><%=rs(0)%></td>
     <td height="30" style="padding-left:10px;"><%=trim(rs("S_name"))%></font></td>
     <td height="30" align="center"><%=trim(rs("s_time"))%></td>
     <td align="center"><input name="orderby<%=trim(rs("id"))%>" type="text" id="orderby" size="4" maxlength="5" value="<%=trim(rs("s_order"))%>" onChange="location='?strIDs=<%=trim(rs("id"))%>&id=<%=id%>&page=<%=page%>&action=orderby&s_pai=<%=s_pai%>&show_add=<%=show_add%>&show_next=<%=show_next%>&show_del=<%=show_del%>&orderby=' + this.value"></td>
     <td height="30" width="11%"><div align="center"> <A HREF="?page=<%=page%>&action=edit&id=<%=trim(rs("id"))%>&s_pai=<%=s_pai%>&show_del=<%=show_del%>&show_img=<%=show_img%>"><img src="../images/detail_off.gif" border=0></A>&nbsp;&nbsp;
       <%if show_del=1 then%>
       <A HREF="?action=del&page=<%=page%>&delid=<%=trim(rs("id"))%>&s_pai=<%=s_pai%>&show_del=<%=show_del%>&show_img=<%=show_img%>"><img src="../images/del.gif" border=0></A>
       <%end if%>
      </div></td>
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
id=rs("id"):s_pai=rs("s_pai"):S_name=rs("S_name"):s_img=rs("s_img"):s_name1=rs("s_name1"):s_name2=rs("s_name2"):s_content=edit_html1(rs("s_content")):s_content1=edit_html1(rs("s_content1")):s_content2=edit_html1(rs("s_content2"))
rs.close:set rs=nothing
else
actstr="?action=save_add" 
end if
%>
    <tr>
     <td colspan="5"><form method="POST" action="<%=actstr%>&page=<%=page%>" name="myform">
       <input type="hidden" name="id" value="<%=trim(id)%>" >
       <input type="hidden" name="s_pai" value="<%=trim(s_pai)%>" >
       <input type="hidden" name="show_del" value="<%=trim(request("show_del"))%>" >
       <input type="hidden" name="show_img" value="<%=trim(request("show_img"))%>" >
       <table width="90%" border="0" cellspacing="2" cellpadding="3" align="center" height="">
        <tr>
         <td height="3" colspan="2"  bgcolor="#ff9900"></td>
        </tr>
        <tr>
         <td width="80" valign="middle" height="20" bgcolor="<%=Color_0%>"><div align="right" class="style55">信息标题:</div></td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_name" type="text" class=input  size="30" maxlength="50" value="<%=S_name%>">
          <%if instr(weblanguage,"1") then%>
          EN:
          <input name="s_name1" type="text" class=input  size="30" maxlength="50" value="<%=s_name1%>">
          <%end if%>
          <%if instr(weblanguage,"2") then%>
          繁体:
          <input name="s_name2" type="text" class=input  size="30" maxlength="50" value="<%=s_name2%>">
          <%end if%>
         </td>
        </tr>
        <%if request("show_img")<>0 then%>
        <tr>
         <td width="80" valign="middle" height="20" bgcolor="<%=Color_0%>"><div align="right" class="style55">上传视频:</div></td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><%=Upload_Init()%><%=Upload_Input("S_img",s_img)%><font color="#FF0000">*请上传正确的FLV格式的视频，若不符合，请转换FLV，谢谢！*</font></td>
        </tr>
        <%end if%>
        <tr>
         <td width="80" valign="bottom" height="36" bgcolor="<%=Color_0%>"><div align="right" class="style55">信息内容:</div></td>
         <td height="36" valign="middle" bgcolor="<%=Color_0%>"><input type=hidden name="s_content" id="s_content" value='<%=s_content%>'>
          <IFRAME ID="txtcontent" src="../htmledit/ewebeditor.htm?id=s_content&style=standard600" frameborder="0" scrolling="no" width="100%" height="250"></IFRAME></td>
        </tr>
        <%if instr(weblanguage,"1") then%>
        <tr bgcolor="<%=Color_0%>">
         <td width="64" height="20" valign="bottom"><div align="right" class="style55">Content:</div></td>
         <td width="572" height="20" valign="middle" bgcolor="<%=Color_0%>"><input type=hidden name="s_content1" value="<%=s_content1%>">
          <IFRAME ID="txtcontent" src="../htmledit/ewebeditor.htm?id=s_content1&style=standard600" frameborder="0" scrolling="no" width="100%" height="250"></IFRAME></td>
        </tr>
        <%end if%>
        <%if instr(weblanguage,"2") then%>
        <tr bgcolor="<%=Color_0%>">
         <td width="64" height="20" valign="bottom"><div align="right" class="style55">繁体:</div></td>
         <td width="572" height="20" valign="middle" bgcolor="<%=Color_0%>"><input type=hidden name="s_content2" value="<%=s_content2%>">
          <IFRAME ID="txtcontent" src="../htmledit/ewebeditor.htm?id=s_content2&style=standard600" frameborder="0" scrolling="no" width="100%" height="250"></IFRAME></td>
        </tr>
        <%end if%>
        <tr>
         <td height="30" colspan="2" valign="middle" bgcolor="<%=Color_0%>"><table width="100%"  border="0" cellspacing="0" cellpadding="3">
           <tr>
            <td width="18%">　</td>
            <td width="65%" valign="middle"><input type="submit" name="button" value=" 确 定 " class=input3 onClick="cmdForm()">
             &nbsp;&nbsp;&nbsp;
             <input name="button" type="button" class=input3 id="button" value=" 返 回 &gt;&gt;" onClick="history.go(-1)">
            </td>
            <td width="17%" valign="middle"></td>
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
