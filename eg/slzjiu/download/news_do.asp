<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
show_img=request("show_img"):if show_img="" then show_img=0 else show_img=cint(show_img)
s_pai=request("s_pai"):action=request("action")
if action="post" then call post()
if s_pai="" then s_pai=0
id=request("id")
if id<>"" then
dim rs,strSql,id
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from A_main where id="&id
rs.open strSql,Conn,1,1 
newsid = rs("id")
classid = rs("classid")
S_name = rs("S_name")
s_name1 = rs("s_name1")
s_name2 = rs("s_name2")
s_name3 = rs("s_name3")
S_down = rs("S_down")
S_bt = rs("S_bt")
S_bt1 = rs("S_bt1")
S_bt2 = rs("S_bt2")
S_gjc = rs("S_gjc")
S_gjc1 = rs("S_gjc1")
S_gjc2 = rs("S_gjc2")
S_ms = rs("S_ms")
S_ms1 = rs("S_ms1")
S_ms2 = rs("S_ms2")
s_content = rs("s_content")
s_content1 = rs("s_content1")
s_content2 = rs("s_content2")
s_time = rs("s_time")
rs.close
set rs=nothing
end if
if trim(s_time)="" or isnull(s_time) then s_time=now()
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<!--#include file="../inc/g_links.asp"-->
<%if webbjq=1 then%>
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
<%end if%>
</head>
<body text="#000000" >
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="tjnrym">
  <tr>
    <td height="40" class="tjnrbt">更新信息</td>
  </tr>
  <tr>
    <td height="40"><table width="96%" border="0" cellpadding="0" cellspacing="0" class="tjnrnk">
      <tr>
        <td><form method="POST" action="?action=post&show_img=<%=show_img%>" name="myform">
    <input type="hidden" name="id" value="<%=id%>" >
    <input type="hidden" name="s_pai" value="<%=s_pai%>" >
    <table width="100%" border="0" cellspacing="2" cellpadding="3" align="center">
     <tr>
      <td width="12%" height="27" align="right" bgcolor="<%=Color_0%>">信息分类:</td>
      <td height="27" valign="middle" bgcolor="<%=Color_0%>" width="88%"><select name="classid" id="classid" class="inputkkys">
        <% call db_ChildID(0,classid,"A_class",s_pai) %>
       </select></td>
     </tr>
     <tr>
      <td width="12%" valign="middle" height="20" bgcolor="<%=Color_0%>" align="right" class="style55">信息名称:</td>
      <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_name" type="text" class="inputkkys" size="50"  id="t2" value="<%=S_name%>">
	   </td>
     </tr>
	 <%if instr(webLanguage,"1") then%>
	 <tr>
      <td width="12%" valign="middle" height="20" bgcolor="<%=Color_0%>" align="right" class="style55">信息名称(en):</td>
      <td height="20" valign="middle" bgcolor="<%=Color_0%>">
       <input name="s_name1" type="text" class="inputkkys"  size="50"  value="<%=s_name1%>">
	   </td>
     </tr>
	 <%end if%>
	 <%if instr(webLanguage,"2") then%>
	 <tr>
      <td width="12%" valign="middle" height="20" bgcolor="<%=Color_0%>" align="right" class="style55">信息名称(日):</td>
      <td height="20" valign="middle" bgcolor="<%=Color_0%>">
       <input name="s_name2" type="text" class="inputkkys"  size="50"  value="<%=s_name2%>">
	   </td>
     </tr>
	 <%end if%>
     <%if show_img=1 then%>
     <tr bgcolor="<%=Color_0%>">
      <td height="20" align="right" valign="middle" bgcolor="<%=Color_0%>">资料上传:</td>
      <td height="20" valign="middle"><%=Upload_Init()%><%=Upload_Input("S_down",S_down)%></td>
     </tr>
     <%end if%>
	 <tr>
         <td width="13%" height="20" align="right" valign="middle" bgcolor="<%=Color_0%>">Title:</td>
         <td width="87%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt" type="text" class=inputkkys style="width:99%" size="60" value="<%=S_bt%>"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="13%" height="20" align="right" valign="middle" bgcolor="<%=Color_0%>">Title(en):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt1" style="width:99%" type="text" class=inputkkys  size="60" value="<%=S_bt1%>"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="13%" height="20" align="right" valign="middle">Title(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt2" style="width:99%" type="text" class=inputkkys  size="60" value="<%=S_bt2%>"></td>
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
         <td width="13%" height="30" align="right" valign="middle">来源:</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="s_name3" type="text" class="inputkkys" value="<%=s_name3%>" size="60"></td>
        </tr>
		<tr>
         <td width="13%" height="30" align="right" valign="middle">加入时间:</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="s_time" type="text" class="inputkkys" value="<%=s_time%>" size="60"></td>
        </tr>
     <tr bgcolor="<%=Color_0%>">
      <td width="12%" height="20" valign="bottom"><div align="right" class="style55">信息内容:</div></td>
      <td width="88%" height="20" valign="middle" bgcolor="<%=Color_0%>">
	  <%if webbjq=1 then%>
	  <textarea name="s_content" style="width:100%;height:400px;visibility:hidden;"><%=s_content%></textarea>
	  <%elseif webbjq=2 then%>
	  <input type=hidden name="s_content" value="<%=edit_html(s_content)%>">
       <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content&style=standard600" frameborder="0" scrolling="no" width="100%" height="400"></IFRAME>
	  <%end if%>
	  </td>
     </tr>
     <%if instr(webLanguage,"1") then%>
     <tr bgcolor="<%=Color_0%>">
      <td width="12%" height="20" valign="bottom"><div align="right" class="style55">Content:</div></td>
      <td width="88%" height="20" valign="middle" bgcolor="<%=Color_0%>">
	  <%if webbjq=1 then%>
	  <textarea name="s_content1" style="width:100%;height:400px;visibility:hidden;"><%=s_content1%></textarea>
	  <%elseif webbjq=2 then%>
	  <input type=hidden name="s_content1" value="<%=edit_html(s_content1)%>">
       <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content1&style=standard600" frameborder="0" scrolling="no" width="100%" height="400"></IFRAME>
	  <%end if%>
	  </td>
     </tr>
     <%end if%>
     <%if instr(webLanguage,"2") then%>
     <tr bgcolor="<%=Color_0%>">
      <td width="12%" height="20" valign="bottom"><div align="right" class="style55">日:</div></td>
      <td width="88%" height="20" valign="middle" bgcolor="<%=Color_0%>">
	  <%if webbjq=1 then%>
	  <textarea name="s_content2" style="width:100%;height:400px;visibility:hidden;"><%=s_content2%></textarea>
	  <%elseif webbjq=2 then%>
	  <input type=hidden name="s_content2" value="<%=edit_html(s_content2)%>">
       <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content2&style=standard600" frameborder="0" scrolling="no" width="100%" height="400"></IFRAME>
	  <%end if%>
	  </td>
     </tr>
     <%end if%>
     <tr>
      <td height="30" colspan="2" valign="middle" bgcolor="<%=Color_0%>"><table width="100%"  border="0" cellspacing="0" cellpadding="3">
        <tr>
         <td width="18%">　</td>
         <td width="65%" valign="middle"><input type="submit" name="button" value=" 确 定 " class=inputkkys onClick="cmdForm()">
          &nbsp;&nbsp;&nbsp;
          <input name="button" type="button" class=inputkkys id="button" value=" 返 回 &gt;&gt;" onClick="history.go(-1)">
         </td>
         <td width="17%" valign="middle"></td>
        </tr>
       </table></td>
     </tr>
    </table>
   </form></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%

sub post()
s_pai=request.form("s_pai")
id = request.form("id")
show_img = request.form("show_img")
classid = changechr(request.form("classid"))
S_name = changechr(trim(request.form("S_name")))
s_name1 = changechr(trim(request.form("s_name1")))
s_name2 = changechr(trim(request.form("s_name2")))
s_name3 = changechr(trim(request.form("s_name3")))
s_bt=changechr(trim(request("s_bt")))
s_bt1=changechr(trim(Request("s_bt1")))
s_bt2=changechr(trim(Request("s_bt2")))	
s_gjc=changechr(trim(request("s_gjc")))
s_gjc1=changechr(trim(Request("s_gjc1")))
s_gjc2=changechr(trim(Request("s_gjc2")))	
s_ms=changechr(trim(request("s_ms")))
s_ms1=changechr(trim(Request("s_ms1")))
s_ms2=changechr(trim(Request("s_ms2")))
s_content = changechr(request.form("s_content"))
s_content1 = changechr(request.form("s_content1"))
s_content2 = changechr(request.form("s_content2"))
s_time = request("s_time")
S_down= replace(trim(request.form("S_down")),"../","")

if S_name = ""  then msg "请输入标题，谢谢",""
if s_time = "" then s_time = "nothing"

	Set rs = server.createobject("ADODB.recordset")
	if id="" then
	sql = "SELECT * FROM A_main where id is null" 
	rs.Open sql,conn,1,3
	rs.addnew
	else
	sql = "SELECT * FROM A_main where id="&id 
	rs.Open sql,conn,1,3
	end if
	rs("s_pai") = s_pai
	rs("classid") =classid
	rs("S_name") = S_name
	rs("s_name1") = s_name1
	rs("s_name2") = s_name2
	rs("s_name3") = s_name3
	rs("S_down") = S_down
	Rs("s_bt")=s_bt
	Rs("s_bt1")=s_bt1
	Rs("s_bt2")=s_bt2
	Rs("s_gjc")=s_gjc
	Rs("s_gjc1")=s_gjc1
	Rs("s_gjc2")=s_gjc2
	Rs("s_ms")=s_ms
	Rs("s_ms1")=s_ms1
	Rs("s_ms2")=s_ms2
	Rs("s_content")=ReplaceStr(s_content,"'","&acute")
	Rs("s_content1")=ReplaceStr(s_content1,"'","&acute"	)
	Rs("s_content2")=ReplaceStr(s_content2,"'","&acute"	)
	Rs("s_time")=s_time
	rs.Update:rs.close:set rs=nothing
	Response.Write "<script Language=Javascript>alert('操作成功!');history.go(-1);</script>"
end sub
%>