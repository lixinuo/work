<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%dim rs,strSql
s_pai=trim(request("s_pai"))
parent_id=trim(request("parent_id"))
id=trim(request("id"))
if s_pai="" then
s_pai=0
else
s_pai=cint(s_pai)
end if

if id<>"" then
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from A_class where id="&id
rs.open strSql,Conn,1,1 
s_name=rs("s_name")
s_name1=rs("s_name1")
s_name2=rs("s_name2")
S_bt = rs("S_bt")
S_bt1 = rs("S_bt1")
S_bt2 = rs("S_bt2")
S_gjc = rs("S_gjc")
S_gjc1 = rs("S_gjc1")
S_gjc2 = rs("S_gjc2")
S_ms = rs("S_ms")
S_ms1 = rs("S_ms1")
S_ms2 = rs("S_ms2")
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
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
.style54 {color: #ff0000}
.style55 {color: #666666}
.STYLE56 {
	color: #FF0000;
	font-weight: bold;
}
.STYLE57 {color: #0000FF}
-->
</style></head>
<!--#include file="../inc/g_links.asp"-->
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="tjnrym">
  <tr>
    <td height="40" class="tjnrbt"><%if id="" then w("增加") else w("修改")%>分类</td>
  </tr>
  <tr>
    <td height="40"><table width="96%" border="0" cellpadding="0" cellspacing="0" class="tjnrnk">
      <tr>
        <td><form method="POST" action="class_save.asp" name="myform">
	<input type="hidden" name="id" value="<%=id%>" >
	<input type="hidden" name="s_pai" value="<%=s_pai%>" >
  <input type="hidden" name="parent_id" value="<%=parent_id%>">
 <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" height="71" bgcolor="#FFFFFF">
  <tr>
    <td width="19%" height="30" align="right" valign="middle">
      分类标题:</td>
    <td width="81%" height="30" valign="middle">
      <input type="text" name="S_name" class="inputkkys" size="40" value="<%=S_name%>">	  </td>
  </tr>
	 <%if instr(webLanguage,"1") then%>
  <tr>
    <td width="19%" height="30" align="right" valign="middle">
      分类标题（en）:</td>
    <td width="81%" height="30" valign="middle">
      <input type="text" name="S_name1" class="inputkkys" size="40" value="<%=S_name1%>">	  </td>
  </tr>
	  <%end if%>
	 <%if instr(webLanguage,"2") then%>
  <tr>
    <td width="19%" height="30" align="right" valign="middle">
      分类标题（日）</td>
    <td width="81%" height="30" valign="middle">
      <input type="text" name="S_name2" class="inputkkys" size="40" value="<%=S_name2%>">	  </td>
  </tr>
	  <%end if%>
	  <tr>
         <td width="19%" height="20" align="right" valign="middle">Title:</td>
         <td width="81%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt" type="text" class=inputkkys style="width:99%" size="60" value="<%=S_bt%>"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="19%" height="20" align="right" valign="middle">Title(en):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt1" style="width:99%" type="text" class=inputkkys  size="60" value="<%=S_bt1%>"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="19%" height="20" align="right" valign="middle">Title(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt2" style="width:99%" type="text" class=inputkkys  size="60" value="<%=S_bt2%>"></td>
        </tr>
        <%end if%>
		<tr>
         <td width="19%" height="30" align="right" valign="middle">Keyword:</td>
         <td width="81%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc" style="width:99%" type="text" class="inputkkys" value="<%=S_gjc%>" size="60"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="19%" height="30" align="right" valign="middle">Keyword(en):</td>
         <td valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc1" type="text" style="width:99%" class="inputkkys" value="<%=S_gjc1%>" size="60"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="19%" height="30" align="right" valign="middle">Keyword(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc2" style="width:99%" type="text" class="inputkkys" value="<%=S_gjc2%>" size="60"></td>
        </tr>
        <%end if%>
		<tr>
         <td width="19%" height="30" align="right" valign="middle">Description:</td>
         <td width="81%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms" style="width:99%" type="text" class="inputkkys" value="<%=S_ms%>" size="60"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="19%" height="30" align="right" valign="middle">Description(en):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms1" type="text" style="width:99%" class="inputkkys" value="<%=S_ms1%>" size="60"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="19%" height="30" align="right" valign="middle">Description(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms2" type="text" style="width:99%" class="inputkkys" value="<%=S_ms2%>" size="60"></td>
        </tr>
        <%end if%>
  <tr>
    <td height="30" colspan="2" valign="middle">
	<table width="100%"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="19%">　</td>
          <td width="54%" valign="middle"><input type="submit" name="button" value=" 确 定 " class=inputkkys onClick="cmdForm()">
&nbsp;            <input name="button" type="button" class=inputkkys id="button" value=" 返 回 &gt;&gt;" onClick="history.go(-1)">                  </td>
          <td width="27%" valign="middle"></td>
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