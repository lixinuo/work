<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<!--#INCLUDE file="../inc/MD5.ASP" -->
<% 
manage_u_name=request.Cookies(Cookies_name)("Name")
manage_u_id=db_f("s_user","id","s_name='"&manage_u_name&"'")
action=request("action")
id=isid(request("id"),0)
if action="post_save" then call post_save()
if action="edit_pwd" then call edit_pwd()
if action="pwd_default" then call pwd_default()
if action="user_del" then call user_del()
if action="edit_me" then call edit_me()

sub post_save()
	Set rs = server.createobject("ADODB.recordset")
	sql = "SELECT * FROM s_user where id is null" 
	rs.Open sql,conn,1,3
	rs.addnew
	rs("s_name") = trim(request("s_name"))
	rs("s_pwd") =md5(request("s_pwd"))
	rs("gradeid") = isid(request("gradeid"),2)
	rs.Update:rs.close:set rs=nothing
msg "操作成功","?"
end sub

sub edit_pwd()
conn.execute("update s_user set s_pwd='"&md5(trim(request("newpwd")))&"' where id="&id)
msg "修改成功","?"
end sub

sub pwd_default()
conn.execute("update s_user set s_pwd='"&md5("123456")&"' where id="&id)
msg "密码已恢复为123456","?"
end sub

sub user_del()
conn.execute("delete * from s_user where id="&id)
msg "删除成功","?"
end sub

sub edit_me()
old_pwd=trim(request("old_pwd"))
new_pwd=trim(request("new_pwd")):if new_pwd="" then msg "新密码不能为空",""
new_pwd1=trim(request("new_pwd1")):if new_pwd1="" then msg "新密码不能为空",""
if new_pwd<>new_pwd1 then msg "两次输入密码不同",""

conn.execute("update s_user set s_pwd='"&md5(new_pwd)&"' where id="&manage_u_id)
msg "修改成功","?"
end sub
%>
<html>
<head>
<title>管理密码</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body text="#000000" >
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
    <tr>
      <td align="center">管理员管理</td>
    </tr>
</table>
<table width="100%"   border="0" cellpadding="5" cellspacing="0">
 
 <tr>
  <td bgcolor="<%=Color_0%>">
<%
if isid(db_f("s_user","gradeid",manage_u_id),0)=1 then 
 call manage_admin()
else
 call edit_admin()
end if
%>
<%sub manage_admin()%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="tjnrym">
  <tr>
    <td height="40" class="tjnrbt">添加管理系统用户</td>
  </tr>
  <tr>
    <td height="40"><table width="96%" border="0" cellpadding="0" cellspacing="0" class="tjnrnk">
      <tr>
        <td><table width="600" height="162" border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#FFFFFF">
 <form name="form1" method="post" action="?action=post_save">
  
  <tr bgcolor="#ffffff">
   <td width="18%" height="30" align="center" bgcolor="<%=Color_0%>">用 户 名</td>
   <td width="82%" bgcolor="<%=Color_0%>">　
    <input name="s_name" type="text" size=20 style="width:150" maxlength="30" ></td>
  </tr>
  <tr bgcolor="#ffffff">
   <td width="18%" height="30" align="center" bgcolor="<%=Color_0%>">初始密码</td>
   <td width="82%" bgcolor="<%=Color_0%>">　
    <input name="s_pwd" type="password" size=20 style="width:150" maxlength="30" ></td>
  </tr>
  <tr bgcolor="#ffffff">
   <td width="18%" align="center" bgcolor="<%=Color_0%>">用户级别</td>
   <td width="82%" bgcolor="<%=Color_0%>" style="padding-left:20px;">
   
   <select name="gradeid" id="gradeid" style="width:129;">
     <%
set Rs=db("select id,GradeName from s_user_grade order by id asc",2)
if not rs.eof then
do while not rs.eof 
%>
     <option value="<%=rs(0)%>"><%=rs(1)%></option>
     <%
rs.movenext
loop
end if
rs.close:set rs=nothing
%>
    </select>
   高级管理员可以添加管理员，普通管理员只能修改自己的密码</td>
  </tr>
  <tr bgcolor="#ffffff" align="center">
   <td height="30" colspan=2 bgcolor="<%=Color_0%>"><input type="submit" name="Submit" value=" 添加管理员 " class=inputkkys></td>
  </tr>
 </form>
</table></td>
      </tr>
    </table></td>
  </tr>
</table>

<br>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
  <tr>
    <td align="center">管理员列表</td>
  </tr>
</table>
<table width="98%" height="94" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
 
 <tr align="center">
  <td width="25%" height="30"><strong>用户名【级别】</strong></td>
  <td width="75%"><strong>操 作</strong></td>
 </tr>
 <%set rs=server.createobject("adodb.recordset")
sql="select * from s_user where s_name<>'kendy520' order by id desc"
rs.open sql,conn,1,1
do while not rs.eof
%>
 <tr align="center" bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
  <td height="30"><%=rs("s_name")%>【<font color="#FF0000"><%=DB_F("s_user_Grade","GradeName",Rs("Gradeid"))%></font>】</td>
  <td align="left"><form name="form<%=rs(0)%>" method="post" action="?action=edit_pwd&id=<%=rs("id")%>">
    <input name="newpwd" type="text" size="15">
    <input type="submit" name="Submit2" class="inputkkys" value="修改密码">
    <input type="button" name="Submit" class="inputkkys" value="恢复密码为123456" onClick="location='?action=pwd_default&id=<%=rs("id")%>'">
    <%if rs("s_name")<>"admin" then%>
    <input type="button" name="Submit" value="删 除" class="inputkkys" onClick="location='?action=user_del&id=<%=rs("id")%>'">
    <%end if%>
   </form></td>
 </tr>
 <%rs.movenext
loop
rs.close
set rs=nothing
%>
</table>
<%end sub%>


<%sub edit_admin()%>
<table width="40%" border="0" align="center" cellpadding="5" cellspacing="1">
 <form method="POST" action="?action=edit_me" >
  <tr>
   <td height="30" align="center" bgcolor="<%=Color_1%>" ><strong>密码更改&nbsp;&nbsp;&nbsp;&nbsp;</strong></td>
  </tr>
  <tr>
   <td height="30" bgcolor="#ffffFF">　</td>
  </tr>
<!--  <tr>
   <td height="30" align="center" bgcolor="#FFFFFF"> 旧的密码： 　
    <input  name="old_pwd" size="30" type="password" ></td>
  </tr>-->
  <tr>
   <td height="30" align="center" bgcolor="#FFFFFF"> 新的密码： 　
    <input  name="new_pwd" size="30"  type="password" ></td>
  </tr>
  <tr>
   <td height="30" align="center" bgcolor="#FFFFFF"> 密码校验： 　
    <input  name="new_pwd1" size="30" type="password"></td>
  </tr>
  <tr>
   <td height="30" bgcolor="<%=Color_0%>">密码最好用字母和数字的组合，不要使用中文</td>
  </tr>
  <tr>
   <td height="30" align="center" ><input type="submit" value=" 确 定 " name="action" class=inputkkys>
    <input type="reset" value=" 重 填  " name="reset" class=inputkkys>
    </p></td>
  </tr>
 </form>
</table>
<%end sub%></td>
 </tr>

</table>
</body>
</html>