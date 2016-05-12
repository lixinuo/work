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
<body>
<div class="aclass">
	<ul>
    	<li class="bold">管理员管理</li>
        <li class="bold" style="text-align:left; text-indent:10px;">添加管理系统用户</li>
        <form name="form1" method="post" action="?action=post_save" onSubmit="return checkPost()">
        <li>
        	<span class="width_30">用 户 名</span>
            <span class="width_70"><input name="s_name" id="s_name" type="text" class="left" value=""></span>
        </li>
        <li>
        	<span class="width_30">初始密码</span>
            <span class="width_70"><input name="s_pwd" id="s_pwd" type="password" class="left" maxlength="15" value=""></span>
        </li>
        <li>
        	<span class="width_30">用户级别</span>
            <span class="width_70">
                <select name="gradeid" id="gradeid" class="left">
                <%
                set Rs=db("select id,s_name from s_user_class order by id asc",2)
                if not rs.eof then
                do while not rs.eof 
                %>
                    <option value="<%=rs(0)%>"><%=rs(1)%></option>
                <%
                rs.movenext
                loop
                end if
                rs.close
				set rs=nothing
                %>
                </select>
                高级管理员可以添加管理员，普通管理员只能修改自己的密码
            </span>
        </li>
        <li>
        	<span><input type="submit" name="Submit" value=" 添加管理员 " class=inputkkys></span>
        </li>
        </form>
        <li></li>
        <li class="bold">管理员列表</li>
        <li>
        	<span class="width_50 bold">用户名【级别】</span>
            <span class="width_50 bold">操 作</span>
        </li>
		<%
        set rs=server.createobject("adodb.recordset")
        sql="select * from s_user where s_name<>'kendy520' order by id desc"
        rs.open sql,conn,1,1
        do while not rs.eof
        %>
        <li>
        	<span class="width_50">
            	<%=rs("s_name")%>【<font color="#FF0000"><%=DB_F("s_user_class","s_name",Rs("Gradeid"))%></font>】
            </span>
            <span class="width_50">
                <form name="form<%=rs(0)%>" method="post" action="?action=edit_pwd&id=<%=rs("id")%>">
                    <input name="newpwd" type="text" size="15">
                    <input type="submit" name="Submit2" class="inputkkys" value="修改密码">
                    <input type="button" name="Submit" class="inputkkys" value="恢复密码为123456" onClick="location='?action=pwd_default&id=<%=rs("id")%>'">
                    <input type="button" name="Submit" value="删 除" class="inputkkys" onClick="location='?action=user_del&id=<%=rs("id")%>'">
                </form>
            </span>
        </li>
        
		<%
        rs.movenext
        loop
        rs.close
        set rs=nothing
        %>
    </ul>
</div>
<script type="text/javascript">
function checkPost(){
	if($("#s_name").val()==""){
		alert("用户名不能为空！");
		$("#s_name").focus();
		return false;	
	}
	if($("#s_pwd").val()==""){
		alert("密码不能为空！");
		$("#s_pwd").focus();
		return false;	
	}	
}
</script>
</body>
</html>