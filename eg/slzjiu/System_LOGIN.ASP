<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<%
if not isnumeric(Session("Count")) then Session("Count")=0
if Session("Count")>=3 then
Session("Count")=Empty
Response.Write("<script>window.opener=null;window.close();</script>")
Response.End()
end if
%>
<html>
<head>
<title>管理员登陆</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="css/login.css"/>
<SCRIPT type="text/javascript">
function SetFocus()
{
if (document.Login.u_name.value=="")
	document.Login.u_name.focus();
else
	document.Login.u_name.select();
}
function CheckForm()
{
	if(document.Login.u_name.value=="")
	{
		alert("请输入用户名！");
		document.Login.u_name.focus();
		return false;
	}
	if(document.Login.password.value == "")
	{
		alert("请输入密码！");
		document.Login.password.focus();
		return false;
	}
	if (document.Login.codenames.value==""){
       alert ("请输入您的验证码！");
       document.Login.codenames.focus();
       return(false);
    }
}
</SCRIPT>
</head>
<body>
<div class="kuang">
    <div class="kcon">
        <div class="ktitle">请登录</div>
        <div class="longin">
        <form name="Login" id="Login" method="post" action="System_chklogin.asp" onSubmit="return CheckForm();">
          <table width="100%" border="0" cellpadding="0" cellspacing="4">
            <tr>
              <td width="30%" height="40" align="right">帐&nbsp;号：</td>
              <td width="70%"><label for="username"></label>
                <input name="u_name" type="text" class='ninput' id="username" maxlength="30"/></td>
            </tr>
            <tr>
              <td height="40" align="right">密&nbsp;码：</td>
              <td><input name="password" type="password" class="ninput" id="password" /></td>
            </tr>
            <tr>
              <td height="40" align="right">验证码：</td>
              <td><input type="text" name="codenames" id="code" class="cinput" />
                &nbsp;&nbsp;<img src="inc/checkcode.asp" onClick="this.src='inc/checkcode.asp?id=Math.random()'" alt="点击更换" style="cursor:pointer" /></td>
            </tr>
            <tr>
              <td height="37" align="right">&nbsp;</td>
              <td><input name="remeber" type="checkbox" class="boxinput" id="remeber" value="1" />
                <label for="remeber">记住密码</label></td>
            </tr>
            <tr>
              <td height="43" colspan="2" align="center"><button type="submit" class="in"></button>&nbsp;&nbsp; <button class="ou" type="reset"></button><input type="hidden" name="action" value="login" /></td>
              </tr>
          </table>
          </form>
        </div>
        <div class="fg"><img src="theme/default/images/fg.jpg" /></div>
    </div>
</div>
</body>
</html>
