<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#include file="../inc/data.asp"-->
<%
id=isid(request.querystring("id"),0)
if id<>0 then
	Set Rs=Server.CreateObject("ADODB.RecordSet") 
	SqlStr="Select * From u_main where id="&id 
	Rs.Open SqlStr,Conn,1,1 
	if not rs.eof then
	 's_type=Rs("s_type") 
	 s_name=Rs("s_name") 
	 s_pwd=Rs("s_pwd") 
	 s_realname=Rs("s_realname") 
	 s_phone=Rs("s_phone") 
	 s_tel=Rs("s_tel") 
	 s_email=Rs("s_email") 
	 s_address=Rs("s_address") 
	 S_Quesion=Rs("S_Quesion") 
	 's_tuijian=Rs("s_tuijian") 
	 s_addtime=Rs("s_addtime") 
	 s_zip=Rs("s_zip") 
	 s_sex=Rs("s_sex") 
	 s_qq=Rs("s_qq") 
	 s_logins=Rs("s_logins")
	 s_jifen=Rs("s_jifen")
	 's_mark=Rs("s_mark")
	end if 
	Rs.Close:Set Rs=Nothing 
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/g_links.asp"-->
</head>
<body>
<table width="100%"   border="0" cellpadding="5" cellspacing="0" bgcolor="<%=Color_0%>">
  <tr>
    <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
    <li>用户资料修改    </li>
    </span></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#004A80"></td>
  </tr>
  <tr>
    <td valign="top"> 
<table class="tableBorder" width="80%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form action="user_save.asp?action=save&id=<%=id%>" method="post" name="form1" id="form1">
<!--<tr bgcolor="#ffffff"> 
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">用户类型：</td>
      <td width="662" bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
	  <input name="s_type" type="radio" value="1"<%=iif(s_type=1," checked","")%>/>普通会员
		<input name="s_type" type="radio" value="2"<%=iif(s_type=2," checked","")%>/>企业会员
		</span></td>
    </tr>-->
<tr bgcolor="#ffffff"> 
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">用 户 名：</td>
      <td width="662" bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
	  <input name="s_name" type="text" id="s_name" value="<%=s_name%>" size="25"/></span></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">用户密码：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="s_pwd" type="password" id="s_pwd" value="<%=s_pwd%>" size="25" />
      </span></td>
    </tr>
    
    
   <!-- <tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">用户积分：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px; color:#F00;"><strong><%=s_mark%></strong></td>
		</tr>
    -->
    
	    <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">名称：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="S_Quesion" type="text" id="S_Quesion" value="<%=S_Quesion%>" size="25" />
      </span></td>
    </tr>
		    <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">联系地址：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="S_address" type="text" id="S_address" value="<%=S_address%>" size="40" />
      </span></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">联系人：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="S_realname" type="text" id="S_realname" value="<%=S_realname%>" size="25" />
      </span></td>
    </tr>

    <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">电话：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input type="text" size="25" name="s_tel" value="<%=s_tel%>" />
      </span></td>
    </tr>
    <tr bgcolor="#ffffff" style="display:none;">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">传真：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input type="text" size="25" name="s_phone" value="<%=s_phone%>" />
      </span> </td>
    </tr>
 <!--   <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">企业名称：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input type="text" size="25" name="s_company" value="<%=s_company%>" />
      </span> </td>
    </tr>-->
    
    
		<tr>
		  <td height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">QQ：</td>
		  <td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="s_qq" type="text" id="s_qq" value="<%=s_qq%>" size="25" /></td>
		  </tr>
		<tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">注册时间：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><%=s_addtime%></td>
		</tr>
		<tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">登陆次数：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><%=s_logins%></td>
		</tr>
<tr>
<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"></td>
<td height="28" colspan="2" bgcolor="<%=Color_0%>" >
<input type="submit" name="Submit" value="确认提交" />
&nbsp;
<input type="button" name="Submit2" value="返回上一页" onclick='javascript:history.go(-1)' /></td>
</tr>
</form>
</table>
</td></tr></table>
</body>
</html>
<%closeconn%>