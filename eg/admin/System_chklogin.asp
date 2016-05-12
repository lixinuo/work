<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE file="inc/MD5.ASP" -->
<!--#INCLUDE FILE="connection.asp" -->
<%
sub we(fcb_str)
 response.write(fcb_str):response.End()
end sub

Function isN(ByVal str)
	isN = False
	Select Case VarType(str)
		Case vbEmpty, vbNull
			isN = True : Exit Function
		Case vbString
			If str="" Then isN = True : Exit Function
		Case vbObject
			If TypeName(str)="Nothing" Or TypeName(str)="Empty" Then isN = True : Exit Function
		Case vbArray,8194,8204,8209
			If Ubound(str)=-1 Then isN = True : Exit Function
	End Select
End Function

function msg(fcb_str,fc_base_url)
	if isN(fc_base_url) then
	  we(js("alert('"&fcb_str&"');history.back();"))
	else
	  we(js("alert('"&fcb_str&"');location='"&fc_base_url&"';"))
	end if
end Function


function js(fcb_str)
js="<script>"&fcb_str&"</script>"
end function

const login_url="System_login.asp"
const codes="'c""c;c=c-c<c>c c\c/c,c.c`c!c@c#c$c%c^c&c*c(c)c|"
	u_name=Lcase(trim(Request.Form("u_name")))
	u_pwd=trim(Request.Form("password"))
	code=split(codes,"c")
	for li = 0 to Ubound(code)
	u_name=Replace(u_name,code(li),"")
	u_pwd=Replace(u_pwd,code(li),"")
	next
	u_pwd=md5(u_pwd)
	codenames=trim(Request.Form("codenames"))


if cstr(Session("CheckCode"))<>codenames then
	msg "登陆验证码错误",login_url
	response.End()
end if

sql="select * from S_User where S_name='"&u_name&"' and s_pwd='"&u_pwd&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not (rs.eof or rs.bof) then
	 pwd_rs=Trim(rs("S_pwd"))
	 if u_pwd<>pwd_rs then 
	 msg "你的密码错误",login_url
	 end if
		userip=Trim(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
		If userip="" Then 
		userip = Trim(Request.ServerVariables("REMOTE_ADDR"))
		end if
		
		if request.Form("remeber")="1" then
		    Response.Cookies(Cookies_name).expires=date+365
		 end if
		 
		Response.Cookies(Cookies_name)("ID")=rs("id")   
		Response.Cookies(Cookies_name)("Name")=u_name   
		Response.Cookies(Cookies_name)("Grade")=Db_s("select gradename from s_user_Grade where id="&rs("Gradeid"))
		Response.Cookies(Cookies_name)("Pwd")=pwd_rs 
		Response.Cookies(Cookies_name)("LastLoginTime")=time
		Session("verifycode") = ""
		Response.Redirect "System_main.asp"
else
   	 	session("CheckCode") = ""
		msg "你的用户名或密码错误",login_url
end if
rs.close:set rs=nothing
conn.close:set conn=nothing
 




	
function db_s(sqlstr) 
	set rs_rt=server.CreateObject("Adodb.recordset")
	rs_rt.open sqlstr,conn,1,1
		if not rs_rt.eof then
				rs_fstr=trim(rs_rt(0))
				db_s = rs_fstr
		else
				db_s = "None"
		end if
end function
%>
