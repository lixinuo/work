<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
dim userid,action
action=request.QueryString("action")
id=isid(request("id"),0)
userid=request("userid")
select case action
case "save"
set rs=server.CreateObject("adodb.recordset")
if id<>0 then
	rs.Open "select * from [u_main] where id="&id,conn,3,3
	if trim(request("s_pwd"))<>rs("s_pwd") then rs("s_pwd")=md5(trim(request("s_pwd")))
else
	rs.Open "select top 1 * from [u_main]",conn,3,3
	rs.addnew()
	rs("s_pwd")=md5(trim(request("s_pwd")))
end if

rs("s_name")=trim(request("s_name"))
rs("s_realname")=trim(request("s_realname"))

rs("s_realname")=trim(request("s_realname"))

rs("S_Quesion")=trim(request("S_Quesion"))

rs("s_qq")=trim(request("s_qq"))

rs("s_email")=trim(request("s_email"))
rs("s_address")=trim(request("s_address"))

rs("s_phone")=trim(request("s_phone"))
rs("s_tel")=trim(request("s_tel"))

rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('操作成功!');history.go(-1);</script>"


case "del"
conn.execute "delete from [u_main] where id in ("&userid&") "
'response.Redirect "manageuser.asp"
response.Redirect request.servervariables("http_referer")
end select
%>
