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
userid=trim(request("userid"))
xdrq=trim(request("xdrq"))
jhrq=trim(request("jhrq"))
jj=trim(request("jj"))
wjm=trim(request("wjm"))
pbfs=trim(request("pbfs"))
cs=trim(request("cs"))
sl=trim(request("sl"))
cens=trim(request("cens"))
bh=trim(request("bh"))
ym=trim(request("ym"))
zf=trim(request("zf"))
dx=trim(request("dx")) 
chsj=trim(request("chsj"))
chfs=trim(request("chfs"))
kddh=trim(request("kddh"))

select case action
case "save"
set rs=server.CreateObject("adodb.recordset")
if id<>0 then
	rs.Open "select * from [d_order] where id="&id,conn,3,3
else
	rs.Open "select top 1 * from [d_order]",conn,3,3
	rs.addnew()
	rs("adddate")=now()
end if

rs("userid")=userid
rs("xdrq")=xdrq
rs("jhrq")=jhrq
rs("jj")=jj
rs("wjm")=wjm
rs("pbfs")=pbfs
rs("cs")=cs
rs("sl")=sl
rs("cens")=cens
rs("bh")=bh
rs("ym")=ym
rs("zf")=zf
rs("dx")=dx
rs("chsj")=chsj
rs("chfs")=chfs
rs("kddh")=kddh

rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('操作成功!');history.go(-1);</script>"


case "del"
conn.execute "delete from [d_order] where id in ("&cid&") "
'response.Redirect "manageuser.asp"
response.Redirect request.servervariables("http_referer")
end select
%>
