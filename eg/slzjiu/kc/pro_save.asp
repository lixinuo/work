<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
id=request("id")
set rs=server.CreateObject("adodb.recordset")
if id="" then
rs.Open "select top 1 * from s_kc",conn,1,3
rs.AddNew
else
rs.Open "select * from s_kc where id="&id,conn,1,3
end if

 
rs("s_xh")=trim(request("s_xh")) 
rs("s_sl")=trim(request("s_sl")) 
rs.Update
rs.Close
set rs=nothing
urlForm=Request("urlForm")

if urlForm<>"" then
response.Write "<script language=javascript>alert('修改成功，请返回！');window.location='"&urlForm&"';</script>"
else
response.Write "<script language=javascript>alert('操作成功，请返回！');history.go(-1);</script>"
end if
response.End

%>