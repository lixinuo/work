<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function

id=request("id")
set rs=server.CreateObject("adodb.recordset")
if id="" then
rs.Open "select top 1 * from P_main",conn,1,3
rs.AddNew
else
rs.Open "select * from P_main where id="&id,conn,1,3
end if

rs("s_pai")=int(trim(request("s_pai"))) 
rs("classid")=int(trim(request("classid")))
'rs("ccid")=int(trim(request("ccid"))) 
rs("S_name")=trim(request("S_name")) 
rs("S_name1")=trim(request("S_name1")) 
rs("S_name2")=trim(request("S_name2")) 
rs("S_name3")=trim(request("S_name3")) 
rs("S_name4")=trim(request("S_name4")) 
rs("S_name5")=trim(request("S_name5")) 
rs("S_bt")=trim(request("S_bt")) 
rs("S_bt1")=trim(request("S_bt1")) 
rs("S_bt2")=trim(request("S_bt2")) 
rs("S_gjc")=trim(request("S_gjc")) 
rs("S_gjc1")=trim(request("S_gjc1")) 
rs("S_gjc2")=trim(request("S_gjc2"))
rs("S_ms")=trim(request("S_ms")) 
rs("S_ms1")=trim(request("S_ms1")) 
rs("S_ms2")=trim(request("S_ms2"))
rs("S_img")=replace(trim(request("S_img")),"../","") 
rs("S_img1")=replace(trim(request("S_img1")),"../","") 
rs("S_down")=replace(trim(request("S_down")),"../.","")  '商品名称
rs("S_price")=0 '商品名称
rs("S_time")=trim(request("S_time"))
'rs("S_price1")=trim(request("S_price1")) '商品名称


rs("s_jifen")=0 '积分

rs("s_time")=Now()



if trim(request("S_content"))="" then
rs("S_content")=" "
else
Rs("s_content")=ReplaceStr(trim(request("S_content")),"'","&acute")
end if
if trim(request("S_content1"))="" then
rs("S_content1")=" "
else
Rs("s_content1")=ReplaceStr(trim(request("S_content1")),"'","&acute")
end if
if trim(request("S_content2"))="" then
rs("S_content2")=" "
else
Rs("s_content2")=ReplaceStr(trim(request("S_content2")),"'","&acute")
end if
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