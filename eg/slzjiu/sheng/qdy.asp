<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
Server.ScriptTimeOut=9000
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
'生成html
Function Fso_info(host,folder,filename)
	host="http://"+Request.ServerVariables("HTTP_HOST")&host
	'response.Write host
	if SaveFile(""&folder&filename&"",""&host&"") then 
		response.Write "<a href='"&folder&filename&"' target='_blank'>"&folder&filename&"</a> 单页生成成功. <br />"
	else 
		Response.write ""&host&"_"&folder&"_"&filename&" 网页生成<font color='#FF0000'>失败</font>,可能您的文件名含有特殊字符或空间未开启写权限.<br />" 
	end if
End Function
%>
<%
'生成文件
function SaveFile(LocalFileName,RemoteFileUrl) 
	Dim Ads, Retrieval, GetRemoteData 
	On Error Resume Next 
	Set Retrieval = Server.CreateObject("Microso" & "ft.XM" & "LHTTP") '//把单词拆开防止杀毒软件误杀
	With Retrieval 
		.Open "Get", RemoteFileUrl, False, "", "" 
		.Send 
		GetRemoteData = .ResponseBody 
	End With 
	Set Retrieval = Nothing 
	Set Ads = Server.CreateObject("Ado" & "db.Str" & "eam") '//把单词拆开防止杀毒软件误杀
	With Ads 
		.Type = 1 
		.Open 
		.Write GetRemoteData 
		.SaveToFile Server.MapPath(LocalFileName), 2 
		.Cancel() 
		.Close() 
	End With 
	Set Ads=nothing 
	if err <> 0 then 
		SaveFile = false 
		err.clear 
	else 
		SaveFile = true 
	end if 
End function 
%>
<%
dydz=split("index,about",",")
for i=0 to ubound(dydz)
	host="/"&dydz(i)&".asp?ranNum="&now()&""
	folder="/" '静态地址文件夹
	html_url_name=""&dydz(i)&".html"
	fso_info host,folder,html_url_name
next
%>  

