﻿
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
		response.Write "<a href='"&folder&filename&"' target='_blank'>"&folder&filename&"</a> 网页生成成功. <br />"
	else 
		Response.write ""&host&"_"&folder&"_"&filename&" 网页生成<font color='#FF0000'>失败</font>,可能您的文件名含有特殊字符或空间未开启写权限.<br />" 
	end if
End Function
%>
<%
'建立文件夹函数
Function CreateFolder(strFolder)'参数为相对路径'首选判断要建立的文件夹是否已经存在
Dim strTestFolder,objFSO
strTestFolder = Server.Mappath(strFolder)
Set objFSO = CreateObject("Scripting.FileSystemObject")'检查文件夹是否存在
If not objFSO.FolderExists(strTestFolder) Then'如果不存在则建立文件夹
objFSO.CreateFolder(strTestFolder)
End If
Set objFSO = Nothing
End function



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
'-------------------单页区域--------------------
	dydz=split("index,about",",")
	for i=0 to ubound(dydz)
		host="/"&dydz(i)&"?ranNum="&now()&""
		folder="/" '静态地址文件夹
		html_url_name=""&dydz(i)&".html"
		fso_info host,folder,html_url_name
	next
	
'-------------------列表区域--------------------
	set rs=server.CreateObject("Adodb.recordset")  
	sql="select * from a_class where s_pai=0 and parent_id=0"  
	rs.open sql,conn,1,1
	if not rs.eof then
	do while not rs.eof
		host="/news.asp?id="&rs(0)&"&ranNum="&now()&""
		folder="/news/" '文件夹
		html_url_name=""&rs(0)&".html"
		fso_info host,folder,html_url_name
	rs.movenext
	loop
	end if
	rs.close
	
	
'-------------------分页区域--------------------
	set rs=server.CreateObject("Adodb.recordset")  
	sql="select * from a_main where s_pai=0"   
	rs.open sql,conn,1,1
	if not rs.eof then
		total=rs(0)  
		if total mod 20 = 0 then
		b=tota/20  
		else  
		b=total/20+1  
		end if
		for i=1 to b
			host="/news.asp?page="&i&"&ranNum="&now()&""
			folder="/分页/" '文件夹
			html_url_name=""&i&".html"
			fso_info host,folder,html_url_name
		next

		
	end if
	rs.close
	set rs=nothing
	
	
'-------------------内容区域--------------------
	set rs=server.CreateObject("Adodb.recordset")  
	sql="select * from a_main where s_pai=0"  
	rs.open sql,conn,1,1
	if not rs.eof then
	do while not rs.eof
		host="/news_details.asp?id="&rs(0)&"&ranNum="&now()&""
		folder="/news/" '文件夹
		html_url_name=""&rs(0)&".html"
		fso_info host,folder,html_url_name
	rs.movenext
	loop
	end if
	rs.close
%>
