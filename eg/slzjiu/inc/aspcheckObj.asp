<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<%
Dim theInstalledObjects(17)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<link rel="stylesheet" type="text/css" href="../images/cssyullhao.css">
</head>

<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#F6FBFF">
<table cellpadding=3 cellspacing=1 border=0 width="95%" align=center>
<tr>
<td width="100%" valign=top>
<p><b>欢迎光临 <a href="<%= Session("Http")%>" target="_blank"><%= Session("Company")%></a>--管理系统</b><br><BR></p>
在这里，您可以控制你所有的管理设置。请在此页的左侧选择您要进行管理的链接。<br><br>
如果你退出系统时，为了系统安全请按页面左上角的“
<a target="_top" href="LOGOUT.ASP?target=exit" onClick="return confirm('是否退出管理？');">退出管理</a>”。<br/>



<p style="color:#F00;">尊敬的后台管理员，您好，如果您在发布信息的时候，要复制粘贴word，Excel文件，请安装格式导入插件！<a href="../Ewebeditor/eWebEditorClientInstall.rar" style="color:#F00; font-weight:bold; font-size:14px;">『下载』</a></p>



</td>
</tr>
</table><br>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
  <th class="tableHeaderText" colspan=2 height=25>网站主机信息统计</th>
<tr>
<tr>
<td width="50%"  class="forumRow" height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRowHighlight">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>站点物理路径：<%
dim path
path=Trim(request.ServerVariables("APPL_PHYSICAL_PATH"))
Response.Write(path)%></td>
<td width="50%" class="forumRowHighlight">数据库地址：<%= Replace(dbimages,path,"")%></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>FSO文本读写：<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="#FF0000"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%" class="forumRowHighlight">数据库使用：<%If Not IsObjInstalled(theInstalledObjects(10)) Then%><font color="#FF0000"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>Jmail组件支持：<%If Not IsObjInstalled(theInstalledObjects(13)) Then%><font color="#FF0000"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%" class="forumRowHighlight">CDONTS组件支持：<%If Not IsObjInstalled(theInstalledObjects(14)) Then%><font color="#FF0000"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
<tr valign="middle">
<td width="50%" height=23 colspan="2" align="center" class="forumRow"><table width="81%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
    <td align="right" valign="middle">
	<!--<button style="height:30; width:100; cursor:hand; " class="button" onClick="parent.location.href='aspcheck.asp';">查看详细信息</button>-->
	</td>
  </tr>
</table></td>
</tr>
</table><br>

</body>
</html>
<%
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>