<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<title>网站管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0" CHARSET="GB2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="<%=Color_0%>">
<%
dim action
dim admin_flag
dim admin_db
action=trim(request("action"))
admin_db = "../"&DbAccessConnection
backup_path="../../database/backup/"

dim dbpath,bkfolder,bkdbname,fso,fso1

select case action
case "CompressData"		'压缩数据
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,username,dateandtime,body
		call CompressData()	

case "BackupData"		'备份数据
		if request("act")="Backup" then
		call updata()
		else
		call BackupData()
		end if

case "RestoreData"		'恢复数据
	dim backpath
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
			response.write "请输入您要恢复成的数据库全名"	
			else
			Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
		
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
			response.write "成功恢复数据！"
			else
			response.write "备份目录下并无您的备份文件！"	
			end if
		else
		call RestoreData()
		end if
case else
 call CompressData()	
 call BackupData()
 call RestoreData()
end select

conn.close
set conn=nothing
response.write"</body></html>"

'====================恢复数据库=========================
sub RestoreData()
%>
<br>
<table border="1"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">		<tr>
	<th height=25 >
   					&nbsp;&nbsp;<B class="fc_Red">恢复站点数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
				<form method="post" action="admin_data.asp?action=RestoreData&act=Restore" onSubmit="return confirm('此操作将会导致现有的所有数据被替换为旧的数据!!!\n\n备份数据库名称以日期命名，该日期表示数据库里的数据是此日期的数据，请慎重操作!!!!\n')"> 				
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;备份数据库路径(相对)：<input type=text size=40 name=DBpath value="<%=request.Cookies("admin_backdata")%>">
  						&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;目标数据库路径(相对)：<input type=text size=40 name=backpath value="<%=admin_db%>"><BR>
  						&nbsp;&nbsp;填写您当前使用的数据库路径，如不想覆盖当前文件，可自行命名（注意路径是否正确），然后修改Connection.asp/data.asp文件，如<br>
                        &nbsp;&nbsp;果目标文件名和当前使用数据库名一致的话，不需修改Connection.asp/data.asp文件<BR>
						&nbsp;&nbsp;<input type=submit value="恢复数据" style="CURSOR:hand"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名,本程序的默认备份数据库文件为:backup\<%= admin_db%>,请按照您的备份文件自行修改。<br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

'====================备份数据库=========================
sub BackupData()
%>
<br>
	<table border="1"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;<B>备份站点数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
  				<form method="post" action="admin_data.asp?action=BackupData&act=Backup" onSubmit="return confirm('如非必要不用经常备份数据库，频繁备份数据库会占用网站空间.')">
  				<tr>
  					<td height="100" class="forumrow">
                        &nbsp;&nbsp;
						当前数据库路径(相对路径)：<input type="text" size="40" name="DBpath" value="<%=admin_db%>">
						<BR>&nbsp;&nbsp;
						备份数据库目录(相对路径)：<input type="text" size="40" name="bkfolder" value="<%=backup_path%>">
						&nbsp;如目录不存在，程序将自动创建<BR>&nbsp;&nbsp;
						备份数据库名称(填写名称)：<input type="text" size="40" name="bkDBname" value="<%=backup_path&"back_"&replace(formatdatetime(now,2),"-","")&".mdb"%>">
						&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>
						&nbsp;&nbsp;<input type="submit" value="确  定" style="CURSOR:hand"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认数据库文件为<%=admin_db%><br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径
  					</td>
  				</tr>	
  				</form>
</table>
<%
end sub

sub updata()
		Dbpath=request.form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.form("bkfolder")
		bkdbname=request.form("bkdbname")
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			response.write "备份数据库成功，您备份的数据库路径为" &bkfolder& "\"& bkdbname
			response.Cookies("admin_backdata")= bkdbname
		Else
			response.write "找不到您所需要备份的文件。"
		End if
end sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function

'====================压缩数据库 =========================
sub CompressData()
%>
<br>
<table border="1"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<tr>
<th height=25>
&nbsp;&nbsp;<B>站点数据压缩</B>
</th>
</tr>
<form action="admin_data.asp?action=CompressData" method="post">
<tr>
<td class="forumrow" height=25><b>注意：</b><br>输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作） </td>
</tr>
<tr>
<td class="forumrow">压缩数据库：<input name="dbpath" type="text" value="<%=request.Cookies("admin_backdata")%>" size="40">
&nbsp;
<input type="submit" value="开始压缩" style="CURSOR:hand"></td>
</tr>
<tr>
<td class="forumrow"><input type="checkbox" name="boolIs97" value="True">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br><br></td>
</tr>
</form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
	End If

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "数据库 " & dbpath & ", 已经压缩成功!" & vbCrLf

Else
	CompactDB = "数据库名称或路径不正确. 请重试!" & vbCrLf
End If

End Function
%>