<%
'函数作用:'简写response.Write
sub w(fcb_str)
	response.write(fcb_str)
end sub
sub wc(fcb_str)
	response.write(fcb_str&vbCrLf)
end sub
sub wn(fcb_str)
	response.write(fcb_str&"<br />")
end sub
sub we(fcb_str)
 response.write(fcb_str):response.End()
end sub
'函数作用:'简写response.End()
sub die()
response.End()
end sub
'函数作用:'三目运算
Function IIf(expr,truepart,falsepart)
	If expr=True Then	IIf=truepart Else IIf=falsepart
End Function
'函数作用:'判断是否是无效
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
'函数作用:'简写response.Write
function sm(fcb_str)
	if isN(fcb_str) then exit function
	sm=server.MapPath(fcb_str)
end function
'函数作用:执行一段js代码
function js(fcb_str)
js="<script>"&fcb_str&"</script>"
end function
'函数作用:弹出对话框提示信息
function msg(fcb_str,fc_base_url)
	if isN(fc_base_url) then
	  we(js("alert('"&fcb_str&"');history.back();"))
	else
	  we(js("alert('"&fcb_str&"');location='"&fc_base_url&"';"))
	end if
end Function
'函数作用:弹出对话框提示信息
function alert(msgstr)
	  w(js("alert('"&msgstr&"');"))
end function
'函数作用:判断是否是数字ID，如果不是赋值
Function IsID(Fc_bacse_Str,Fc_bacse_Num)
	If IsN(Fc_bacse_Str) Or not isNumeric(Fc_bacse_Str) Then
		IsID=Int(Fc_bacse_Num)
	Else
		IsID=Int(Fc_bacse_Str)
	End If
End Function
'函数作用:判断是否是浮点型，如果不是赋
function Isfloat(Fc_bacse_Str,Fc_bacse_Num)
	If IsN(Fc_bacse_Str) Then
		Isfloat=Round(Fc_bacse_Str,2)
	Else
		Isfloat=Round(Fc_bacse_Num,2)
	End If
end function
'函数作用：获取伪静态的参数值
'备注：paras为参数名，如果含有.html的话参数名为html///
'例如地址为content.asp?/html/10/   content.asp?/html/10.html
function rq(paras)
	quest_str=trim(Request.querystring):get_paras=""
	If left(quest_str,1)<>Chr(47) Then quest_str = Chr(47) & quest_str
	if ubound(split(quest_str,Chr(47)))<=1  then
		get_paras=replace(replace(quest_str,Chr(47),""),".html",""):if get_paras<>"" then get_paras=trim(get_paras)
	else
			quest_a=a_noemtpy(quest_str,Chr(47))
			for Fc_bacse_i=0 to ubound(quest_a)
					if quest_a(Fc_bacse_i)=paras then get_paras=quest_a(Fc_bacse_i+1)
			next
	end if
	rq=get_paras
end function
'函数作用:取得当前或者获取的url，并且可以通过参数调整url字符串
Function GetUrl(fcb_url,fcb_str)
	script_name =iif(isn(fcb_url),Request.ServerVariables("SCRIPT_NAME"),fcb_url)
	fcb_url= right(script_name,(len(script_name)-InstrRev(script_name,"/")))
	'wn "fcb_url= "&fcb_url:wn "len(fcb_url)= "&len(fcb_url):wn "instr(fcb_url,""?"")= "&instr(fcb_url,"?")
	if len(fcb_url)=instr(fcb_url,"?") then fcb_url=left(fcb_url,len(fcb_url)-1)
	fcb_name= left(script_name,InstrRev(script_name,"?"))
	fcb_path= left(script_name,InstrRev(script_name,"/"))
	fcb_string=right(script_name,(len(script_name)-InstrRev(script_name,"?")))
	If fcb_str="0" or isn(fcb_str) Then : GetUrl = fcb_url : Exit Function
	If fcb_str="1" Then : GetUrl = script_name&"?"&fcb_string : Exit Function
	If fcb_str="2" Then : GetUrl = fcb_path : Exit Function
	fcb_out =iif(InStr(fcb_str,":")>0,Mid(fcb_str,2),fcb_str) 
	If Not IsN(fcb_string) Then
		fcb_temp="":fcb_i=0:fcb_out=","&fcb_out&",":fcb_arr=split(fcb_string,"&")
		if not isn(fcb_arr) then
			For fcb_j=0 to ubound(fcb_arr)
			  if not isn(fcb_arr(fcb_j)) then
					fcb_arr1=split(fcb_arr(fcb_j),"=")
					If InStr(fcb_out,"-")>0 Then
						 if InStr(fcb_out,",-"&fcb_arr1(0)&",")>0 then fcb_get=0 else fcb_get=1
					else
						 if InStr(fcb_out,","&fcb_arr1(0)&",")>0 then fcb_get=1 else fcb_get=0
					end if
					if fcb_get=1 then
						If fcb_i<>0 Then fcb_temp = fcb_temp & "&"
						fcb_temp = fcb_temp&fcb_arr1(0)&"="&fcb_arr1(1)
						fcb_i = fcb_i + 1
					End If
				end if
			Next
		end if
		  fcb_url=fcb_name&iif(instr(fcb_name,"?")>0,"&","?")&fcb_temp
	End If
	fcb_url=replace(fcb_url,"?&","?"):GetUrl = fcb_url
End Function

Function GetUrlWith(fcb_url,fcb_str,fcb_pv)
	fcb_fnurl = GetUrl(fcb_url,fcb_str)
	fcb_fnurl1 = GetUrl(fcb_url,0)
	GetUrlWith = fcb_fnurl&iif(instr(fcb_fnurl,"?")>0,"&","?")& fcb_pv
End Function



'--------------------------------------- 19
'函数名：ImgWriter
'功  能：给上传的图片加自己的水印
'参  数：SaveImgPath 图片路径和图片名
'用  法：ImgWriter("images\q18.jpg") 里面的参数就传你上传成功或者数据库读取的路径
'---------------------------------------
Sub ImgWriter(SaveImgPath)    '添加水印
  dim jpeg,Text
  Text = "原图版权.你的网站名"       ''水印字符串 
  Set Jpeg = Server.CreateObject("Persits.Jpeg") ''需要安装aspjpeg组建
      Jpeg.Open Server.MapPath(SaveImgPath)
  dim x,y  '水印在右下角位置
  x = Jpeg.Width      ''参数需根据水印字符串长度调整
  y = Jpeg.Height     ''水印位置Y高 参数根据水印字体大小调整
  Jpeg.Canvas.Font.Color = &HFFFFFF ''水印字体颜色
  Jpeg.Canvas.Font.Family = "宋体"  ''水印字体 如“宋体”
  Jpeg.Canvas.Font.Size = 13        ''水印字体大小 数值如:18
  Jpeg.Canvas.Font.ShadowColor = &H111111  ''水印背景颜色
  Jpeg.Canvas.Font.ShadowXoffset = 1 ''背景色X坐标偏移像素值
  Jpeg.Canvas.Font.ShadowYoffset = 1 ''背景色Y坐标偏移像素值
  Jpeg.Canvas.Font.Bold = false      ''True=粗体 False=正常
  Jpeg.Canvas.Print x,y,Text         ''文字水印位置 x，y
  Jpeg.Save Server.MapPath(SaveImgPath) ''保存加水印后的图片
  Jpeg.close
  set Jpeg=nothing
End Sub 

'函数作用:JMAIL发送邮件'括号里面分别是：发送邮件服务器，邮件接收人，发送人，登录邮箱的用户名，登录邮箱的密码，邮件主题，邮件内容

function Upload_Init()
 Upload_Class_path="../../inc/upload/"'设置上传程序的路径
 Upload_Save_path="../../Uploadfiles"'设置文件保存路径
 dim s:s=""
 s =s & "<script type=""text/javascript"" src="""&Upload_Class_path&"js/AnPlus.js""></script>"&vbcr
 s =s & "<script type=""text/javascript"" src="""&Upload_Class_path&"js/AjaxUploader.js""></script>"&vbcr
 s =s & "<div id=""uploadContenter"" style=""background-color:#eeeeee;position:absolute;border:1px #555555 solid;padding:3px;""></div>"&vbcr
 s =s & "<iframe style=""display:none;"" name=""AnUploader""></iframe>"&vbcr
 s =s & "<script type=""text/javascript"">"&vbcr
 s =s & "	var AjaxUp=new AjaxProcesser(""uploadContenter"");"&vbcr
 s =s & "	AjaxUp.target=""AnUploader"";  "&vbcr
 s =s & "	AjaxUp.url="""&Upload_Class_path&"upload.asp""; "&vbcr
 s =s & "	AjaxUp.savePath="""&Upload_Save_path&""";"&vbcr
 s =s & "	var contenter=document.getElementById(""uploadContenter"");"&vbcr
 s =s & "	contenter.style.display=""none""; //隐藏容器"&vbcr
 s =s & "	"&vbcr
 s =s & "	function showUploader(objID,srcElement){"&vbcr
 s =s & "		AjaxUp.reset();  "&vbcr
 s =s & "		contenter.style.display=""block""; "&vbcr
 s =s & "		var ps=_.abs(srcElement);"&vbcr
 s =s & "		contenter.style.top=(ps.y + 20) + ""px"";  "&vbcr
 s =s & "		contenter.style.left=ps.x + ""px"";"&vbcr
 s =s & "		AjaxUp.succeed=function(files){"&vbcr
 s =s & "		    var fujian=document.getElementById(objID);"&vbcr
 s =s & "			fujian.value=AjaxUp.savePath + ""/"" + files[0].NewName;"&vbcr
 s =s & "			contenter.style.display=""none"";"&vbcr
 s =s & "			alert(""文件上传成功!\n\n文件大小:"" + files[0].size + ""字节\n原文件名:"" + files[0].LocalName + ""\n新文件名:"" + files[0].NewName + """");"&vbcr
 s =s & "		}"&vbcr
 s =s & "		AjaxUp.faild=function(msg){alert(""失败原因:"" + msg);contenter.style.display=""none"";}"&vbcr
 s =s & "	}"&vbcr
 s =s & "</script>"&vbcr
 Upload_Init=s
end function



'函数作用:获取文件的大小

Function getFileSize(FileName)
    '判断文件名是不是为空
 if FileName="" then
  getFileSize="0KB"
  Exit Function
 end if
 Dim oFso,oFile,sFile
 sFile=FileName
 Set oFso=Server.CreateObject("Scripting.FileSystemObject")
  '判断获取文件大小的文件是否存在
 If oFso.FileExists(Server.MapPath(sFile)) Then  
 Set oFile=oFso.GetFile(Server.MapPath(sFile))
  '判断获取文件大小
 getFileSize= CStr( CDbl( FormatNumber( oFile.Size / 1024))) & "KB"
 else
     getFileSize="0KB"
  Exit Function
 end if
 Set oFile=nothing
 Set oFso=nothing
End Function


'函数作用:JMAIL发送邮件'括号里面分别是：发送邮件服务器，邮件接收人，发送人，登录邮箱的用户名，登录邮箱的密码，邮件主题，邮件内容

function Upload_input(Input_name,Input_value)
 dim s:s=""
 s =s & "<input type=""text"" name="""&Input_name&""" id="""&Input_name&"""  readonly=""true"" size=""30""  class=""inputkkys"" value="""&Input_value&""" />"&vbcr
 s =s & "<input type=""button"" value=""上传附件"" onclick=""showUploader('"&Input_name&"',this);""  class=""inputkkys"" />"&vbcr
 Upload_input=s
end function

'函数作用:JMAIL发送邮件'括号里面分别是：发送邮件服务器，邮件接收人，发送人，登录邮箱的用户名，登录邮箱的密码，邮件主题，邮件内容

function Jmail_sendmail(jm_smtp,jm_sendto,jm_from,jm_user,jm_pwd,jm_subject,jm_body)
	Jmail_sendmail=true
	Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象
	jmail.silent = FALSE '屏蔽例外错误，返回FALSE跟TRUE两值j
	jmail.logging = true '启用邮件日志
	jmail.Charset = "GB2312" '邮件的文字编码为国标
	jmail.ContentType = "text/html" '邮件的格式为HTML格式
	jmail.AddRecipient jm_sendto '邮件收件人的地址
	jmail.From = jm_from '发件人的E-MAIL地址
	jmail.MailServerUserName = jm_user '登录邮件服务器所需的用户名
	jmail.MailServerPassword = jm_pwd '登录邮件服务器所需的密码
	jmail.Subject = jm_subject '邮件的标题
	jmail.Body = jm_body '邮件的内容
	jmail.Priority = 3 '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
	'jmail.Send(jm_smtp) '执行邮件发送（通过邮件服务器地址）
			if jmail.send(jm_smtp)=false then
					Jmail_sendmail=false
					jmail.close
			end if
	jmail.Close
end function


function sendmail(smtp,sendto,from,user,pwd,subject,body)
	sendmail=true
	Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象
	jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值j
	jmail.logging = true '启用邮件日志
	jmail.Charset = "GB2312" '邮件的文字编码为国标
	jmail.ContentType = "text/html" '邮件的格式为HTML格式
	jmail.AddRecipient sendto '邮件收件人的地址
	jmail.From = from '发件人的E-MAIL地址
	jmail.MailServerUserName = user '登录邮件服务器所需的用户名
	jmail.MailServerPassword = pwd '登录邮件服务器所需的密码
	jmail.Subject = subject '邮件的标题
	jmail.Body = body '邮件的内容
	jmail.Priority = 3 '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
	'jmail.Send(smtp) '执行邮件发送（通过邮件服务器地址）
			if jmail.send(smtp)=false then
					sendmail=false
					msg "发送邮件失败！\n请确认你输入的邮件地址是正确的！",""
					jmail.close
					response.End
			end if
	jmail.Close
end function


'函数作用://1//数字  //2//英文  //3//数字英文及字符  //4//英文/数字级字符-  //5//Email地址  //6//http,https,ftp地址  //7//日期时间验证      //8//IP地址验证   //9//http,https,ftp开头的图片路径,不支持中文  //10//整数,带格式化,如:100,000

Function CheckData(ChkStr,ChkType)
	Dim Pattern,RegEx
	Set RegEx=New RegExp
	Select Case Cstr(ChkType)
		Case "1"  : Pattern = "^\d+$" ' 数字
		Case "2"  : Pattern = "^[A-Za-z]+$" ' 英文
		Case "3"  : Pattern = "^[a-zA-Z0-9\,\/\-\_\[\]]+$" ' 英文/数字及字符[,/-_]
		Case "4"  : Pattern = "^[A-Za-z0-9\_\-]+$" ' 英文/数字级字符-_
		Case "5"  : Pattern = "^[A-Za-z0-9]+$" ' 由数字和26个英文字母组成的字符串
		Case "6"  : Pattern = "^\w+((-\w+)|(.\w+))*@[A-Za-z0-9]+((.|-)[A-Za-z0-9]+)*.[A-Za-z0-9]+$" ' Email地址
		Case "7"  : Pattern = "^(http|https|ftp):(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\=\?\+\-~`@\':!%#]|(&)|&)+" ' http,https,ftp地址
		Case "8"  : Pattern = "^((((1[6-9]|[2-9]\d)\d{2})-(0?[13578]|1[02])-(0?[1-9]|[12]\d|3[01]))|(((1[6-9]|[2-9]\d)\d{2})-(0?[13456789]|1[012])-(0?[1-9]|[12]\d|30))|(((1[6-9]|[2-9]\d)\d{2})-0?2-(0?[1-9]|1\d|2[0-8]))|(((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))-0?2-29-)) (20|21|22|23|[0-1]?\d):[0-5]?\d:[0-5]?\d$" ' 日期时间验证
		Case "9"  : Pattern = "^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$" ' IP地址验证
		Case "10" : Pattern = "^((http|https|ftp):(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\=\?\+\-~`@\':!%#]|(&)|&)+|\/([\w\/\\\.\=\?\+\-~`@\':!%#]|(&)|&)+)\.(jpeg|jpg|gif|png|bmp)$" ' http,https,ftp开头的图片路径,不支持中文
		Case "11" : Pattern = "^\w+\.(\w){1,30}$" ' 文件名格式,不支持中文
		Case "12" : Pattern = "^[0-9\,\.]+$" ' 整数,带格式化,如:100,000
		Case Else Pattern = ValidType
	End Select
	RegEx.Pattern= Pattern
	CheckData = RegEx.Test(Trim(ChkStr))
	Set RegEx = Nothing
End Function

Function CheckObjStr(str)
on error resume next
	Dim I1,I2
	Set I2 = Server.CreateObject(str)
	IF I1 = 0 OR I1 = -2147221477 Then
		CheckObjStr=true
	ElseIF I1 = 1 OR I1 = -2147221005 Then
		CheckObjStr=false
	End IF
	Set I2 = Nothing
End Function


Function ClearAllHTML(strHTML)   
    if strHTMl="" or isnull(strHTML) then    
    exit Function  
    end if   
    StrHtml = Replace(StrHtml,vbCrLf,"")   
    StrHtml = Replace(StrHtml,Chr(13)&Chr(10),"")   
    StrHtml = Replace(StrHtml,Chr(13),"")   
    StrHtml = Replace(StrHtml,Chr(10),"")   
    StrHtml = Replace(StrHtml," ","")   
    StrHtml = Replace(StrHtml,"    ","")   
     Dim objRegExp, Match, Matches    
     Set objRegExp = New Regexp   
     objRegExp.IgnoreCase = True  
     objRegExp.Global = True   
     objRegExp.Pattern = "<style(.+?)/style>"  
     Set Matches = objRegExp.Execute(strHTML)    
     For Each Match in Matches    
     strHtml=Replace(strHTML,Match.Value,"")   
     Next  
     objRegExp.Pattern = "<script(.+?)/script>"   
     Set Matches = objRegExp.Execute(strHTML)   
     For Each Match in Matches    
     strHtml=Replace(strHTML,Match.Value,"")   
     Next    
     objRegExp.Pattern = "<.+?>"  
     Set Matches = objRegExp.Execute(strHTML)     
     For Each Match in Matches    
     strHtml=Replace(strHTML,Match.Value,"")   
     Next  
     ClearAllHTML=strHTML   
     Set objRegExp = Nothing  
End Function


Function ClearHtml(str)
Dim re
str=replace(str,"<br>","{br}")
str=replace(str,"</p>","{p}")
Set re = new RegExp
re.IgnoreCase = True
re.Global = True
re.Pattern = "<(.[^>]*)>"
str = re.Replace(str,"")
set re = Nothing
str=Replace(str,chr(10),"")
str=Replace(str,chr(13),"")
str=Replace(str,"　","")
str=Replace(str," ","") 
str=Replace(str,"&nbsp;","")

str=replace(str,"</p>","{p}")
str=replace(str,"{br}","<br>")
ClearHtml = str
End Function

 
Function cutbadchar(str)  
badstr="不|文|明|字|符|列|表|格|式"  
badword=split(badstr,"|")  
For i=0 to Ubound(badword)  
If instr(str,badword(i)) > 0 then  
str=Replace(str,badword(i),"***")  
End If  
Next  
cutbadchar=str  
End Function  

function Checkstr(Str) 
If Isnull(Str) Then 
CheckStr = "" 
Exit function 
End If 
Str = Replace(Str,Chr(0),"", 1, -1, 1) 
Str = Replace(Str, """", "&quot;", 1, -1, 1) 
Str = Replace(Str,"<","&lt;", 1, -1, 1) 
Str = Replace(Str,">","&gt;", 1, -1, 1) 
Str = Replace(Str, "script", "&#115;cript", 1, -1, 0) 
Str = Replace(Str, "SCRIPT", "&#083;CRIPT", 1, -1, 0) 
Str = Replace(Str, "Script", "&#083;cript", 1, -1, 0) 
Str = Replace(Str, "script", "&#083;cript", 1, -1, 1) 
Str = Replace(Str, "object", "&#111;bject", 1, -1, 0) 
Str = Replace(Str, "OBJECT", "&#079;BJECT", 1, -1, 0) 
Str = Replace(Str, "Object", "&#079;bject", 1, -1, 0) 
Str = Replace(Str, "object", "&#079;bject", 1, -1, 1) 
Str = Replace(Str, "applet", "&#097;pplet", 1, -1, 0) 
Str = Replace(Str, "APPLET", "&#065;PPLET", 1, -1, 0) 
Str = Replace(Str, "Applet", "&#065;pplet", 1, -1, 0) 
Str = Replace(Str, "applet", "&#065;pplet", 1, -1, 1) 
Str = Replace(Str, "[", "&#091;") 
Str = Replace(Str, "]", "&#093;") 
Str = Replace(Str, """", "", 1, -1, 1) 
Str = Replace(Str, "=", "&#061;", 1, -1, 1) 
Str = Replace(Str, "'", "''", 1, -1, 1) 
Str = Replace(Str, "select", "sel&#101;ct", 1, -1, 1) 
Str = Replace(Str, "execute", "&#101xecute", 1, -1, 1) 
Str = Replace(Str, "exec", "&#101xec", 1, -1, 1) 
Str = Replace(Str, "join", "jo&#105;n", 1, -1, 1) 
Str = Replace(Str, "union", "un&#105;on", 1, -1, 1) 
Str = Replace(Str, "where", "wh&#101;re", 1, -1, 1) 
Str = Replace(Str, "insert", "ins&#101;rt", 1, -1, 1) 
Str = Replace(Str, "delete", "del&#101;te", 1, -1, 1) 
Str = Replace(Str, "update", "up&#100;ate", 1, -1, 1) 
Str = Replace(Str, "like", "lik&#101;", 1, -1, 1) 
Str = Replace(Str, "drop", "dro&#112;", 1, -1, 1) 
Str = Replace(Str, "create", "cr&#101;ate", 1, -1, 1) 
Str = Replace(Str, "rename", "ren&#097;me", 1, -1, 1) 
Str = Replace(Str, "count", "co&#117;nt", 1, -1, 1) 
Str = Replace(Str, "chr", "c&#104;r", 1, -1, 1) 
Str = Replace(Str, "mid", "m&#105;d", 1, -1, 1) 
Str = Replace(Str, "truncate", "trunc&#097;te", 1, -1, 1) 
Str = Replace(Str, "nchar", "nch&#097;r", 1, -1, 1) 
Str = Replace(Str, "char", "ch&#097;r", 1, -1, 1) 
Str = Replace(Str, "alter", "alt&#101;r", 1, -1, 1) 
Str = Replace(Str, "cast", "ca&#115;t", 1, -1, 1) 
Str = Replace(Str, "exists", "e&#120;ists", 1, -1, 1) 
Str = Replace(Str,Chr(13),"<br>", 1, -1, 1) 
CheckStr = Replace(Str,"'","''", 1, -1, 1) 
End function



Function FilterHTML(str)
    Dim re,cutStr
    Set re=new RegExp
    re.IgnoreCase =True
    re.Global=True
    re.Pattern="<(.[^>]*)>"
    str=re.Replace(str,"")    
    set re=Nothing
    str=Replace(str," ","")
    str=Replace(str,chr(10),"")
    str=Replace(str,chr(13),"")
    str=Replace(str," ","")
    str=Replace(str,"　","") 
    Dim l,t,c,i
    l=Len(str)
    t=0
    For i=1 to l
        c=Abs(Asc(Mid(str,i,1)))
        If c>255 Then
            t=t+2
        Else
            t=t+1
        End If
        cutStr=str
    Next
    
    FilterHTML= cutStr
End Function



'把长的数字用逗号隔开显示
'　　文字格式：12345678
'　　格式化数字：12,345,678
'　　自定义函数：


Function Comma(str) 
If Not(IsNumeric(str)) Or str = 0 Then 
Result = 0 
ElseIf Len(Fix(str)) < 4 Then 
Result = str 
Else 
Pos = Instr(1,str,".") 
If Pos > 0 Then 
Dec = Mid(str,Pos) 
End if 
Res = StrReverse(Fix(str)) 
LoopCount = 1 
While LoopCount <= Len(Res) 
TempResult = TempResult + Mid(Res,LoopCount,3) 
LoopCount = LoopCount + 3 
If LoopCount <= Len(Res) Then 
TempResult = TempResult + "," 
End If 
Wend 
Result = StrReverse(TempResult) + Dec 
End If 
Comma = Result 
End Function 


%>