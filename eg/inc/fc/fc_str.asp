<%
'=============================================================================== 
'函数作用:过滤提交参数 
'=============================================================================== 
Function Kill_sql(fcs_str)
	Kill_sql=trim(Request(fcs_str)):Kill_sql = LCase(fcs_str) : Kill_sql = Replace(Kill_sql," ","") : Kill_sql = Replace(Kill_sql,"'","") : Kill_sql = Replace(Kill_sql,"""","") : Kill_sql = Replace(Kill_sql,"=","") : Kill_sql = Replace(Kill_sql,"*","")
End Function
'=============================================================================== 
'函数作用：'剔除所有的HTML标签
'=============================================================================== 
Function RemoveAllHTML( strText )
Dim TAGLIST
TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;" &_
"BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;" &_
"COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;" &_
"FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;I;IFRAME;IMG;" &_
"INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;" &_
"MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;p;PARAM;PLAINTEXT;" &_
"PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;" &_
"TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;UL;VAR;WBR;XMP;"

Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"

Dim nPos1
Dim nPos2
Dim nPos3
Dim strResult
Dim strTagName
Dim bRemove
Dim bSearchForBlock

nPos1 = InStr(strText, "<")
Do While nPos1 > 0
nPos2 = InStr(nPos1 + 1, strText, ">")
If nPos2 > 0 Then
strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ")

nPos3 = InStr(strTagName, " ")
If nPos3 > 0 Then
strTagName = Left(strTagName, nPos3 - 1)
End If

If Left(strTagName, 1) = "/" Then
strTagName = Mid(strTagName, 2)
bSearchForBlock = False
Else
bSearchForBlock = True
End If

If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
bRemove = True
If bSearchForBlock Then
If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
nPos2 = Len(strText)
nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
If nPos3 > 0 Then
nPos3 = InStr(nPos3 + 1, strText, ">")
End If

If nPos3 > 0 Then
nPos2 = nPos3
End If
End If
End If
Else
bRemove = False
End If

If bRemove Then
strResult = strResult & Left(strText, nPos1 - 1)
strText = Mid(strText, nPos2 + 1)
Else
strResult = strResult & Left(strText, nPos1)
strText = Mid(strText, nPos1 + 1)
End If
Else
strResult = strResult & strText
strText = ""
End If

nPos1 = InStr(strText, "<")
Loop
strResult = strResult & strText

RemoveAllHTML = strResult
End Function



Function RemoveHTML( strText )
	Set fcs_RegEx = New RegExp
	fcs_RegEx.Pattern = "<[^>]*>"
	fcs_RegEx.Global = True
	RemoveHTML = fcs_RegEx.Replace(strText, "")
	RemoveHTML = replace(RemoveHTML, "&lsquo;", "‘")
	RemoveHTML = replace(RemoveHTML, "&rsquo;", "’")
	RemoveHTML = replace(RemoveHTML, "&ldquo;", "“")
	RemoveHTML = replace(RemoveHTML, "&rdquo;", "”")
	RemoveHTML = replace(RemoveHTML, chr(34), "")
	RemoveHTML = replace(RemoveHTML, chr(39), "")
End Function

'=============================================================================== 
'函数作用：'剔除所有的HTML标签,
'=============================================================================== 
Function RemoveHTMLCode(ByVal Str,ByVal cHas)
	Str=""&Str:if cHas="" OR Str="" then:RemoveHTMLCode=Cstr(Str):exit function:end if
	Dim i:cHas=split(trimall(UCASE(cHas)),"|")
	for i=0 to UBound(cHas)
		SELECT CASE cHas(i)
			CASE "TABLE":Str=RegReplace(Str,"/<\/?(table|thead|tbody|tr|th|td).*>/ig","")
			CASE "OBJECT":Str=RegReplace(Str,"/<\/?(object|param|embed).*>/ig","")
			CASE "SCR"&"IPT":Str=RegReplace(Str,"/<scr"&"ipt.*>[\w\W]+?<\/scr"&"ipt>/ig",""):Str=RegReplace(Str,"/on[\w]+=[\'\""].+?[\'\""](\s|>)/ig","$1")
			CASE "STYLE":Str=RegReplace(Str,"/<style.*>[\w\W]+?<\/style>/ig","")
			CASE "CLASS":Str=RegReplace(Str,"/\sclass=.+?(\s|>)/ig","")
			CASE "*":Str=RemoveHTMLCode(Str,"SCR"&"IPT|STYLE"):Str=RegReplace(Str,"/<.*?>/ig",""):Str=Replace(Replace(Str,"<","&lt;"),">","&gt;")
			CASE ELSE:Str=RegReplace(Str,"/<\/?"&addcslashes(cHas(i))&".*?>/ig","")
		End SELECT
	next
	RemoveHTMLCode=Replace(Str,"&nbsp;"," ")
End Function



Function HtmlEncode(fcs_str)
	If Not IsN(fcs_str) Then
		fcs_str = Replace(fcs_str, Chr(38), "&#38;")
		fcs_str = Replace(fcs_str, "<", "&lt;")
		fcs_str = Replace(fcs_str, ">", "&gt;")
		fcs_str = Replace(fcs_str, Chr(39), "&#39;")
		fcs_str = Replace(fcs_str, Chr(32), "&nbsp;")
		fcs_str = Replace(fcs_str, Chr(34), "&quot;")
		fcs_str = Replace(fcs_str, Chr(9), "&nbsp;&nbsp; &nbsp;")
		fcs_str = Replace(fcs_str, vbCrLf, "<br />")
	End If
	HtmlEncode = fcs_str
End Function
Function HtmlDecode(fcs_str)
		If Not IsN(fcs_str) Then
		Set fcs_RegEx = New RegExp
		fcs_RegEx.Pattern = "<br\s*/?\s*>"
		fcs_RegEx.Global = True
		fcs_str = fcs_RegEx.Replace(fcs_str, vbCrLf)
		fcs_str = Replace(fcs_str, "&nbsp;&nbsp; &nbsp;", Chr(9))
		fcs_str = Replace(fcs_str, "&quot;", Chr(34))
		fcs_str = Replace(fcs_str, "&nbsp;", Chr(32))
		fcs_str = Replace(fcs_str, "&#39;", Chr(39))
		fcs_str = Replace(fcs_str, "&apos;", Chr(39))
		fcs_str = Replace(fcs_str, "&gt;", ">")
		fcs_str = Replace(fcs_str, "&lt;", "<")
		fcs_str = Replace(fcs_str, "&amp;", Chr(38))
		fcs_str = Replace(fcs_str, "&#38;", Chr(38))
		HtmlDecode = fcs_str
	End If
End Function
'=============================================================================== 
'函数作用：'在新闻标题列表等应用中，只取一定长度的字符，若超过这个长度，则加上... 
'=============================================================================== 
function str_cut(str_sss,cut_len) 
if str_sss="" then str_cut="":exit function
str_sss=replace(replace(replace(replace(str_sss,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
fc_str_l=len(str_sss):fc_str_t=0
for fc_str_i=1 to fc_str_l
	fc_str_c=Abs(Ascw(Mid(str_sss,fc_str_i,1)))
	if fc_str_c>255 then fc_str_t=fc_str_t+2 else fc_str_t=fc_str_t+1
	if fc_str_t>=cut_len then str_cut=left(str_sss,fc_str_i) & "…":exit for else str_cut=str_sss
next
str_cut=replace(replace(replace(replace(str_cut," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function
'函数作用：'在新闻标题列表等应用中，只取一定长度的字符，若超过这个长度，则加上... 
'=============================================================================== 
function str_eweb(str_sss) 
if str_isnull(str_sss) then str_sss = "&nbsp;" 
str_eweb = server.HTMLEncode(str_sss)
end function
'=============================================================================== 
'函数作用：将搜索关键字加亮
'备注：'搜索关键字（词）加亮
'===============================================================================
Function str_skey(fcs_str,fcs_key)
If str_isnull(fcs_str) Then Exit Function
If str_isnull(fcs_key) Then Exit Function
		fcs_str = Cstr(fcs_str):fcs_key = Cstr(fcs_key)
		If Len(Trim(fcs_key))>0 Then
		fcs_str=replace(fcs_str,fcs_key,"<font color='red'><b>"&fcs_key&"</b></font>")
		End If
		str_skey=fcs_str
End Function
'=============================================================================== 
'函数作用：生成随机字符串
'===============================================================================
function Str_rand(intLength)
	fcs_str="":strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"
	seedLength=len(strSeed)
Randomize
	for fc_str_i=1 to intLength
		fcs_str=fcs_str+mid(strSeed,int(seedLength*rnd)+1,1)
	next
Str_rand=fcs_str
end Function
'=============================================================================== 
'函数作用：生成随机字符串
'===============================================================================
function Str_randtime(fcs_str,intlength)
 ranNum=""
 randomize
 for fc_str_i=1 to intlength
 ranNum=ranNum&int(9*rnd)+1
 next
times=replace(now(),":",""):times=replace(times,"-",""):times=replace(times," ","")
Str_randtime=fcs_str&times&cint(ranNum)
end Function
'=============================================================================== 
'函数作用：格式化日期格式
'备注：'str_time(2006-05-20,".")输出为(2006.05.20)
'===============================================================================
Function str_date(fcs_str,char)
	If str_isnull(fcs_str) Then Exit Function
  syear=year(fcs_str):smonth=right("0"&month(fcs_str),2):sday=right("0"&day(fcs_str),2)
	str_date=syear&char&smonth&char&sday
End Function

Function str_date1(fcs_str,char)
	If str_isnull(fcs_str) Then Exit Function
  syear=year(fcs_str):smonth=right("0"&month(fcs_str),2):sday=right("0"&day(fcs_str),2)
	str_date1=smonth&char&sday
End Function
'=============================================================================== 
'函数作用：判断是否是无效的字符串
'备注：'str_time(2006-05-20,".")输出为(2006.05.20)
'===============================================================================
Function str_isnull(fcs_str)
	str_isnull=False
	If IsArray(fcs_str) Then Exit Function
	If fcs_str="" or IsNull(fcs_str) or IsEmpty(fcs_str) Then str_isnull=True
End Function
'=============================================================================== 
'函数作用：判断字符串sstr中是否含有fcs_key，返回布尔值
'备注：'str_time(2006-05-20,".")输出为(2006.05.20)
'===============================================================================
function a_instr(fcs_str,fcs_key)
 a_instr=false:fcs_arr=split(fcs_str,",")
	for fc_str_i=0 to ubound(fcs_arr)
	  if cint(fcs_arr(fc_str_i))=fcs_key then a_instr=true
	next
end function
'=============================================================================== 
'函数作用：清除函数中的空值
'备注：传入需要分割为数组的字符串和分隔符，返回无空值的数组
'===============================================================================
function a_noemtpy(a_str,split_str)
	new_a_str="":a_str=replace(a_str,".html","")
	old_a=split(a_str,split_str)
			for fc_str_i=0 to ubound(old_a)
					if old_a(fc_str_i)<>"" then new_a_str=new_a_str&split_str&old_a(fc_str_i)
			next
	new_a_str = Mid(new_a_str,2)
	a_noemtpy=split(new_a_str,split_str)
end function  

'=============================================================================== 
'函数作用：正则替换函数
'备注：oldStr表示要替换的字符串,patrn表示要替换的字符，replStr表示用什么来替换
'str = "aaa/hheeh"

'Response.write(ReplaceStr(str,"hh","##"))
'返回的就是aaa/##eeh
'
'===============================================================================

Function ReplaceSTR(oldStr,patrn, replStr)
   Dim regEx, str1             ' 建立变量。
   str1 = oldstr
   Set regEx = New RegExp             ' 建立正则表达式。
   regEx.Global = True              '设置为全局
   regEx.Pattern = patrn             ' 设置模式。
   regEx.IgnoreCase = true             ' 设置是否区分大小写。
   ReplaceStr= regEx.Replace(str1, replStr)       ' 作替换。
   set regEx = nothing
End Function



'-----------添加成功转向Url-----------
Sub AddSuccess(Url)
    Response.Write "<html>"
    Response.Write "<head>"
    Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
    Response.Write "<title>"&SiteName&"</title>"
    Response.Write "<meta http-equiv=""refresh"" content=""1;URL="&Url&""">"
    Response.Write "</head>"
    Response.Write "<body>"
    Response.Write "<p style=""color:red;font-size:14px;"">添加成功, 等待返回... </p>"
    Response.Write "</body>"
    Response.Write "</html>"
	Response.End()
End Sub
'-----------修改成功转向Url-----------
Sub EditSuccess(Url)
    Response.Write "<html>"
    Response.Write "<head>"
    Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
    Response.Write "<title>"&SiteName&"</title>"


    Response.Write "<meta http-equiv=""refresh"" content=""1;URL="&Url&""">"
    Response.Write "</head>"
    Response.Write "<body>"
    Response.Write "<p style=""color:red;font-size:14px;"">更新成功, 等待返回... </p>"
    Response.Write "</body>"
    Response.Write "</html>"
	Response.End()
End Sub
'-----------删除成功转向Url-----------
Sub DelSuccess(Url)
    Response.Write "<html>"
    Response.Write "<head>"
    Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
    Response.Write "<title>"&SiteName&"</title>"
	
    Response.Write "<meta http-equiv=""refresh"" content=""1;URL="&Url&""">"
    Response.Write "</head>"
    Response.Write "<body>"
    Response.Write "<p style=""color:red;font-size:14px;"">删除成功, 等待返回... </p>"
    Response.Write "</body>"
    Response.Write "</html>"
	Response.End()
End Sub

'-----------图片的展示-----------


'参数：Title 图片的名称  path 图片路径  ID 图片ID Typ 图片类型
Public Function CheckImg(Title,Path,ID,Typ)
Select Case Typ
Case "-1"
Typtxt="小图"
Case "1"
Typtxt="大图"
Case "0"
Typtxt="图片"
End Select
FSO="Scripting.FileSystemObject"
set objfso=server.CreateObject(Fso)
if objfso.fileexists(server.MapPath("../../"&Path)) then	
Str="<a href=../../"& Path &" rel=gb_imageset["& ID &"] title="& Title & Typtxt &"><img src=../../"& Path &" height=""60"" border='0' alt=点击查看"& Typtxt &"></a>"
else
Str=Typtxt&"不存在"
end if
if Path="" then
Str="暂无"&Typtxt
end if
CheckImg=Str
End Function
%>