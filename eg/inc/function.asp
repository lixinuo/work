<%
Function RemoveHTML( strText )
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

RemoveHTML = strResult
End Function
'============================================================================================================================
'函数作用:'以下定义出错提示 
'传人参数： 
'msg: 显示的错误信息
'tourl:显示错误后要跳转的页面
'返回值：
'备注：
'============================================================================================================================  
function w(str)
response.write(str)
end function
'============================================================================================================================
'函数作用:'以下定义出错提示 
'传人参数： 
'msg: 显示的错误信息
'tourl:显示错误后要跳转的页面
'返回值：
'备注：
'============================================================================================================================  
function msg(msgstr,tourl)
if tourl="" then
response.write("<script>alert('"&msgstr&"');history.back();</script>")
response.end()
else
response.write("<script>alert('"&msgstr&"');location='"&tourl&"';</script>")
response.end()
end if
end function
'============================================================================================================================ 
'函数作用：'在新闻标题列表等应用中，只取一定长度的字符，若超过这个长度，则加上... 
'传人参数： 
'title: 文章标题
'content:文章内容
'length:截取长度
'返回值：
'备注：默认截取长度是8 
'============================================================================================================================ 
function GetTitle(title,length) 
If length = 0 Then length = 8 
if title = "" or isnull(title) then title = "没有内容" 

If Len(title) > length Then 
GetTitle = Left(title, length) & ".." 
Else 
GetTitle = title 
End If 
end function
'============================================================================================================================ 
'函数作用：'在新闻标题列表等应用中，只取一定长度的字符，若超过这个长度，则加上... 
'传人参数： 
'title: 文章标题
'content:文章内容
'length:截取长度
'返回值：
'备注：默认截取长度是8 
'============================================================================================================================ 
function mysplit(spstr,splitstr,arrid) 
if arrid = "" or isnull(arrid) then arrid = 0 
if splitstr = "" or isnull(splitstr) then splitstr = "|" 
arr_spstr=split(spstr,splitstr)
if ubound(arr_spstr)<1 then
mysplit=spstr
else
if ubound(arr_spstr)<arrid then
mysplit=spstr
else
mysplit=trim(arr_spstr(arrid))
end if
end if 
end function
'============================================================================================================================ 
'函数作用：'在新闻标题列表等应用中，只取一定长度的字符，若超过这个长度，则加上... 
'传人参数： 
'title: 文章标题
'content:文章内容
'length:截取长度
'返回值：
'备注：默认截取长度是8 
'============================================================================================================================ 
function Rt(sqlstr) 
set rs_rt=server.CreateObject("adodb.recordset")
rs_rt.Open sqlstr,conn,1,1
if not rs_rt.eof then
Classname=trim(rs_rt(0))
Rt = Classname
else
Rt = "暂无内容"
end if
rs_rt.close
set rs_rt=nothing
end function

function Rta(sqlstr) 
dotnum=instr(sqlstr,",")
set rs_rt=server.CreateObject("adodb.recordset")
rs_rt.Open sqlstr,conn,1,1
if not rs_rt.eof then
rsnum=rs_rt.recordcount
Classname=rs_rt.getrows(rsnum)
Rta = Classname
else
Rta(0,0)="暂无内容"
end if
rs_rt.close
set rs_rt=nothing
end function
'============================================================================================================================
'函数作用：'调用数据列表
'传人参数： 
'newid:新闻分类
'lcount:循环显示多少条
'titlecounts:标题显示多少字数
'isen:判断中英文
'classstyle:链接样式
'返回值：
'备注：call getlist("select top 5 id,s_name"&language&",s_time from news_list_content where newid=10 order by s_time desc,id desc",15,15,25,"images/tu_03.jpg","","black12")
'============================================================================================================================ 
Function Getlist(sqlstr,ulclass,tlength,istime)
dim list_text
list_text=""
set RsList=server.createobject("adodb.recordset")
RsList.open sqlstr,conn,1,1
if not(RsList.eof or RsList.bof) then
list_text=list_text&"<ul class='"&ulclass&"'>"
 do while not RsList.eof  
  list_text=list_text&"<li>"
	list_text=list_text&"<a href='news_detail.asp?id="&RsList(0)&"' title='标题"&RsList(1)&"'>"&GetTitle(RsList(1),tlength)&"</a>"
	if istime=1 then list_text=list_text&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style='color:red;'>("&formatdatetime(RsList(2),2)&")</span>"
	list_text=list_text&"</li>"
  if RsList.eof then exit do
  RsList.movenext
 loop
list_text=list_text&"</ul>"
else
list_text=list_text&"<ul class='"&ulclass&"'><li>暂无内容！</li></ul>"
end if

Getlist=list_text
RsList.close
set RsList=nothing
End Function
'============================================================================================================================
'函数作用：'调用数据列表
'传人参数： 
'newid:新闻分类
'lcount:循环显示多少条
'titlecounts:标题显示多少字数
'isen:判断中英文
'classstyle:链接样式
'返回值：
'备注：call Get_link(sqlstr,clength,tdnum,cstyle)
'============================================================================================================================ 
Function Get_link(sqlstr,clength,tdnum,cstyle)
dim list_text,i,tdh
list_text=""
tdh=25
tdsw=100/tdnum
tdsw="'"&tdsw&"%'"
set RsList=server.createobject("adodb.recordset")
RsList.open sqlstr,conn,1,1
if not(RsList.eof or RsList.bof) then
list_text=list_text&"<table width=""100%"" border='0' cellspacing='0' cellpadding='0'><tr>"
i=1
do while not RsList.eof  
list_text=list_text&"<td width="&tdsw&" height='"&tdh&"'>"
list_text=list_text&"<a href="&RsList(0)&" class="&cstyle&" title="&RsList(2)&">"&GetTitle(RsList(1),clength)&"</a>"
list_text=list_text&"</td>"

 if RsList.eof then exit do
 if i mod tdnum= 0 then list_text=list_text&"</tr><tr>"
 RsList.movenext
 
if RsList.eof then
  if i>tdnum then
   t=i mod tdnum
  else
   t=tdnum-i
  end if
 for j=tdnum to t
  list_text=list_text&"<td width="&tdsw&">&nbsp;</td>"
 next
end if
 
i=i+1
loop
list_text=list_text&"</tr></table>"
else
list_text=list_text&"<table><tr><td align='center' height=40><span class='"&cstyle&"'>暂无内容！</span></td></tr></table>"
 
end if

Get_link=list_text
RsList.close
set RsList=nothing
End Function
'============================================================================================================================
'函数作用：'调用数据列表
'传人参数： 
'newid:新闻分类
'lcount:循环显示多少条
'titlecounts:标题显示多少字数
'isen:判断中英文
'classstyle:链接样式
'返回值：
'备注：call getlist("select top 5 id,s_name"&language&",s_time from news_list_content where newid=10 order by s_time desc,id desc",15,15,25,"images/tu_03.jpg","","black12")
'============================================================================================================================ 

function getalllist(sqlstr,titlelength,tdwidth,tdwidth1,tdheight,imgurl,classstyle,int_RPP,int_showNumberLink_)
set Rs=server.createobject("adodb.recordset")
Rs.open sqlstr,conn,1,1
if not(Rs.eof or Rs.bof) then

        int_RPP=int_RPP '设置每页显示数目
		int_showNumberLink_=int_showNumberLink_ '数字导航显示数目
		showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
		str_nonLinkColor_="#999999" '非热链接颜色
		toF_="<font face=webdings>9</font>"   			'首页 
		toP10_=" <font face=webdings>7</font>"			'上十 
		toP1_=" <font face=webdings>3</font>"			'上一
		toN1_=" <font face=webdings>4</font>"			'下一
		toN10_=" <font face=webdings>8</font>"			'下十
		toL_="<font face=webdings>:</font>"				'尾页
		rs.PageSize=int_RPP
		cPageNo=Request.QueryString("Page")
		If cPageNo="" or not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<1 Then cPageNo=1
		If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
		rs.AbsolutePage=cPageNo
count=0 

list_text=list_text&"<table width='100%' border='0' cellspacing='0' cellpadding='0'>" 

do while not (Rs.eof or Rs.bof) and count<int_RPP 
 
list_text=list_text&"<tr><td height="&tdheight&"><table width='100%' border='0' cellspacing='0' cellpadding='0'><tr>"
list_text=list_text&"<td width='"&tdwidth&"%' align='center'>"
if imgurl<>"" then
list_text=list_text&"<img src='"&imgurl&"'>"
end if
list_text=list_text&"</td><td width='"& 100 - tdwidth - tdwidth1 &"%' class='"&classstyle&"'><a href='news_detail.asp?/news_"&Rs(0)&".html' class='"&classstyle&"' title='标题"&Rs(1)&"'>"&GetTitle(Rs(1),titlelength)&"</a></td><td width='"&tdwidth1&"%' class='"&classstyle&"'>"&formatdatetime(Rs(2),2)&"</td>"
list_text=list_text&"</tr></table></td></tr>"
 

 if Rs.eof then exit do
 
 i=i+1
 count=count+1
 Rs.movenext
 
loop
list_text=list_text&"<tr><td align='center' height=30><table><tr><td class="&classstyle&">"&fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)&"</td></tr></table></td></tr>"
list_text=list_text&"</table>"
else
list_text=list_text&"<table><tr><td align='center' height=40><table><tr><td><span class='"&classstyle&"'>暂无内容！</span></tr></table></td></tr></table>"

end if

getalllist=list_text
Rs.close
set Rs=nothing
end function
'============================================================================================================================
'函数作用：'调用数据列表
'传人参数： 
'newid:新闻分类
'lcount:循环显示多少条
'titlecounts:标题显示多少字数
'isen:判断中英文
'classstyle:链接样式
'返回值：
'备注：call getlist("select top 5 id,s_name"&language&",s_time from news_list_content where newid=10 order by s_time desc,id desc",15,15,25,"images/tu_03.jpg","","black12")
'============================================================================================================================ 
Function GetAD(adid,adw,adh)
dim f_text
f_text=""
set RsList=server.createobject("adodb.recordset")
RsList.open "select pic,url from ad where id="&adid,conn,1,1
f_text=f_text&"<a href='"&RsList(1)&"' target='_blank'><img src='"&RsList(0)&"' width='"&adw&"' height='"&adh&"' border='0'></a>"
GetAD=f_text
RsList.close
set RsList=nothing
End Function
'============================================================================================================================
'函数作用：'调用数据列表
'传人参数： 
'newid:新闻分类
'lcount:循环显示多少条
'titlecounts:标题显示多少字数
'isen:判断中英文
'classstyle:链接样式
'返回值：
'备注：call getlist("select top 5 id,s_name"&language&",s_time from news_list_content where newid=10 order by s_time desc,id desc",15,15,25,"images/tu_03.jpg","","black12")
'============================================================================================================================ 
function GetAllChildID(id)
	'取得FolderID为id的目录下所有子目录的FolderID，以半角逗号分开
	dim arrID
	arrID = id
	Set rsdir = Conn.Execute("Select ID from pro_class where Parent_ID = " & id & "")
	if rsdir.eof or rsdir.bof then
		set rsdir = nothing
		GetAllChildID = arrID
		exit function
	else
		while not rsdir.eof		
			arrID = arrID&","&GetAllChildID(rsdir("ID"))
		rsdir.movenext
		wend
	end if
	set rsdir = nothing
	GetAllChildID = arrID
end function
%>