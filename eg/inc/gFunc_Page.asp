<%''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''调用例子 code by awen ueuo.cn 最好的网络收藏夹
'Dim int_RPP,int_Start,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo
'int_RPP=2 '设置每頁显示数目
'int_showNumberLink_=8 '数字导航显示数目
'showMorePageGo_Type_ = 1 '是下拉菜单还是输入值跳转，当多次调用时只能选1
'str_nonLinkColor_="#999999" '非热链接颜色
'toF_="<font face=webdings>9</font>"  			'首頁 
'toP10_=" <font face=webdings>7</font>"			'上十
'toP1_=" <font face=webdings>3</font>"			'上一
'toN1_=" <font face=webdings>4</font>"			'下一
'toN10_=" <font face=webdings>8</font>"			'下十
'toL_="<font face=webdings>:</font>"				'尾頁

'============================================
'这段代码一定要在VClass_Rs.Open 与 for循环之间
'	Set VClass_Rs = CreateObject(G_FS_RS)
'	VClass_Rs.Open This_Fun_Sql,User_Conn,1,1
'	IF not VClass_Rs.eof THEN 
'	VClass_Rs.PageSize=int_RPP
'	cPageNo=NoSqlHack(Request.QueryString("Page"))
'	If cPageNo="" Then cPageNo = 1
'	If not isnumeric(cPageNo) Then cPageNo = 1
'	cPageNo = Clng(cPageNo)
'	If cPageNo<=0 Then cPageNo=1
'	If cPageNo>VClass_Rs.PageCount Then cPageNo=VClass_Rs.PageCount 
'	VClass_Rs.AbsolutePage=cPageNo
'	  FOR int_Start=1 TO int_RPP 
	  ''++++++++++
	  '加循环体显示数据
	  ''++++++++++
'		VClass_Rs.MoveNext
'		if VClass_Rs.eof or VClass_Rs.bof then exit for
'      NEXT
'	END IF	  
'============================================
'response.Write "<p>"&  fPageCount(VClass_Rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)

''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'*********************************************************
' 目的：分頁的頁面参数保持
'          提交查询的一致性
' 输入：moveParam：分頁参数
'         removeList：要移除的参数
' 返回：分頁Url
'*********************************************************
Function TestRegExp(myString) 
Set r = New RegExp
r.IgnoreCase = True 
r.Global = True 
r.Pattern = "(\?/.*/[0-9]{1,}.html?)((/page/[0-9]{1,})|)"
str2 = r.Replace(myString,"$1")
TestRegExp = str2 
End Function 
Function TestRegExp1(myString) 
Set r = New RegExp
r.IgnoreCase = True 
r.Global = True 
r.Pattern = "(\?)(((/|)page/[0-9]{1,})|)"
str2 = r.Replace(myString,"$1")
TestRegExp1 = str2 
End Function 

Function PageUrl(moveParam,removeList)
	dim strName
	dim KeepUrl,KeepForm,KeepMove
	removeList=removeList&","&moveParam
	KeepForm=""
	For Each strName in Request.Form 
		'判断form参数中的submit、空值
		if not InstrRev(","&removeList&",",","&strName&",", -1, 1)>0 and Request.Form(strName)<>"" then
			KeepForm=KeepForm&"&"&strName&"="&Server.URLencode(Request.Form(strName))
		end if
		removeList=removeList&","&strName
	Next
	
	KeepUrl=""
	For Each strName In Request.QueryString
		If not (InstrRev(","&removeList&",",","&strName&",", -1, 1)>0) Then
			KeepUrl = KeepUrl & "&" & strName & "=" & Server.URLencode(Request.QueryString(strName))
		End If
	Next
	
	KeepMove=KeepForm&KeepUrl
	
	If (KeepMove <> "") Then 
	  KeepMove = Right(KeepMove, Len(KeepMove) - 1)
	  KeepMove = Server.HTMLEncode(KeepMove) & "&"
	End If
	
	'PageUrl = replace(Request.ServerVariables("URL"),"/products.asp","/products.asp") & "?" & KeepMove & moveParam & "="
	PageUrl =  "?" & KeepMove & moveParam & "="
	
if instr(PageUrl,"/") then
	PageUrl=replace(PageUrl,"&","")
	PageUrl=replace(PageUrl,"=","/")
	PageUrl=TestRegExp(PageUrl)
	PageUrl=TestRegExp1(PageUrl)
end if	
End Function 



Function fPageCount(Page_Rs,showNumberLink_,nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,Page)

Dim This_Func_Get_Html_,toPage_,p_,sp2_,I,tpagecount
Dim NaviLength,StartPage,EndPage

This_Func_Get_Html_ = ""  : I = 1   
NaviLength=showNumberLink_ 

if IsEmpty(showMorePageGo_Type_) then showMorePageGo_Type_ = 1
tpagecount=Page_Rs.pagecount
If tPageCount<1 Then tPageCount=1 

if not Page_Rs.eof or not Page_Rs.bof then

toPage_ = PageUrl("Page","submit,GetType,no-cache,_")
if Page=1 then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""首页"">"&toF_&"</font> " &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&"1 title=""首页"">"&toF_&"</a> " &vbNewLine
End If 
if Page<NaviLength then
	StartPage = 1
else
	StartPage = fix(Page / NaviLength) * NaviLength	
end if	
EndPage=StartPage+NaviLength-1 
If EndPage>tPageCount Then EndPage=tPageCount 

If StartPage>1 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_& Page - NaviLength &" title=""上"&int_showNumberLink_&"頁"">"&toP10_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""上"&int_showNumberLink_&"頁"">"&toP10_&"</font> "  &vbNewLine
End If 

If Page <> 1 and Page <>0 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&(Page-1)&"  title=""上一页"">"&toP1_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""上一页"">"&toP1_&"</font> "  &vbNewLine
End If 

For I=StartPage To EndPage 
	If I=Page Then 
		This_Func_Get_Html_=This_Func_Get_Html_& "<font style=""font-family:""humnst777_btbold""""><B>"&I&"</B></font>"  &vbNewLine
	Else 
		This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&I&">" &I& "</a>"  &vbNewLine
	End If 
	If I<>tPageCount Then This_Func_Get_Html_=This_Func_Get_Html_& vbNewLine
Next 

If Page <> Page_Rs.PageCount and Page <>0 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_&(Page+1)&" title=""下一页"">"&toN1_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""下一页"">"&toN1_&"</font> "  &vbNewLine
End If 

If EndPage<tpagecount Then  
	This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_& Page + NaviLength &"  title=""下"&int_showNumberLink_&"頁"">"&toN10_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& " <font color="&nonLinkColor_&"  title=""下"&int_showNumberLink_&"頁"">"&toN10_&"</font> "  &vbNewLine
End If 

if Page_Rs.PageCount<>Page then  
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&Page_Rs.PageCount&" title=""尾页"">"&toL_&"</a>"  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" title=""尾页"">"&toL_&"</font>"  &vbNewLine
End If 

'If showMorePageGo_Type_ = 1 then 
'	Dim Show_Page_i
'	if Show_Page_i > tPageCount then Show_Page_i = 1
'	This_Func_Get_Html_=This_Func_Get_Html_& "<input type=""text"" size=""4"" maxlength=""10"" name=""Func_Input_Page"" onmouseover=""this.focus();"" onfocus=""this.value='"&Show_Page_i&"';"" onKeyUp=""value=value.replace(/[^1-9]/g,'')"" onbeforepaste=""clipboardData.setData('text',clipboardData.getData('text').replace(/[^1-9]/g,''))"">" &vbNewLine _
'		&"<input type=""button"" value=""Go"" onmouseover=""Func_Input_Page.focus();"" onclick=""javascript:var Js_JumpValue;Js_JumpValue=document.all.Func_Input_Page.value;if(Js_JumpValue=='' || !isNaN(Js_JumpValue)) location='"&topage_&"'+Js_JumpValue; else location='"&topage_&"1';"">"  &vbNewLine
'
'Else 
'
'	This_Func_Get_Html_=This_Func_Get_Html_& " 跳转:<select NAME=menu1 onChange=""var Js_JumpValue;Js_JumpValue=this.options[this.selectedIndex].value;if(Js_JumpValue!='') location=Js_JumpValue;"">"
'	for i=1 to tPageCount
'		This_Func_Get_Html_=This_Func_Get_Html_& "<option value="&topage_&i
'		if Page=i then This_Func_Get_Html_=This_Func_Get_Html_& " selected style='color:#0000FF'"
'		This_Func_Get_Html_=This_Func_Get_Html_& ">第"&cstr(i)&"頁</option>" &vbNewLine
'	next
'	This_Func_Get_Html_=This_Func_Get_Html_& "</select>" &vbNewLine
'
'End if

This_Func_Get_Html_="<div class=""in_fyk_cen"">PAGE"&This_Func_Get_Html_

else
	'没有记录
end if
fPageCount = This_Func_Get_Html_
End Function
%>

