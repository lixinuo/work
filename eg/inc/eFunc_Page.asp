<%
'*********************************************************
' 目的：分页的页面参数保持
'          提交查询的一致性
' 输入：moveParam：分页参数
'         removeList：要移除的参数
' 返回：分页Url
'*********************************************************
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
	
	'PageUrl = replace(Request.ServerVariables("URL"),"/Search.asp","/Search.html") & "?" & KeepMove & moveParam & "="
	PageUrl =  "?" & KeepMove & moveParam & "="
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
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" >"&toF_&"</font> " &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&"1 >"&toF_&"</a> " &vbNewLine
End If 
if Page<NaviLength then
	StartPage = 1
else
	StartPage = fix(Page / NaviLength) * NaviLength	
end if	
EndPage=StartPage+NaviLength-1 
If EndPage>tPageCount Then EndPage=tPageCount 

If StartPage>1 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_& Page - NaviLength &" >"&toP10_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" >"&toP10_&"</font> "  &vbNewLine
End If 

If Page <> 1 and Page <>0 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&(Page-1)&"  >"&toP1_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" >"&toP1_&"</font> "  &vbNewLine
End If 

For I=StartPage To EndPage 
	If I=Page Then 
		This_Func_Get_Html_=This_Func_Get_Html_& "<b>"&I&"</b>"  &vbNewLine
	Else 
		This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&I&">" &I& "</a>"  &vbNewLine
	End If 
	If I<>tPageCount Then This_Func_Get_Html_=This_Func_Get_Html_& vbNewLine
Next 

If Page <> Page_Rs.PageCount and Page <>0 Then 
	This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_&(Page+1)&" >"&toN1_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&" >"&toN1_&"</font> "  &vbNewLine
End If 

If EndPage<tpagecount Then  
	This_Func_Get_Html_=This_Func_Get_Html_& " <a href="&toPage_& Page + NaviLength &"  >"&toN10_&"</a> "  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& " <font color="&nonLinkColor_&" >"&toN10_&"</font> "  &vbNewLine
End If 

if Page_Rs.PageCount<>Page then  
	This_Func_Get_Html_=This_Func_Get_Html_& "<a href="&toPage_&Page_Rs.PageCount&" >"&toL_&"</a>"  &vbNewLine
Else 
	This_Func_Get_Html_=This_Func_Get_Html_& "<font color="&nonLinkColor_&">"&toL_&"</font>"  &vbNewLine
End If 

If showMorePageGo_Type_ = 1 then 
	Dim Show_Page_i
	Show_Page_i = Page + 1
	if Show_Page_i > tPageCount then Show_Page_i = 1
	This_Func_Get_Html_=This_Func_Get_Html_& "<input type=""text"" size=""4"" maxlength=""10"" name=""Func_Input_Page"" onmouseover=""this.focus();"" onfocus=""this.value='"&Show_Page_i&"';"" onKeyUp=""value=value.replace(/[^1-9]/g,'')"" onbeforepaste=""clipboardData.setData('text',clipboardData.getData('text').replace(/[^1-9]/g,''))"">" &vbNewLine _
		&"<input type=""button"" value=""Go"" onmouseover=""Func_Input_Page.focus();"" onclick=""javascript:var Js_JumpValue;Js_JumpValue=document.all.Func_Input_Page.value;if(Js_JumpValue=='' || !isNaN(Js_JumpValue)) location='"&topage_&"'+Js_JumpValue; else location='"&topage_&"1';"">"  &vbNewLine

Else 

	This_Func_Get_Html_=This_Func_Get_Html_& " Jump:<select NAME=menu1 onChange=""var Js_JumpValue;Js_JumpValue=this.options[this.selectedIndex].value;if(Js_JumpValue!='') location=Js_JumpValue;"">"
	for i=1 to tPageCount
		This_Func_Get_Html_=This_Func_Get_Html_& "<option value="&topage_&i
		if Page=i then This_Func_Get_Html_=This_Func_Get_Html_& " selected style='color:#0000FF'"
		This_Func_Get_Html_=This_Func_Get_Html_& ">Page"&cstr(i)&"</option>" &vbNewLine
	next
	This_Func_Get_Html_=This_Func_Get_Html_& "</select>" &vbNewLine

End if

This_Func_Get_Html_=This_Func_Get_Html_& p_&sp2_&" <b>"&Page_Rs.PageSize&"</b>items/page,<b><span class=""tx"">"&sp2_&Page&"</span>/"&tPageCount&"</b>page,<b><span id='recordcount'>"&sp2_&Page_Rs.recordCount&"</span></b>items"

else
	'没有记录
end if
fPageCount = This_Func_Get_Html_
End Function
%>

