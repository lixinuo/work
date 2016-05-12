<!--#include file="../../inc/fc_include.asp"-->
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!-- #Include File="config.asp" -->
<Script language="JavaScript">
if (top==self)
{
alert("对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。");
self.location.href="../index.asp";
}
</Script>
<%
	Response.Write "<link rel=""stylesheet"" rev=""stylesheet"" href=""../inc/greybox/gb_styles.css"" type=""text/css"" media=""all"" />" & vbCrLf
	Response.Write "<script language=""javascript"">var GB_ROOT_DIR ='../inc/greybox/';</script>" & vbCrLf
	Response.Write "<script language=""javascript"" src=""../inc/greybox/AJS.js""></script>" & vbCrLf
	Response.Write "<script language=""javascript"" src=""../inc/greybox/AJS_fx.js""></script>" & vbCrLf
	Response.Write "<script language=""javascript"" src=""../inc/greybox/gb_scripts.js""></script>" & vbCrLf

%>
<%'call hacker()
if request.Cookies(Cookies_name)("Grade")="" then 
	'Response.Clear()
	'SERVER.Transfer("error.asp")
	Response.Redirect("error.asp")
	Response.End()
end if

dim conn   
dim connstr,dbimages

	dbimages = "../"&DbAccessConnection
	dbimages = Trim(Server.MapPath(""&dbimages&""))
	Set conn = Server.CreateObject("ADODB.connection")
	connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&dbimages
	 if err.number<>0 then 
	  Response.Write("<script>alert('"&err.description&"');</script>")
	  err.clear
	  Response.End()
   else
        conn.open connstr
        if err.number<>0 then
		  Response.Write("<script>alert('"&err.description&"');</script>")
	      err.clear
		  Response.End()
		end if  
   end if
	
	sub CloseConn()
		conn.close
		set conn=nothing
	end sub	
weblanguage=db_s("select S_Language from s_Main ")
webbjq=db_s("select s_bjq from s_Main ")
webname=db_s("select w_name from s_Main ")
'---------------------------函数列表-----------------------
function edit_html(str)
    edit_html=str
		if edit_html<>"" then
    edit_html=replace(edit_html,"'","")
    edit_html=server.HTMLEncode(edit_html)
		else
		edit_html="&nbsp;"
		end if
end function
'---------------------------函数列表-----------------------
function changechr(str) 
    changechr=trim(str)
	'changechr=replace(changechr,chr(13),"<br>")
   ' changechr=replace(changechr,"'","’")
	'changechr=replace(changechr,",","，")
    'changechr=replace(changechr,mid(" "" ",2,1),"&quot;")
end function


sub hacker()
myurl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
if myurl="" then
response.write "<script>alert('对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。');</script>"
response.write "<script>location.href='../index.asp';</script>"
else
outurl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
if mid(myurl,len(outurl)+1,1)=":" then
outurl=outurl & ":" & Request.ServerVariables("SERVER_PORT")
end if
outurl=lcase(outurl & request.ServerVariables("SCRIPT_NAME"))
if lcase(left(myurl,instrrev(myurl,"/")))<>lcase(left(outurl,instrrev(outurl,"/"))) then
response.write "<script>alert('对不起，为了系统安全，不允许从外部链接地址访问本系统的后台管理页面。')</script>"
response.write "<script>location.href='../index.asp';</script>"
end if
end if
end sub

'***********************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'***********************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function
'***********************************************
'函数名：showpage
'作  用：显示分页信息
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'***********************************************
	sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    'strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & ""
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location.href='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">" & i & "</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	response.write strTemp
end sub
%>