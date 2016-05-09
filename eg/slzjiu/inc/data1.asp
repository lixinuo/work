<!--#include file="../../inc/fc_include.asp"-->
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!-- #Include File="config.asp" -->




<%'call hacker()
'if request.Cookies(Cookies_name)("Grade")="" then 
'	'Response.Clear()
'	'SERVER.Transfer("error.asp")
'	Response.Redirect("error.asp")
'	Response.End()
'end if

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
webname=db_s("select w_name from s_Main ")
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
response.write "<script>location.href='index.asp';</script>"
else
outurl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
if mid(myurl,len(outurl)+1,1)=":" then
outurl=outurl & ":" & Request.ServerVariables("SERVER_PORT")
end if
outurl=lcase(outurl & request.ServerVariables("SCRIPT_NAME"))
if lcase(left(myurl,instrrev(myurl,"/")))<>lcase(left(outurl,instrrev(outurl,"/"))) then
response.write "<script>alert('对不起，为了系统安全，不允许从外部链接地址访问本系统的后台管理页面。')</script>"
response.write "<script>location.href='index.asp';</script>"
end if
end if
end sub
%>