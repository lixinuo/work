<!--#include file="../../inc/fc_include.asp"-->
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!-- #Include File="config.asp" -->
<%'call hacker()


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
%>