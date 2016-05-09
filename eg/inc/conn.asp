<!--#include file="fc_include.asp"-->
<%
sqldata=0 '''数据库类型，1为sql数据库，0为access数据库
if sqldata=0 then
dbstr="database/#%slzjiu#com.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath(dbstr)
connstr="provider=microsoft.jet.oledb.4.0;data source=" &DBPath
conn.Open connstr

else

'sql数据库连接参数：数据库名、用户密码、用户名、连接名（本地用local，外地用IP）
   Dim sql_databasename,sql_password,sql_username,sql_localname
		sql_localname = "127.0.0.1" ''服务器ip
		sql_databasename = "gq.fyasp.com"
		sql_username = "gq.fyasp.com"
		sql_password = "gq.fyasp.com"
		connstr = "Provider = Sqloledb; User ID = " & sql_username & "; Password = " & sql_password & "; Initial Catalog = " & sql_databasename & "; Data Source = " & sql_localname & ";"
	Set conn=Server.CreateObject("ADODB.connection")
	conn.open connstr

end if

language=""

'取得网站配置开始
my_config=db_a("select top 1 s_name,s_name1,s_name2,s_content,s_content1,s_content2,s_keywords,s_copy,S_Description,s_keywords1,s_keywords2,S_Description1,S_Description2,s_img,s_znfd,s_fdkf,s_fddm,s_tjkg,s_tjdm,s_img1,s_ljdz from s_main")
web_name=my_config(0,0)
web_name1=my_config(1,0)
web_name2=my_config(2,0)
web_copyright=my_config(3,0)
web_copyright1=my_config(4,0)
web_copyright2=my_config(5,0)
web_keywords=my_config(6,0)
web_copy=my_config(7,0)
web_Description=my_config(8,0)
web_keywords1=my_config(9,0)
web_keywords2=my_config(10,0)
web_Description1=my_config(11,0)
web_Description2=my_config(12,0)
web_img=my_config(13,0)
web_znfd=my_config(14,0)
web_fdkf=my_config(15,0)
web_fddm=my_config(16,0)
web_tjkg=my_config(17,0)
web_tjdm=my_config(18,0)

web_img1=my_config(19,0)
web_ljdz=my_config(20,0)
'取得网站配置结束

Function ClearAllHTML(strHTML)               

    if strHTMl="" or isnull(strHTML) then    
    exit Function  
    end if   
    StrHtml = Replace(StrHtml,vbCrLf,"")   
    'StrHtml = Replace(StrHtml,Chr(13)&Chr(10),"")   
    StrHtml = Replace(StrHtml,Chr(13),"")   
    StrHtml = Replace(StrHtml,Chr(10),"")   
    StrHtml = Replace(StrHtml," ","")   
    'StrHtml = Replace(StrHtml,"    ","")   
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


function rqq(drq)
'response.Write drq
'response.End()
if drq<>"" then
	yue=(month(drq))
	ri=(day(drq))
	xs=(hour(drq))
	fz=(minute(drq))
	rqq=yue&"月"&ri&"日&nbsp;"&xs&":"&fz
end if
end function

function ywrq(ywq)
	nian=(year(ywq))
	ri=(day(ywq))
	yue=(month(ywq))
	if yue="1" then 
		yuef="January"
	elseif yue="2" then 
		yuef="February"
	elseif yue="3" then 
		yuef="March"
	elseif yue="4" then 
		yuef="April"
	elseif yue="5" then 
		yuef="May"
	elseif yue="6" then 
		yuef="June"
	elseif yue="7" then 
		yuef="July"
	elseif yue="8" then 
		yuef="August"
	elseif yue="9" then 
		yuef="September"
	elseif yue="10" then 
		yuef="October"
	elseif yue="11" then 
		yuef="November"
	elseif yue="12" then 
		yuef="December"
	end if
	ywrq=yuef&","&ri&","&nian
end function

function wzban(ymid)

	set rs=server.CreateObject("adodb.recordset")
	sql="select id,s_img from O_ad where id="&ymid&""
	rs.open sql,conn,1,1
	if not rs.eof then
			response.Write"<img src="""&rs(1)&""" width=""924"" height=""200"" alt=""""/>"
	end if
	rs.close
end function
%>