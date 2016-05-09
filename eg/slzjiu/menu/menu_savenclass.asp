<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%dim action,url,i,abc,anclassid,anclass
anclassid=request("anclassid")
anclass=request.QueryString("anclass")
url="http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))
action=request.QueryString("action")


if anclassid="" then
response.Write("<script language='javascript'>alert('请选择类别！');history.back();</script>")
response.End()
end if

'//添加新数据
select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from menu_nclass",conn,1,3
rs.AddNew
rs("nclass")=trim(request("nclass2"))
rs("nclassidorder")=int(request("nclassidorder2"))
rs("anclassid")=int(request("anclassid"))
rs("nclassurl")=trim(request("nclassurl2"))
rs.Update
rs.Close
set rs=nothing
response.redirect url&"menu_nclass.asp?id="&anclassid&"&anclass="&anclass
'//修改数据
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from menu_nclass where nclassid="&request.QueryString("id"),conn,1,3
rs("nclass")=trim(request("nclass"))
rs("nclassidorder")=int(request("nclassidorder"))
rs("nclassurl")=trim(request("nclassurl"))
rs.update
rs.close
set rs=nothing
response.redirect url&"menu_nclass.asp?id="&anclassid&"&anclass="&anclass
'//删除数据
case "del"
anclassid=request.QueryString("anclassid")
conn.execute ("delete from menu_nclass where nclassid="&request.QueryString("id"))
'conn.execute ("delete from shop_books where nclassid="&request.QueryString("id"))
response.redirect url&"menu_nclass.asp?id="&anclassid&"&anclass="&anclass
end select
%>
<%
Function finddir(filepath)
	finddir=""
	for i=1 to len(filepath)
	if left(right(filepath,i),1)="/" or left(right(filepath,i),1)="\" then
	  abc=i
	  exit for
	end if
	next
	if abc <> 1 then
	finddir=left(filepath,len(filepath)-abc+1)
	end if
end Function
%>