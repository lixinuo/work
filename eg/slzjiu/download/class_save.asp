<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
dim S_name,sql,rs,num
id=trim(request.form("id"))
S_name=changechr(trim(request.form("S_name")))
S_name1=changechr(trim(request.form("S_name1")))
S_name2=changechr(trim(request.form("S_name2")))
s_bt=changechr(trim(request("s_bt")))
s_bt1=changechr(trim(Request("s_bt1")))
s_bt2=changechr(trim(Request("s_bt2")))	
s_gjc=changechr(trim(request("s_gjc")))
s_gjc1=changechr(trim(Request("s_gjc1")))
s_gjc2=changechr(trim(Request("s_gjc2")))	
s_ms=changechr(trim(request("s_ms")))
s_ms1=changechr(trim(Request("s_ms1")))
s_ms2=changechr(trim(Request("s_ms2")))
parent_id=trim(request.form("parent_id"))
s_pai=trim(request.form("s_pai"))
s_time=Now()

if parent_id="" then parent_id=0

if id="" then
'增加数据
if parent_id=0 then
	class_depth=0
else
	class_depth=db_f("a_class","class_depth","id="&parent_id&"")+1
end if
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from a_class"
	rs.open sql,conn,1,3
	rs.addnew
		rs("s_pai") = s_pai
		rs("class_depth") =class_depth
		rs("parent_id") = parent_id
		rs("S_name") = S_name
		rs("s_name1") = s_name1
		rs("s_name2") = s_name2
		Rs("s_bt")=s_bt
		Rs("s_bt1")=s_bt1
		Rs("s_bt2")=s_bt2
		Rs("s_gjc")=s_gjc
		Rs("s_gjc1")=s_gjc1
		Rs("s_gjc2")=s_gjc2
		Rs("s_ms")=s_ms
		Rs("s_ms1")=s_ms1
		Rs("s_ms2")=s_ms2
	rs.update
else
'修改数据
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from a_class where id="&id&""
	rs.open sql,conn,1,3
	if not rs.eof then
		rs("S_name") = S_name
		rs("s_name1") = s_name1
		rs("s_name2") = s_name2
		Rs("s_bt")=s_bt
		Rs("s_bt1")=s_bt1
		Rs("s_bt2")=s_bt2
		Rs("s_gjc")=s_gjc
		Rs("s_gjc1")=s_gjc1
		Rs("s_gjc2")=s_gjc2
		Rs("s_ms")=s_ms
		Rs("s_ms1")=s_ms1
		Rs("s_ms2")=s_ms2
	rs.update
	end if
end if
	Response.Write "<script Language=Javascript>alert('操作成功!');</script>"
	response.write "<script>location.href='class_list.asp?id="&parent_id&"&s_pai="&s_pai&"';</script>"

%>