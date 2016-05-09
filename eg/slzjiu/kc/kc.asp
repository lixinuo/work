<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
Server.ScriptTimeOut=9000
%>
<%
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=0;DBQ=" & Server.MapPath("kc_up.xls") 

SQL1="select * from [sheet1$]" 
Set rs = Server.CreateObject("ADODB.Recordset") 
rs.Open SQL1, conn, 3, 3 

curDir = Server.MapPath("../../database/#%slzjiu#com.mdb") 
Set conn1 = Server.CreateObject("ADODB.Connection") 
conn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & curDir 
Set rs1 = Server.CreateObject("ADODB.Recordset") 
'Dim sql_databasename,sql_password,sql_username,sql_localname
'		sql_localname = "114.112.59.208" ''服务器ip
'		sql_databasename = "net1378482"
'		sql_username = "net1378482"
'		sql_password = "r2d3g6e8"
'connstr = "Provider = Sqloledb; User ID = " & sql_username & "; Password = " & sql_password & "; Initial Catalog = " & sql_databasename & "; Data Source = " & sql_localname & ";"
'Set conn1=Server.CreateObject("ADODB.connection")
'conn1.open connstr

Set rs1 = Server.CreateObject("ADODB.Recordset") 
Set rs1.ActiveConnection = conn1 	
rs1.Source = "select * from s_kc" 
rs1.CursorType = 3 ' adOpenKeyset 
rs1.LockType = 3 'adLockOptimistic 
rs1.Open
Do While Not rs.Eof 
'response.Write rs(0)
'response.End()
set rsx=server.CreateObject("adodb.recordset")
sqlx="select id,s_xh from s_kc where s_xh='"&Trim(rs(0))&"'"
rsx.open sqlx,conn1,1,3
if rsx.eof and rsx.bof then
	rs1.AddNew 
	rs1("s_xh")=Trim(rs(0))
	rs1("s_sl")=Trim(rs(1))
	rs1.Update 
end if
rs.MoveNext 
'j=j+1 
Loop 


'response.End()

rs.Close 
rs1.Close 
conn.Close 
conn1.Close 
Set rs=nothing 
Set conn=nothing
%>
数据导入成功！【<a href="pro_list.asp">点击查看管理页</a>】