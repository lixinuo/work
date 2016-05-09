<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%
dim id
id=trim(request.QueryString("id"))
o_state=trim(request.QueryString("o_state"))
if id="" or id=null then
response.Write("<script language=javascript>alert('没有该定单！');window.close;</script>")
response.End()
else
id=cint(id)
end if

if o_state<>"" then
o_state=cint(o_state)
conn.execute("update s_orders set o_state="&o_state&" where id="&id)
call msg("订单状态修改成功！","?id="&id)
end if

sql="select * from s_orders where id="&id
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Conn,1,1
%>
<html>
<head>
<title>订单处理</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
 background-color: <%=Color_0%>;
}
.style1 {
	color: #000000;
	font-weight: bold;
	font-size: 14px;
}
.style55 {
	color: #666666
}
body, td, th {
	color: #666666;
}
.style56 {
	color: #FF0000
}
-->
</style>
<link href="../images/cssyullhao.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%"   border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
 <tr>
  <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
   <li>订单处理</li>
   </span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
 
 <td valign="top" bgcolor="<%=Color_0%>">
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#DDEEFF">
   <form name="form1" method="post" action="">
    <tr>
     <td height="100"><table width="80%" height="336" border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#FFFFFF">
       <tr>
        <td height=30 colspan="2" align="center" bgcolor="#DDEEFF" class="style1">订单<%= rs("orderid") %> 处理</td>
        </tr>
       <tr>
        <td width="25%" height=25 bgcolor="#DDEEFF">订单编号</td>
        <td width="75%" bgcolor="#DDEEFF"><%= rs("orderid") %></td>
       </tr>
       <tr>
        <td height=25 bgcolor="#DDEEFF">订单类型</td>
        <td bgcolor="#DDEEFF"><%if rs("o_type")=0 then w("个人业务") else w("公司业务") %></td>
       </tr>
       <tr>
        <td height=25 bgcolor="#DDEEFF">申请者/公司</td>
        <td bgcolor="#DDEEFF"><%= rs("o_name") %></td>
       </tr>
       <tr>
        <td height=25 bgcolor="#DDEEFF">申请业务</td>
        <td bgcolor="#DDEEFF"><%= rs("o_yewu") %></td>
       </tr>
<%if rs("o_type")=1 then%>
       <tr>
        <td height=25 bgcolor="#DDEEFF">部门</td>
        <td bgcolor="#DDEEFF"><%= rs("o_part") %></td>
       </tr>
<%end if%>
       <tr>
        <td height=25 bgcolor="#DDEEFF">姓名</td>
        <td bgcolor="#DDEEFF"><%= rs("o_realname") %></td>
       </tr>
       <tr>
        <td height=25 bgcolor="#DDEEFF">手机号码</td>
        <td bgcolor="#DDEEFF"><%= rs("o_phone") %></td>
       </tr>
       <tr>
        <td height=25 bgcolor="#DDEEFF">其他说明</td>
        <td bgcolor="#DDEEFF"><%= rs("o_note") %></td>
       </tr>
       <tr>
        <td height=25 bgcolor="#DDEEFF">订单时间</td>
        <td bgcolor="#DDEEFF"><%= rs("o_time") %></td>
       </tr>
       <tr>
        <td height=25 bgcolor="#DDEEFF">订单状态</td>
        <td bgcolor="#DDEEFF"><select name="state" id="state" onChange="location=document.form1.state.options[document.form1.state.selectedIndex].value;" >
         <option value="?o_state=0&id=<%=id%>" <% if rs("o_State")="0" then w("selected")%>>未处理</option>
         <option value="?o_state=1&id=<%=id%>" <% if rs("o_State")="1" then w("selected")%>>无法处理</option>
         <option value="?o_state=2&id=<%=id%>" <% if rs("o_State")="2" then w("selected")%>>处理中</option>
         <option value="?o_state=3&id=<%=id%>" <% if rs("o_State")="3" then w("selected")%>>已处理</option>
        </select></td>
       </tr>
       <tr>
        <td height=20 bgcolor="#DDEEFF">&nbsp;</td>
        <td height="30" bgcolor="#DDEEFF">
         <input type="submit" name="button" id="button" value="返回订单列表" onClick="location='s_order_manage.asp';"></td>
       </tr>
      </table></td>
   </form>
   
  </table></td>
 </tr>
</table>
</body>
</html>
<script>
function test()
{
  if(!confirm('确认删除吗？')) return false;
}
</script>
<script language=javascript>
function mm()
{
   var a = document.getElementsByTagName("input");
   for (var i=1; i<a.length; i++)
			a[i].checked = a[a.length-1].checked;
}
</script>
