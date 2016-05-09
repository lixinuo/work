<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%
order_id=trim(request("orderid"))


 w orderid
sql="select * from o_orders where orderid='"&order_id&"'"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql,Conn,1,1
if not rs.eof then
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
   <li>会员积分详细信息</li>
   </span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
 
 <td valign="top" bgcolor="#DDEEFF">
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#DDEEFF">
    <tr>
     <td height="100" colspan="2"><table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
       <td colspan="4" align="center" background="../images/admin_bg_1.gif" bgcolor="#CCCCCC"><b><font color="#ffffff">处理产品订单</font></b></td>
      </tr>
      <tr bgcolor="#EFF5FE">
       <td colspan="2" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
          <td width="89%" align="center"> 订单号：<%=rs("orderid")%> ,详细资料如下： </td>
          <td width="11%" align="center"><input type="button" name="Submit4" value="打 印" onClick="javascript:window.print()">
          </td>
         </tr>
       </table></td>
      </tr>
     
     
       </table></td>
      </tr>
      <tr>
       <td width="20%" bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>商品列表：</td>
       <td bgcolor="#EFF5FE"><table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
         <tr>
          <td bgcolor="#EFF5FE" align="center">商品名称</td>
          <td bgcolor="#EFF5FE" align="center">订购数量</td>
          <td bgcolor="#EFF5FE" align="center">价 格</td>
           <td bgcolor="#EFF5FE" align="center">所获得积分</td>
          <td bgcolor="#EFF5FE" align="center">金额小计</td>
         </tr>
         <%
				 set rss=db("select * from o_orders where orderid='"&rs("orderid")&"'",2)
				 zongji=0
				 z_jifen=0
				 do while not rss.eof
				 %>
         <tr>
          <td bgcolor="#EFF5FE" style='PADDING-LEFT: 5px'>
					<a href="../../pro_detail.asp?id=<%=rss("pid")%>" target="_blank"> <%=str_cut(db_name("p_main",rss("pid")),12)%></a></td>
          <td bgcolor="#EFF5FE"><div align="center"><%=trim(rss("s_num"))%></div></td>
          <td bgcolor="#EFF5FE"><div align="center"><%=Rss("p_price")%>元</div></td>
          <td bgcolor="#EFF5FE"><div align="center"><%=Rss("p_jifen")%></div></td>
          <td bgcolor="#EFF5FE"><div align="center"><%=rss("s_count")&"元"%></div></td>
         </tr>
         <%
					zongji=rss("s_count")+zongji
					z_jifen=z_jifen+Rss("p_jifen")				
					rss.movenext
					loop
					rss.movefirst
					rss.close:set rss=nothing
				%>
         <tr>
          <td colspan="5" bgcolor="#EFF5FE"><div align="right">该订单所获得的总积分：<%=z_jifen%> &nbsp;&nbsp;&nbsp;&nbsp;订单总额共计：<%=zongji%>元 
           &nbsp;&nbsp;&nbsp;&nbsp;</div></td>
         </tr>
       </table></td>
      </tr>
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>用 户 名：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=db_name("u_main",rs("userid"))%></td>
      </tr>
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>收货人姓名：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=trim(rs("s_realname"))%></td>
      </tr>
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>收货地址：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=trim(rs("s_address"))%></td>
      </tr>
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>邮 编：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=trim(rs("s_zip"))%></td>
      </tr>
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>电子邮件：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=trim(rs("s_email"))%></td>
      </tr>
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>联系电话：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=trim(rs("s_phone"))%></td>
      </tr>
   <!--   <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>送货方式：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=db_name("o_order_post",rs("post_method"))%>
       </td>
      </tr>
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>支付方式：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=db_name("o_order_pay",rs("pay_method"))%>
       </td>
      </tr>
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>需要发票：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>
			 <%     if rs("fapiao")&""="1" then
							response.Write("<font color=red>需要</font>")
							else
							response.Write("<font color=red>不需要</font>")
							end if
							%></td>
      </tr>-->
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>备注：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=trim(rs("s_note"))%></td>
      </tr>
      
      <tr>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>推荐人姓名：</td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=Rs("s_tel")%></td>
      </tr>
      
      <tr>
       <td height="20" bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>下单日期：</td>
       <td height="20" bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'><%=rs("order_time")%></td>
      </tr>
      <tr>
       <td height="30" bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'></td>
       <td bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>&nbsp;
        <input type="button" name="Submit2" value="返回订单" onClick="location='order_list.asp'">
       </td></tr>
     </table></td>
   
  </table></td>
 </tr>
</table>
</body>
</html>
<%end if%>
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
