<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%
order_id=trim(request("orderid"))
s_state=isid(trim(request("s_state")),0)

if s_state<>0 then
conn.execute("update o_orders set s_state="&s_state&" where orderid='"&order_id&"'")
if s_state=3 then call point_add(order_id)
end if




''''积分的处理×××××××××××××××××××××××××××××××××××××××××××××
sub point_add(order_id)
userid=db_f("o_orders","top 1 userid","orderid='"&order_id&"'")
Set rs =db("select pid,p_jifen,s_type from o_orders where orderid='"&order_id&"'",2)
if not rs.eof then
 do while not rs.eof
 if rs(2)=0 then
	 points=db_f("p_main","s_jifen",rs(0))
	 s_content="在订单<b>"&order_id&"</b>中购买<b>"&str_cut(db_f("p_main","s_name",rs(0)),30)&"</b>产品获得<b>"&points&"</b>积分"
	 	 
	 
	 conn.execute("Insert Into u_points(s_points,userid,orderid,s_note) Values("&points&","&userid&",'"&order_id&"','"&s_content&"')")
'	else
'	 	 points=rs(1)
'	 s_content="在订单<b>"&order_id&"</b>中成功充值<b>"&points&"</b>元"
'	 conn.execute("Insert Into u_money(points,userid,orderid,s_note) Values("&points&","&userid&",'"&order_id&"','"&s_content&"')")
'	 conn.execute("update u_main set s_money=s_money+"&points&" where id="&userid)
	end if
  rs.movenext
 loop
end if
rs.close:set rs=nothing
msg "订单处理成功","?orderid="&order_id
end sub
''''积分的处理×××××××××××××××××××××××××××××××××××××××××××××


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
   <li>订单处理</li>
   </span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
 
 <td valign="top" bgcolor="<%=Color_0%>">
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#DDEEFF">
    <tr>
     <td height="100" colspan="2"><table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
       <td colspan="4" align="center" background="../images/admin_bg_1.gif" bgcolor="#CCCCCC"><b><font color="#ffffff">处理产品订单</font></b></td>
      </tr>
      <tr bgcolor="#EFF5FE">
       <td colspan="2" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
          <td width="89%" align="center"> 订单号为：<%=rs("orderid")%> ,详细资料如下： </td>
          <td width="11%" align="center"><input type="button" name="Submit4" value="打 印" onClick="javascript:window.print()">
          </td>
         </tr>
       </table></td>
      </tr>
      <tr>
       <td width="20%" bgcolor="#EFF5FE" style='PADDING-LEFT: 10px'>订单状态：</td>
       <td width="80%" bgcolor="#EFF5FE"><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
          <td >
			<form name="ff" id="ff" action="?action=s_state">
       <input type="hidden" name="orderid" value="<%=rs("orderid")%>"/>
       <select name="s_state" id="s_state">
			 <%call db_optionid(0,rs("s_state"),"o_order_state")%>
       </select>
			 <input name="ggogo" type="submit" value="处理订单">
			 </form>
          </td>
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
          <td bgcolor="#EFF5FE" align="center">金额小计</td>
         </tr>
         <%
				 set rss=db("select * from o_orders where orderid='"&rs("orderid")&"'",2)
				 zongji=0
				 do while not rss.eof
				 %>
         <tr>
          <td bgcolor="#EFF5FE" style='PADDING-LEFT: 5px'>
					<a href="../../pro_detail.asp?id=<%=rss("pid")%>" target="_blank"> <%=str_cut(db_name("p_main",rss("pid")),12)%></a></td>
          <td bgcolor="#EFF5FE"><div align="center"><%=trim(rss("s_num"))%></div></td>
          <td bgcolor="#EFF5FE"><div align="center"><%=Rss("p_price")%>元</div></td>
          <td bgcolor="#EFF5FE"><div align="center"><%=rss("s_count")&"元"%></div></td>
         </tr>
         <%
					zongji=rss("s_count")+zongji				
					rss.movenext
					loop
					rss.movefirst
					rss.close:set rss=nothing
				%>
         <tr>
          <td colspan="4" bgcolor="#EFF5FE"><div align="right">订单总额共计：<%=zongji%>元 
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
