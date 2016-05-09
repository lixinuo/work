<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#include file="../inc/data.asp"-->
<%
function yonghu(toid)
	Set Rsy=Server.CreateObject("ADODB.RecordSet") 
	SqlStr="Select id,s_name From u_main where id="&toid 
	Rsy.Open SqlStr,Conn,1,1
	if not rsy.eof then
		yonghu=rsy(1)
	end if
	rsy.close
end function
id=isid(request.querystring("id"),0)
if id<>0 then
	Set Rs=Server.CreateObject("ADODB.RecordSet") 
	SqlStr="Select * From d_order where id="&id 
	Rs.Open SqlStr,Conn,1,1 
	if not rs.eof then
	 userid=Rs("userid") 
	 xdrq=Rs("xdrq") 
	 jhrq=Rs("jhrq") 
	 jj=Rs("jj") 
	 wjm=Rs("wjm") 
	 pbfs=Rs("pbfs") 
	 cs=Rs("cs") 
	 sl=Rs("sl") 
	 cens=Rs("cens") 
	 bh=Rs("bh") 
	 ym=Rs("ym") 
	 zf=Rs("zf") 
	 dx=Rs("dx") 
	 chsj=Rs("chsj") 
	 chfs=Rs("chfs") 
	 kddh=Rs("kddh") 
	end if 
	Rs.Close:Set Rs=Nothing 
end if
if xdrq="" or isnull(xdrq) then xdrq=date()
if jhrq="" or isnull(jhrq) then jhrq=date()
if chsj="" or isnull(chsj) then chsj=date()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/g_links.asp"-->
</head>
<body>
<table width="100%"   border="0" cellpadding="5" cellspacing="0" bgcolor="<%=Color_0%>">
  <tr>
    <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
    <li>订单修改    </li>
    </span></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#004A80"></td>
  </tr>
  <tr>
    <td valign="top"> 
<table class="tableBorder" width="80%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form action="order_save.asp?action=save&id=<%=id%>" method="post" name="form1" id="form1">
<tr bgcolor="#ffffff"> 
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">客户编号：</td>
      <td width="662" bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px"><select name="userid" id="userid" >
	  <%if userid<>"" then%><option value="<%=userid%>" selected><%=yonghu(userid)%></option><%end if%>
<%
if userid<>"" then cuid=" and id<>"&userid&" "
Set Rs=Server.CreateObject("ADODB.RecordSet") 
SqlStr="Select id,s_name From u_main where 1=1 "&cuid&" order by id desc"
Rs.Open SqlStr,Conn,1,1
if not rs.eof then
do while not rs.eof
%>
		  <option value="<%=rs(0)%>"><%=rs(1)%></option>
<%
rs.movenext
loop
end if
rs.close
%>
          </select></span></td>
    </tr>
<tr bgcolor="#ffffff"> 
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">下单日期：</td>
      <td width="662" bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
	  <input name="xdrq" type="text" id="xdrq" value="<%=xdrq%>" size="40"/>
      </span></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">交货日期：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="jhrq" type="text" id="jhrq" value="<%=jhrq%>" size="40" />
      </span></td>
    </tr>
	    <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">加急：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="jj" type="text" id="jj" value="<%=jj%>" size="40" />
      </span></td>
    </tr>
		    <tr bgcolor="#ffffff">           
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">文件名：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="wjm" type="text" id="wjm" value="<%=wjm%>" size="40" />
      </span></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">排版方式：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="pbfs" type="text" id="pbfs" value="<%=pbfs%>" size="40" />
      </span></td>
    </tr>

    <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">测式：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="cs" type="text" id="cs" value="<%=cs%>" size="40" />
      </span></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td width="125" height="30" align="left" bgcolor="<%=Color_0%>" style="PADDING-LEFT: 8px">数量：</td>
      <td bgcolor="<%=Color_0%>"><span style="PADDING-LEFT: 8px">
        <input name="sl" type="text" id="sl" value="<%=sl%>" size="40" />
      </span> </td>
    </tr>
	<tr>
		  <td height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">层数：</td>
		  <td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="cens" type="text" id="cens" value="<%=cens%>" size="40" /></td>
		  </tr>
		<tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">板厚：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="bh" type="text" id="bh" value="<%=bh%>" size="40" /></td>
		</tr>
		<tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">油墨：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="ym" type="text" id="ym" value="<%=ym%>" size="40"/></td>
		</tr>
		<tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">字符：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="zf" type="text" id="zf" value="<%=zf%>" size="40"/></td>
		</tr>
		<tr>     
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">大小：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="dx" type="text" id="dx" value="<%=dx%>" size="40"/></td>
		</tr>
		<tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">出货时间：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="chsj" type="text" id="chsj" value="<%=chsj%>" size="40"/></td>
		</tr>
		<tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">出货方式：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="chfs" type="text" id="chfs" value="<%=chfs%>" size="40"/></td>
		</tr>
		<tr> 
			<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px">快递单号：</td>
			<td colspan="2" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"><input name="kddh" type="text" id="kddh" value="<%=kddh%>" size="40"/></td>
		</tr>
<tr>
<td width="125" height="30" bgcolor="<%=Color_0%>"  style="PADDING-LEFT: 8px"></td>
<td height="28" colspan="2" bgcolor="<%=Color_0%>" >
<input type="submit" name="Submit" value="确认提交" />
&nbsp;
<input type="button" name="Submit2" value="返回上一页" onclick='javascript:history.go(-1)' /></td>
</tr>
</form>
</table>
</td></tr></table>
</body>
</html>
<%closeconn%>