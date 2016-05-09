<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
dim id
id=request.QueryString("id")
if not isnumeric(id) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if

if id<>"" then
urlForm=Request.ServerVariables("HTTP_REFERER")
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from S_kc where id="&id,conn,1,1
	s_xh=rs("s_xh")
	s_sl=rs("s_sl")
rs.close
set rs=nothing
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" S_content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body text="#000000" >
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="tjnrym">
  <tr>
    <td height="40" class="tjnrbt">产品管理</td>
  </tr>
  <tr>
    <td height="40"><table width="96%" border="0" cellpadding="0" cellspacing="0" class="tjnrnk">
      <tr>
        <td><form name="myform" method="post" action="pro_save.asp">
       <input name="id" type="hidden" value="<%=id%>">
       <input name="urlForm" type="hidden" value="<%=urlForm%>">       
       <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        

        <tr >
         <td width="14%" height="30" align="right"><font color="#000000">型号：</font></td>
         <td width="86%" bgcolor="<%=Color_0%>">
          <input name="S_xh" type="text" class="inputkkys" id="S_xh" size="60" value="<%=S_xh%>">         </td>
        </tr>
        <tr >
         <td height="30" align="right">数量：</td>
         <td width="86%" bgcolor="<%=Color_0%>">

           <input name="S_sl" type="text" class="inputkkys" id="S_name1" value="<%=S_sl%>" size="60">         </td>
        </tr>
        <tr >
         <td height="40" align="right" bgcolor="<%=Color_0%>"></td>
         <td height="30" bgcolor="<%=Color_0%>">
          <input type="submit" name="Submit" class="inputkkys" value=" 提 交 " onClick="return check();">         </td>
        </tr>
       </table>
      </form></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
conn.close
set conn=nothing
%>
</body>
</html>
<script>
	function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
</script>
<%
function HTMLEncode(fString)
	fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
	fString = Replace(fString, "<BR>", CHR(10))
	HTMLEncode = fString
end function
%>
