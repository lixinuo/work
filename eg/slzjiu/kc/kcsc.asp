<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" S_content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>
<body text="#000000" >
<table width="100%" border="0" cellpadding="5" cellspacing="0">
 <tr>
  <td height="50" bgcolor="<%=Color_1%>"><span class="style1"> </span>
   <li></li>
   <span class="style1">库存管理</span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
  <td bgcolor="<%=Color_0%>">
   <table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#DDEEFF">
    <tr>
     <td>
            
       <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
	   <tr >
         <td width="13%" align="right" bgcolor="<%=Color_0%>">&nbsp;</td>
         <td width="87%" bgcolor="<%=Color_0%>">【<a href="kc.xls" target="_blank">范例下载</a>】</td>
        </tr>
      <%=Upload_Init()%>
        <tr >
         <td height="30" align="right" bgcolor="<%=Color_0%>"><strong>上传excel文件</strong></td>
         <td colspan="3" bgcolor="<%=Color_0%>"><form action="shangchuan/save.asp" method="post" enctype="multipart/form-data">  

		<input type="file" name="eimage" id="eimage" />  
		
		<input name="commit" type="submit" value="提交" />  
		
		</form> </td>
        </tr>

       </table>
     </td>
    </tr>
   </table>
   </td>
 </tr>
</table>
<%
conn.close
set conn=nothing
%>
</body>
</html>

