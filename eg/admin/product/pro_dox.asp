<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
s_pai=request("s_pai")
if s_pai="" then
s_pai=0
end if
session("s_pai")=s_pai
dim id
id=request.QueryString("id")
if not isnumeric(id) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if

if id<>"" then
urlForm=Request.ServerVariables("HTTP_REFERER")
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from P_main where id="&id,conn,1,1
classid=rs("classid")
S_name=rs("S_name")
S_name1=rs("S_name1")
S_name2=rs("S_name2")
S_img=rs("S_img")
S_price=rs("S_price")
S_price1=rs("S_price1")
S_img1=rs("S_img1")
S_down=rs("S_down")
S_content=rs("S_content")
S_content1=rs("S_content1")
S_content2=rs("S_content2")


S_jifen=Rs("S_jifen")''积分
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
<table width="100%" border="0" cellpadding="5" cellspacing="0">
 <tr>
  <td height="50" bgcolor="<%=Color_1%>"><span class="style1"> </span>
   <li></li>
   <span class="style1">产品管理</span></td>
 </tr>
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <tr>
  <td bgcolor="<%=Color_0%>">
   <table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#DDEEFF">
    <tr>
     <td>
      <form name="myform" method="post" action="pro_save.asp">
       <input name="id" type="hidden" value="<%=id%>">
       <input name="s_pai" type="hidden" value="<%=s_pai%>">
       <input name="urlForm" type="hidden" value="<%=urlForm%>">       
       <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
        <tr >
         <td width="15%" align="right" bgcolor="<%=Color_0%>"><strong><b><font color="#000000">产品分类</font></b></strong></td>
         <td width="85%" bgcolor="<%=Color_0%>">
          <select name="classid" id="classid" >
            <option value="1">加盟店</option>
            <option value="2">合作伙伴</option>
          
          </select>
         </td>
        </tr>
        <tr >
         <td align="right" bgcolor="<%=Color_0%>"><strong><b><font color="#000000">产品名称</font></b></strong></td>
         <td bgcolor="<%=Color_0%>">
          <input name="S_name" type="text" id="S_name" size="40" maxlength="60" value="<%=S_name%>">

         </td>
        </tr>
        
          <tr >
         <td align="right" bgcolor="<%=Color_0%>"><strong><b><font color="#000000">产品编号</font></b></strong></td>
         <td bgcolor="<%=Color_0%>">

          <input name="S_name1" type="text" id="S_name1" size="60"  value="<%=S_name1%>">

         </td>
        </tr>      
             <tr >
         <td align="right" bgcolor="<%=Color_0%>"><strong><b><font color="#000000">产品规格</font></b></strong></td>
         <td bgcolor="<%=Color_0%>">

          <input name="S_name2" type="text" id="S_name2" size="60"  value="<%=S_name2%>">

         </td>
        </tr>         
        


		<tr >
		  <td align="right" bgcolor="#DDEEFF"><strong><b><font color="#000000">产品价格</font></b></strong></td>
		  <td bgcolor="#DDEEFF">￥<input name="S_price" type="text" id="S_price" onKeyUp="this.value=this.value.replace(/[^\d]/g,'');" onafterpaste="this.value=this.value.replace(0,'');" size="12" maxlength="10" value="<%=S_price%>">元
		    <!--会员价：
		    <input name="S_price1" type="text" id="S_price1" size="12" maxlength="10" value="<%=S_price1%>">-->
		    </td>
		  </tr>


		<tr>
		  <td align="right" bgcolor="#DDEEFF"><strong><b><font color="#000000">产品积分</font></b></strong></td>
		  <td bgcolor="#DDEEFF"><input name="S_jifen" type="text" id="S_jifen" onKeyUp="this.value=this.value.replace(/[^\d]/g,'');" onafterpaste="this.value=this.value.replace(0,'');" size="12" maxlength="10" value="<%=S_jifen%>">

		    <!--会员价：
		    <input name="S_price1" type="text" id="S_price1" size="12" maxlength="10" value="<%=S_price1%>">-->
		    </td>
		  </tr>
          
		<!--<tr >
		  <td align="right" bgcolor="#DDEEFF"><strong><b><font color="#000000">产品积分</font></b></strong></td>
		  <td bgcolor="#DDEEFF"><input name="S_jifen" type="text" id="S_jifen" onKeyUp="this.value=this.value.replace(/[^\d]/g,'');" onafterpaste="this.value=this.value.replace(0,'');" size="12" maxlength="10" value="<%=S_jifen%>">
<font color="#FF0000">*购买该产品时，所获得的积分数目*</font>
		    </td>
		  </tr>-->
          
      <%=Upload_Init()%>
        <tr >
         <td align="right" bgcolor="<%=Color_0%>"><strong>上传图片</strong></td>
         <td bgcolor="<%=Color_0%>"><%=Upload_Input("S_img",s_img)%>
         </td>
        </tr>
        

        
        
<!--              <tr >
         <td align="right" bgcolor="<%=Color_0%>"><strong>上传大图</strong></td>
         <td bgcolor="<%=Color_0%>"><%=Upload_Input("S_img1",s_img1)%>
         </td>
        </tr> 

        <tr bgcolor="<%=Color_0%>">
         <td height="20" align="right" valign="middle"><strong>上传文档</strong></td>
         <td width="569" height="20" valign="middle">
          <input name="S_down" type="text" id="S_down" value="<%=S_down%>" size="30">
          <input type="button" name="Submit23" value="上传产品" onClick="window.open('../inc/upfile.asp?formname=myform&editname=S_down','','status=no,scrollbars=yes,top=20,left=110,width=500,height=200')">
         </td>
        </tr>-->
        <!--			  
              <tr >
                <td align="right" bgcolor="<%=Color_0%>"><strong>产品简介</strong></td>
                <td bgcolor="<%=Color_0%>"><textarea name="guige" cols="40" rows="4"><%=guige%></textarea>                </td>
              </tr>
 -->
        <tr >
         <td align="right" bgcolor="<%=Color_0%>"><strong>详细介绍</strong></td>
         <td bgcolor="<%=Color_0%>">
          <input type=hidden name="S_content" value="<%=edit_html(S_content)%>">
          <iframe id="txtcontent" src="../htmledit/ewebeditor.htm?id=S_content&style=standard600" frameborder="0" scrolling="no" width="100%" height="300"></iframe>
         </td>
        </tr>
        
        <tr bgcolor="<%=Color_0%>">
         <td align="right" valign="top" bgcolor="#DDEEFF"><strong>规格参数:</strong></td>
         <td bgcolor="#DDEEFF">
          <input type=hidden name="S_content1" value="<%=edit_html(S_content1)%>">
          <IFRAME ID="txtcontent" src="../htmledit/ewebeditor.htm?id=S_content1&style=standard600" frameborder="0" scrolling="no" width="100%" height="250"></IFRAME>
         </td>
        </tr>
    
        <%if instr(webLanguage,"2") then%>
        <tr bgcolor="<%=Color_0%>">
         <td align="right" valign="top" bgcolor="#DDEEFF"><strong>繁体:</strong></td>
         <td bgcolor="#DDEEFF">
          <input type=hidden name="S_content2" value="<%=edit_html(S_content2)%>">
          <IFRAME ID="txtcontent" src="../htmledit/ewebeditor.htm?id=S_content2&style=standard600" frameborder="0" scrolling="no" width="100%" height="250"></IFRAME>
         </td>
        </tr>
        <%end if%>
        <tr >
         <td align="right" bgcolor="<%=Color_0%>"></td>
         <td height="30" bgcolor="<%=Color_0%>">
          <input type="submit" name="Submit" value=" 提 交 " onClick="return check();">
         </td>
        </tr>
       </table>
      </form>
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
