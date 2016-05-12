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
ccid=rs("ccid")
S_name=rs("S_name")
S_name1=rs("S_name1")
S_name2=rs("S_name2")
S_name3=rs("S_name3")
S_name4=rs("S_name4")
S_name5=rs("S_name5")
S_time=rs("S_time")
S_bt = rs("S_bt")
S_bt1 = rs("S_bt1")
S_bt2 = rs("S_bt2")
S_gjc = rs("S_gjc")
S_gjc1 = rs("S_gjc1")
S_gjc2 = rs("S_gjc2")
S_ms = rs("S_ms")
S_ms1 = rs("S_ms1")
S_ms2 = rs("S_ms2")
S_img=rs("S_img")
S_img1=rs("S_img1")
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

if trim(S_time)="" then S_time=now()
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" S_content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
<%if webbjq=1 then%>
<script charset="utf-8" src="../htmledit/kindeditor.js"></script>
<script charset="utf-8" src="../htmledit/lang/zh_CN.js"></script>
<script>
KindEditor.ready(function(K) {
	K.create('textarea[name="s_content"]', {
		allowFileManager : true
	});
	K.create('textarea[name="s_content1"]', {
		allowFileManager : true
	});
	K.create('textarea[name="s_content2"]', {
		allowFileManager : true
	});
});
</script>
<%end if%>
</head>
<body text="#000000" >
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="tjnrym">
  <tr>
    <td height="40" class="tjnrbt">信息管理</td>
  </tr>
  <tr>
    <td height="40"><table width="96%" border="0" cellpadding="0" cellspacing="0" class="tjnrnk">
      <tr>
        <td><form name="myform" method="post" action="pro_save.asp">
       <input name="id" type="hidden" value="<%=id%>">
       <input name="s_pai" type="hidden" value="<%=s_pai%>">
       <input name="urlForm" type="hidden" value="<%=urlForm%>">       
       <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr >
         <td width="14%" height="30" align="right"><font color="#000000">信息分类：</font></td>
         <td width="86%" bgcolor="<%=Color_0%>">
          <select name="classid" id="classid" >
           <% call db_childid(0,classid,"p_class",s_pai) %>
          </select>         </td>
        </tr>

        <tr >
         <td height="30" align="right"><font color="#000000">信息名称(en)：</font></td>
         <td width="86%" bgcolor="<%=Color_0%>">
          <input name="S_name" type="text" class="inputkkys" id="S_name" size="60" value="<%=S_name%>">         </td>
        </tr>
        
        <tr >
         <td height="30" align="right"><font color="#000000">信息名称(cn)：</td>
         <td width="86%" bgcolor="<%=Color_0%>">

           <input name="S_name1" type="text" class="inputkkys" id="S_name1" value="<%=S_name1%>" size="60">         </td>
        </tr>
		<tr >
         <td height="30" align="right"><font color="#000000">发表：</font></td>
         <td width="86%" bgcolor="<%=Color_0%>">

           <input name="S_name2" type="text" class="inputkkys" id="S_name2" value="<%=S_name2%>" size="60">         </td>
        </tr>  
<!--		<tr >
         <td height="80" align="right"><font color="#000000">简略说明：</font></td>
         <td width="86%" bgcolor="<%=Color_0%>">

           <textarea name="S_name3" cols="60" class="inputkkys" rows="5" id="S_name3">< %=S_name3%></textarea>         </td>
        </tr>-->
		<tr >
         <td height="30" align="right"><font color="#000000">类型：</font></td>
         <td width="86%" bgcolor="<%=Color_0%>">

           <input name="S_name4" type="text" class="inputkkys" id="S_name4" value="<%=S_name4%>" size="60">        </td>
        </tr>
		<%if instr(webLanguage,"2") then%>
		<tr >
         <td height="80" align="right"><font color="#000000">简略说明(日)：</font></td>
         <td width="86%" bgcolor="<%=Color_0%>">

           <textarea name="S_name5" cols="60" class="inputkkys" rows="5" id="S_name5"><%=S_name5%></textarea>         </td>
        </tr>
		<%end if%>
      <%=Upload_Init()%>
        <tr >
         <td height="30" align="right">上传图片：</td>
         <td width="86%" bgcolor="<%=Color_0%>"><%=Upload_Input("S_img",s_img)%> (370*555)(370*245)        </td>
        </tr>
		<tr >
         <td height="30" align="right">上传视频：</td>
         <td width="86%" bgcolor="<%=Color_0%>"><%=Upload_Input("S_img1",S_img1)%>  格式为.MP4       </td>
        </tr>
		<tr>
         <td width="14%" height="20" align="right" valign="middle">Title:</td>
         <td width="81%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt" type="text" class=inputkkys style="width:99%" size="60" value="<%=S_bt%>"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="14%" height="20" align="right" valign="middle">Title(en):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt1" style="width:99%" type="text" class=inputkkys  size="60" value="<%=S_bt1%>"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="14%" height="20" align="right" valign="middle">Title(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_bt2" style="width:99%" type="text" class=inputkkys  size="60" value="<%=S_bt2%>"></td>
        </tr>
        <%end if%>
		<tr>
         <td width="14%" height="30" align="right" valign="middle">Keyword:</td>
         <td width="81%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc" style="width:99%" type="text" class="inputkkys" value="<%=S_gjc%>" size="60"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="14%" height="30" align="right" valign="middle">Keyword(en):</td>
         <td valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc1" type="text" style="width:99%" class="inputkkys" value="<%=S_gjc1%>" size="60"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="14%" height="30" align="right" valign="middle">Keyword(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_gjc2" style="width:99%" type="text" class="inputkkys" value="<%=S_gjc2%>" size="60"></td>
        </tr>
        <%end if%>
		<tr>
         <td width="14%" height="30" align="right" valign="middle">Description:</td>
         <td width="81%" height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms" style="width:99%" type="text" class="inputkkys" value="<%=S_ms%>" size="60"></td>
        </tr>
		 <%if instr(weblanguage,"1") then%>
		<tr>
         <td width="14%" height="30" align="right" valign="middle">Description(en):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms1" type="text" style="width:99%" class="inputkkys" value="<%=S_ms1%>" size="60"></td>
		</tr>
		<%end if%>
		<%if instr(weblanguage,"2") then%>
		<tr>
         <td width="14%" height="30" align="right" valign="middle">Description(日):</td>
         <td height="20" valign="middle" bgcolor="<%=Color_0%>"><input name="S_ms2" type="text" style="width:99%" class="inputkkys" value="<%=S_ms2%>" size="60"></td>
        </tr>
        <%end if%>
<!--        <tr >
         <td align="right" bgcolor="<%=Color_0%>"><strong>上传大图片</strong></td>
         <td bgcolor="<%=Color_0%>">< %=Upload_Input("S_img1",s_img1)%>         </td>
        </tr>-->
		<tr >
         <td height="30" align="right"><font color="#000000">日期：</font></td>
         <td width="86%" bgcolor="<%=Color_0%>">

           <input name="S_time" type="text" class="inputkkys" id="S_time" value="<%=S_time%>" size="60">         </td>
        </tr>  
        <tr >
         <td height="410" align="right">信息说明：</td>
         <td width="86%" bgcolor="<%=Color_0%>">
          <%if webbjq=1 then%>
	  <textarea name="s_content" style="width:100%;height:400px;visibility:hidden;"><%=s_content%></textarea>
	  <%elseif webbjq=2 then%>
	  <input type=hidden name="s_content" value="<%=edit_html(s_content)%>">
       <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content&style=standard600" frameborder="0" scrolling="no" width="100%" height="400"></IFRAME>
	  <%end if%>
	           </td>
        </tr>
		<%if instr(webLanguage,"1") then%>
		<tr >
         <td height="410" align="right">信息说明（en）：</td>
         <td width="86%" bgcolor="<%=Color_0%>">
         <%if webbjq=1 then%>
	  <textarea name="s_content1" style="width:100%;height:400px;visibility:hidden;"><%=s_content1%></textarea>
	  <%elseif webbjq=2 then%>
	  <input type=hidden name="s_content1" value="<%=edit_html(s_content1)%>">
       <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content1&style=standard600" frameborder="0" scrolling="no" width="100%" height="400"></IFRAME>
	  <%end if%>
	           </td>
        </tr>
		<%end if%>
		<%if instr(webLanguage,"2") then%>
		<tr >
         <td height="410" align="right">信息说明（日）：</td>
         <td width="86%" bgcolor="<%=Color_0%>">
         <%if webbjq=1 then%>
	  <textarea name="s_content2" style="width:100%;height:400px;visibility:hidden;"><%=s_content2%></textarea>
	  <%elseif webbjq=2 then%>
	  <input type=hidden name="s_content2" value="<%=edit_html(s_content2)%>">
       <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content2&style=standard600" frameborder="0" scrolling="no" width="100%" height="400"></IFRAME>
	  <%end if%>
		         </td>
        </tr>
		<%end if%>
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
