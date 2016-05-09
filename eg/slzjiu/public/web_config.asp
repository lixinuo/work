<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
action=request.QueryString("action")
show_yuyan=request("show_yuyan")
show_mail=isid(request("show_mail"),0)
show_count=request("show_count")
if show_yuyan="" then show_yuyan=0
'//保存信息
if action="save" then
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from S_Main ",conn,1,3
rs("S_Keywords")=trim(request("S_Keywords"))
rs("S_Keywords1")=trim(request("S_Keywords1"))
rs("S_Keywords2")=trim(request("S_Keywords2"))

rs("S_Description")=trim(request("S_Description"))
rs("S_Description1")=trim(request("S_Description1"))
rs("S_Description2")=trim(request("S_Description2"))
rs("S_img")=replace(trim(request("S_img")),"../","") 
rs("S_img1")=replace(trim(request("S_img1")),"../","") 

rs("s_ljdz")=trim(request("s_ljdz"))

rs("W_name")=trim(request("W_name"))
rs("S_name")=trim(request("S_name"))
rs("S_name1")=trim(request("S_name1"))
rs("S_name2")=trim(request("S_name2"))
rs("S_Language")=trim(request("S_Language"))
rs("s_mail_smtp")=trim(request("s_mail_smtp"))
rs("s_mail_from")=trim(request("s_mail_from"))
rs("s_mail_to")=trim(request("s_mail_to"))
rs("s_mail_user")=trim(request("s_mail_user"))
rs("s_mail_pwd")=trim(request("s_mail_pwd"))
if show_count=1 then
rs("s_count")=trim(request("s_count"))
end if

rs("s_bjq")=trim(request("s_bjq"))
rs("s_znfd")=trim(request("s_znfd"))
rs("s_fdkf")=trim(request("s_fdkf"))
rs("s_fddm")=trim(request("s_fddm"))
rs("s_tjkg")=trim(request("s_tjkg"))
rs("s_tjdm")=trim(request("s_tjdm"))

rs("S_Content")=trim(request("S_Content"))
rs("S_Content1")=trim(request("S_Content1"))
rs("S_Content2")=trim(request("S_Content2"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('网站资料修改成功！');history.go(-1);</script>"
end if%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
<link rel="stylesheet" href="../htmledit/themes/default/default.css" />
<%if webbjq=1 then%>
<script charset="utf-8" src="../htmledit/kindeditor.js"></script>
<script charset="utf-8" src="../htmledit/lang/zh_CN.js"></script>
<script>
KindEditor.ready(function(K) {
	K.create('textarea[name="s_content"]', {
		resizeType : 1,
		allowPreviewEmoticons : false,
		allowImageUpload : false,
		items : [
			'fontname', 'fontsize', '|', 'textcolor', 'bgcolor', 'bold', 'italic', 'underline',
			'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
			'insertunorderedlist', '|', 'emoticons', 'image', 'link']
	});
	K.create('textarea[name="s_content1"]', {
		resizeType : 1,
		allowPreviewEmoticons : false,
		allowImageUpload : false,
		items : [
			'fontname', 'fontsize', '|', 'textcolor', 'bgcolor', 'bold', 'italic', 'underline',
			'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
			'insertunorderedlist', '|', 'emoticons', 'image', 'link']
	});
		K.create('textarea[name="s_content2"]', {
		resizeType : 1,
		allowPreviewEmoticons : false,
		allowImageUpload : false,
		items : [
			'fontname', 'fontsize', '|', 'textcolor', 'bgcolor', 'bold', 'italic', 'underline',
			'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
			'insertunorderedlist', '|', 'emoticons', 'image', 'link']
	});
	
});

</script>
<%end if%></head>
<body>
<table width="100%"  border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
 <tr>
  <td valign="top" bgcolor="#FFFFFF"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="tjnrym">
    <tr>
      <td height="40" class="tjnrbt">网站信息设置</td>
    </tr>
    <tr>
      <td height="40"><table width="96%" border="0" cellpadding="0" cellspacing="0" class="tjnrnk">
        <tr>
          <td><table width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="tableBorder">
            <form name="form1" method="post" action="?action=save">
              <input type="hidden" name="show_yuyan" value="<%=show_yuyan%>"> <input type="hidden" name="s_znfd" value="0">
              <%set rs=server.CreateObject("adodb.recordset")
		rs.Open "select * from S_Main",conn,1,1%>
			<tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">后台编辑器：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input name="s_bjq" type="radio" value="1" <%if rs("s_bjq")=1 then response.Write" checked"%>>
                  KindEditor
                  <input type="radio" name="s_bjq" value="2"<%if rs("s_bjq")=2 then response.Write" checked"%>>
                  Ewebeditor</td>
              </tr>
              <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">公司名称：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input name="W_name" class="inputkkys" type="text" id="W_name" value="<%=trim(rs("W_name"))%>" size="60" >
                </td>
              </tr>
              <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">网站名称：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input class="inputkkys" name="S_name" type="text" id="S_name" value="<%=trim(rs("S_name"))%>" size="60" >
                </td>
              </tr>
              <%if instr(weblanguage,"1") then%>
              <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">Web Name：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input class="inputkkys" name="S_name1" type="text" id="S_name1" value="<%=trim(rs("S_name1"))%>" size="60" >
                </td>
              </tr>
              <%end if%>
              <%if instr(weblanguage,"2") then%>
              <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">网站名称（日）：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input class="inputkkys" name="S_name2" type="text" id="S_name2" value="<%=trim(rs("S_name2"))%>" size="60" ></td>
              </tr>
              <%end if%>
              <tr >
                <td height="80" align="right" bgcolor="<%=Color_0%>">关健词：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><textarea class="inputkkys" name="S_Keywords" cols="60" rows="5"><%=trim(rs("S_Keywords"))%></textarea>
                </td>
              </tr>
			  <%if instr(weblanguage,"1") then%>
			  <tr >
			  <td align="right" bgcolor="<%=Color_0%>">关健词(en)：</td>
			  <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><textarea class="inputkkys" name="S_Keywords1" cols="60" rows="5"><%=trim(rs("S_Keywords1"))%></textarea>
			  </td>
			 </tr>
			 <%end if%>
			 <%if instr(weblanguage,"2") then%>
			 <tr >
			  <td align="right" bgcolor="<%=Color_0%>">关健词(日)：</td>
			  <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><textarea class="inputkkys" name="S_Keywords2" cols="60" rows="5"><%=trim(rs("S_Keywords2"))%></textarea>
			  </td>
			 </tr>
			 <%end if%>
              <tr >
                <td height="80" align="right" bgcolor="<%=Color_0%>">网站描述：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><textarea class="inputkkys" name="S_Description" cols="60" rows="5"><%=trim(rs("S_Description"))%></textarea>
                </td>
              </tr>
			  <%if instr(weblanguage,"2") then%>
              <tr >
                <td height="80" align="right" bgcolor="<%=Color_0%>">网站描述(en)：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><textarea class="inputkkys" name="S_Description1" cols="60" rows="5"><%=trim(rs("S_Description1"))%></textarea>
                </td>
              </tr>
			  <%end if%>
			  <%if instr(weblanguage,"1") then%>
			  <tr >
                <td height="80" align="right" bgcolor="<%=Color_0%>">网站描述(日)：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><textarea class="inputkkys" name="S_Description2" cols="60" rows="5"><%=trim(rs("S_Description2"))%></textarea>
                </td>
              </tr>
			  <%end if%>
              <%=Upload_Init()%>
              <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">LOGO上传： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><%=Upload_Input("S_img",rs("s_img"))%> (163*47)</td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>"><strong>下方链接</strong></td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px">例：http://www.baidu.com</td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">链接1： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input class="inputkkys" name="S_Keywords1" type="text" id="S_Keywords1" value="<%=trim(rs("S_Keywords1"))%>" size="60" >
                </td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">链接2： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input class="inputkkys" name="S_Keywords2" type="text" id="S_Keywords2" value="<%=trim(rs("S_Keywords2"))%>" size="60" >
                </td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">链接3： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input class="inputkkys" name="S_Description1" type="text" id="S_Description1" value="<%=trim(rs("S_Description1"))%>" size="60" >
                </td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">链接4： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input class="inputkkys" name="S_Description2" type="text" id="S_Description2" value="<%=trim(rs("S_Description2"))%>" size="60" >
                </td>
              </tr>
			  <!--<tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>"><strong>站内浮动客服</strong> </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px">&nbsp;</td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">浮动开关： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><select name="s_znfd" id="s_znfd">
                  <option value="1" < %if rs("s_znfd")=1 then response.Write" selected"%>>开</option>
                  <option value="0"< %if rs("s_znfd")=0 then response.Write" selected"%>>关</option>
                </select>
                </td>
              </tr>-->
			   <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>"><STRONG>第三方浮动客服</STRONG></td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px">&nbsp;</td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">浮动开关： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><select name="s_fdkf" id="s_fdkf">
                  <option value="1"<%if rs("s_fdkf")=1 then response.Write" selected"%>>开</option>
                  <option value="0"<%if rs("s_fdkf")=0 then response.Write" selected"%>>关</option>
                </select>
                </td>
              </tr>
			   <tr >
                <td height="100" align="right" bgcolor="<%=Color_0%>">浮动代码： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><textarea name="s_fddm" cols="60" rows="5" class="inputkkys" id="s_fddm"><%=rs("s_fddm")%></textarea>
                  <br>
                  (在这里输入53客服，TQ客服等客服代码)</td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>"><STRONG>第三方流量统计</STRONG></td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px">&nbsp;</td>
              </tr>
			  <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">统计开关： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><select name="s_tjkg" id="s_tjkg">
                  <option value="1"<%if rs("s_tjkg")=1 then response.Write" selected"%>>开</option>
                  <option value="0"<%if rs("s_tjkg")=0 then response.Write" selected"%>>关</option>
                </select>
                </td>
              </tr>
			   <tr >
                <td height="100" align="right" bgcolor="<%=Color_0%>">统计代码： </td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><textarea name="s_tjdm" cols="60" rows="5" class="inputkkys" id="s_tjdm"><%=rs("s_tjdm")%></textarea>
                  <br>
                  (在这里输入CNZZ，51，百度等统计代码)</td>
              </tr>
              <%if show_count=1 then%>
              <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">计数器：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input name="s_count" class="inputkkys" type="text" id="s_count" value="<%=trim(rs("s_count"))%>" size="60" >
                </td>
              </tr>
              <%end if%>
              <%if show_mail=1 then%>
              <tr >
                <td height="112" align="right" bgcolor="<%=Color_0%>">邮件发送配置：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><strong>发送邮件使用SMTP</strong>
                    <input name="s_mail_smtp" type="text" id="s_mail_smtp" value="<%=trim(rs("s_mail_smtp"))%>" size="20" >
                  例如smtp.126.com具体请咨询邮箱服务商<br>
                  <strong>收件显示来自邮箱</strong>
                  <input name="s_mail_from" type="text" id="s_mail_from" value="<%=trim(rs("s_mail_from"))%>" size="20" >
                  用户收到邮件时显示的发送邮件地址 例如：XXX@126.com<br>
                  <strong>用于收件邮箱地址</strong>
                  <input name="s_mail_to" type="text" id="s_mail_to" value="<%=trim(rs("s_mail_to"))%>" size="20" >
                  用户在本站发送邮件，你的收件地址 例如：XXX@126.com<br>
                  <strong>发送邮箱登录帐号</strong>
                  <input name="s_mail_user" type="text" id="s_mail_user" value="<%=trim(rs("s_mail_user"))%>" size="20" >
                  发送邮件所使用的邮箱帐号，与SMTP一致需同时提供密码<br>
                  <strong>发送邮箱登录密码</strong>
                  <input name="s_mail_pwd" type="password" id="s_mail_pwd" value="<%=trim(rs("s_mail_pwd"))%>" size="20" >
                  发送邮件所使用的邮箱帐号，与SMTP一致需同时提帐号 </td>
              </tr>
              <%end if%>
              <%if show_yuyan=0 then%>
              <input type="hidden" name="S_Language" value="<%=rs("S_Language")%>">
              <%else%>
              <tr >
                <td height="30" align="right" bgcolor="<%=Color_0%>">网站语言：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input name="S_Language" type="checkbox" id="S_Language" value="0" <%if instr(rs("S_Language"),"0") then response.Write("checked")%>>
                  中文
                  <input name="S_Language" type="checkbox" id="S_Language" value="1" <%if instr(rs("S_Language"),"1") then response.Write("checked")%>>
                  英文
                  <input name="S_Language" type="checkbox" id="S_Language" value="2" <%if instr(rs("S_Language"),"2") then response.Write("checked")%>>
                  日文 </td>
              </tr>
              <%end if%>
			  <%if instr(weblanguage,"1") then%>
              <tr >
                <td height="210" align="right" bgcolor="<%=Color_0%>">版权信息：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px">
				<%if webbjq=1 then%>
				  <textarea name="s_content" style="width:100%;height:200;visibility:hidden;"><%=server.HTMLEncode(rs("s_content"))%></textarea>
				  <%elseif webbjq=2 then%>
				  <input type=hidden name="s_content" value="<%=server.HTMLEncode(rs("s_content"))%>">
    			 <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content&style=mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME>
				  <%end if%>
                </td>
              </tr>
              
              <tr >
                <td height="210" align="right" bgcolor="<%=Color_0%>">版权信息(en)：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px">
				<%if webbjq=1 then%>
				  <textarea name="s_content1" style="width:100%;height:200;visibility:hidden;"><%=server.HTMLEncode(rs("s_content1"))%></textarea>
				  <%elseif webbjq=2 then%>
				  <input type=hidden name="s_content1" value="<%=server.HTMLEncode(rs("s_content1"))%>">
    			 <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content1&style=mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME>
				  <%end if%>
                </td>
              </tr>
              <%end if%>
              <%if instr(weblanguage,"2") then%>
              <tr >
                <td height="210" align="right" bgcolor="<%=Color_0%>">版权信息（日）：</td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px">
				<%if webbjq=1 then%>
				  <textarea name="s_content2" style="width:100%;height:200;visibility:hidden;"><%=server.HTMLEncode(rs("s_content2"))%></textarea>
				  <%elseif webbjq=2 then%>
				  <input type=hidden name="s_content2" value="<%=server.HTMLEncode(rs("s_content2"))%>">
    			 <IFRAME ID="txtcontent" src="../Ewebeditor/ewebeditor.htm?id=s_content2&style=mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME>
				  <%end if%>
                </td>
              </tr>
              <%end if%>
              <tr >
                <td height="32" bgcolor="<%=Color_0%>"></td>
                <td bgcolor="<%=Color_0%>" style="PADDING-LEFT: 10px"><input type="submit" class="inputkkys" name="Submit" value=" 修改保存 ">
                  &nbsp;
                  <input type="reset" name="Submit2" class="inputkkys" value=" 重新添写 ">
                </td>
              </tr>
            </form>
          </table></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
    
 <tr>
  <td height="2" bgcolor="#004A80"></td>
 </tr>
 <td height="0"></td>
 
 <td height="0"></tr>
</table>
<%rs.Close
  set rs=nothing
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
