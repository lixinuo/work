<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../images/cssyullhao.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {
	font-size: 12pt;
	font-weight: bold;
}
body {
	margin-top: 15px;
}
-->
</style>
</head>
<body>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td height="22" bgcolor="f1f1f1"> 
      <div align="center" class="STYLE1"><font color="#000000">查看信息反馈</font></div>    </td>
  </tr>
  <tr> 
    <td height="169" align="center" valign="top" bgcolor="#FFFFFF"> 
      <%
Function unHtml(content)
	ON ERROR RESUME NEXT
	unHtml=content
	IF content <> "" Then
		unHtml=Server.HTMLEncode(unHtml)
		unHtml=Replace(unHtml,vbcrlf,"<br>")
		unHtml=Replace(unHtml,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
		unHtml=Replace(unHtml," ","&nbsp;")
	End IF
	IF Err.Number <>0 Then
		unHtml= "HTML转换中出错请联系管理员<br>"
		Err.Clear
	End IF
End Function
%>
      <%
			
if request("action")="s_ok" then
s_ok=Trim(Request("s_ok"))
id=Trim(Request("id"))
conn.execute("update o_gbook set s_ok="&s_ok&" where id="&id)
end if

set rs=server.createobject("adodb.recordset")
rs.open "select * from o_gbook order by id desc",conn,3,2
'dim page
'dim e_page
'e_page=5 '每页显示留言数
'rs.pagesize=e_page
'if request.querystring("page")="" or request.querystring("page")=0 then
'page=1
'else
'page=request.querystring("page")
'rs.absolutepage=trim(request.querystring("page"))
'end if
%>
      <SCRIPT language=JavaScript>

function is_number(str)
{
	exp=/[^0-9()-]/g;
	if(str.search(exp) != -1)
	{
		return false;
	}
	return true;
}

function CheckInput(){

	if(form.name.value==''){
		alert("您没有填写昵称！");
		form.name.focus();
		return false;
	}
	if(form.name.value.length>20){
		alert("昵称不能超过20个字符！");
		form.name.focus();
		return false;
	}

	if(!is_number(document.form.qq.value)){
		alert("QQ号必须是数字！");
		form.qq.focus();
		return false;
	}

	if(form.content.value==''){
		alert("您没有填写留言内容！");
		form.content.focus();
		return false;
	}
	if(form.content.value.length>255){
		alert("留言内容不能超过255个字符！");
		form.content.focus();
		return false;
	}
	
	return true;
}
</SCRIPT>
      <br> 
      <%
if rs.eof and rs.bof then
%>
      <table width="380" border="0" align="center" cellpadding="4" >
        <tr> 
          <td height="40" align="center"> <p>还没有任何留言！</p></td>
        </tr>
      </table>
      <%else%>
      <%
	  	rs.PageSize =5 '每页记录条数
		iCount=rs.RecordCount '记录总数
		iPageSize=rs.PageSize
    	maxpage=rs.PageCount 
    	page=request("page")
    	per_page=rs.PageSize
    
    if Not IsNumeric(page) or page="" then
        page=1
    else
        page=cint(page)
    end if
    
    if page<1 then
        page=1
    elseif  page>maxpage then
        page=maxpage
    end if
    
    rs.AbsolutePage=Page

	if page=maxpage then
		x=iCount-(maxpage-1)*iPageSize
	else
		x=iPageSize
	end if
%>
      <table width="100%" border="0">
        <tr> 
          <td colspan="12" height="25" align="center" bgcolor="#FFFFFF" > 
            <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>
          </td>
        </tr>
      </table>
      <%
For i=1 To x
'do while not rs.eof and e_page>0
%>
      <TABLE width=98% border=0 align="center" cellPadding=0 cellSpacing=1 bgcolor="#CCCCCC">
        <TBODY>
          <TR> 
            <TD bgcolor="f1f1f1"><TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                <TBODY>
                  <TR bgcolor="f1f1f1"> 
                    <TD width="67" height="23" align="center" bgcolor="f1f1f1"><font color="#FF0000"><strong>第<%=i%>条</strong></font></TD>
                    <TD width="464" align="left"> 
                       <%if rs("s_sex")<>"" then%>
                      性别：<%=rs("s_sex")%> 
                      <%end if%>
                      &nbsp;
                      <%if rs("s_tel")<>"" then%>
                      手机：
                      <%=rs("s_tel")%> 
                      <%end if%>
                      &nbsp; 
                       <%if rs("s_phone")<>"" then%>
                      电话：<%=rs("s_phone")%> 
                      <%end if%>
                      &nbsp; 
                      <%if rs("s_email")<>"" then%>
                      <a href="mailto:<%=rs("s_email")%>" target="_blank" title=Email：<%=rs("s_email")%>>
                      邮箱：<%=rs("s_email")%></a> 
                      </a> 
                      <%end if%>
                      &nbsp;
                      <%if rs("s_address")<>"" then%>
                      地址：<%=rs("s_address")%> 
                      <%end if%>
                      &nbsp;
                      <%if rs("s_fax")<>"" then%>
                      传真：<%=rs("s_fax")%> 
                      <%end if%>
                      &nbsp;
                      <%if rs("s_company")<>"" then%>
                      公司名称：<%=rs("s_company")%> 
                      <%end if%></TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
          </TR>
          <TR> 
            <TD><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY>
                  <TR> 
                    <TD bgcolor="#FFFFFF"> <TABLE width="100%" border=0 style="table-layout:fixed;word-break:break-all">
                        <TBODY>
                          <TR> 
                            <TD width="64"></TD>
                            <TD colspan="2"> <strong><font color="">姓名：</font></strong><font color="<%=flzhzt%>"><%=rs("s_name")%></font><br> <br>
                              <strong><font color="<%=flzhbt%>">内容：</font></strong><font color="<%=flzhzt%>"><%=unHtml(rs("s_content"))%></font></TD>
                          </TR>
                          <TR> 
                            <TD colspan="2">
							<%
							if rs("s_ok") then
							w "&nbsp; [<a href=""?action=s_ok&s_ok=0&id="&rs("id")&"""><font color=blue>已审核</font></a>] "
							else
							w "&nbsp; [<a href=""?action=s_ok&s_ok=1&id="&rs("id")&"""><font color=red>审核</font></a>] "
							end if
							%>
													&nbsp; [<a href="guestbook_reply.asp?id=<%=rs("id")%>">回复</a>] 
                              &nbsp; [<a href="guestbook_del.asp?id=<%=rs("id")%>">删除</a>] 
                            </TD>
                            <TD width="250" align="right"><%=rs("s_time")%></TD>
                          </TR>
                        </TBODY>
                      </TABLE></TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
          </TR>
          <%if rs("s_reply")<>"" then%>
          <TR> 
            <TD bgColor=#f2f2f2> <table width="100%" border="0" style="table-layout:fixed;word-break:break-all">
                <tr> 
                  <td width="10">&nbsp;</td>
                  <td width="567"><font color="#FF0000">管理员回复：</font><br> <%=unHtml(rs("s_reply"))%></td>
                </tr>
              </table></TD>
          </TR>
          <%end if%>
        </TBODY>
      </TABLE>
      <br> 
      <%
'e_page=e_page-1
'rs.movenext
'loop
		RS.MoveNext
next
%>
      <table width="100%" border="0">
        <tr> 
          <td> 
            <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>
          </td>
        </tr>
      </table>
      <%
rs.close
set rs=nothing
end if
%>
    </td>
  </tr>
</table>
</body>
</html><%
Sub PageControl(iCount,pagecount,page,table_style,font_style,per_page)
'生成上一页下一页链接
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
	
    temp=""

    Response.Write("<table " & Table_style & ">" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<TD align=right>" & vbCrLf )
    Response.Write(font_style & vbCrLf )    
        
    if page<=1 then
        Response.Write ("首页 " & vbCrLf)        
        Response.Write ("上页 " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>首页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">上页</A> " & vbCrLf)
    end if

    if page>=pagecount then
        Response.Write ("下页 " & vbCrLf)
        Response.Write ("尾页 " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">下页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">尾页</A> " & vbCrLf)            
    end if

    Response.Write(" 页次：" & page & "/" & pageCount & "页" &  vbCrLf)
    Response.Write(" 共有" & iCount & "条/每页"&per_page&"条" &  vbCrLf)
    Response.Write(" 转到" & "<INPUT TYEP=TEXT NAME=page SIZE=4 Maxlength=8 VALUE=" & page & ">" & "页"  & vbCrLf & "<INPUT type=submit style=""font-size: 9pt"" value=GO class=b2>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%>
