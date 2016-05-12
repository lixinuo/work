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
      <div align="center" class="STYLE1"><font color="#000000">查看产品信息反馈</font></div>    </td>
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
conn.execute("update o_pbook set s_ok="&s_ok&" where id="&id)
end if

set rs=server.createobject("adodb.recordset")
rs.open "select * from o_pbook order by id desc",conn,3,2

%>
     
      <br> 
      <%
if rs.eof and rs.bof then
%>
      <table width="380" border="0" align="center" cellpadding="4" >
        <tr> 
          <td height="40" align="center"> <p>还没有任何反馈！</p></td>
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
                    <TD width="80" height="29" align="center" bgcolor="f1f1f1"><font color="#FF0000"><strong>第<%=i%>条</strong></font></TD>
                    <TD width="744" align="left" valign="middle"> 
                    
                     <strong>产品名称：</strong><%
					 sqlstr="select id,s_name from p_main where s_pai=0 and id="& rs("cpid") &""
					 set Rs_l=db(sqlstr,3)
					 if not Rs_l.Eof then
					 	response.Write"<a href='../../pro_details.asp?id="&Rs_l(0)&"' target='_blank'>"&Rs_l(1)&"</a>"
					 end if
					 %>                </TD>
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
                            <TD width="96"></TD>
                            <TD colspan="2"> <%if rs("S_company")<>"" then%>
                       <strong>国家：</strong><%=rs("S_company")%>
                      <%end if%>&nbsp;&nbsp;<%if rs("S_name")<>"" then%>
                       <strong>姓名：</strong><%=rs("S_name")%>
                      <%end if%>&nbsp;&nbsp;
					   <%if rs("S_www")<>"" then%>
                       <strong>MSN：</strong><%=rs("S_www")%>
                      <%end if%>&nbsp;&nbsp;
					   <%if rs("S_email")<>"" then%>
                      <a href="mailto:<%=rs("S_email")%>" target="_blank" title=Email：<%=rs("S_email")%>>
                       <strong>邮箱：</strong><%=rs("S_email")%></a> 
                      <%end if%> <br> <!--<strong><font color="">姓名：</font></strong><font color="< %=flzhzt%>">< %=rs("s_name")%></font><br> <br>-->  <strong><font color="<%=flzhbt%>">主题：</font></strong><font color="<%=flzhzt%>"><%=rs("S_address")%></font>
                              <br><strong><font color="<%=flzhbt%>">内容：</font></strong><font color="<%=flzhzt%>"><%=unHtml(rs("s_content"))%></font></TD>
                          </TR>
                          <TR> 
                            <TD colspan="2">
							<%
'							if rs("s_ok") then
'							w "&nbsp; [<a href=""?action=s_ok&s_ok=0&id="&rs("id")&"""><font color=blue>已审核</font></a>] "
'							else
'							w "&nbsp; [<a href=""?action=s_ok&s_ok=1&id="&rs("id")&"""><font color=red>审核</font></a>] "
'							end if
							%>
	<!--													&nbsp; [<a href="guestbook_reply.asp?id=<%=rs("id")%>">回复</a>] -->
                              &nbsp; [<a href="guestbook_del.asp?id=<%=rs("id")%>">删除</a>] &nbsp;&nbsp;
                            [
							<%
							if rs("s_ok")=0 then
								response.Write"<a href=""order_list.asp?id="&rs("id")&"&action=s_ok&s_ok=1"">未处理</a>"
							else
								response.Write"<a href=""order_list.asp?id="&rs("id")&"&action=s_ok&s_ok=0"">已处理</a>"
							end if
							%>
							
							]</TD>
                            <TD width="250" align="right"><%=rs("s_time")%></TD>
                          </TR>
                        </TBODY>
                      </TABLE></TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
          </TR>
         
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
