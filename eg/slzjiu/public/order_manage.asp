<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<html><head><title>表格管理</title>
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
.style55 {color: #666666}
body,td,th {
	color: #666666;
}
.style56 {color: #FF0000}
-->
</style>
<link href="../images/cssyullhao.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
dim selectm,selectkey,selectbookid
selectkey=trim(request(trim("selectkey")))
selectm=trim(request("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectkey="请输入关键字" then
selectkey=""
end if
'//删除商品
if selectm="" then
selectm=request.QueryString("selectm")
end if
selectbookid=request("selectbookid")
if selectbookid<>"" then
conn.execute "delete from o_s_orders where id in ("&selectbookid&")"
response.Redirect "order_manage.asp"
response.End
end if
%>
<table width="100%"   border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="50" bgcolor="<%=Color_1%>"><span class="style1">
    <li>信息列表</li>
    </span></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#004A80"></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="<%=Color_0%>">    
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#DDEEFF">
<tr> 
<form name="form1" method="post" action="">
<td height="100"> 
				<%'开始分页
				Const MaxPerPage=20
   				dim totalPut   
   				dim CurrentPage
   				dim TotalPages
   				dim j
   				dim sql
                dim rs
    				if Not isempty(request("page")) then
      				currentPage=Cint(request("page"))
   				else
      				currentPage=1
   				end if 
      				s_type=Cint(request("s_type")) 
			set rs=server.CreateObject("adodb.recordset")

           		rs.open "select * from o_s_orders where s_type="&s_type&" order by id desc",conn,1,1

		
		   	if err.number<>0 then
				response.write "数据库中无数据"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> 数据库中无数据！</p>"
   				else
	  				totalPut=rs.recordcount

      				if currentpage<1 then
          				currentpage=1
      				end if

      				if (currentpage-1)*MaxPerPage>totalput then
	   					if (totalPut mod MaxPerPage)=0 then
	     					currentpage= totalPut \ MaxPerPage
	   					else
	      					currentpage= totalPut \ MaxPerPage + 1
	   					end if
      				end if

       				if currentPage=1 then
            			showContent
            			showpage totalput,MaxPerPage,"order_manage.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"order_manage.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"order_manage.asp"
	      				end if
	   				end if
   				   	end if

   				sub showContent
       				dim i
	   			i=0%>

                <table width="500" border="0" align="center"cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr bgcolor="#DDEEFF"> 
<td width="20%" align="center">信息序号</td>
<td width="32%" align="center">阅读信息</td>

<td width="11%" align="center">选 择</td>
</tr>
<%do while not rs.eof
%>
<tr bgcolor="#DDEEFF"> 
<td align="center"> <a href=order_read.asp?id=<%=rs("id")%>><%=rs("id")%></a></td>
<td width="32%" align="center" STYLE='PADDING-LEFT: 10px'><a href=# onClick="javascript:window.open('order_read.asp?id=<%=rs("id")%>','_blank','width=400,height=300,scrollbars=yes')">阅读信息</a>
</td>
<td align="center">
<input name="selectbookid" type="checkbox" id="selectbookid" value="<%=rs("id")%>"></td></tr>
          <%i=i+1
		if i>=MaxPerPage then Exit Do
		rs.movenext
		loop
		rs.close
		set rs=nothing%>
<tr bgcolor="#f3f3f3"> 
<td height="30" colspan="6" align="right" bgcolor="#DDEEFF">全选 
<input type="checkbox" name="checkbox" value="Check All" onClick="mm()">
<input type="submit" name="Submit" value="删 除" onClick="return test();">
&nbsp;</td>
</tr>
</table>
				<%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				
				Response.Write "<form method=Post action="&filename&"?s_type="&s_type&"&selectm="&selectm&"&selectkey="&selectkey&" >"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>首页 上一页</font> "  
				Else  
					Response.Write "<a href="&filename&"?s_type="&s_type&"&page=1&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?s_type="&s_type&"&page="&CurrentPage-1&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>上一页</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 尾页</font>"  
				Else  
					Response.Write "<a href="&filename&"?s_type="&s_type&"&page="&(CurrentPage+1)&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?s_type="&s_type&"&page="&n&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>尾页</a>"  
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有"&totalnumber&"条信息 " 
					Response.Write "<font class='contents'>转到：</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='GO' name='cndok' ></form>"  
				End Function  
			%>
<table width="12" height="7" border="0" cellpadding="0" cellspacing="0">
<tr> 
<td height=7></td>
</tr>
</table>
</td>
</form>
</tr>
</table></td></tr></table>
</body>
</html>
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
   for (var i=0; i<a.length; i++)
			a[i].checked = a[a.length-4].checked;
}
</script>