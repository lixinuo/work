<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
id=trim(request("id"))
action=trim(request("action"))
if id<>"" then
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from s_user_Grade where id="&id
rs.open strSql,Conn,1,1 
gradename=rs("gradename")
gradealist=rs("gradealist")
gradenlist=rs("gradenlist")
rs.close
set rs=nothing
end if

if action="save" then
	gradename=request("gradename")
	gradealist=request("alist")
	gradenlist=request("nlist")
if id="" then
	conn.execute("Insert into s_user_Grade(gradename,gradealist,gradenlist) values ('"&gradename&"','"&gradealist&"''"&gradenlist&"')")
else


	conn.execute("update s_user_Grade set gradename='"&gradename&"',gradealist='"&gradealist&"',gradenlist='"&gradenlist&"' Where id=" & id & "")
end if
	'msg "信息提交成功！","Grade_manage.asp"
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<!--#include file="../inc/g_links.asp"-->
</head>
<body text="#000000" >
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
  <tr>
    <td align="center"> <%if id="" then w("增加") else w("修改")%>权限</td>
  </tr>
</table>
<table width="98%"   border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td>
 <form method="POST" action="?action=save" name="myform">
 <input type="hidden" name="id" value="<%=id%>">
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
      <tr>
        <td class="lmtjdt">权限名称:<input type="text" name="gradename" id="gradename" value="<%=gradename%>"></td>
      </tr>
</table>

  <table width="100%" height="103" border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#FFFFFF">
  
  <tr>
   <td valign="middle" height="30" bgcolor="#DDEEFF">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
<%

set rs=server.CreateObject("adodb.recordset")
rs.Open "select id,s_name from S_menu_class where S_ok=1 and parent_id=0 order by S_order ",conn,1,1
if rs.EOF and rs.BOF then
response.Write "<div align=center><font color=red>还没有分类</font></center>"
else
do while not rs.EOF

%>
     <tr>
      <td class="ff_bold" style="border-top:dashed 2px #FFFFFF;"><%=rs(1)%><input id="alist<%=rs(0)%>" name="alist" type="checkbox" value="<%=rs(0)%>" onClick="mm(<%=rs(0)%>,<%=db_s("select count(id) from S_menu_class where parent_id="&rs(0)&" and s_ok=1")%>)" 
       <%if gradealist<>"" then%>   <%
		  		
				 aa=split(gradealist,",")
				 for ai=0 to ubound(aa)
				 if rs(0)&""=trim(aa(ai)) then w("checked='checked'")
         next
				 %> <%end if%>  >				 </td>
     </tr>
     
<%if gradealist<>"" then%>
     <tr>
      <td>
       <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
       <%
        set rs_s=server.CreateObject("adodb.recordset")
        rs_s.Open "select id,s_name from S_menu_class where parent_id="&rs(0)&" and s_ok=1 order by s_order",conn,1,1
        if rs_s.EOF and rs_s.BOF then
		    response.Write "<div align=center><font color=red>还没有菜单</font></center>"
		    else
				i=1
				tdnum=6
				tdwidth="17%"
        do while not rs_s.EOF
        %>
         <td><%=rs_s("s_name")%><input id="nlist<%=rs(0)%><%=i%>" name="nlist" type="checkbox" value="<%=rs_s(0)%>" 
				 <%
				 nn=split(gradenlist,",")
				 for ni=0 to ubound(nn)
				 if rs_s(0)&""=trim(nn(ni)) then w("checked='checked'")
         next
				 %>></td>
				<%
        if i mod tdnum = 0 then w("</tr><tr><td height=10></td></tr><tr>")
        rs_s.movenext 
        
        if rs_s.eof then
        if i>tdnum then t=i mod tdnum else t=tdnum-i
        for j=1 to t
        w("<td width='"&tdwidth&"'>&nbsp;</td>")
        next
        end if
        i=i+1
        loop
        end if
				rs_s.close
				set rs_s=nothing
        %>
         </tr>
       </table>      </td>
     </tr>
 <%else%>    
     
     
 <tr>
      <td>
       <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
       <%
        set rs_s=server.CreateObject("adodb.recordset")
        rs_s.Open "select * from S_menu_class where parent_id="&Rs(0)&" and s_ok=1 order by s_order",conn,1,1
        if rs_s.EOF and rs_s.BOF then
		    response.Write "<div align=center><font color=red>还没有菜单</font></center>"
		    else
				i=1
				tdnum=6
				tdwidth="17%"
        do while not rs_s.EOF
        %>
         <td><%=rs_s(1)%><input id="nlist<%=Rs(0)%><%=i%>" name="nlist" type="checkbox" value="<%=rs_s(0)%>" 
				 ></td>
				<%
        if i mod tdnum = 0 then w("</tr><tr><td height=10></td></tr><tr>")
        rs_s.movenext 
        
        if rs_s.eof then
        if i>tdnum then t=i mod tdnum else t=tdnum-i
        for j=1 to t
        w("<td width='"&tdwidth&"'>&nbsp;</td>")
        next
        end if
        i=i+1
        loop
        end if
				rs_s.close
				set rs_s=nothing
        %>
         </tr>
       </table>      </td>
     </tr>    
     
     
<%
end if
rs.MoveNext
loop
end if
%>
    </table>   </td>
  </tr>
  
  
     
   
  
  <tr>
    <td height="30" valign="middle" bgcolor="#DDEEFF">        <table width="100%"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="18%">　</td>
          <td width="54%" valign="middle"><input type="submit" name="button" value=" 确 定 " class=input3 onClick="cmdForm()">
&nbsp;</td>
          <td width="26%" valign="middle"></td>
        </tr>
      </table></td>
    </tr>
</table>
</form></td>
 </tr>
  <tr>
    <td height="2" bgcolor="#004A80"></td>
  </tr>
</table>
</body>
</html>
<script language=javascript>
function mm(aid,anum)
{
     var alist=aid;
	 var anum=anum;
	 var a= document.getElementById("alist"+alist);
	 window.status=a.checked;
   if( a.checked==true){
   for (var i=0; i<anum; i++)
      if (document.getElementById("nlist"+alist+(i+1)).type == "checkbox"){
			document.getElementById("nlist"+alist+(i+1)).checked = true;}
			
   }
   else
   {
   for (var i=0; i<anum; i++)
      if (document.getElementById("nlist"+alist+(i+1)).type == "checkbox"){
			document.getElementById("nlist"+alist+(i+1)).checked = false;}
   }
}
</script>