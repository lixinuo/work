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
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站管理系统</title>
<!--#include file="../inc/g_links.asp"-->

<script src="../js/jquery-1.3.1.js" type="text/javascript"></script>

<script type="text/javascript">
function change_alist(alistid,pid)
{
 var	aid=alistid;
 var    pid=pid;
 var	aid_str="#alist"+aid;
 var tn=document.getElementById("alist"+aid).value;
// alert(tn);
// return;
  $.ajax({
   url:"ajax.asp",type:"post",dataType:"html",data:"tn="+tn+"&tm="+pid+"&t="+new Date().getTime(),
	  beforeSend:function(){$("#mmResult").html("AjAX……");},
	  error:function(){$("#mmResult").html("Failed!");},
	  success:function(text){$("#mmResult").html(text);}
  });
}

</script>

</head>
<body text="#000000" >
<table width="100%"   border="0" cellpadding="5" cellspacing="0">
  <tr>
    <td height="50" bgcolor="<%=Color_1%>"><span class=" fs_14px fc_Black ff_bold">
      <li></li>
    <%if id="" then w("增加") else w("修改")%>权限</span></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#004A80"></td>
  </tr>
  <tr>
    <td bgcolor="#DDEEFF">
 <form method="POST" action="?action=save" name="myform">
 <input type="hidden" name="id" value="<%=id%>">
  <table width="90%" height="103" border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
    <td height="3"  bgcolor="#ff9900">      </td>
  </tr>
  <tr>
    <td height="30" valign="middle" bgcolor="#DDEEFF" class="ff_bold">
    权限名称:<input type="text" name="gradename" id="gradename" value="<%=gradename%>"></td>
    </tr>
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
      <td class="ff_bold" style="border-top:dashed 2px #FFFFFF;"><%=rs(1)%><input id="alist<%=rs(0)%>" name="alist" onClick="change_alist(<%=id%>,<%=Rs(0)%>)" type="checkbox" value="<%=rs(0)%>" 
    
 <%if gradealist<>"" then%>   <%
		  		
				 aa=split(gradealist,",")
				 for ai=0 to ubound(aa)
				 if rs(0)&""=trim(aa(ai)) then w("checked='checked'")
         next
				 %> <%end if%>   >     
    	 
 				 
				
				 </td>
     </tr>
     

     
     
<%
rs.MoveNext
loop
end if
%>
 
    </table>
   </td>
  </tr>
  
  
  <tr>
    <td height="30" valign="middle" bgcolor="#DDEEFF">        <table width="100%"  border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="18%">　</td>
          <td width="54%" valign="middle">
          
          
          <div id="mmResult"></div>
          
          </td>
          <td width="26%" valign="middle"></td>
        </tr>
      </table></td>
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