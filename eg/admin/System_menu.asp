<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"

dim bc1,bc2
bc1="#6E0200"
bc2="#fff"
if request.Cookies(Cookies_name)("Name")="" then 
	response.Write(" ")
else
  Response.Expires = 0  
  Response.ExpiresAbsolute = Now() - 1  
  Response.cachecontrol = "no-cache" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="images/cssyullhao.css">
<base target="main">
<style type="text/css">
.style58 {color: #fff;font-weight: bold;font-family:宋体;font-size:12px;}
.style59 {color: #000000;padding-left:25px;}
.name_grade { color:red; font-weight:bold;}
a:hover { text-decoration:none;}
</style>
<SCRIPT type="text/javascript">
function menuChange(obj,menu)
{
	if(menu.style.display=="")
	{ 
	  menu.style.display="none";
	  obj.background="images/menudown.gif" ; 
	}else{
	   
	   menu.style.display="";
	  obj.background="images/menuup.gif" ;
	}
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" oncontextmenu="self.event.returnValue=false;" onselectstart="return false;" onMouseOver="window.status='欢迎光临 <%= request.Cookies(Cookies_name)("Webname")%>--网站后台管理系统';return true;" >

 <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
   <tr>
     <td width="100%" align="center"><table width="179" border="0" align="center" cellpadding="0" cellspacing="0" id="table1">
       <tr>
         <td align="right"><br>
             <table align="center" width="179" border="0" cellpadding="0" cellspacing="0">
               <tr>
                 <TD width="179" height="25" align="center" style=" cursor:hand; background:url(images/forum_footer.gif) no-repeat;"><span class="style58">｜网站管理系统｜</span></td>
               </tr>
               <tr>
                 <TD width="166" height="30" align="left"><table width="179" border="0" cellpadding="0" cellspacing="0" bgcolor="#F6FBFF">
                     <tr>
                       <td width="179"><table width="155" border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="#ffffff">
                           <tr>
                             <td width="175" align="left" ><span class="style59">用户:</span><span class="name_grade"><%=request.Cookies(Cookies_name)("Name")%></span></td>
                           </tr>
                           <tr>
                             <td align="left" ><span class="style59">身份:</span><span class="name_grade"><%=request.Cookies(Cookies_name)("Grade")%></span></td>
                           </tr>
                           <tr align="center">
                             <td><a href="SyStem_first.ASP">返回首页</a> ｜ <a target="_top" href="SyStem_LOGOUT.ASP?target=exit" onClick="return confirm('是否退出管理？');">退出管理</a> </td>
                           </tr>
                       </table></td>
                     </tr>
                     <tr>
                       <td height="8"></td>
                     </tr>
                 </table></td>
               </tr>
               <tr>
                 <td><table width="179" border="0" cellpadding="0" cellspacing="0" bgcolor="#F6FBFF">
                     <tr>
                       <td width="179"><table width="166" border="0" align="center" cellpadding="0" cellspacing="0">
                           <!--#include file="menu/menu_class.asp" -->
                       </table></td>
                     </tr>
                 </table></td>
               </tr>
               <tr>
                 <td width="100%" border="0" cellpadding="2">&nbsp;</td>
               </tr>
           </table></td>
       </tr>
     </table></td>
   </tr>
</table>
</body>
</html>
<%end if%>
