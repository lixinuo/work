<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="connection.asp" -->
<html>
<head>
<title><%= webname%>网站管理系统</title>
<style type="text/css">
.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt;}
.a2{BACKGROUND-COLOR: A4B6D7;}
.body_toper{ width:100%; background:url(images/top_bg.jpg) repeat-x; height:30px; line-height:28px; padding-top:2px; font-size:12px;}
.body_toper_left{ float:left; padding-left:50px; background:url(images/icon_member.jpg) no-repeat 30px 2px;}
.body_header{ width:100%; height:116px; margin:14px 10px 6px 10px; overflow:hidden; background-color:#000; }
</style>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<BODY leftMargin=0 topMargin=0 rightMargin=0 bgcolor="#3A96BB">
<div class="body_toper">
   <div class="body_toper_left">尊敬的 <%=request.Cookies(Cookies_name)("Name")%> 您好！今天是 <%=now()%> 您上次登录时间<%=request.Cookies(Cookies_name)("LastLoginTime")%></div>
   <div class="clear"></div>
</div>
</body>
<body scroll="no" onResize="javascript:parent.carnoc.location.reload();">
<script type="text/javascript">
if(self!=top){top.location=self.location;}
function switchSysBar(){
	if (switchPoint.innerText==3){
		switchPoint.innerText=4
		document.all("frmTitle").style.display="none"
	}else{
	switchPoint.innerText=3
	document.all("frmTitle").style.display=""
	}
}
if(window.screen.width<'1024'){switchSysBar()}
</script>
<table border="0" cellPadding="0" cellSpacing="0" height="93%" width="100%">
    <tr>
        <td align="middle" noWrap vAlign="center" id="frmTitle">
        <iframe frameBorder="0" id="carnoc"  name="carnoc" src="System_menu.asp" style="HEIGHT: 100%; margin:0px; padding:0px; VISIBILITY: inherit; WIDTH: 196px; Z-INDEX: 2" target="main">
        </iframe>
    </td>
    <td style="WIDTH: 9pt">
        <table border="0" cellPadding="0" cellSpacing="0" height="100%">
            <tr>
            <td style="HEIGHT: 100%" onClick="switchSysBar()">
            <font style="FONT-SIZE: 9pt; CURSOR: default; COLOR: #ffffff">
            <span class="navPoint" id="switchPoint" title="关闭/打开左栏">3</span><br>
            屏幕切换 </font></td>
            </tr>
        </table>
    </td>
    <td style="WIDTH: 100%">
        <iframe frameborder="0" id="main" name="main" src="System_first.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"></iframe></td>
    </tr>
</table>
</body>
</html>