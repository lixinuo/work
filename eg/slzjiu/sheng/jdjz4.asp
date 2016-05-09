<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>页面生成进度加载</title>
</head>

<body>

<script language="javascript">  
window.onload=function(){    
var a = document.getElementById("loading");    
a.parentNode.removeChild(a);    
}   
document.write('<div id="loading" style="background:#FF0000;color:#FFFFFF;width:250px;">一键生成页面正在生成中，请稍候……</div>');    
</script>  
<!--把下面代码改为您的网页内容-->  
<iframe src="qdy4.asp" width="100%" height="500" frameborder="0" marginwidth="0" marginheight="0"></iframe>  
</body>
</html>
