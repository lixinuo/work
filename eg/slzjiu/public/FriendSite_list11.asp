<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<%
const C_url="FriendSite_list.asp"
dim strFileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim sql,rs,ID,LinkType
dim Action,FoundErr,ErrMsg
Show_img=request("Show_img")
if Show_img="" then Show_img=0 else Show_img=cint(Show_img)
Action=trim(request("Action"))
ID=Trim(Request("ID"))
LinkType=trim(request("LinkType"))
strFileName=C_url&"?LinkType=" & LinkType&"&Show_img="&Show_img

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

if ID<>"" then
	if Action="Check" then
		conn.execute "Update O_links set IsOK=True where ID=" & CLng(ID)
	elseif Action="CancelCheck" then
		conn.execute "Update O_links set IsOK=False Where ID=" & CLng(ID)
	elseif Action="Good" then
		conn.execute "Update O_links set IsGood=True Where ID=" & CLng(ID)
	elseif Action="CancelGood" then
		conn.execute "Update O_links set IsGood=False Where ID=" & CLng(ID)
	elseif Action="Del" then
		conn.execute "Delete From O_links Where ID=" & CLng(ID)
	end if
end if
%>
<html>
<head>
<title>友情链接管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="../CSS/styleyullhao.css">
<script LANGUAGE="javascript">
function Check() {
if (document.AddLink.s_name.value=="")
	{
	  alert("请输入链接名称！")
	  document.AddLink.s_name.focus()
	  return false
	 }
if (document.AddLink.s_url.value=="")
	{
	  alert("请输入链接地址！")
	  document.AddLink.s_url.focus()
	  return false
	 }
}

function ConfirmDel()
{
   if(confirm("确定要删除此友情链接站点吗？"))
     return true;
   else
     return false;
}
</script>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
.STYLE56 {
	color: #FF0000;
	font-weight: bold;
}
.STYLE57 {color: #0000FF}
.STYLE58 {color: #FF0000}
-->
</style>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" bgcolor="<%=Color_0%>">
<%=Upload_Init()%>
<div align="center">
 <table width="99%" border="0" cellpadding="2" cellspacing="1" Class="border" bgcolor="#FFFFFF">
  <tr class="topbg">
   <td height="22" colspan=2 align=center bgcolor="<%=Color_1%>"><b> 友 情 链 接</b></td>
  </tr>
  <tr bgcolor="<%=Color_0%>" class="tdbg">
   <td width="70" height="30">管理导航：</td>
   <td height="30"><a href="FriendSite_list.asp?Action=Add&Show_img=<%=Show_img%>">添加友情链接</a>&nbsp;</td>
  </tr>
 </table>
</div>
<br>
<%
if Action="Add" then
	call Add()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
else
	call main()
end if
call CloseConn()

sub main()
	sql="select * from O_links where id<>0 "
	sql=sql & "order by id desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	
  	if rs.eof and rs.bof then
		response.write "目前共有 0 个友情链接"
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
        	showpage strFileName,totalput,MaxPerPage,true,true,"个友情链接"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"个友情链接"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"个友情链接"
	    	end if
		end if
	end if
	rs.close
	set rs=nothing
end sub

sub showContent
   	dim i
    i=0
%>
<div align="center">
 <table width="99%" border="0" cellspacing="1" Class="border" bgcolor="#FFFFFF">
  <tr bgcolor="<%=Color_0%>" class="title">
	<%if Show_img=1 then%>
   <td width="100" align="center">LOGO</td>
	 <%end if%>
   <td width="100" height="22" align="center">链接名称</td>
   <td width="193" height="22" align="center">链接简介</td>
   <td width="131" height="22" align="center">操作</td>
  </tr>
  <%
do while not rs.eof

S_logo=rs("s_img")
if instr(S_logo,"Uploadfiles") then S_logo="../../"&S_logo
%>
  <tr class="tdbg">
	<%if Show_img=1 then%>
   <td width="100" align="center" bgcolor="<%=Color_0%>"><img src="<%=S_logo %>" width="88" height="33" alt=""></td>
	 <%end if%>
   <td width="100" bgcolor="<%=Color_0%>"><a href="<%=rs("s_url")%>" target='blank' title="<%=rs("s_url")%>"><%=rs("s_name")%></a></td>
   <td width="193" bgcolor="<%=Color_0%>"><%=left(trim(rs("S_content")),13)%>...</td>
   <td width="131" align="center" bgcolor="<%=Color_0%>"><%
	  response.write "<a href='"&C_url&"?Action=Modify&Show_img="&Show_img&"&ID=" & rs("ID") & "'>修改</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      response.write "<a href='"&C_url&"?Action=Del&Show_img="&Show_img&"&ID=" & rs("ID") & "' onclick='return ConfirmDel();'>删除</a>"
	  %>
   </td>
  </tr>
  <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
 </table>
</div>
<%
end sub

sub Add()
%>
<form method="post" name="AddLink" onSubmit="return Check()" action="FriendSite_list.asp?Show_img=<%=Show_img%>">
 <div align="center">
  <table border="0" cellpadding="2" cellspacing="1" width="99%" class="border" bgcolor="#FFFFFF">
   <tr class="title">
    <td height="22" colspan="2" align="center" bgcolor="<%=Color_1%>">添加友情链接</td>
   </tr>
   <tr class="tdbg">
    <td width="198" height="25" valign="middle" bgcolor="<%=Color_0%>">链接名称：</td>
    <td width="691" height="25" bgcolor="<%=Color_0%>"><input name="s_name" class="input" title="这里请输入您的网站名称，最多为20个汉字" size="30"  maxlength="20">
     <font color="#FF0000"> *</font> </td>
   </tr>
   <%if instr("1",webLanguage) then%>
   <tr class="tdbg">
    <td width="198" height="25" valign="middle" bgcolor="<%=Color_0%>">En：</td>
    <td width="691" height="25" bgcolor="<%=Color_0%>"><input name="s_name1" class="input" title="这里请输入您的网站名称，最多为20个汉字" size="30"  maxlength="20">
     <font color="#FF0000"> *</font> </td>
   </tr>
   <%end if%>
   <tr class="tdbg">
    <td width="198" height="25" bgcolor="<%=Color_0%>">链接地址：</td>
    <td height="25" bgcolor="<%=Color_0%>"><input name="s_url" type="text" class="input" title="这里请输入您的网站地址，最多为50个字符，前面必须带http://"  value="http://" size="30"  maxlength="100">
     <font color="#FF0000">*</font></td>
   </tr>
   <%if Show_img=1 then%>
   <tr class="tdbg">
    <td width="198" height="25" bgcolor="<%=Color_0%>">链接LOGO：</td>
    <td height="25" bgcolor="<%=Color_0%>">
    <%=Upload_Input("S_logo","")%>
     <span class="STYLE58"> 填写了LOGO为图片链接，不填写为文字链接 </span></td>
   </tr>
   <%end if%>
   <tr class="tdbg">
    <td width="198" bgcolor="<%=Color_0%>">链接简介：</td>
    <td valign="middle" bgcolor="<%=Color_0%>"><textarea name="S_content" cols="40" rows="5" class="input" id="S_content" title="这里请输入您的网站的简单介绍"></textarea></td>
   </tr>
   <%if instr("1",webLanguage) then%>
   <tr class="tdbg">
    <td width="198" bgcolor="<%=Color_0%>">En：</td>
    <td valign="middle" bgcolor="<%=Color_0%>"><textarea name="S_content1" cols="40" rows="5" class="input" id="S_content1" title=""></textarea></td>
   </tr>
   <%end if%>
   <tr class="tdbg">
    <td height="40" colspan="2" align="center" bgcolor="<%=Color_0%>"><input name="Action" type="hidden" id="Action" value="SaveAdd">
     <input name="cmdOk" type="submit" class="input" value=" 确 定 ">
     &nbsp;
     <input name="cmdReset" type="reset" class="input" value=" 重 填 ">
    </td>
   </tr>
  </table>
 </div>
</form>
<%
end sub

sub Modify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定友情站点ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	dim sqlLink,rsLink
	sqlLink="select * from O_links where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到友情链接！</li>"
		rsLink.close
		set rsLink=nothing
		exit sub
	end if

%>
<form method="post" name="AddLink" onSubmit="return Check()" action="FriendSite_list.asp?Show_img=<%=Show_img%>">
 <div align="center">
  <table border="0" cellpadding="2" cellspacing="1" width="99%" class="border" bgcolor="#FFFFFF">
   <tr class="title">
    <td height="22" colspan="2" align="center" bgcolor="<%=Color_1%>">修改友情链接</td>
   </tr>
   <tr class="tdbg">
    <td width="199" height="25" valign="middle" bgcolor="<%=Color_0%>">链接名称：</td>
    <td width="690" height="25" bgcolor="<%=Color_0%>"><input name="s_name" class="input" title="这里请输入您的链接名称，最多为20个汉字" value="<%=rsLink("s_name")%>" size="30"  maxlength="20">
     <font color="#FF0000"> *</font>
    </td>
   </tr>
     <%if instr("1",webLanguage) then%>
   <tr class="tdbg">
    <td width="199" height="25" valign="middle" bgcolor="<%=Color_0%>">EN：</td>
    <td width="690" height="25" bgcolor="<%=Color_0%>"><input name="s_name1" class="input" value="<%=rsLink("s_name1")%>" size="30"  maxlength="20">
     <font color="#FF0000"> *</font>
    </td>
   </tr>
     <%end if%>
   <tr class="tdbg">
    <td width="199" height="25" bgcolor="<%=Color_0%>">链接地址：</td>
    <td height="25" bgcolor="<%=Color_0%>"><input name="s_url" type="text" class="input" title="这里请输入您的链接地址，最多为50个字符，前面必须带http://"  value="<%=rsLink("s_url")%>" size="30"  maxlength="100">
     <font color="#FF0000">*</font></td>
   </tr>
	  <%if Show_img=1 then%>
   <tr class="tdbg">
    <td width="199" height="25" bgcolor="<%=Color_0%>">Logo地址：</td>
    <td height="25" bgcolor="<%=Color_0%>"><%=Upload_Input("S_logo",trim(rsLink("S_img")))%>
     <span class="STYLE58"> 填写了LOGO为图片链接，不填写为文字链接 </span></td>
   </tr>
	 <%end if%>
   <tr class="tdbg">
    <td width="199" bgcolor="<%=Color_0%>">链接简介：</td>
    <td valign="middle" bgcolor="<%=Color_0%>"><textarea name="S_content" cols="40" rows="5" class="input" id="S_content" title="这里请输入您的链接的简单介绍"><%=rsLink("S_content")%></textarea></td>
   </tr>
<%if instr("1",webLanguage) then%>
   <tr class="tdbg">
    <td width="199" bgcolor="<%=Color_0%>">EN：</td>
    <td valign="middle" bgcolor="<%=Color_0%>"><textarea name="S_content" cols="40" rows="5" class="input" id="S_content1" ><%=rsLink("S_content1")%></textarea></td>
   </tr>
<%end if%>
   <tr class="tdbg">
    <td height="40" colspan="2" align="center" bgcolor="<%=Color_0%>"><input name="ID" type="hidden" id="ID" value="<%=rsLink("ID")%>">
     <input name="Action" type="hidden" id="Action" value="SaveModify">
     <input name="cmdOk" type="submit" class="input" value=" 确 定 ">
     &nbsp;&nbsp;
     <input name="BT" type="button" class="input" id="BT" style="cursor:hand; height:20; " onClick="history.go(-1);" value=" 返 回 &gt;&gt;">
    </td>
   </tr>
  </table>
 </div>
</form>
<%
	rsLink.close
	set rsLink=nothing
end sub
%>
</body>
 
</html>
<%

sub SaveAdd()
	LinkSiteName=trim(request("s_name"))
	LinkSiteName1=trim(request("s_name1"))
	LinkSiteUrl=trim(request("s_url"))
	LInkSiteIntro=trim(request("S_content"))
	LInkSiteIntro1=trim(request("S_content1"))
	LInkLogoUrl=replace(trim(request("S_logo")),"../","")
	if LinkSiteName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>链接名称不能为空！</li>"
	end if
	if LinkSiteUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>链接地址不能为空！</li>"
	end if
	if FoundErr<>True then
		dim sqlLink,rsLink
		sqlLink="select top 1 * from O_links where s_name='" & (LinkSiteName) & "' and s_url='" & (LinkSiteUrl) & "'"
		set rsLink=Server.CreateObject("Adodb.RecordSet")
		rsLink.open sqlLink,conn,3,3
		if not (rsLink.bof and rsLink.eof) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>你要添加的链接已经存在！</li>"
		else
			rsLink.Addnew
			rsLink("s_name")=(LinkSiteName)
			rsLink("s_name1")=(LinkSiteName1)
			rsLink("s_url")=(LinkSiteUrl)
			rsLink("S_content")=(LinkSiteIntro)
			rsLink("S_content1")=(LinkSiteIntro1)
			rsLink("S_img")=(LInkLogoUrl)
			rsLink.update
			rsLink.close
			set rsLink=nothing
			call CloseConn()
			Response.Redirect strFileName
		end if
		rsLink.close
		set rsLink=nothing
	end if
end sub

sub SaveModify()
	if ID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定链接ID</li>"
		exit sub
	else
		ID=Clng(ID)
	end if
	LinkSiteName=trim(request("s_name")):LinkSiteName1=trim(request("s_name1"))
	LinkSiteUrl=trim(request("s_url"))
	LInkSiteIntro=trim(request("S_content")):LInkSiteIntro1=trim(request("S_content1"))
	LInkLogoUrl=replace(trim(request("S_logo")),"../","")
	if LinkSiteName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>链接名称不能为空！</li>"
	end if
	if LinkSiteUrl="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>链接地址不能为空！</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	dim sqlLink,rsLink
	sqlLink="select * from O_links where ID=" & ID
	set rsLink=Server.CreateObject("Adodb.RecordSet")
	rsLink.open sqlLink,conn,1,3
	if rsLink.bof and rsLink.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到链接！</li>"
	else
		rsLink("s_name")=(LinkSiteName)
		rsLink("s_name1")=(LinkSiteName1)
		rsLink("s_url")=(LinkSiteUrl)
		rsLink("S_content")=(LinkSiteIntro)
		rsLink("S_content1")=(LinkSiteIntro1)
		rsLink("S_img")=(LInkLogoUrl)
		rsLink.update
		rsLink.close
		set rsLink=nothing
		call CloseConn()
		Response.Redirect strFileName
	end if
	rsLink.close
	set rsLink=nothing
end sub
%>
