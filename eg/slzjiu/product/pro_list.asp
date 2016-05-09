<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data.asp" -->
<!--#INCLUDE FILE="../inc/gFunc_Page.asp" -->
<%
page=trim(request("page"))
classid=trim(request("classid"))
ccid=trim(request("ccid"))
skey=trim(request("skey"))
stype=trim(request("stype"))
s_pai=trim(request("s_pai"))
show_hot=trim(request("show_hot"))
session("show_hot")=1
if s_pai="" then
s_pai=0
else
s_pai=cint(s_pai)
end if
url="Pro_list.asp?s_pai="&s_pai&"&show_hot="&show_hot&"&page="&page

%>


<html>
<head>
<title>商品管理</title>
<meta http-equiv="Content-Type" S_content="text/html; charset=utf-8">

<link rel="stylesheet" rev="stylesheet" href="../inc/Greybox/gb_styles.css" type="text/css" media="all">
<script language="javascript" >var GB_ROOT_DIR ="../inc/Greybox/";</script>
<script language="javascript" src="../inc/Greybox/AJS.js" ></script>

<script language="javascript" src="../inc/Greybox/AJS_fx.js" ></script>
<script language="javascript" src="../inc/Greybox/gb_scripts.js" ></script>


<!--#include file="../inc/g_links.asp"-->
<style type="text/css">
<!--
.STYLE1 {font-weight: bold}
-->
</style>
</head>
<body text="#000000" >
<%
if request("action")="orderby" then
  strIDs=request("id")
	orderby=request("orderby")
	strSQL="update P_main set S_order="&orderby&" Where id=" & strIDs & ""
	conn.execute strSQL
end if
if request("action")="s_ok" then
  strIDs=request("id")
	s_ok=request("s_ok")
	strSQL="update P_main set s_ok="&s_ok&" Where id=" & strIDs & ""
	conn.execute strSQL
end if
if request("action")="s_xp" then
  strIDs=request("id")
	s_xp=request("s_xp")
	strSQL="update P_main set s_xp="&s_xp&" Where id=" & strIDs & ""
	conn.execute strSQL
end if


if request("action")="change_class" then'//删除全部
	selectid=request("selectid")
	change_classid=request("change_classid")
	if selectid<>"" and change_classid<>"" then
	  conn.execute "update P_main set classid="&change_classid&" where id in ("&selectid&")"
		response.Redirect url
		response.End
	end if
end if

'//删除商品图片
selectid=request("selectid")
if selectid<>"" then
set rspic=db("select S_img,S_img1 from P_main where id in ("&selectid&")",2)
if not rspic.eof then
do while not rspic.eof
 delpic1="../../"&rspic(0):delpic2="../../"&rspic(1)
 deletefile delpic1:deletefile delpic2
rspic.movenext
loop
end if
rspic.close:set rspic=nothing

''删除内容介绍中的在服务器上的图片
set rspic=db("select s_content from P_main where id in ("&selectid&")",2)
if not rspic.eof then
do while not rspic.eof

   Content=rspic(0) '''内容介绍

   arrayPic=Split(getPicUrl(Content),"|||")
   For I=0 To Ubound(arrayPic)
   deletefile arrayPic(I)
   Next
   
rspic.movenext
loop
end if
rspic.close:set rspic=nothing



conn.execute "delete from P_main where id in ("&selectid&")"
response.Redirect url
response.End
end if


'''创建目录函数
function creatFolder(foldername)

Set fs=Server.CreateObject("Scripting.FileSystemObject") 
if (fs.FolderExists(foldername)) then'判断今天的文件夹是否存在 
else 
fs.CreateFolder(foldername) 
end if 
set fs = Nothing
if err.number<>0 then err.clear

end function


'删除文件操作函数
function deletefile(filedir)
'on error resume next
dim fso
set fso = Server.CreateObject("Scripting.FileSystemObject")
if (fso.fileexists(SM(filedir))) then fso.deletefile(SM(filedir))
set fso = Nothing
if err.number<>0 then err.clear
end Function


'Function DelImagesFile(Images)
'FileName=server.MapPath(Images) '必须转换成绝对路径的
'dim Fso
'Set Fso = CreateObject("scripting.filesystemobject")
'if Fso.FileExists(filename) then '判断文件是否存在
'Fso.DelImagesFile filename 
'end if
'Set Fso = nothing
'End Function

''''得到图片的路径××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
function getPicUrl(str) '////////////////得到图片的路径
 dim content,regstr,url
 content=str&""
 regstr="src=.+?.(gif|jpg|png|swf)"
 url=Replace(Replace(Replace(RegExp_Execute(regstr,content),"'",""),"""",""),"src=","")
 getPicUrl=url
end function

'''''''''''''匹配图片的内容中图片的路径
Function RegExp_Execute(patrn, strng)
 Dim regEx, Match, Matches,values '建立变量。
 Set regEx = New RegExp '建立正则表达式。
 regEx.Pattern = patrn '设置模式。
 regEx.IgnoreCase = true '设置是否区分字符大小写。
 regEx.Global = True '设置全局可用性。
 Set Matches = regEx.Execute(strng) '执行搜索。
 For Each Match in Matches '遍历匹配集合。
 values=values&Match.Value&"|||" '将图片用特别符合隔开“|||”
 Next
 RegExp_Execute = values
End Function
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmdh">
    <tr>
      <td align="center">资料管理</td>
    </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
          <form name="search" method="post" action="<%=url%>">
            <tr>
              <td width="142" height="30" align="left" class="lmtjdt">按关键字查询：</td>
              <td width="204" class="lmtjdt"><input name="skey" type="text" id="skey" onFocus="this.value=''" value="请输入关键字">              </td>
              <td width="125" align="left" class="lmtjdt"><select name="stype" id="stype">
                  <option value="S_name">按资料名称</option>
                  <option value="S_content">按资料说明</option>
                  <option value="id">按资料序号</option>
                </select>              </td>
              <td width="223" height="30" align="left" class="lmtjdt"><input type="submit" class="inputkkys" name="Submit2" value="查询">              </td>
              <td width="292" align="left" class="lmtjdt"><input type="button" name="bb" class="inputkkys" onClick="location='pro_do.asp?s_pai=<%= s_pai %>'" value="添加信息">              </td>
            </tr>
            <tr>
              <td height=30 class="lmtjdt">按信息分类查询：</td>
              <td class="lmtjdt">
                  <select name="classid" id="classid"  onChange="location='<%=url%>&classid='+this.options[this.selectedIndex].value ;">
                    <option value="">请选择分类</option>
                    <% call db_childid(0,classid,"p_class",s_pai) %>
                  </select>              </td>
              <td class="lmtjdt"><!--按尺寸分类查询：--></td>
              <td colspan=2 class="lmtjdt"><!--<select name="ccid" id="ccid"  onChange="location='< %=url%>&ccid='+this.options[this.selectedIndex].value ;">
                  <option value="">请选择分类</option>
                < % call db_childid(0,ccid,"p_class",1) %>
               </select>--></td>
            </tr>
          </form>
        </table>
<table width="100%" border="0" cellpadding="5" cellspacing="0">
 <tr>
  <td bgcolor="<%=Color_0%>">
   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
     <td>
      <%'开始分页
sqlstr="select * from P_main where s_pai="&s_pai
if classid<>"" then
sqlstr=sqlstr&" and classid in ("&db_allid(classid,"p_class")&")"  '''搜索所在类别下的信息
end if
if ccid<>"" then
sqlstr=sqlstr&" and ccid in ("&db_allid(ccid,"p_class")&")"  '''搜索所在类别下的信息
end if
if skey<>"" then
sqlstr=sqlstr&" and "&stype&" like '%"&skey&"%'"
end if
sqlstr=sqlstr&" order by S_order asc,id desc"		
			
set rs=server.CreateObject("adodb.recordset")
			rs.open sqlstr,conn,1,1
if err.number<>0 then
response.write "NO data!"
end if

	if rs.eof And rs.bof then
			Response.Write "<p align='center' class='contents'> 数据库中无数据！</p><br>"
	else
    int_RPP=20 '设置每页显示数目
		int_showNumberLink_=6 '数字导航显示数目
		showMorePageGo_Type_ = 0 '是下拉菜单还是输入值跳转，当多次调用时只能选1
		str_nonLinkColor_="#000000" '非热链接颜色
		toF_="首页"   			'首页 
		toP10_=""			'上十 
		toP1_=" 上一页"			'上一
		toN1_=" 下一页"			'下一
		toN10_=""			'下十
		toL_="尾页"				'尾页
		rs.PageSize=int_RPP
		cPageNo=Request.QueryString("Page")
		If cPageNo="" or not isnumeric(cPageNo) Then cPageNo = 1
		cPageNo = Clng(cPageNo)
		If cPageNo<1 Then cPageNo=1
		If cPageNo>rs.PageCount Then cPageNo=rs.PageCount 
		rs.AbsolutePage=cPageNo
		   
count=0 
i=1
%>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="tableBorder">
       <form name="form1" method="post" action="<%=url%>">
	   <tr>
         <td colspan="6"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="lmtj">
         <tr>
         <td width="123" height="30" align="center"><strong>图片</strong></td>
         <td width="125" align="center"><strong>信息分类</strong></td>
         
         <td width="299" align="left"><strong>名称</strong></td>
         <td width="130" align="center"><strong>加入时间</strong></td>
         <td width="93" align="center"><strong>修改</strong></td>
         <td width="113" align="center"><strong>选 择/排序</strong></td>
        </tr>
        <%do while not (rs.eof or rs.bof) and count<rs.PageSize
		
if i mod 2=0 then 
mm_color="#C2DBFF"
else
mm_color="#F1F5FA"
end if
		
		
		%>
        <tr bgcolor="#FFFFFF" onMouseOver="this.bgColor='#CDE6FF'" onMouseOut="this.bgColor='#FFFFFF'">
         <td height="64" align="center">
         
         <%=CheckImg(Rs("s_name"),Rs("s_img"),Rs("id"),0)%>         </td>
         <td align="center"><%=db_name("p_class",rs("classid"))%></td>
        
         <td><%=str_cut(rs("s_name"),35)%></td>
         <td align="center"> <%=rs("S_time")%></td>
         <td align="center">
          <%
		response.Write("<A title=修改内容 HREF='pro_do.asp?id="&rs("id")&"&s_pai="&s_pai&"'>编辑</A>&nbsp;")
		if show_hot=1 then
		 if rs("s_ok")=1 then
		    response.Write("<a href='"&url&"&action=s_ok&s_ok=0&id="&rs("id")&"'><font color=red>推荐</font></a>&nbsp;")
		 else
		   response.Write("<a href='"&url&"&action=s_ok&s_ok=1&id="&rs("id")&"'><font color=blue>推荐</font></a>&nbsp;")
		 end if 
		 end if
'		 if rs("s_xp")=1 then
'		    response.Write("<a href='"&url&"&action=s_xp&s_xp=0&id="&rs("id")&"'><font color=red>滚动</font></a>")
'		 else
'		   response.Write("<a href='"&url&"&action=s_xp&s_xp=1&id="&rs("id")&"'><font color=black>滚动</font></a>")
'		 end if
		%>         </td>
         <td align="center">
          <input name="selectid" type="checkbox" id="selectid" value="<%=rs("id")%>">
          <input name="orderby<%=trim(rs("id"))%>" type="text" id="orderby" size="4" maxlength="5" value="<%=trim(rs("S_order"))%>" onChange="location='?id=<%=trim(rs("id"))%>&action=orderby&s_pai=<%=s_pai%>&show_hot=<%=show_hot%>&orderby=' + this.value">         </td>
        </tr>
        <%
rs.movenext
i=i+1
count=count+1
loop
%>
        <tr>
         <td height="30" colspan="8" align="center">
          <%response.Write( fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
       <select name="change_classid" size="1" >
        <option value="">全部</option>
				<% call db_ChildID(0,classid,"P_class",s_pai) %>
       </select>
          <input type="submit" name="Submit" value="批量归类" class="inputkkys" onClick="this.form.action='<%=url%>&action=change_class';this.form.submit();">
			      <input type="submit" name="Submit" value="批量删除" class="inputkkys" onClick="return test();">
      <input type="checkbox" name="checkbox" value="Check All" onClick="mm()">全选/反选&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
<%    
end if   
%>
         </table></td>
         </tr>
 
       </form>
      </table>
     </td>
    </tr>
   </table>
  </td>
 </tr>
</table>
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
   for (var i=1; i<a.length; i++)
			a[i].checked = a[a.length-1].checked;
}
</script>
