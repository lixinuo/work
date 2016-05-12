<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/gFunc_Page.asp"-->
<%
page=request("page")
if request("action")="edit" then
  strIDs=request("strIDs")

	typename=trim(request("Gradename"))
	conn.execute("update s_user_Grade set Gradename='"&Gradename&"' Where id=" & strIDs & "")
end if
if request("action")="del" then
  Gradeid=request("Gradeid")
    if Gradeid="" then msg "没有选择删除项！","Grade_manage.asp"
	conn.execute("delete from s_user_Grade Where id in(" & Gradeid & ")")
end if
if request("action")="add" then
	addname=request("addname")
	conn.execute("Insert into s_user_Grade(Gradename) values ('"&addname&"')")
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
</head>
<body>
<div class="aclass">
	<ul>
    	<li class="bold">权限管理</li>
        <form name="form2" method="post" action="?action=add">
        <li style=" text-align:left; text-indent:10px;">
        	<span>
                增加权限&nbsp;<input name="addname" type="text" id="addname">&nbsp;
                <input type="submit" name="button" id="button" value="提交">
            </span>
        </li>
        </form>
        <li>
        	<span class="width_10">权限ID</span>
            <span class="width_30">权限名称</span>
            <span class="width_20">权限人数</span>
            <span class="width_20">修改权限</span>
            <span class="width_20">操作</span>
        </li>
        <form name="form1" method="post" action="?action=del">
		<%
        sqlstr="select id,GradeName from [s_user_Grade] where id<>0"
        sqlstr=sqlstr&" order by id desc"
        set rs=server.CreateObject("adodb.recordset")
        rs.open sqlstr,conn,1,1
        if rs.eof then
			w("<li><span>没有权限！</span></li>")
        else
			int_RPP=20 '设置每页显示数目
			int_showNumberLink_=5 '数字导航显示数目
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
			do while not (rs.eof or rs.bof) and count<rs.PageSize
		%>
		<li>
        	<span class="width_10"><%=rs(0)%></span>
            <span class="width_30"><%=rs(1)%></span>
            <span class="width_20"><%=db_s("select count(id) from s_user where Gradeid="&rs("id"))%></span>
            <span class="width_20"><a href="grade_edit.asp?id=<%=rs(0)%>"><img src="../images/detail_off.gif" border=0></a></span>
            <span class="width_20">
			<%
            r_shum=db_s("select count(id) from s_user where Gradeid="&rs("id"))
            if r_shum>0 then
				w "不能删除"
            else
            %>
                <input name="Gradeid" type="checkbox" id="Gradeid" value="<%=rs("id")%>">
            <%end if%>
            </span>
        </li>
		<%
			rs.movenext
			i=i+1
			count=count+1
			loop
        end if
        %>
    	<li>
			<%response.Write( fPageCount(rs,int_showNumberLink_,str_nonLinkColor_,toF_,toP10_,toP1_,toN1_,toN10_,toL_,showMorePageGo_Type_,cPageNo)  & vbcrlf )%>
            &nbsp;&nbsp;
            <input type="submit" name="Submit" class="inputkkys" value="删除" onClick="javascript:this.form.submit();">
            <input type="checkbox" name="checkbox2" value="Check All" onClick="mm()">全选 
        </li>
        </form>
    </ul>
</div>

</body>
</html>
<script type="text/javascript">
function mm()
{
   var a = document.getElementsByTagName("input");
   for (var i=0; i<a.length; i++)
			a[i].checked = a[a.length-1].checked;
}
</script>
