<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#INCLUDE FILE="../inc/data1.asp" -->
<%
table_name="s_menu_class"
id=trim(request("id"))
o_id=isid(request.QueryString("o_id"),0)
action=request("action")

id_a=db_a("select id from "&table_name&" where parent_id="&o_id&" order by s_order asc,id desc")
  pre_id=get_content(get_positon(id)-1)
  next_id=get_content(get_positon(id)+1)

if action="s_order" then
	strSQL="update "&table_name&" set s_order="&request("s_order")&" Where id=" &id& ""
	conn.execute strSQL
end if

if action="s_order_up" then

	if pre_id<>"" then pre_id_order=db_f(table_name,"s_order",pre_id)
	if next_id<>"" then next_id_order=db_f(table_name,"s_order",next_id)
	now_id_order=db_f(table_name,"s_order",id)

	strSQL="update "&table_name&" set s_order="&now_id_order&" Where id=" &pre_id& ""
	conn.execute strSQL
	strSQL="update "&table_name&" set s_order="&pre_id_order&" Where id=" &id& ""
	conn.execute strSQL
end if

if action="s_order_down" then

	if pre_id<>"" then pre_id_order=db_f(table_name,"s_order",pre_id)
	if next_id<>"" then next_id_order=db_f(table_name,"s_order",next_id)
	now_id_order=db_f(table_name,"s_order",id)

	strSQL="update "&table_name&" set s_order="&now_id_order&" Where id=" &next_id& ""
	conn.execute strSQL
	strSQL="update "&table_name&" set s_order="&next_id_order&" Where id=" &id& ""
	conn.execute strSQL
end if


if action="edit" then
	strSQL="update "&table_name&" set s_name='"&request("s_name")&"',s_url='"&request("s_url")&"' Where id=" &id& ""
	conn.execute strSQL
end if
if action="s_ok" then
	strSQL="update "&table_name&" set s_ok="&request("s_ok")&" Where id=" &id& ""
	conn.execute strSQL
end if
if action="add" then
	strSQL="insert into "&table_name&" (s_name,s_url,s_order,parent_id) values ('"&request("s_name")&"','"&request("s_url")&"',"&request("s_order")&","&o_id&")"
	conn.execute strSQL
end if
if action="move" then
conn.execute "update "&table_name&" set parent_id="&o_id&" where id="&id
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="../inc/g_links.asp"-->
  
 <script type="text/ecmascript"> 
    function opacity(id, opacStart, opacEnd, millisec) { 
  //speed for each frame 
  var speed = Math.round(millisec / 100); 
  var timer = 0; 
  //determine the direction for the blending, if start and end are the same nothing happens 
  if(opacStart > opacEnd) { 
  for(i = opacStart; i >= opacEnd; i--) { 
  setTimeout("changeOpac(" + i + ",'" + id + "')",(timer * speed)); 
  timer++; 
  } 
  } else if(opacStart < opacEnd) { 
  for(i = opacStart; i <= opacEnd; i++) 
  { 
  setTimeout("changeOpac(" + i + ",'" + id + "')",(timer * speed)); 
  timer++; 
  } 
  } 
} 
//change the opacity for different browsers 
function changeOpac(opacity, id) { 
    var obj = document.getElementById(id); 
    if (obj) { 
        var s = obj.style; 
        s.opacity = (opacity / 100); 
        s.MozOpacity = (opacity / 100); 
        s.KhtmlOpacity = (opacity / 100); 
        s.filter = "alpha(opacity=" + opacity + ")"; 
        s = null; 
    } 
} 
function shiftOpacity(id, millisec) { 
  //if an element is invisible, make it visible, else make it ivisible 
  if(document.getElementById(id).style.opacity == 0) { 
  opacity(id, 0, 100, millisec); 
  } else { 
  opacity(id, 100, 0, millisec); 
  } 
} 
function blendimage(divid, imageid, imagefile, millisec) { 
  var speed = Math.round(millisec / 100); 
  var timer = 0; 
  //set the current image as background 
  document.getElementById(divid).style.backgroundImage = "url(" + document.getElementById(imageid).src + ")"; 
  //make image transparent 
  changeOpac(0, imageid); 
  //make new image 
  document.getElementById(imageid).src = imagefile; 
  //fade in image 
  for(i = 0; i <= 100; i++) { 
  setTimeout("changeOpac(" + i + ",'" + imageid + "')",(timer * speed)); 
  timer++; 
  } 
} 
function currentOpac(id, opacEnd, millisec) { 
  //standard opacity is 100 
  var currentOpac = 100; 
  //if the element has an opacity set, get it 
  if(document.getElementById(id).style.opacity < 100) { 
  currentOpac = document.getElementById(id).style.opacity * 100; 
  } 
  //call for the function that changes the opacity 
  opacity(id, currentOpac, opacEnd, millisec) 
} 

function showContent(i, event){ 
    showid = "content" + i; 
    var target = document.getElementById(showid); 
    target.style.position = "absolute"; 
    if(navigator.appName!="Netscape"){ 
        event=window.event; 
        event.srcElement.style.fontWeight = "700"; 
    } else { 
        event.target.style.fontWeight = "700"; 
    } 
    target.style.top = event.clientY + 22 +"px"; 
    target.style.left = event.clientX + 12 + "px"; 

    //复制一个背景 
    var bg = target.cloneNode(true); 
    if (bg) { 
        bg.id="bg1"; 
        if (bg.style.backgroundColor.length==0) { 
            bg.style.backgroundColor ="#FFFFE1"; 
        } 
        bg.style.filter = "alpha(opacity=0)"; 
        bg.style.opacity = 0; 
        target.parentNode.appendChild(bg); 
         
        opacity("bg1", 0, 90, 300); 
        bg.style.display="block"; 
    } 

    target.style.display = "block"; 
} 

function hiddenContent(i, event){ 
    if(navigator.appName!="Netscape"){ 
        event=window.event; 
        event.srcElement.style.fontWeight = "400"; 
    } else { 
        event.target.style.fontWeight = "400"; 
    } 
    hiddenid = "content" + i; 
    document.getElementById(hiddenid).style.display = "none"; 
    var bg = document.getElementById("bg1"); 
    if (bg) { 
        bg.parentNode.removeChild(bg); 
    } 
} 
</script> 
</head>
<body>
<div class="aclass">
	<ul>
    	<li class="bold">菜单管理</li>
    	<li style="text-align:left;">
        	<select name="select" class="chose_item" onChange="location='?o_id='+this.options[this.selectedIndex].value" >
                <option >选择栏目</option>
				<%call my_optionid(0,o_id,table_name)%>
            </select>
        </li>
        <li>
        	<span class="width_15">菜单名称</span>
            <span class="width_50">链接地址</span>
            <span class="width_15">菜单排序</span>
            <span class="width_20">确定操作</span>
        </li>
		<%
        if o_id=0 then
        else
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from "&table_name&" where parent_id="&o_id&" order by s_order asc,id desc",conn,1,1
			if rs.EOF and rs.BOF then
				response.Write "<div align=center><font color=red>还没有菜单</font></center>"
				paixu=0
			else
				formi=0
				do while not rs.EOF
        %>
        <form name="formlist<%=formi%>" method="post" action="?action=edit&id=<%=rs(0)%>&o_id=<%= o_id %>">
        <li>
        	<span class="width_15">
            	<input name="s_name" type="text" id="s_name" size="12" value="<%=trim(rs("s_name"))%>">
            </span>
            <span class="width_50">
            	<input name="s_url" type="text" id="s_url" size="45" value="<%=trim(rs("s_url"))%>">
            </span>
            <span class="width_15">
            	<%if Isid(db_F(table_name,"top 1 s_order","s_order<"&Rs("s_order")&" and s_ok=1 and parent_id<>0  order by s_order desc"),0)<>0 then%><a href="?action=s_order_up&id=<%=trim(rs("id"))%>&o_id=<%= o_id %>">上</a><%else%>顶<%end if%>
                <input name="s_order<%=rs(0)%>" type="text"  size="2" value="<%=int(rs("s_order"))%>" onChange="location='?action=s_order&id=<%=trim(rs("id"))%>&o_id=<%= o_id %>&s_order=' + this.value">
				<%if Isid(db_F(table_name,"top 1 s_order","s_order>"&Rs("s_order")&" and s_ok=1 and parent_id<>0  order by s_order asc"),0)<>0 then%><a href="?action=s_order_down&id=<%=trim(rs("id"))%>&o_id=<%= o_id %>">下</a><%else%>底<%end if%>
            </span>
            <span class="width_20">
          <input type="submit" class="inputkkys" name="Submit" value="修改">&nbsp;|&nbsp;<%
		 if rs("s_ok") then
		    response.Write("<a title=取消显示 href=?action=s_ok&s_ok=0&id="&rs(0)&"&O_id="&O_id&" ><font color=blue>显示</font></a>")
		 else
		    response.Write("<a title=取消隐藏' href=?action=s_ok&s_ok=1&id="&rs(0)&"&O_id="&O_id&" ><font color=red>隐藏</font></a>")
		 end if
%>&nbsp;|&nbsp;<a href="?id=<%=int(rs("id"))%>&action=del&o_id=<%=o_id%>" onClick="return confirm('您确定进行删除操作吗？')">删除</a>
            </span>
        </li>
        </form>
        <%
				rs.movenext:formi=formi+1
				loop
				paixu=rs.RecordCount
				rs.close
				set rs=nothing
			end if
        end if
		%>
        <li></li>
        <li class="bold">菜单添加</li>
        <li>
        <form name="form2" method="post" action="?action=add&o_id=<%= o_id %>">
        <span class="width_15"><input name="s_name" type="text" id="s_name" size="12"></span>
        <span class="width_50"><input name="s_url" type="text" id="s_url" size="45"></span>
        <span class="width_15"><input name="s_order" type="text" id="s_order" size="4" value="<%=paixu+1%>"></span> 
        <span class="width_20"><input type="submit" class="inputkkys" name="Submit2" value="添加菜单"></span>
        </form>
        </li>
    </ul>
</div>
<script>
function append_url(idnum)
{
	var add_item_value=document.getElementById("add_item"+idnum).value;
	alert(add_item_value);
	return;
	var s_url=document.forms[idnum].s_url;
	s_url.value=s_url.value+"&"+ add_item;
	}


function change_me(objform,obj,str)
{
	var s_url=objform.s_url;
	var url_name=obj.name;
	var url_value=obj.value;
	var new_str=url_name+"="+url_value;
	var all_str=s_url.value;
  
	var str1=all_str.split("?");//取得？之后的数据
	if (str1.length>1)
	{
		var str2=str1[1].split("&");
		for(i=0;i<str2.length;i++)
		  {
				var str3=str2[i].split("=");
				if(str3[0]==url_name){str=str3[0]+"="+str3[1]}
				}
		}
	
	//s_url.value=all_str.replace("/"+url_name+"=(.*?)\&/gi",url_name+"="+url_value+"&");
	s_url.value=all_str.replace(str,new_str);
	}
</script>
<%
closeconn

function my_optionid(optionid,old_optionid,tablename)
	arrID=Cint(optionid)
	Set rsdir = Conn.Execute("Select ID,s_name,Parent_ID,class_depth from "&tablename&" where s_ok=1 and Parent_ID="&arrID&" order by s_order asc,id desc")
	if rsdir.eof or rsdir.bof then
		set rsdir = nothing:db_OptionID = arrID:exit function
	else
		do while not rsdir.eof
        for j=1 to rsdir(3)
          brstr="&nbsp;&nbsp;"&brstr
        next
			arrID = "<option value="&rsdir(0)
			if cint(rsdir(0))=Cint(old_optionid) then 
			 arrID=arrID&" selected "
			end if
			arrID=arrID&">"&brstr&"|-"&trim(rsdir(1))&"</option> "
			response.Write(arrID)			
		rsdir.movenext:brstr=""
		loop
	end if
	set rsdir = nothing
end function
%>
</body>
</html>
<%
function get_positon(id)
for i=0 to ubound(id_a,2)
	 if id_a(0,i)&""=id&"" then get_positon=i
next
end function

function get_content(id)
if id>ubound(id_a,2) or id<0 then
 get_content=""
else
 for i=0 to ubound(id_a,2)
	 if i=id then get_content=id_a(0,i)
 next
end if
end function




function url_manage(str)
str_w="":str1=split(str,"?"):if ubound(str1)<1 then exit function
str2=split(str1(1),"&")
for i=0 to ubound(str2)
 str3=split(str2(i),"=")
 str_w=str_w&"<div style='width:250px;'><input name='"&str3(0)&"' value='"&str3(1)&"' onChange=""change_me(this.form,this,'"&str3(0)&"="&str3(1)&"')"">"&str3(0)&"</div>"
next
url_manage=str_w
end function
%>