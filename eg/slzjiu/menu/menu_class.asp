<!--#INCLUDE FILE="../data.asp" -->
<%
AllGrade=Request.Cookies(Cookies_name)("Grade")

GradeAlist=Db_s("select GradeAlist from S_user_Grade where gradename='"&AllGrade&"'")
Gradenlist=Db_s("select GradeNlist from S_user_Grade where gradename='"&AllGrade&"'")

set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from S_menu_class where s_ok=1 and Parent_id=0 and id in("&GradeAlist&") order by S_order asc,id desc",conn,1,1
if rs_s.recordcount=0 then 
else
	i=1
	do while not rs_s.eof
%>
<tr>
 <TD onMouseMove="this.className='menu_title2';" onMouseOut="this.className='';"   onclick="menuChange(this,mu0<%=i%>)" width="100%" height="30" align="left" bgcolor="#DDDDDD" style="cursor:hand;" background="images/menudown.gif"> <span class="style58">&nbsp;<%=rs_s("S_name")%></span></td>
</tr>
<tr>
 <td align="center">
  <DIV id="mu0<%=i%>" >
   <div align="center">
    <table width="100%" border="0" align="center" cellpadding="2" style="border: 1px solid #CADDE4; padding: 0" bgcolor="#F6FBFF" >
     <tr>
     <td width="100%" align="left" >&nbsp;
		 <%
		set rs_s1=server.CreateObject("adodb.recordset")
		rs_s1.open "select * from S_menu_class where Parent_id=" & rs_s(0) & " and s_ok=1 and id in("&GradeNlist&") order by s_order asc,id desc",conn,1,1
		if rs_s1.recordcount=0 then 
			response.Write("<tr><td>None</td></tr>")
		else
			n=1
			tdnum=2
			tdwidth="50%"
			do while not rs_s1.eof
				if (n mod 2=0) or n>=rs_s1.recordcount then ssplit="" else ssplit="|"
%>
                 <a href="<%=rs_s1("S_url")%>"><%=rs_s1("S_name")%></a><%=ssplit%>
<%
				if n mod tdnum = 0 and n<rs_s1.recordcount  then response.Write("</td></tr><tr><td>&nbsp;")
			rs_s1.movenext 
			n=n+1
			loop
		end if
		rs_s1.close
		set rs_s1=nothing
%>
</td></tr></table>
    <br>
   </div>
  </div>
 </td>
</tr>
<%
	rs_s.movenext
	i=i+1
loop
end if
rs_s.close
set rs_s=nothing
%>
