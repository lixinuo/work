<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#include file="../inc/data.asp"-->
<%
tn=Request("tn"):tm=Request("tm")
if tn<>"" then tn=cint(tn)
if tm<>"" then tm=cint(tm)

gradealist_nn=DB_F("S_user_Grade","GradeAList",tn)
gradenlist_nn=DB_F("S_user_Grade","GradeNList",tn)

%>
<%if gradealist_nn<>"" then%>
     <tr>
      <td>
       <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
       <%
        set rs_s=server.CreateObject("adodb.recordset")
        rs_s.Open "select id,s_name from S_menu_class where parent_id="&tm&" and s_ok=1 order by s_order",conn,1,1
        if rs_s.EOF and rs_s.BOF then
		    response.Write "<div align=center><font color=red>还没有菜单</font></center>"
		    else
				i=1
				tdnum=6
				tdwidth="17%"
        do while not rs_s.EOF
        %>
         <td><%=rs_s("s_name")%><input id="nlist<%=i%>" name="nlist" type="checkbox" value="<%=rs_s(0)%>" 
				 <%
				 nn=split(gradenlist_nn,",")
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
       </table>
      </td>
     </tr>
 <%else%>    
     
     
 <tr>
      <td>
       <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
       <%
        set rs_s=server.CreateObject("adodb.recordset")
        rs_s.Open "select * from S_menu_class where parent_id="&tm&" and s_ok=1 order by s_order",conn,1,1
        if rs_s.EOF and rs_s.BOF then
		    response.Write "<div align=center><font color=red>还没有菜单</font></center>"
		    else
				i=1
				tdnum=6
				tdwidth="17%"
        do while not rs_s.EOF
        %>
         <td><%=rs_s(1)%><input id="nlist<%=i%>" name="nlist" type="checkbox" value="<%=rs_s(0)%>" 
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
       </table>
      </td>
     </tr>  
     
 <%end if%>
