<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>


<%
'所有变量声明/配置:
dim list_parent
list_parent = "../"
Const DbAccessConnection = "../database/#%slzjiu#com.mdb"  
Const Cookies_name="afei_oms"
Const Color_Style = "s_blue"
Const Color_0 = "#ffffff"
Const Color_1 = "#F2F9FD"
Const Color_2 = "#DDEEF1"
Const yu_sjpx_xs = "yu_sjpx-xs"
Const Shtml_TypeName = ".shtml"
dim IframeSrc, no_thing,IframeSrc1
IframeSrc = "../htmledit/ewebeditor.htm?id=content&style=standard600"
IframeSrc1 = "../htmledit/ewebeditor.htm?id=content1&style=standard600"
IframeSrc2 = "../htmledit/ewebeditor.htm?id=content2&style=standard600"
no_thing = "<tr bgcolor='"&Color_0&"'><td colspan='5'><font color='#ff0000'>还没有任何东东</font></td></tr>"
%>
