<%
'===================================================================================
'函数作用：创建RS，0执行sql无返回，1执行sql返回不可操作rs，2执行sql返回只读rs,3执行sql返回可读写rs,  
'===================================================================================
Function Db(fcd_sqlstr,fcd_type)
	on error resume next
	select case fcd_type
	case 0
		conn.execute(fcd_sqlstr)
	case 1
		set db = conn.execute(fcd_sqlstr)
	case 2
		set db = server.createobject("adodb.recordset")
		db.open fcd_sqlstr,conn,1,1
	case 3
		set db = server.createobject("adodb.recordset")
		db.open fcd_sqlstr,conn,1,3
	end Select
	if err then 
		response.clear
		w "时间：" & now & "<br>"
		w "描述：" & err.description & "<br>" : err.clear 
		w "参考：" & dbSQL &"<br>":response.End()
	end if
End Function
'===================================================================================
'函数作用：'关闭数据连接  
'===================================================================================
sub CloseConn()
		conn.close:set conn=nothing
end sub	
'===================================================================================
'函数作用：'从数据库中读取一条数据出来  
'===================================================================================
function db_s(fcd_sqlstr) 
	set rs_rt=db(fcd_sqlstr,1)
	db_s=iif(rs_rt.eof,"None",rs_rt(0))
end function
'===================================================================================
'函数作用：'返回一个数据记录数组
'===================================================================================
function db_a(fcd_sqlstr) 
 set rs_rt=db(fcd_sqlstr,1)
		if not rs_rt.eof then
			 rs_fa=rs_rt.getrows(rs_rt.recordcount)
		else
       redim rs_fa(0,0):rs_fa(0,0)="None"
		end if
 db_a=rs_fa
end function
'===================================================================================
'函数作用：'//取得任意字段的一个值，可以直接填写ID也可以填写其他查询条件 
'===================================================================================
function Db_F(fcd_table,fcd_Feild,fcd_Where)
	fcd_sqlstr="select "&fcd_Feild&" from "&fcd_table&" Where id<>0"
	fcd_sqlstr=iif(isnumeric(fcd_Where),fcd_sqlstr&" and id="&fcd_Where,fcd_sqlstr&" and "&fcd_Where)
	Db_F=db_s(fcd_sqlstr)
end function
function Db_Filed(fcd_table,fcd_Feild,fcd_Where)
	Db_Filed=Db_F(fcd_table,fcd_Feild,fcd_Where)
end function
'===================================================================================
'函数作用：'//取得任意几个字段的一个值，可以直接填写ID也可以填写其他查询条件 
'===================================================================================
function Db_Fs(fcd_table,fcd_Feild,fcd_Where)
	fcd_sqlstr="select top 1 "&fcd_Feild&" from "&fcd_table&" Where id<>0"
	fcd_sqlstr=iif(isnumeric(fcd_Where),fcd_sqlstr&" and id="&fcd_Where,fcd_sqlstr&" and "&fcd_Where)
	fcd_arr=db_a(fcd_sqlstr):redim Db_arr(ubound(fcd_arr))
	for fcd_i=0 to ubound(fcd_arr)
	  Db_arr(fcd_i)=fcd_arr(fcd_i,0)
	next
	Db_Fs=Db_arr
end function
'===================================================================================
'函数作用：'//函数的部分简写
'===================================================================================
function DB_Name(fcd_table,fcd_id)'//取得name值
 DB_Name=Db_F(fcd_table,"s_name",fcd_id)
end function
function DB_Name1(fcd_table,fcd_id)'//取得name值
 DB_Name1=Db_F(fcd_table,"s_name1",fcd_id)
end function
function DB_Name2(fcd_table,fcd_id)'//取得name值
 DB_Name2=Db_F(fcd_table,"s_name2",fcd_id)
end function

function DB_bt(fcd_table,fcd_id)'//取得name值
 DB_bt=Db_F(fcd_table,"s_bt",fcd_id)
end function
function DB_gjc(fcd_table,fcd_id)'//取得name值
 DB_gjc=Db_F(fcd_table,"s_gjc",fcd_id)
end function
function DB_ms(fcd_table,fcd_id)'//取得name值
 DB_ms=Db_F(fcd_table,"s_ms",fcd_id)
end function

function DB_bt1(fcd_table,fcd_id)'//取得name值
 DB_bt1=Db_F(fcd_table,"s_bt1",fcd_id)
end function
function DB_gjc1(fcd_table,fcd_id)'//取得name值
 DB_gjc1=Db_F(fcd_table,"s_gjc",fcd_id)
end function
function DB_ms1(fcd_table,fcd_id)'//取得name值
 DB_ms1=Db_F(fcd_table,"s_ms1",fcd_id)
end function

function DB_bt2(fcd_table,fcd_id)'//取得name值
 DB_bt2=Db_F(fcd_table,"s_bt2",fcd_id)
end function
function DB_gjc2(fcd_table,fcd_id)'//取得name值
 DB_gjc2=Db_F(fcd_table,"s_gjc2",fcd_id)
end function
function DB_ms2(fcd_table,fcd_id)'//取得name值
 DB_ms2=Db_F(fcd_table,"s_ms2",fcd_id)
end function


function Db_Content(fcd_table,fcd_id)
 Db_Content=Db_F(fcd_table,"s_content",fcd_id)
end function
function Db_Content1(fcd_table,fcd_id)
 Db_Content1=Db_F(fcd_table,"s_content1",fcd_id)
end function
function Db_Content2(fcd_table,fcd_id)
 Db_Content2=Db_F(fcd_table,"s_content2",fcd_id)
end function

function DB_img(fcd_table,fcd_id)'//取得name值
 DB_img=Db_F(fcd_table,"s_img",fcd_id)
end function

function DB_ad_img(fcd_id)'//取得name值
 DB_ad_img=Db_F("o_ad","s_img",fcd_id)
end function
function DB_ad_url(fcd_id)'//取得name值
 DB_ad_url=Db_F("o_ad","s_url",fcd_id)
end function

function DB_info_name(fcd_id)'//取得name值
 DB_info_name=Db_F("a_info","s_name",fcd_id)
end function
function DB_info_content(fcd_id)'//取得name值
 DB_info_content=Db_F("a_info","s_content",fcd_id)
end function


'===================================================================================
'函数作用：''直接取得fcd_table表中的最顶单一个id
'===================================================================================

function DB_FirstID(fcd_table,fcd_Where)'//取得name值
 DB_FirstID=Db_F(fcd_table,"top 1 id",fcd_Where)
end function



'===================================================================================
'函数作用：'直接导入Excel的全部数据，导入到指定的数据库字段"select id,s_name,s_time from info|s_name=afei,s_time=123"
'注意需要设置默认值的字段请放在SQL最后按顺序写
'===================================================================================
function db_excelin(excel,accsql)
 if instr(accsql,"|") then accsql_str=split(accsql,"|")(0):varstr=split(accsql,"|")(1) else accsql_str=accsql:varstr=""
 db_excelin=0:ecConnStr = "Driver={Microsoft Excel Driver (*.xls)};Dbq=" & sm(excel)
	Set ecConn = Server.CreateObject("ADODB.Connection")
	ecConn.Open ecConnStr
	set rsfrom=ecConn.execute("Select * From [Sheet1$]")
	set rsto=db(accsql_str,3)
	if rsfrom.eof then '判断Excel力是否有数据
		msg "Excel里无数据","":exit function
	else
		rsf_a=rsfrom.getrows(rsfrom.recordcount) '构建Excel表格的数组
			for fc_db_i=0 to ubound(rsf_a,2) '获得数组的二维大小，即有多少行,并循环
					rsto.addnew()
					for fc_db_j=0 to ubound(rsf_a,1) '获得数组的一维大小，即有多少列，并循环
							rsto(fc_db_j)=rsf_a(fc_db_j,fc_db_i) '存入数据
							if err then:msg("Excel格式不符合"):exit function
					next
					if instr(varstr,"=") then
						varstr_f=split(varstr,",")'取得要设置的默认值的字符串，以,为分隔符处理为数组
						for fc_db_ii=0 to ubound(varstr_f)'循环设置数组中的字段
						rsto(split(varstr_f(fc_db_ii),"=")(0))=split(varstr_f(fc_db_ii),"=")(1)'格式为 字段=默认值，
						next
					end if
					rsto.update
			next
	end if
	db_excelin=fc_db_i
	ecConn.close:set ecConn=nothing
end function
'============================================================================================================================
'函数作用：读取数据库里相应字段，导入到指定的Excel表中
'===================================================================================
function db_excelout(excel,accsql)
 db_excelout=0:rsnums=0:filename = sm(excel):set rsin=db(accsql,2)			 
	for C_tbi=0 to rsin.Fields.Count-1                        '构建Excel表格字段创建语句
	C_tb=C_tb&rsin.fields(C_tbi).name&" text ,"
	next
	C_tb_sql="Create fcd_table Sheet1("&left(C_tb,len(C_tb)-1)&")" '构建结束
	set fso=Server.CreateObject ("Scripting.FileSystemObject") '判断是否存在此Excel文件，如果有则删除 
	if fso.FileExists(filename) then
		Set fs = CreateObject("Scripting.FileSystemObject")
		Set thisfile = fs.GetFile(filename)
		thisfile.delete true
	end if         		                                         '判断结束
	Set Econn = Server.CreateObject("ADODB.Connection")
	Econn.Open "Driver={Microsoft Excel Driver (*.xls)};Readonly=0;DBQ=" & filename
	Econn.Execute(C_tb_sql)'调用ODBC的Excel组件创建Excel文件
		for fc_db_ii=0 to rsin.recordcount-1
			for fc_db_i=0 to rsin.Fields.Count-1'循环写出字段名和字段内容
			In_tb_name=In_tb_name&rsin.fields(fc_db_i).name&",":In_tb=In_tb&"'"&rsin(fc_db_i)&"',"
			Next 
			In_sql="Insert into Sheet1("&left(In_tb_name,len(In_tb_name)-1)&")values("&left(In_tb,len(In_tb)-1)&")"
			Econn.Execute(In_sql):In_sql ="":In_tb="":In_tb_name=""
			if err.number<>0 then w err.description:die'显示错误信息结束
			rsin.MoveNext
		Next
	if err.number<>0 then w err.description:die'显示错误信息开始
	rsin.close:set rsin=nothing:db_excelout=fc_db_ii
end function 
'============================================================================================================================
'函数作用：'取得FolderID为id的目录下所有子目录的FolderID，以半角逗号分开
'===================================================================================
function db_allid(fc_db_id,fcd_table)
 arrID=fc_db_id
	Set rsdir = db("Select ID from "&fcd_table&" Where Parent_ID = " & fc_db_id & "",1)
	if rsdir.eof or rsdir.bof then
		 set rsdir = nothing:db_allid = arrID:exit function
	else
			while not rsdir.eof		
					arrID = arrID&","&db_allid(rsdir("ID"),fcd_table)
					rsdir.movenext
			wend
	end if
	set rsdir=nothing:db_allid=arrID
end function



'============================================================================================================================
'函数作用：'取得最顶端父类的id，如果只有一级分类返回本身id
'===================================================================================
function db_parentid(fc_db_id,fcd_table)
 arrID=fc_db_id
	Set rsdir = db("Select parent_id from "&fcd_table&" Where id = " & fc_db_id & "",1)
	if rsdir("parent_id")=0 then
		 set rsdir = nothing:db_parentid = fc_db_id:exit function
	else
			while not rsdir.eof		
					arrID = db_parentid(rsdir("parent_id"),fcd_table)
					rsdir.movenext
			wend
	end if
	set rsdir=nothing:db_parentid=arrID
end function


'============================================================================================================================
'函数作用：'取得上一级父类的id，如果只有一级分类返回本身id
'===================================================================================
function db_sparentid(fc_db_id,fcd_table)
 arrID=fc_db_id
	Set rsdir = db("Select parent_id from "&fcd_table&" Where id = " & fc_db_id & "",1)
	if rsdir("parent_id")=0 then
		 set rsdir = nothing:db_sparentid = fc_db_id:exit function
	else
		 arrID=rsdir("parent_id")	
	end if
	set rsdir=nothing:db_sparentid=arrID
end function







'============================================================================================================================
'函数作用：'取得FolderID为id的目录下所有子目录的FolderID，以半角逗号分开
'===================================================================================
function db_ChildID(fc_db_id,oldid,fcd_table,fcd_spai)
  arrID=fc_db_id:ChildStr=""
	strchild="Select ID,S_name,Parent_ID,class_depth from "&fcd_table&" Where s_pai="&fcd_spai
	if fc_db_id=0 then strchild=strchild&"and Parent_ID=0" else strchild=strchild&"and Parent_ID="&arrID&""
	strchild=strchild&" order by s_order asc,id desc"
	Set rsdir = Conn.Execute(strchild)
	if rsdir.eof and rsdir.bof then
		set rsdir = nothing:db_ChildID = "":exit function
	else
		do while not rsdir.eof
			for fc_db_j=1 to rsdir(3)
			brstr="&nbsp;"&brstr
			next
			ChildStr = "<option value="&rsdir(0)
			if rsdir(0)&""= oldid&"" then ChildStr=ChildStr&" selected "
			ChildStr=ChildStr&">"&brstr&"|--"&trim(rsdir(1))&"</option> "
			wc(ChildStr)			
		  db_ChildID rsdir("ID"),oldid,fcd_table,fcd_spai
		  rsdir.movenext:brstr=""
		loop
	end if
	set rsdir = nothing
end function
'============================================================================================================================
'函数作用：	'取得FolderID为id的目录下所有子目录的FolderID，以半角逗号分
'===================================================================================
function db_OptionID(optionid,old_optionid,fcd_table)
	arrID=Cint(optionid)
	Set rsdir = Conn.Execute("Select ID,s_name,Parent_ID,class_depth from "&fcd_table&" Where Parent_ID="&arrID)
	if rsdir.eof or rsdir.bof then
		set rsdir = nothing:db_OptionID = arrID:exit function
	else
		do while not rsdir.eof
        for fc_db_j=1 to rsdir(3)
          brstr="&nbsp;&nbsp;"&brstr
        next
			arrID = "<option value="&rsdir(0)
			if rsdir(0)&""=old_optionid&"" then arrID=arrID&" selected "
			arrID=arrID&">"&brstr&"|-"&trim(rsdir(1))&"</option> "
			wc(arrID)		
		call db_OptionID(rsdir(0),old_optionid,fcd_table)
		rsdir.movenext:brstr=""
		loop
	end if
	set rsdir=nothing
end function
%>