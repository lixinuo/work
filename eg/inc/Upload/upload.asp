<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
%>
<!--#include file="upload_class.inc"-->
<%
 on error resume next
 Server.ScriptTimeout = 9999999
 Dim Upload,successful,thisFile,upPath,path
 set Upload=new AnUpLoad
 Upload.openProcesser=true  '打开进度条显示
 Upload.SingleSize=512*1024*1024  '设置单个文件最大上传限制,按字节计；默认为不限制，本例为512M
 Upload.MaxSize=1024*1024*1024 '设置最大上传限制,按字节计；默认为不限制，本例为1G
 Upload.Exe="jpg|gif|png|rar|zip|pdf|doc|txt|xls|ppt|flv|mp3|wma|wav|mp4"  '设置允许上传的扩展名
 Upload.CharSet="utf-8"  '设置上传的字符集
 Upload.GetData()
 if Upload.ErrorID>0 then
 	upload.setApp "faild",1,0 ,Upload.description
 else
 	if Upload.files(-1).count>0 then
 	        upPath=request.querystring("path")
    		path=server.mappath(upPath)
    		set tempCls=Upload.files("file1") 
            upload.setApp "saving",Upload.TotalSize,Upload.TotalSize,tempCls.FileName
    		successful=tempCls.SaveToFile(path,0)
			thisFile="{name:'" & tempCls.FileName & "',size:" & tempCls.Size & ",LocalName:'" & tempCls.LocalName & "',NewName:'" & tempCls.NewName & "'}"
			allFiles=allFiles & thisFile & ","
    		set tempCls=nothing
		upload.setApp "saved",Upload.TotalSize,Upload.TotalSize,left(allFiles,len(allFiles)-1)
    else
        upload.setApp "faild",1,0,"没有上传任何文件"
 	end if
 end if
 if err then upload.setApp "faild",1,0,err.description
 set Upload=nothing
 response.end
%>
