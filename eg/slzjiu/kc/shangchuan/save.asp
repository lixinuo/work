<!--#include file="upload_class.asp"-->  

 <%  

 Dim Upload,path,tempCls  

 set Upload=new AnUpLoad  

 Upload.SingleSize=1024*1024*1024  

 Upload.MaxSize=1024*1024*1024  

 Upload.Exe="xls" 

 Upload.Charset="utf-8" 

 Upload.openProcesser=false  

 Upload.GetData()  

if Upload.ErrorID>0 then  

     response.write Upload.Description  

 else  

     dim said,stitle,scontent  

       

     if Upload.files(-1).count>0 then  

         'path=server.mappath("/uploads")&"/"&Year(date())&"/"&Month(date())&"/"&Day(date())&"/" 
		 
		path=server.mappath("../")
        set tempCls=Upload.files("eimage")   

         tempCls.SaveToFile path,0  

        simage_name=tempCls.FileName  

        'simage_path="/uploads/"&Year(date())&"/"&Month(date())&"/"&Day(date())&"/" 

            simage_path="../"

        simage_url=simage_path & simage_name  

    

         response.write "文件:" & tempCls.FileName & "上传完毕,大小为" & Upload.getsize(tempCls.Size) & ";原文件名" & tempCls.LocalName & "!<br />" 
		 response.Write"<a href='../kc.asp'>【点击录入到系统】</a>"

            

        set tempCls=nothing  

       

 end if  

        

end if  

set Upload=nothing  

%> 
