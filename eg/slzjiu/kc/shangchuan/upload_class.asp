<%@LANGUAGE="VBSCRIPT" CodePage="65001"%>
<%
Response.CharSet = "utf-8"
Session.CodePage = "65001"
Server.ScriptTimeOut=9000
%>
<%
Dim StreamT  

Class AnUpLoad  

     Private Form, Fils  

     Private vCharSet, vMaxSize, vSingleSize, vErr, vVersion, vTotalSize, vExe, pID,vOP  

        

     Public Property Let MaxSize(ByVal value)  

         vMaxSize = value  

     End Property 

        

     Public Property Let SingleSize(ByVal value)  

         vSingleSize = value  

     End Property 

        

     Public Property Let Exe(ByVal value)  

         vExe = LCase(value)  

     End Property 

        

     Public Property Let CharSet(ByVal value)  

         vCharSet = value  

     End Property 

        

     Public Property Get ErrorID()  

         ErrorID = vErr  

     End Property 

        

     Public Property Get Description()  

         Description = GetErr(vErr)  

     End Property 

        

     Public Property Get Version()  

         Version = vVersion  

     End Property 

        

     Public Property Get TotalSize()  

         TotalSize = vTotalSize  

     End Property 

        

     Public Property Get ProcessID()  

         ProcessID = pID  

     End Property 

        

     Public Property Let openProcesser(ByVal value)  

         vOP = value  

     End Property 

        

     Private Sub Class_Initialize()  

         set StreamT=server.createobject("ADODB.STREAM")  

         set Form = server.createobject("Scripting.Dictionary")  

         set Fils = server.createobject("Scripting.Dictionary")  

         vVersion = "V9.9.9" 

         vMaxSize = -1  

         vSingleSize = -1  

         vErr = -1  

         vExe = "" 

         vTotalSize = 0  

         vCharSet = "utf-8" 

         vOP=false  

         pID="Upload" 

         setApp "",0,0,"" 

     End Sub 

        

     Private Sub Class_Terminate()  

         Form.RemoveAll()  

         Fils.RemoveAll()  

         Set Form = Nothing 

         Set Fils = Nothing 

         Set StreamT = Nothing 

     End Sub 

        

     Public Sub GetData()  

         If vMaxSize > 0 And Request.TotalBytes > vMaxSize Then 

             vErr = 1  

             Exit Sub 

         End If 

         if vOP then pID=request.querystring("processid")  

         Dim value, str, bcrlf, fpos, sSplit, slen, istart  

         Dim TotalBytes,BytesRead,ChunkReadSize,PartSize,DataPart,tempdata,formend, formhead, startpos, endpos, formname, FileName, fileExe, valueend, NewName,localname,type_1,contentType  

         If checkEntryType = True Then 

             vTotalSize = 0  

             StreamT.Type = 1  

             StreamT.Mode = 3  

             StreamT.Open  

             TotalBytes = Request.TotalBytes  

             BytesRead = 0  

             ChunkReadSize = 1024 * 36  

             Do While BytesRead < TotalBytes  

                 PartSize = ChunkReadSize  

                 If PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead  

                 DataPart = Request.BinaryRead(PartSize)  

                 StreamT.Write DataPart  

                 BytesRead = BytesRead + PartSize  

                 setApp "uploading",TotalBytes,BytesRead,"" 

             Loop 

             setApp "uploaded",TotalBytes,BytesRead,"" 

             StreamT.Position = 0  

             tempdata = StreamT.Read  

             bcrlf = ChrB(13) & ChrB(10)  

             fpos = InStrB(1, tempdata, bcrlf)  

             sSplit = MidB(tempdata, 1, fpos - 1)  

             slen = LenB(sSplit)  

             istart = slen + 2  

             Do 

                 formend = InStrB(istart, tempdata, bcrlf & bcrlf)  

                 formhead = MidB(tempdata, istart, formend - istart)  

                 str = Bytes2Str(formhead)  

                 startpos = InStr(str, "name=""") + 6  

                 endpos = InStr(startpos, str, """")  

                 formname = LCase(Mid(str, startpos, endpos - startpos))  

                 valueend = InStrB(formend + 3, tempdata, sSplit)  

                 If InStr(str, "filename=""") > 0 Then 

                     startpos = InStr(str, "filename=""") + 10  

                     endpos = InStr(startpos, str, """")  

                     type_1=instr(endpos,lcase(str),"content-type")  

                     contentType=trim(mid(str,type_1+13))  

                     FileName = Mid(str, startpos, endpos - startpos)  

                     If Trim(FileName) <> "" Then 

                         LocalName = FileName  

                         FileName = Replace(FileName, "/", "\")  

                         FileName = Mid(FileName, InStrRev(FileName, "\") + 1)  

                         setApp "processing",TotalBytes,BytesRead,FileName  

                         If instr(FileName,".")>0 Then 

                             fileExe = Split(FileName, ".")(UBound(Split(FileName, ".")))  

                         else  

                             fileExe = "" 

                         End If 

                         If vExe <> "" Then 

                             If checkExe(fileExe) = True Then 

                                 vErr = 3  

                                 Exit Sub 

                             End If 

                         End If 

                         NewName = Getname()  

                         NewName = NewName & "." & fileExe  

                         vTotalSize = vTotalSize + valueend - formend - 6  

                         If vSingleSize > 0 And (valueend - formend - 6) > vSingleSize Then 

                             vErr = 5  

                             Exit Sub 

                         End If 

                         If vMaxSize > 0 And vTotalSize > vMaxSize Then 

                             vErr = 1  

                             Exit Sub 

                         End If 

                         If Fils.Exists(formname) Then 

                             vErr = 4  

                             Exit Sub 

                         Else 

                             Dim fileCls:set fileCls=New fileAction  

                             fileCls.ContentType=contentType  

                             fileCls.Size = (valueend - formend - 6)  

                             fileCls.Position = (formend + 3)  

                             fileCls.NewName = NewName  

                             fileCls.LocalName = FileName  

                             Fils.Add formname, fileCls  

                             Form.Add formname, LocalName  

                             Set fileCls = Nothing 

                         End If 

                     End If 

                 Else 

                     value = MidB(tempdata, formend + 4, valueend - formend - 6)  

                     If Form.Exists(formname) Then 

                         Form(formname) = Form(formname) & "," & Bytes2Str(value)  

                     Else 

                         Form.Add formname, Bytes2Str(value)  

                     End If 

                 End If 

                 istart = valueend + 2 + slen  

             Loop Until (istart + 2) >= LenB(tempdata)  

             vErr = 0  

         Else 

             vErr = 2  

         End If 

         setApp "processed",TotalBytes,BytesRead,"" 

         if err then setApp "faild",1,0,err.description  

     End Sub 

        

     Public sub setApp(stp,total,current,desc)  

         Application.lock()  

         Application(pID)="{ID:""" & pID & """,step:""" & stp & """,total:" & total & ",now:" & current & ",description:""" & desc & """,dt:""" & now() & """}" 

         Application.unlock()  

     end sub  

        

     Private Function checkExe(ByVal ex)  

         Dim notIn: notIn = True 

         If vExe="*" then  

             notIn=false  

         elseIf InStr(1, vExe, "|") > 0 Then 

             Dim tempExe: tempExe = Split(vExe, "|")  

             Dim I: I = 0  

             For I = 0 To UBound(tempExe)  

                 If LCase(ex) = tempExe(I) Then 

                     notIn = False 

                     Exit For 

                 End If 

             Next 

         Else 

             If vExe = LCase(ex) Then 

                 notIn = False 

             End If 

         End If 

         checkExe = notIn  

     End Function 

        

     Public Function GetSize(ByVal Size)  

         If Size < 1024 Then 

             GetSize = FormatNumber(Size, 2) & "B" 

         ElseIf Size >= 1024 And Size < 1048576 Then 

             GetSize = FormatNumber(Size / 1024, 2) & "KB" 

         ElseIf Size >= 1048576 Then 

             GetSize = FormatNumber((Size / 1024) / 1024, 2) & "MB" 

         End If 

     End Function 

        

     Private Function Bytes2Str(ByVal byt)  

         If LenB(byt) = 0 Then 

             Bytes2Str = "" 

             Exit Function 

         End If 

         Dim mystream, bstr  

         Set mystream =server.createobject("ADODB.Stream")  

         mystream.Type = 2  

         mystream.Mode = 3  

         mystream.Open  

         mystream.WriteText byt  

         mystream.Position = 0  

         mystream.CharSet = vCharSet  

         mystream.Position = 2  

         bstr = mystream.ReadText()  

         mystream.Close  

         Set mystream = Nothing 

         Bytes2Str = bstr  

     End Function 

        

     Private Function GetErr(ByVal Num)  

         Select Case Num  

         Case 0  

             GetErr = "数据处理完毕!" 

         Case 1  

             GetErr = "上传数据超过" & GetSize(vMaxSize) & "限制!可设置MaxSize属性来改变限制!" 

         Case 2  

             GetErr = "未设置上传表单enctype属性为multipart/form-data或者未设置method属性为Post,上传无效!" 

         Case 3  

             GetErr = "含有非法扩展名文件!只能上传扩展名为" & Replace(vExe, "|", ",") & "的文件" 

         Case 4  

             GetErr = "对不起,程序不允许使用相同name属性的文件域!" 

         Case 5  

             GetErr = "单个文件大小超出" & GetSize(vSingleSize) & "的上传限制!" 

         End Select 

     End Function 

        

     Private Function Getname()  

         Dim y, m, d, h, mm, S, r  

         Randomize  

         y = Year(Now)  

         m = Month(Now): If m < 10 Then m = "0" & m  

         d = Day(Now): If d < 10 Then d = "0" & d  

         h = Hour(Now): If h < 10 Then h = "0" & h  

         mm = Minute(Now): If mm < 10 Then mm = "0" & mm  

         S = Second(Now): If S < 10 Then S = "0" & S  

         r = 0  

         r = CInt(Rnd() * 1000)  

         If r < 10 Then r = "00" & r  

         If r < 100 And r >= 10 Then r = "0" & r  

         'Getname = y & m & d & h & mm & S & r  
		 
			Getname="kc_up"
     End Function 

        

     Private Function checkEntryType()  

         Dim ContentType, ctArray, bArray,RequestMethod  

         RequestMethod=trim(LCase(Request.ServerVariables("REQUEST_METHOD")))  

         if RequestMethod="" or RequestMethod<>"post" then  

             checkEntryType = False 

             exit function  

         end if  

         ContentType = LCase(Request.ServerVariables("HTTP_CONTENT_TYPE"))  

         ctArray = Split(ContentType, ";")  

         if ubound(ctarray)>=0 then  

             If Trim(ctArray(0)) = "multipart/form-data" Then 

                 checkEntryType = True 

             Else 

                 checkEntryType = False 

             End If 

         else  

             checkEntryType = False 

         end if  

     End Function 

        

     Public Function Forms(ByVal formname)  

         If trim(formname) = "-1" Then 

             Set Forms = Form  

         Else 

             If Form.Exists(LCase(formname)) Then 

                 Forms = Form(LCase(formname))  

             Else 

                 Forms = "" 

             End If 

         End If 

     End Function 

        

     Public Function Files(ByVal formname)  

         If trim(formname) = "-1" Then 

             Set Files = Fils  

         Else 

             If Fils.Exists(LCase(formname)) Then 

                 Set Files = Fils(LCase(formname))  

             Else 

                 Set Files = Nothing 

             End If 

         End If 

     End Function 

        

     Public Function SaveAs(ByVal formname,ByVal path, ByVal saveType )  

         dim vfileAction  

         set vfileAction=Files(formname)  

         if vfileAction.FileName<>"" then  

             if vfileAction.SaveToFile(path,saveType) then  

                 SaveAs=vfileAction.FileName  

             else  

                 SaveAs="Has Error!" 

             end if  

         end if  

         set vfileAction=nothing  

     end function  

 End Class 

    

 Class fileAction  

     Private vSize, vPosition, vName, vNewName, vLocalName, vPath, saveName,vContentType  

        

     Public Property Let NewName(ByVal value)  

         vNewName = value  

     End Property 

        

     Public Property Get NewName()  

         NewName = vNewName  

     End Property 

        

     Public Property Let ContentType(vData)  

         vContentType = vData  

     End Property 

        

     Public Property Get ContentType()  

         ContentType = vContentType  

     End Property 

        

     Public Property Let LocalName(ByVal value)  

         vLocalName = value  

         vName = value  

     End Property 

        

     Public Property Get LocalName()  

         LocalName = vLocalName  

     End Property 

        

     Public Property Get FileName()  

         FileName = vName  

     End Property 

        

     Public Property Let Position(ByVal value)  

         vPosition = value  

     End Property 

        

     Public Property Let Size(ByVal value)  

         vSize = value  

     End Property 

        

     Public Property Get Size()  

         Size = vSize  

     End Property 

        

     Public Function SaveToFile(ByVal path, ByVal saveType)  

         On Error Resume Next 

         Err.Clear  

         vPath = Replace(path, "/", "\")  

         If Right(vPath, 1) <> "\" Then vPath = vPath & "\"  

         CreateFolder vPath  

         Dim mystream  

         Set mystream =server.createobject("ADODB.Stream")  

         mystream.Type = 1  

         mystream.Mode = 3  

         mystream.Open  

         StreamT.Position = vPosition  

         StreamT.CopyTo mystream, vSize  

         vName = vNewName  

         If saveType = 1 Then vName = vLocalName  

         mystream.SaveToFile vPath & vName, 2  

         mystream.Close  

         Set mystream = Nothing 

         If Err Then 

             SaveToFile = False 

         Else 

             SaveToFile = True 

         End If 

     End Function 

        

     Public Function GetBytes()  

         StreamT.Position = vPosition  

         GetBytes = StreamT.Read(vSize)  

     End Function 

        

     Private Function CreateFolder(ByVal FolderPath)  

         on error resume next  

         Dim FolderArray  

         Dim i  

         Dim DiskName  

         Dim Created  

         Dim FSO : Set FSO = Server.CreateObject("Scripting.FileSystemObject")  

         If FSO.FolderExists(FolderPath) Then 

             Set Fso = Nothing 

             Exit Function 

         End If 

         FolderPath = Replace(FolderPath, "/", "\")  

         If Mid(FolderPath, Len(FolderPath), 1) = "\" Then FolderPath = Mid(FolderPath, 1, Len(FolderPath) - 1)  

         FolderArray = Split(FolderPath, "\")  

         DiskName = FolderArray(0)  

         Created = DiskName  

         For i = 1 To UBound(FolderArray)  

             Created = Created & "\" & FolderArray(i)  

             If Not FSO.FolderExists(Created) Then FSO.CreateFolder Created  

         Next 

         Set FSO = Nothing 

         err.clear  

     End Function 

 End Class 
%>