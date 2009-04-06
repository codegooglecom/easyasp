<%
Class EasyASP_Fso
	Public oFso
	Private Fso
	Private s_fsoName,s_err,s_sizeformat,s_charset
	Private b_debug,b_force,b_overwrite
	
	Private Sub Class_Initialize
		s_fsoName 	= "Scripting.FilesyStemObject"
		s_charset	= "UTF-8"
		Set Fso 	= Server.CreateObject(s_fsoName)
		Set oFso 	= Fso		'FSO原型接口
		b_debug		= False
		b_force		= True
		b_overwrite	= True
		s_sizeformat= "K"
	End Sub

	Private Sub Class_Terminate
		Set Fso 	= Nothing
		Set oFso 	= Nothing
	End Sub
	'属性：FSO组件名称
	Public Property Let fsoName(Byval str)
		s_fsoName = str
		Set Fso = Server.CreateObject(s_fsoName)
		Set oFso = Fso
	End Property
	'属性：文件编码
	Public Property Let CharSet(Byval str)
		s_charset = Ucase(str)
	End Property
	'属性：是否开启调试状态
	Public Property Let Debug(Byval bool)
		b_debug = bool
	End Property
	'属性：是否删除只读文件
	Public Property Let Force(Byval bool)
		b_force = bool
	End Property
	'属性：是否覆盖原有文件
	Public Property Let OverWrite(Byval bool)
		b_overwrite = bool
	End Property
	'属性：文件大小显示格式(G,M,K,b,auto)
	Public Property Let SizeFormat(Byval str)
		s_sizeformat = str
	End Property
	'属性：显示错误信息
	Public Property Get ShowErr()
		ShowErr = s_err
	End Property
	
	'文件或文件夹是否存在
	Public Function isExists(ByVal path)
		path = absPath(path) : isExists = False
		If Fso.FileExists(path) or Fso.FolderExists(path) Then isExists = True			
	End Function
	'文件是否存在
	Public Function isFile(ByVal filePath)
		filePath = absPath(filePath) : isFile = False
		If Fso.FileExists(filePath) Then isFile = True
	End Function
	'读取文件内容
	Public Function Read(ByVal filePath)
		Dim p, f, o_strm, tmpStr : p = absPath(filePath)
		If isFile(p) Then
			If s_charset = "GB2312" Then
				Set f = Fso.OpenTextFile(p)
				tmpStr = f.ReadAll
				f.Close()
				Set f = Nothing
			Else
				Set o_strm = Server.CreateObject("ADODB.Stream")
				With o_strm
					.Type = 2
					.Mode = 3
					.Open
					.LoadFromFile p
					.Charset = s_charset
					.Position = 2
					tmpStr = .ReadText
					.Close
				End With
				Set o_strm = Nothing
			End If
		Else
			tmpStr = ""
			ErrMsg "读取文件错误！", "文件未找到(" & filePath & ")"
		End If
		Read = tmpStr
	End Function
	'创建文件并写入内容
	Public Function CreateFile(ByVal filePath, ByVal fileContent)
		On Error Resume Next
		Dim f,p,t : p = absPath(filePath)
		CreateFile = MD(Left(p,InstrRev(p,"\")-1))
		If CreateFile Then
			If s_charset = "GB2312" Then
				Set f = Fso.CreateTextFile(p,b_overwrite)
				f.Write fileContent
				f.Close()
				Set f =Nothing
			Else
				Set o_strm = Server.CreateObject("ADODB.Stream")
				With o_strm
					.Type = 2
					.Open
					.Charset = s_charset
					.Position = o_strm.Size
					.WriteText = fileContent
					.SaveToFile p,Easp_IIF(b_overwrite,2,1)
					.Close
				End With
				Set o_strm = Nothing
			End If
		End If
		If Err.Number<>0 Then
			CreateFile = False
			ErrMsg "写入文件错误！", Err.Description & "("&filePath&")"
		End If
		Err.Clear()
	End Function
	'按正则表达式更新文件内容
	Public Function UpdateFile(ByVal filePath, ByVal rule, ByVal result)
		Dim tmpStr : filePath = absPath(filePath)
		tmpStr = Easp_Replace(Read(filePath),rule,result,0)
		UpdateFile = CreateFile(filePath,tmpStr)
	End Function
	'追加文件内容
	Public Function AppendFile(ByVal filePath, ByVal fileContent)
		Dim tmpStr : filePath = absPath(filePath)
		tmpStr = Read(filePath) & fileContent
		AppendFile = CreateFile(filePath,tmpStr)
	End Function
	'文件夹是否存在
	Public Function isFolder(ByVal folderPath)
		folderPath = absPath(folderPath) : isFolder = False
		If Fso.FolderExists(folderPath) Then isFolder = True
	End Function
	'创建文件夹 MD
	Public Function CreateFolder(ByVal folderPath)
		On Error Resume Next
		Dim p,arrP,i : CreateFolder = True
		p = absPath(folderPath)
		arrP = Split(p,"\") : p = ""
		For i = 0 To Ubound(arrP)
			p = p & arrP(i) & "\"
			If Not isFolder(p) Then Fso.CreateFolder(p)
		Next
		If Err.Number<>0 Then
			CreateFolder = False
			ErrMsg "创建文件夹错误！", Err.Description & "("&folderPath&")"
		End If
		Err.Clear()
	End Function
	Public Function MD(ByVal folderPath)
		MD = CreateFolder(folderPath)
	End Function
	'列出文件夹下的所有文件夹、文件
	Public Function Dir(ByVal folderPath)
		Dir = List(folderPath,0)
	End Function
	'列出文件夹下的所有文件夹或文件
	Public Function List(ByVal folderPath, ByVal fileType)
		On Error Resume Next
		Dim f,fs,k,arr(),i,l
		folderPath = absPath(folderPath) : i = 0
		Select Case LCase(fileType)
			Case "0","" l = 0
			Case "1","file" l = 1
			Case "2","folder" l = 2
			Case Else l = 0
		End Select
		Set f = Fso.GetFolder(folderPath)
		If l = 0 Or l = 2 Then
			Set fs = f.SubFolders
			ReDim Preserve arr(4,fs.Count-1)
			For Each k In fs
				arr(0,i) = k.Name & "/"
				arr(1,i) = formatSize(k.Size,s_sizeformat)
				arr(2,i) = k.DateLastModified
				arr(3,i) = Attr2Str(k.Attributes)
				arr(4,i) = k.Type
				i = i + 1
			Next
		End If
		If l = 0 Or l = 1 Then
			Set fs = f.Files
			ReDim Preserve arr(4,fs.Count+i-1)
			For Each k In fs
				arr(0,i) = k.Name
				arr(1,i) = formatSize(k.Size,s_sizeformat)
				arr(2,i) = k.DateLastModified
				arr(3,i) = Attr2Str(k.Attributes)
				arr(4,i) = k.Type
				i = i + 1
			Next
		End If
		Set fs = Nothing
		Set f = Nothing
		List = arr
		If Err.Number<>0 Then
			ErrMsg "读取文件列表失败！", Err.Description & "("&folderPath&")"
		End If
		Err.Clear()
	End Function
	'设置文件或文件夹属性
	Public Function Attr(ByVal path, ByVal attrType)
		On Error Resume Next
		Dim p,a,i,n,f,at : p = absPath(path) : n = 0 : Attr = True
		If not isExists(p) Then
			Attr = False : ErrMsg "设置属性失败！", "文件不存在("&path&")" : Exit Function
		End If
		If isFile(p) Then
			Set f = Fso.GetFile(p)
		ElseIf isFolder(p) Then
			Set f = Fso.GetFolder(p)
		End If
		at = f.Attributes : a = UCase(attrType)
		If Instr(a,"+")>0 Or Instr(a,"-")>0 Then
			a = Easp_IIF(Instr(a," ")>0,Split(a," "),Split(a,","))
			For i = 0 To Ubound(a)
				Select Case a(i)
					Case "+R" at = Easp_IIF(at And 1,at,at+1)
					Case "-R" at = Easp_IIF(at And 1,at-1,at)
					Case "+H" at = Easp_IIF(at And 2,at,at+2)
					Case "-H" at = Easp_IIF(at And 2,at-2,at)
					Case "+S" at = Easp_IIF(at And 4,at,at+4)
					Case "-S" at = Easp_IIF(at And 4,at-4,at)
					Case "+A" at = Easp_IIF(at And 32,at,at+32)
					Case "-A" at = Easp_IIF(at And 32,at-32,at)
				End Select
			Next
			f.Attributes = at
		Else
			For i = 1 To Len(a)
				Select Case Mid(a,i,1)
					Case "R" n = n + 1
					Case "H" n = n + 2
					Case "S" n = n + 4
				End Select
			Next
			f.Attributes = Easp_IIF(at And 32,n+32,n)
		End If
		Set f = Nothing
		If Err.Number<>0 Then
			Attr = False
			ErrMsg "设置属性失败！", Err.Description & "("&path&")"
		End If
		Err.Clear()
	End Function
	'获取文件或文件夹属性
	Public Function getAttr(ByVal path, ByVal attrType)
		Dim f,s : p = absPath(path)
		If isFile(p) Then
			Set f = Fso.GetFile(p)
		ElseIf isFolder(p) Then
			Set f = Fso.GetFolder(p)
		Else
			getAttr = "" : ErrMsg "获取属性失败！", "文件不存在("&path&")"
			Exit Function
		End If
		Select Case LCase(attrType)
			Case "0","name" : s = f.Name
			Case "1","date", "datemodified" : s = f.DateLastModified
			Case "2","datecreated" : s = f.DateCreated
			Case "3","dateaccessed" : s = f.DateLastAccessed
			Case "4","size" : s = formatSize(f.Size,s_sizeformat)
			Case "5","attr" : s = Attr2Str(f.Attributes)
			Case "6","type" : s = f.Type
			Case Else s = ""
		End Select
		Set f = Nothing
		getAttr = s
	End Function
	'复制文件(支持通配符*和?)
	Public Function CopyFile(ByVal fromPath, ByVal toPath)
		CopyFile = FOFO(fromPath,toPath,0,0)
	End Function
	'复制文件夹(支持通配符*和?)
	Public Function CopyFolder(ByVal fromPath, ByVal toPath)
		CopyFolder = FOFO(fromPath,toPath,1,0)
	End Function
	'复制文件或文件夹
	Public Function Copy(ByVal fromPath, ByVal toPath)
		Dim ff,tf : ff = absPath(fromPath) : tf = absPath(toPath)
		If isFile(ff) Then
			Copy = CopyFile(fromPath,toPath)
		ElseIf isFolder(ff) Then
			Copy = CopyFolder(fromPath,toPath)
		Else
			Copy = False : ErrMsg "复制失败！","源文件不存在("&fromPath&")"
		End If
	End Function
	'移动文件(支持通配符*和?)
	Public Function MoveFile(ByVal fromPath, ByVal toPath)
		MoveFile = FOFO(fromPath,toPath,0,1)
	End Function
	'移动文件夹(支持通配符*和?)
	Public Function MoveFolder(ByVal fromPath, ByVal toPath)
		MoveFolder = FOFO(fromPath,toPath,1,1)
	End Function
	'移动文件或文件夹
	Public Function Move(ByVal fromPath, ByVal toPath)
		Dim ff,tf : ff = absPath(fromPath) : tf = absPath(toPath)
		If isFile(ff) Then
			Move = MoveFile(fromPath,toPath)
		ElseIf isFolder(ff) Then
			Move = MoveFolder(fromPath,toPath)
		Else
			Move = False : ErrMsg "移动失败！","源文件不存在("&fromPath&")"
		End If
	End Function
	'删除文件(支持通配符*和?)
	Public Function DelFile(ByVal path)
		DelFile = FOFO(path,"",0,2)
	End Function
	'删除文件夹(支持通配符*和?)
	Public Function DelFolder(ByVal path)
		DelFolder = FOFO(path,"",1,2)
	End Function
	Public Function RD(ByVal path)
		RD = DelFolder(path)
	End Function
	'删除文件或文件夹
	Public Function Del(ByVal path)
		Dim p : p = absPath(path)
		If isFile(p) Then
			Del = DelFile(path)
		ElseIf isFolder(p) Then
			Del = DelFolder(path)
		Else
			Del = False : ErrMsg "删除失败！", "文件不存在" & "("&path&")"
		End If
		Err.Clear()
	End Function
	'文件或文件夹更名
	Public Function Rename(ByVal path, ByVal newname)
		Dim p,n : p = absPath(path) : Rename = True
		n = Left(p,InstrRev(p,"\")) & newname
		If Not isExists(p) Then
			Rename = False : ErrMsg "重命名失败！","源文件不存在("&path&")"
			Exit Function
		End If
		If isExists(n) Then
			Rename = False : ErrMsg "重命名失败！","已存在同名文件("&newname&")"
			Exit Function
		End If
		Copy p,n : Del p
	End Function
	Public Function Ren(ByVal path, ByVal newname)
		Ren = Rename(path,newname)
	End Function
	'===私有方法===
	'取文件夹绝对路径
	Private Function absPath(ByVal path)
		Dim p : p = path
		If Instr(p,":") = 0 Then p = Server.MapPath(p)
		If Right(p,1) = "\" Then p = Left(p,Len(p)-1)
		absPath = p
	End Function
	'路径是否包含通配符
	Private Function isWildcards(ByVal path)
		isWildcards = False
		If Instr(path,"*")>0 Or Instr(path,"?")>0 Then isWildcards = True
	End Function
	'文件或文件夹操作原型
	Private Function FOFO(ByVal fromPath, ByVal toPath, ByVal FOF, ByVal MOC)
		On Error Resume Next
		Dim ff,tf,oc,of,oi,ot,os
		ff = absPath(fromPath) : tf = absPath(toPath)
		If FOF = 0 Then
			oc = isFile(ff) : of = "File" : oi = "文件"
		ElseIf FOF = 1 Then
			oc = isFolder(ff) : of = "Folder" : oi = "文件夹"
		End If
		If MOC = 0 Then
			ot = "Copy" : os = "复制"
		ElseIf MOC = 1 Then
			ot = "Move" : os = "移动"
		ElseIf MOC = 2 Then
			ot = "Delete" : os = "删除"
		End If
		If oc Then
			If MOC<>2 Then
				If FOF = 0 Then
					If Right(toPath,1)="/" or Right(toPath,1)="\" Then
						FOFO = MD(tf) : tf = tf & "\"
					Else
						FOFO = MD(Left(tf,InstrRev(tf,"\")-1))
					End If
				ElseIf FOF = 1 Then
					FOFO = MD(tf)
				End If
				Execute("Fso."&ot&of&" ff,tf"&Easp_IIF(MOC=0,",b_overwrite",""))
			Else
				Execute("Fso."&ot&of&" ff,b_force")
			End If
			If Err.Number<>0 Then
				FOFO = False
				ErrMsg os&oi&"失败！", Err.Description & "( "&frompath&" "&Easp_IIF(MOC=2,"",os&"到 "&toPath)&" )"
			End If
		ElseIf isWildcards(ff) Then
			If Not isFolder(Left(ff,InstrRev(ff,"\")-1)) Then
				FOFO = False
				ErrMsg os&oi&"失败！", Easp_IIF(MOC=2,"","源")&oi&"不存在( "&frompath&" )"
			End If
			If MOC<>2 Then
				FOFO = MD(tf)
				Execute("Fso."&ot&of&" ff,tf"&Easp_IIF(MOC=0,",b_overwrite",""))
			Else
				Execute("Fso."&ot&of&" ff,b_force")
			End If
			If Err.Number<>0 Then
				FOFO = False
				ErrMsg os&oi&"失败！", Err.Description & "( "&frompath&" "&Easp_IIF(MOC=2,"",os&"到 "&toPath)&" )"
			End If
		Else
			FOFO = False
			ErrMsg os&oi&"失败！", Easp_IIF(MOC=2,"","源")&oi&"不存在( "&frompath&" )"
		End If
		Err.Clear()
	End Function
	'格式化文件大小
	Private Function formatSize(Byval fileSize, ByVal level)
		Dim s : s = Int(fileSize) : level = UCase(level)
		formatSize = Easp_IIF(s/(1073741824)>0.01,FormatNumber(s/(1073741824),2,-1,0,-1),"0.01") & " GB"
		If s = 0 Then formatSize = "0 GB"
		If level = "G" Or (level="AUTO" And s>1073741824) Then Exit Function
		formatSize = Easp_IIF(s/(1048576)>0.1,FormatNumber(s/(1048576),1,-1,0,-1),"0.1") & " MB"
		If s = 0 Then formatSize = "0 MB"
		If level = "M" Or (level="AUTO" And s>1048576) Then Exit Function
		formatSize = Easp_IIF((s/1024)>1,Int(s/1024),1) & " KB"
		If s = 0 Then formatSize = "0 KB"
		If Level = "K" Or (level="AUTO" And s>1024) Then Exit Function
		If level = "B" or level = "AUTO" Then
			formatSize = s & " bytes"
		Else
			formatSize = s
		End If
	End Function
	'格式化文件属性
	Private Function Attr2Str(ByVal attrib)
		Dim a,s : a = Int(attrib)
		If a>=2048 Then a = a - 2048
		If a>=1024 Then a = a - 1024
		If a>=32 Then : s = "A" : a = a- 32 : End If
		If a>=16 Then a = a- 16
		If a>=8 Then a = a - 8
		If a>=4 Then : s = "S" & s : a = a- 4 : End If
		If a>=2 Then : s = "H" & s : a = a- 2 : End If
		If a>=1 Then : s = "R" & s : a = a- 1 : End If
		Attr2Str = s
	End Function
	'输出错误信息
	Private Sub ErrMsg(e,d)
		s_err = "<div id=""easp_err"">" & e
		If Not Easp_isN(d) Then s_err = s_err & "<br/>错误信息:" & d
		s_err = s_err & "</div>"
		If b_debug Then
			Response.Write s_err
			Response.End()
		End If
	End Sub
End Class
%>