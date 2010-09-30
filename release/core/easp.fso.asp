<%
'######################################################################
'## easp.fso.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp FileSystemObject Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/09/30 00:03:09
'## Description :   EasyAsp Files System Operator
'##
'######################################################################
Class EasyAsp_Fso
	Public oFso, IsVirtualHost
	Private Fso
	Private b_force,b_overwrite
	Private s_fsoName,s_sizeformat,s_charset

	Private Sub Class_Initialize
		s_fsoName 	= Easp.FsoName
		s_charset	= Easp.CharSet
		Set Fso 	= Server.CreateObject(s_fsoName)
		Set oFso 	= Fso
		IsVirtualHost = True
		b_force		= True
		b_overwrite	= True
		s_sizeformat= "K"
		Easp.Error(52) = "写入文件错误！"
		Easp.Error(53) = "创建文件夹错误！"
		Easp.Error(54) = "读取文件列表失败！"
		Easp.Error(55) = "设置属性失败，文件不存在！"
		Easp.Error(56) = "设置属性失败！"
		Easp.Error(57) = "获取属性失败，文件不存在！"
		Easp.Error(58) = "复制失败，源文件不存在！"
		Easp.Error(59) = "移动失败，源文件不存在！"
		Easp.Error(60) = "删除失败，文件不存在！"
		Easp.Error(61) = "重命名失败，源文件不存在！"
		Easp.Error(62) = "重命名失败，已存在同名文件！"
		Easp.Error(63) = "文件或文件夹操作错误！"
	End Sub

	Private Sub Class_Terminate
		Set Fso 	= Nothing
		Set oFso 	= Nothing
	End Sub
	Public Property Let fsoName(Byval str)
		s_fsoName = str
		Set Fso = Server.CreateObject(s_fsoName)
		Set oFso = Fso
	End Property
	Public Property Let CharSet(Byval str)
		s_charset = Ucase(str)
	End Property
	Public Property Let Force(Byval bool)
		b_force = bool
	End Property
	Public Property Let OverWrite(Byval bool)
		b_overwrite = bool
	End Property
	Public Property Let SizeFormat(Byval str)
		s_sizeformat = str
	End Property
	Public Property Get ShowErr()
		ShowErr = s_err
	End Property
	Public Function isExists(ByVal path)
		isExists = False
		If isFile(path) or isFolder(path) Then isExists = True
	End Function
	Public Function isFile(ByVal filePath)
		filePath = absPath(filePath) : isFile = False
		If Fso.FileExists(filePath) Then isFile = True
	End Function
	Public Function Read(ByVal filePath)
		Dim p, f, o_strm, tmpStr, s_char
		s_char = s_charset
		If Instr(filePath,">")>0 Then
			s_char = UCase(Trim(Easp.CRight(filePath,">")))
			filePath = Trim(Easp.CLeft(filePath,">"))
		End If
		p = absPath(filePath)
		If isFile(p) Then
			Set o_strm = Server.CreateObject("ADODB.Stream")
			With o_strm
				.Type = 2
				.Mode = 3
				.Open
				.LoadFromFile p
				.Charset = s_char
				.Position = 2
				tmpStr = .ReadText
				.Close
			End With
			Set o_strm = Nothing
			If s_char = "UTF-8" Then
				Select Case Easp.FileBOM
					Case "keep"
						'Do Nothing
					Case "remove"
						If Easp.Test(tmpStr, "^\uFEFF") Then
							tmpStr = Easp.RegReplace(tmpStr, "^\uFEFF", "")
						End If
					Case "add"
						If Not Easp.Test(tmpStr, "^\uFEFF") Then
							tmpStr = Chrw(&hFEFF) & tmpStr
						End If
				End Select
			End If
		Else
			tmpStr = ""
			Easp.Error.Msg = "(" & filePath & ")"
			Easp.Error.Raise 2
		End If
		Read = tmpStr
	End Function
	Public Function SaveAs(ByVal filePath, ByVal fileContent)
		On Error Resume Next
		Dim f,p,t, o_strm
		p = absPath(filePath)
		SaveAs = MD(Left(p,InstrRev(p,"\")-1))
		If SaveAs Then
			Set o_strm = Server.CreateObject("ADODB.Stream")
			With o_strm
				.Type = 1
				.Open
				.Write fileContent
				.SaveToFile p,Easp.IIF(b_overwrite,2,1)
				.Close
			End With
			Set o_strm = Nothing
		End If
		If Err.Number<>0 Then
			SaveAs = False
			Easp.Error.Msg = "(" & filePath & ")"
			Easp.Error.Raise 52
		End If
		Err.Clear()
	End Function
	Public Function CreateFile(ByVal filePath, ByVal fileContent)
		On Error Resume Next
		Dim f,p,t, s_char, o_strm
		s_char = s_charset
		If Instr(filePath,">")>0 Then
			s_char = UCase(Trim(Easp.CRight(filePath,">")))
			filePath = Trim(Easp.CLeft(filePath,">"))
		End If
		p = absPath(filePath)
		CreateFile = MD(Left(p,InstrRev(p,"\")-1))
		If CreateFile Then
			Set o_strm = Server.CreateObject("ADODB.Stream")
			With o_strm
				.Type = 2
				.Open
				.Charset = s_char
				.Position = o_strm.Size
				.WriteText = fileContent
				.SaveToFile p,Easp.IIF(b_overwrite,2,1)
				.Close
			End With
			Set o_strm = Nothing
		End If
		If Err.Number<>0 Then
			CreateFile = False
			Easp.Error.Msg = "(" & filePath & ")"
			Easp.Error.Raise 52
		End If
		Err.Clear()
	End Function
	Public Function UpdateFile(ByVal filePath, ByVal rule, ByVal result)
		Dim tmpStr : filePath = absPath(filePath)
		tmpStr = Easp.regReplace(Read(filePath),rule,result)
		UpdateFile = CreateFile(filePath,tmpStr)
	End Function
	Public Function AppendFile(ByVal filePath, ByVal fileContent)
		Dim tmpStr : filePath = absPath(filePath)
		tmpStr = Read(filePath) & fileContent
		AppendFile = CreateFile(filePath,tmpStr)
	End Function
	Public Function isFolder(ByVal folderPath)
		folderPath = absPath(folderPath) : isFolder = False
		If Fso.FolderExists(folderPath) Then isFolder = True
	End Function
	Public Function CreateFolder(ByVal folderPath)
		On Error Resume Next
		Dim p,arrP,i : CreateFolder = True
		p = absPath(folderPath)
		arrP = Split(p,"\") : p = ""
		For i = 0 To Ubound(arrP)
			p = p & arrP(i) & "\"
			If IsVirtualHost Then
				If Instr(p, absPath("/") & "\")>0 Then
					If Not isFolder(p) And i>0 Then Fso.CreateFolder(p)
				End If
			Else
				If Not isFolder(p) And i>0 Then Fso.CreateFolder(p)
			End If
		Next
		If Err.Number<>0 Then
			CreateFolder = False
			Easp.Error.Msg = "(" & folderPath & ")"
			Easp.Error.Raise 53
		End If
		Err.Clear()
	End Function
	Public Function MD(ByVal folderPath)
		MD = CreateFolder(folderPath)
	End Function
	Public Function Dir(ByVal folderPath)
		Dir = List(folderPath,0)
	End Function
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
			Easp.Error.Msg = "(" & folderPath & ")"
			Easp.Error.Raise 54
		End If
		Err.Clear()
	End Function
	Public Function NameOf(ByVal f)
		NameOf = GetNameOf(f, 0)
	End Function
	Public Function ExtOf(ByVal f)
		ExtOf = GetNameOf(f, 1)
	End Function
	Private Function GetNameOf(ByVal f, ByVal t)
		Dim re,na,ex
		If Easp.isN(f) Then GetNameOf = "" : Exit Function
		f = Replace(f,"\","/")
		If Right(f,1) = "/" Then
			re = Split(f,"/")
			GetNameOf = Easp.IIF(t=0,re(Ubound(re)-1),"")
			Exit Function
		ElseIf Instr(f,"/")>0 Then
			re = Split(f,"/")(Ubound(Split(f,"/")))
		Else
			re = f
		End If
		If Instr(re,".")>0 Then
			na = Left(re,InstrRev(re,".")-1)
			ex = Mid(re,InstrRev(re,"."))
		Else
			na = re
			ex = ""
		End If
		If t = 0 Then
			GetNameOf = na
		ElseIf t = 1 Then
			GetNameOf = ex
		End If
	End Function
	Public Function Attr(ByVal path, ByVal attrType)
		On Error Resume Next
		Dim p,a,i,n,f,at : p = absPath(path) : n = 0 : Attr = True
		If not isExists(p) Then
			Attr = False
			Easp.Error.Msg = "(" & path & ")"
			Easp.Error.Raise 55
			Exit Function
		End If
		If isFile(p) Then
			Set f = Fso.GetFile(p)
		ElseIf isFolder(p) Then
			Set f = Fso.GetFolder(p)
		End If
		at = f.Attributes : a = UCase(attrType)
		If Instr(a,"+")>0 Or Instr(a,"-")>0 Then
			a = Easp.IIF(Instr(a," ")>0,Split(a," "),Split(a,","))
			For i = 0 To Ubound(a)
				Select Case a(i)
					Case "+R" at = Easp.IIF(at And 1,at,at+1)
					Case "-R" at = Easp.IIF(at And 1,at-1,at)
					Case "+H" at = Easp.IIF(at And 2,at,at+2)
					Case "-H" at = Easp.IIF(at And 2,at-2,at)
					Case "+S" at = Easp.IIF(at And 4,at,at+4)
					Case "-S" at = Easp.IIF(at And 4,at-4,at)
					Case "+A" at = Easp.IIF(at And 32,at,at+32)
					Case "-A" at = Easp.IIF(at And 32,at-32,at)
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
			f.Attributes = Easp.IIF(at And 32,n+32,n)
		End If
		Set f = Nothing
		If Err.Number<>0 Then
			Attr = False
			Easp.Error.Msg = "(" & path & ")"
			Easp.Error.Raise 56
		End If
		Err.Clear()
	End Function
	Public Function getAttr(ByVal path, ByVal attrType)
		Dim f,s,p : p = absPath(path)
		If isFile(p) Then
			Set f = Fso.GetFile(p)
		ElseIf isFolder(p) Then
			Set f = Fso.GetFolder(p)
		Else
			getAttr = ""
			Easp.Error.Msg = "(" & path & ")"
			Easp.Error.Raise 57
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
	Public Function CopyFile(ByVal fromPath, ByVal toPath)
		CopyFile = FOFO(fromPath,toPath,0,0)
	End Function
	Public Function CopyFolder(ByVal fromPath, ByVal toPath)
		CopyFolder = FOFO(fromPath,toPath,1,0)
	End Function
	Public Function Copy(ByVal fromPath, ByVal toPath)
		Dim ff,tf : ff = absPath(fromPath) : tf = absPath(toPath)
		If isFile(ff) Then
			Copy = CopyFile(fromPath,toPath)
		ElseIf isFolder(ff) Then
			Copy = CopyFolder(fromPath,toPath)
		Else
			Copy = False
			Easp.Error.Msg = "(" & fromPath & ")"
			Easp.Error.Raise 58
		End If
	End Function
	Public Function MoveFile(ByVal fromPath, ByVal toPath)
		MoveFile = FOFO(fromPath,toPath,0,1)
	End Function
	Public Function MoveFolder(ByVal fromPath, ByVal toPath)
		MoveFolder = FOFO(fromPath,toPath,1,1)
	End Function
	Public Function Move(ByVal fromPath, ByVal toPath)
		Dim ff,tf : ff = absPath(fromPath) : tf = absPath(toPath)
		If isFile(ff) Then
			Move = MoveFile(fromPath,toPath)
		ElseIf isFolder(ff) Then
			Move = MoveFolder(fromPath,toPath)
		Else
			Move = False
			Easp.Error.Msg = "(" & fromPath & ")"
			Easp.Error.Raise 59
		End If
	End Function
	Public Function DelFile(ByVal path)
		DelFile = FOFO(path,"",0,2)
	End Function
	Public Function DelFolder(ByVal path)
		DelFolder = FOFO(path,"",1,2)
	End Function
	Public Function RD(ByVal path)
		RD = DelFolder(path)
	End Function
	Public Function Del(ByVal path)
		Dim p : p = absPath(path)
		If isFile(p) Then
			Del = DelFile(path)
		ElseIf isFolder(p) Then
			Del = DelFolder(path)
		Else
			Del = False
			Easp.Error.Msg = "(" & path & ")"
			Easp.Error.Raise 60
		End If
		Err.Clear()
	End Function
	Public Function Rename(ByVal path, ByVal newname)
		Dim p,n : p = absPath(path) : Rename = True
		n = Left(p,InstrRev(p,"\")) & newname
		If Not isExists(p) Then
			Rename = False
			Easp.Error.Msg = "(" & path & ")"
			Easp.Error.Raise 61
			Exit Function
		End If
		If isExists(n) Then
			Rename = False
			Easp.Error.Msg = "(" & newname & ")"
			Easp.Error.Raise 62
			Exit Function
		End If
		If isFolder(p) Then
			Fso.MoveFolder p,n
		ElseIf isFile(p) Then
			Copy p, n
			Del p
		End If
	End Function
	Public Function Ren(ByVal path, ByVal newname)
		Ren = Rename(path,newname)
	End Function
	Private Function absPath(ByVal p)
		Dim pt
		If Easp.IsN(p) Then absPath = "" : Exit Function
		If Mid(p,2,1)<>":" Then
			If isWildcards(p) Then
				p = Replace(p,"*","[.$.[e.a.s.p.s.t.a.r].#.]")
				p = Replace(p,"?","[.$.[e.a.s.p.q.u.e.s].#.]")
				p = Server.MapPath(p)
				p = Replace(p,"[.$.[e.a.s.p.q.u.e.s].#.]","?")
				p = Replace(p,"[.$.[e.a.s.p.s.t.a.r].#.]","*")
			Else
				p = Server.MapPath(p)
			End If
		End If
		If Right(p,1) = "\" Then p = Left(p,Len(p)-1)
		absPath = p
	End Function
	Public Function MapPath(p)
		MapPath = absPath(p)
	End Function
	Public Function formatSize(Byval fileSize, ByVal level)
		Dim s : s = Int(fileSize) : level = UCase(level)
		formatSize = Easp.IIF(s/(1073741824)>0.01,FormatNumber(s/(1073741824),2,-1,0,-1),"0.01") & " GB"
		If s = 0 Then formatSize = "0 GB"
		If level = "G" Or (level="AUTO" And s>1073741824) Then Exit Function
		formatSize = Easp.IIF(s/(1048576)>0.1,FormatNumber(s/(1048576),1,-1,0,-1),"0.1") & " MB"
		If s = 0 Then formatSize = "0 MB"
		If level = "M" Or (level="AUTO" And s>1048576) Then Exit Function
		formatSize = Easp.IIF((s/1024)>1,Int(s/1024),1) & " KB"
		If s = 0 Then formatSize = "0 KB"
		If Level = "K" Or (level="AUTO" And s>1024) Then Exit Function
		If level = "B" or level = "AUTO" Then
			formatSize = s & " bytes"
		Else
			formatSize = s
		End If
	End Function
	Private Function isWildcards(ByVal path)
		isWildcards = False
		If Instr(path,"*")>0 Or Instr(path,"?")>0 Then isWildcards = True
	End Function
	Private Function FOFO(ByVal fromPath, ByVal toPath, ByVal FOF, ByVal MOC)
		On Error Resume Next
		FOFO = True
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
					tf = tf & "\"
					FOFO = MD(tf)
				End If
				Execute("Fso."&ot&of&" ff,tf"&Easp.IfThen(MOC=0,",b_overwrite"))
			Else
				Execute("Fso."&ot&of&" ff,b_force")
			End If
			If Err.Number<>0 Then
				FOFO = False
				Easp.Error.Msg = "<br />" & os & oi & "失败！" & "( "&frompath&" "&Easp.IIF(MOC=2,"",os&"到 "&toPath)&" )"
				Easp.Error.Raise 63
			End If
		ElseIf isWildcards(ff) Then
			If MOC<>2 Then
				FOFO = MD(tf)
				Execute("Fso."&ot&of&" ff,tf"&Easp.IIF(MOC=0,",b_overwrite",""))
			Else
				Execute("Fso."&ot&of&" ff,b_force")
			End If
			If Err.Number<>0 Then
				FOFO = False
				Easp.Error.Msg = "<br />" & os & oi & "失败！" & "( "&frompath&" "&Easp.IIF(MOC=2,"",os&"到 "&toPath)&" )"
				Easp.Error.Raise 63
			End If
		Else
			FOFO = False
			Easp.Error.Msg = "<br />" & os & oi & "失败！" & Easp.IIF(MOC=2,"","源")&oi&"不存在( "&frompath&" )"
			Easp.Error.Raise 63
		End If
		Err.Clear()
	End Function
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
End Class
%>