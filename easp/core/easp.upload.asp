<%
'######################################################################
'## easp.upload.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp Upload Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/01/26 16:08:30
'## Description :   Upload file(s) with EasyASP
'##
'######################################################################
Class EasyAsp_upload
	Public  File, Form
	Private o_strm
	Private s_charset,s_allowed,s_denied,s_filename,s_savepath
	Private i_maxsize,i_totalmaxsize,i_filecount
	Private b_automd,b_random
	
	Private Sub Class_Initialize
		i_totalmaxsize	= 0
		i_maxsize	= 0
		i_filecount = 0
		s_charset	= Easp.CharSet
		s_allowed	= ""
		s_denied	= ""
		s_filename	= ""
		s_savepath	= ""
		b_automd	= False
		b_random	= False
		Easp.Error(71) = "表单类型错误，表单只能是""multipart/form-data""类型！"
		Easp.Error(72) = "请先选择要上传的文件！"
		Easp.Error(73) = "上传文件失败，上传文件总大小超过了限制！"
		Easp.Error(74) = "上传文件失败，上传文件不能为空！"
		Easp.Error(75) = "上传文件失败，文件大小超过了限制！"
		Easp.Error(76) = "上传文件失败，不允许上传此类型的文件！"
		Easp.Error(77) = "上传文件失败！"
		Easp.Error(78) = "获取文件失败！"
		Set File = Server.CreateObject("Scripting.Dictionary")
		Set Form = Server.CreateObject("Scripting.Dictionary")
		File.CompareMode = 1
		Form.CompareMode = 1
		Set o_strm  = Server.CreateObject("ADODB.Stream")
		o_strm.Type	= 1
		o_strm.Mode	= 3
		o_strm.Open
	End Sub
	
	Private Sub Class_Terminate
		o_strm.Close
		Set o_strm = Nothing
		Form.RemoveAll
		Set Form = Nothing
		File.RemoveAll
		Set File = Nothing
	End Sub
	'属性：文件编码
	Public Property Let CharSet(ByVal str)
		s_charset = UCase(str)
	End Property
	'属性：单个文件最大尺寸
	Public Property Let MaxSize(ByVal n)
		i_maxsize = n
	End Property
	'属性：所有文件最大总尺寸
	Public Property Let TotalMaxSize(ByVal n)
		i_totalmaxsize = n
	End Property
	'属性：允许上传的文件类型
	Public Property Let Allowed(ByVal str)
		s_allowed = str
	End Property
	'属性：禁止上传的文件类型
	Public Property Let Denied(ByVal str)
		s_denied = str
	End Property
	'属性：文件上传后保存的路径(相对或绝对)
	Public Property Let SavePath(ByVal str)
		s_savepath = str
	End Property
	'属性：是否自动创建不存在的文件夹
	Public Property Let AutoMD(ByVal bool)
		b_automd = bool
	End Property
	'属性：是否重命名上传文件为随机文件名
	Public Property Let Random(ByVal bool)
		b_random = bool
	End Property	
	'属性：显示上传成功的文件数
	Public Property Get FileCount()
		FileCount = i_filecount
	End Property
		
	Public Property Get RadomName
		If s_filename = "" Then
			s_filename = GetNewFileName
		End If
		RadomName	= s_filename
	End Property
	
	Public Property Get FileSize
		If (Request.TotalBytes-400) >(1024*1024) Then
			FileSize = CInt((Request.TotalBytes-400)/1024/1024*100)/100 & "M"
		Else
			FileSize = CInt((Request.TotalBytes-400)/1024*10)/10 & "K"
		End If
	End Property

	Public Sub StartUpload()
		Dim aCType : aCType = Split(Request.ServerVariables("HTTP_CONTENT_TYPE"), ";")
		If LCase(aCType(0)) <> "multipart/form-data" Then
			Easp.Error.Raise 71
			Exit Sub
		End If
		Dim nTotalSize : nTotalSize	= Request.TotalBytes
		If nTotalSize < 1 Then
			Easp.Error.Raise 72
			Exit Sub
		End If
		If i_totalmaxsize > 0 And nTotalSize > i_totalmaxsize Then
			Easp.Error.Raise 73
			Exit Sub
		End If
		o_strm.Write Request.BinaryRead(nTotalSize)
		o_strm.Position = 0
		
		Dim oTotalData, oFormStream, sFormHeader, sFormName, sFormValue, bCrLf, nBoundLen, nFormStart, nFormEnd, nPosStart, nPosEnd, sBoundary
		
		oTotalData	= o_strm.Read
		bCrLf		= ChrB(13) & ChrB(10)
		sBoundary	= MidB(oTotalData, 1, InStrB(1, oTotalData, bCrLf) - 1)
		nBoundLen	= LenB(sBoundary) + 2
		nFormStart	= nBoundLen
		
		Set oFormStream = Server.CreateObject("ADODB.Stream")
		
		Do While (nFormStart + 2) < nTotalSize
			nFormEnd	= InStrB(nFormStart, oTotalData, bCrLf & bCrLf) + 3
			
			With oFormStream
				.Type	= 1
				.Mode	= 3
				.Open
				o_strm.Position = nFormStart
				o_strm.CopyTo oFormStream, nFormEnd - nFormStart
				.Position	= 0
				.Type		= 2
				.CharSet	= s_charset
				sFormHeader	= .ReadText
				.Close
			End With
			
			nFormStart	= InStrB(nFormEnd, oTotalData, sBoundary) - 1
			nPosStart	= InStr(22, sFormHeader, " name=", 1) + 7
			nPosEnd		= InStr(nPosStart, sFormHeader, """")
			sFormName	= Mid(sFormHeader, nPosStart, nPosEnd - nPosStart)
			
			If InStr(45, sFormHeader, " filename=", 1) > 0 Then
				Set File(sFormName)	= New Easp_upload_FileInfo
				File(sFormName).FormName = sFormName
				File(sFormName).Start = nFormEnd
				File(sFormName).Size = nFormStart - nFormEnd - 2
				nPosStart = InStr(nPosEnd, sFormHeader, " filename=", 1) + 11
				nPosEnd = InStr(nPosStart, sFormHeader, """")
				File(sFormName).ClientPath = Mid(sFormHeader, nPosStart, nPosEnd - nPosStart)
				File(sFormName).Name = Mid(File(sFormName).ClientPath, InStrRev(File(sFormName).ClientPath, "\") + 1)
				File(sFormName).Ext = LCase(Mid(File(sFormName).Name, InStrRev(File(sFormName).Name, ".") + 1))
				nPosStart = InStr(nPosEnd, sFormHeader, "Content-Type: ", 1) + 14
				nPosEnd = InStr(nPosStart, sFormHeader, vbCr)
				File(sFormName).MIME = Mid(sFormHeader, nPosStart, nPosEnd - nPosStart)
			Else
				With oFormStream
					.Type = 1
					.Mode = 3
					.Open
					o_strm.Position = nPosEnd
					o_strm.CopyTo oFormStream, nFormStart - nFormEnd - 2
					.Position = 0
					.Type = 2
					.CharSet = s_charset
					sFormValue = .ReadText
					Form(sFormName)	= Easp.IIF(Form.Exists(sFormName),Form(sFormName)&", "&sFormValue,sFormValue)
					.Close
				End With
			End If
			nFormStart	= nFormStart + nBoundLen
		Loop
		oTotalData = ""
		Set oFormStream = Nothing
		'##CheckFile()
	End Sub
	
	Public Sub SaveAs(sItem, sFileName)
		If File(sItem).Size < 1 Then
			Easp.Error.Raise 74
			Exit Sub
		ElseIf i_maxsize > 0 And File(sItem).Size > i_maxsize Then
			Easp.Error.Raise 75
			Exit Sub
		End If
		If Not isAllowed(File(sItem).Ext) Then
			Easp.Error.Raise 76
			Exit Sub
		End If
		
		Dim oFileStream
		Set oFileStream = Server.CreateObject("ADODB.Stream")
		With oFileStream
			.Type		= 1
			.Mode		= 3
			.Open
			o_strm.Position = File(sItem).Start
			o_strm.CopyTo oFileStream, File(sItem).Size
			.Position	= 0
			'##.SaveToFile sFileName, 2   '暂时不保存文件
			.Close
		End With
		Set oFileStream = Nothing
	End Sub
	
	Function SaveFile(ByVal sItem, ByVal sPath)
		On Error Resume Next
		SaveFile = True
		Dim fileStream, sName
		Set fileStream = Server.CreateObject("ADODB.Stream")
		With fileStream
			.Type = 1
			.Mode = 3
			.Open
			o_strm.Position = File(sItem).Start
			o_strm.CopyTo fileStream, File(sItem).Size
			.Position = 0
			.SaveToFile sPath, 2
			.Close
		End With
		Set fileStream = Nothing
		If Err.Number<>0 Then
			Easp.Error.Raise 77
			SaveFile = False
		End If
		Err.Clear
	End Function
	
	Function CheckFile()
		CheckFile = True : i_filecount = 0
		If File.Count > 0 Then
			For Each item In File
				If File(item).Size > 0 Then
					If CheckOneFile(item) Then
						i_filecount = i_filecount + 1
					Else
						File.Remove(item) : CheckFile = False
					End If
				Else
					File.Remove(item)
				End If
			Next
		Else
			CheckFile = False
		End If
	End Function
	
	Function CheckOneFile(ByVal sItem)
		CheckOneFile = True
		If Not File.Exists(sItem) Then
			Easp.Error.Msg = "表单控件("""&sItem&""")不存在"
			Easp.Error.Raise 78
			CheckOneFile = False : Exit Function
		End If
		Dim cp : cp = File(sItem).ClientPath
		If Not isAllowed(File(sItem).Ext) Then
			Easp.Error.Msg = "不允许上传此类型的文件("&File(sItem).ClientPath&")"
			Easp.Error.Raise 77
			CheckOneFile = False : Exit Function
		End If
		If File(sItem).Size < 1 Then
			Easp.Error.Msg = "上传文件不能为空("&Easp.IIF(Easp.isN(cp),sItem,cp)&")"
			Easp.Error.Raise 77
			CheckOneFile = False : Exit Function
		ElseIf i_maxsize > 0 And File(sItem).Size > i_maxsize Then
			Easp.Error.Msg = "文件大小超过了限制("&File(sItem).ClientPath&")"
			Easp.Error.Raise 77
			CheckOneFile = False : Exit Function
		End If
	End Function
	
	'取文件夹绝对路径
	Private Function absPath(ByVal path)
		Dim p : p = path
		If Instr(p,":") = 0 Then p = Server.MapPath(p)
		If Right(p,1) <> "\" Then p = p & "\"
		absPath = p
	End Function
	
	Private Function isAllowed(ByVal sExt)
		If Easp.isN(s_allowed) Then
			isAllowed = Easp.isN(s_denied) Or Not Easp.Test(sExt,"^("&s_denied&")$")
		Else
			isAllowed = Easp.Test(sExt,"^("&s_allowed&")$")
		End If
	End Function
	
	Private Function GetRandomName(ByVal sItem)
		GetRandomName = Easp.DateTime(Now(),"yymmddhhiiss") & Easp.Rand(10000,99999)' & "." & File(sItem).Ext
	End Function
	
	Private Function GetNewFileName()
		dim ranNum
		randomize
		ranNum = int(90000*rnd) + 10000
		GetNewFileName = year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & ranNum
	End Function
End Class

Class Easp_upload_FileInfo
	Public FormName, ClientPath, Path, Name, Ext, Content, Size, MIME, Start
End Class
%>