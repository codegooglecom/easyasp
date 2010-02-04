<%
'#################################################################################
'##	easp.upload.asp
'##	------------------------------------------------------------------------------
'##	Feature		:	EasyAsp Upload Class
'##	Version		:	v2.2 Alpha
'##	Author		:	Coldstone(coldstone[at]qq.com)
'##	Update Date	:	2010/02/03 11:11:14
'##	Description	:	Upload file(s) with EasyASP
'#################################################################################
Dim EasyAsp_o_updata
Class EasyAsp_Upload
	Public Form, File, Count
	'Private o_prog
	Private s_charset,s_allowed,s_denied,s_filename,s_savepath, s_jsonPath
	Private i_maxsize,i_totalmaxsize,i_filecount,i_blockSize
	Private b_automd,b_random
	'构造函数
    Private Sub Class_Initialize 
		s_jsonPath = ""
		s_charset	= Easp.CharSet
		'分块上传默认每次上传64K
		i_blockSize = 64 * 1024
		Easp.Error(71) = "表单类型错误，表单只能是""multipart/form-data""类型！"
		Easp.Error(72) = "请先选择要上传的文件！"
		Easp.Error(73) = "上传文件失败，上传文件总大小超过了限制！"
		Easp.Error(74) = "上传文件失败，上传文件不能为空！"
		Easp.Error(75) = "上传文件失败，文件大小超过了限制！"
		Easp.Error(76) = "上传文件失败，不允许上传此类型的文件！"
		Easp.Error(77) = "上传文件失败！"
		Easp.Error(78) = "获取文件失败！"
		Set Form = Server.CreateObject("Scripting.Dictionary")
		Set File = Server.CreateObject("Scripting.Dictionary")
		Count = 0
    End Sub
	
	'属性：文件编码
	Public Property Let CharSet(ByVal s)
		s_charset = UCase(s)
	End Property
	'进度条Json文件保位置：
	Public Property Let JsonPath(ByVal s)
		s_jsonPath = Easp.IIF(Mid(s,2,1)=":", s, Server.MapPath(s))
	End Property
	Public Property Get JsonPath()
		JsonPath = s_jsonPath
	End Property
	'属性：分块上传大小，单位K
	Public Property Let BlockSize(ByVal i)
		i_blockSize = Int(i) * 1024
	End Property
	
	'初始化：
	Public Sub StartUpload
		Dim o_strm, o_prog, o_file
		Dim s_total, s_block, s_blockData, s_start, s_formName, s_formValue, s_fileName, s_data
		Dim i_loaded, i_block, i_formStart, i_formEnd, i_Start, i_End, i_dataStart, i_dataEnd
		'取得表单总大小
		s_total = Request.TotalBytes
		'如果表单的内容为空，则退出上传程序
		If s_total < 1 Then Easp.Error.Raise 72 : Exit Sub
		Set o_strm = Server.CreateObject("ADODB.Stream")
		'临时数据储存区
		Set EasyAsp_o_updata = Server.CreateObject("ADODB.Stream")
		EasyAsp_o_updata.Type = 1
		EasyAsp_o_updata.Mode =3
		EasyAsp_o_updata.Open
		'已读取的大小
		i_loaded = 0
		'记录进度到Json文件
		Set o_prog = New Easp_Upload_Progress
		o_prog.ProgressInit(s_jsonPath)
		o_prog.UpdateProgress s_total,0
		'循环分块读取
		Do While i_loaded < s_total
			i_block = i_blockSize
			If i_block + i_loaded > s_total Then i_block = s_total - i_loaded
			s_block = Request.BinaryRead(i_block)
			i_loaded = i_loaded + i_block
			'写入分块数据
			EasyAsp_o_updata.Write s_block
			'更新进度条文件
			o_prog.UpdateProgress s_total,i_loaded 
		Loop
		'EasyAsp_o_updata.Write  Request.BinaryRead(s_total)
		'将数据块读出处理
		EasyAsp_o_updata.Position = 0
		s_blockData = EasyAsp_o_updata.Read
		i_formStart = 1
		i_formEnd = LenB(s_blockData)
		CrLf = chrB(13) & chrB(10)
		s_start = MidB(s_blockData,1, InStrB(i_formStart,s_blockData,CrLf)-1)
		i_start = LenB(s_start)
		i_formStart = i_formStart + i_start + 1
		While (i_formStart + 10) < i_formEnd 
			i_End = InStrB(i_formStart,s_blockData,CrLf & CrLf)+3
			o_strm.Type = 1
			o_strm.Mode =3
			o_strm.Open
			EasyAsp_o_updata.Position = i_formStart
			EasyAsp_o_updata.CopyTo o_strm, i_End-i_formStart
			o_strm.Position = 0
			o_strm.Type = 2
			o_strm.Charset = s_charset
			s_data = o_strm.ReadText
			o_strm.Close
			'Easp.w s_data
			'取得表单项目名称
			i_formStart = InStrB(i_End,s_blockData,s_start)
			i_dataStart = InStr(22,s_data,"name=""",1) + 6
			i_dataEnd = InStr(i_dataStart,s_data,"""",1)
			s_formName = lcase(Mid(s_data,i_dataStart,i_dataEnd-i_dataStart))
			'如果是文件
			If InStr(45,s_data,"filename=""",1) > 0 Then
				Set o_file = New Easp_Upload_FileInfo
				'取得文件大小
				o_file.Size = i_formStart - i_End - 3
				'如果不为空
				If o_file.Size > 0 Then
					'取得文件名
					i_dataStart = InStr(i_dataEnd,s_data,"filename=""",1) + 10
					i_dataEnd = InStr(i_dataStart,s_data,"""",1)
					s_fileName = Mid(s_data,i_dataStart,i_dataEnd-i_dataStart)
					o_file.Client = s_fileName
					o_file.OldPath = Left(s_fileName, InstrRev(s_fileName, "\"))
					o_file.Name = Mid(s_fileName, InstrRev(s_fileName, "\")+1)
					o_file.Ext = Mid(o_file.Name, InstrRev(o_file.Name,".")+1)
					'取得文件类型
					i_dataStart = InStr(i_dataEnd,s_data,"Content-Type: ",1) + 14
					i_dataEnd = InStr(i_dataStart,s_data,vbCr)
					o_file.MIME = Mid(s_data,i_dataStart,i_dataEnd-i_dataStart)
					o_file.Start = i_End
					o_file.FormName = s_formName
					Count = Count + 1
				End If
				If NOT File.Exists(s_formName) Then
					File.Add s_formName, o_file
				End If
				Set o_file = Nothing
			Else
				'如果是表单项目
				o_strm.Type = 1
				o_strm.Mode = 3
				o_strm.Open
				EasyAsp_o_updata.Position = i_End 
				EasyAsp_o_updata.CopyTo o_strm, i_formStart-i_End-3
				o_strm.Position = 0
				o_strm.Type = 2
				o_strm.Charset = s_charset
				s_formValue = o_strm.ReadText 
				o_strm.Close
				If Form.Exists(s_formName) Then
					Form(s_formName) = Form(s_formName) & ", " & s_formValue
				Else
					Form.Add s_formName, s_formValue
				End If
			End If
			i_formStart = i_formStart + i_start + 1
		Wend
		s_blockData = ""
		Set o_strm = Nothing
		Set o_prog = Nothing
	End Sub
    
    Private Sub Class_Terminate  
		If Request.TotalBytes > 0 Then
			Form.RemoveAll
			File.RemoveAll
			EasyAsp_o_updata.Close
			Set EasyAsp_o_updata = Nothing
		End If
		Set Form=Nothing
		Set File=Nothing
		Easp.Use "Fso"
		Easp.Fso.DelFile s_jsonPath
    End Sub
End Class
'上传文件信息
Class Easp_Upload_FileInfo
	Public FormName, Client, OldPath, NewPath, Name, Ext, Size, MIME, Start
	Private Sub Class_Initialize 
		FormName = ""
		Client = ""
		OldPath = ""
		NewPath = ""
		Name = ""
		Ext = ""
		Size = 0
		Start = 0
		MIME = ""
	End Sub
	Public Function SaveAs(ByVal p)
		Dim o_strm
		SaveAs = True
		'如果上传文件为空
		If Size<=0 Then
			SaveAs = False
			Easp.Error.Raise 74
			Exit Function
		End If
		'如果上传的不是文件或保存路径为空
		If Easp.IsN(p) Or Easp.IsN(Name) Or Start = 0 Or Right(p,1)="/" Then
			SaveAs = False
			Exit Function
		End If
		'保存文件
		Set o_strm = Server.CreateObject("Adodb.Stream")
		o_strm.Mode = 3
		o_strm.Type = 1
		o_strm.Open
		EasyAsp_o_updata.position = Start
		EasyAsp_o_updata.copyto o_strm, Size
		o_strm.SaveToFile p, 2
		o_strm.Close
		Set o_strm = Nothing
	End Function
End Class

Class Easp_Upload_Progress
  Dim objDom,xmlPath
    Dim startTime
  Private Sub Class_Initialize
    End Sub
    
    Public Sub ProgressInit(xmlPathTmp)
      Dim objRoot,objChild
        Dim objPI
        xmlPath = xmlPathTmp
        Set objDom = Server.CreateObject("Microsoft.XMLDOM")
        Set objRoot = objDom.createElement("progress")
        objDom.appendChild objRoot
        
        Set objChild = objDom.createElement("totalbytes")
        objChild.Text = "0"
        objRoot.appendChild objChild
        Set objChild = objDom.createElement("uploadbytes")
        objChild.Text = "0"
        objRoot.appendChild objChild
        Set objChild = objDom.createElement("uploadpercent")
        objChild.Text = "0%"
        objRoot.appendChild objChild
        Set objChild = objDom.createElement("uploadspeed")
        objChild.Text = "0"
        objRoot.appendChild objChild
        Set objChild = objDom.createElement("totaltime")
        objChild.Text = "00:00:00"
        objRoot.appendChild objChild
        Set objChild = objDom.createElement("lefttime")
        objChild.Text = "00:00:00"
        objRoot.appendChild objChild
        
        Set objPI = objDom.createProcessingInstruction("xml","version='1.0' encoding='utf-8'")
        objDom.insertBefore objPI, objDom.childNodes(0)
		Easp.wn "进度条文件地址：" & xmlPath
        objDom.Save xmlPath
        Set objPI = Nothing
        Set objChild = Nothing
        Set objRoot = Nothing
        Set objDom = Nothing
    End Sub
    
    Sub UpdateProgress(tBytes,rBytes)
      Dim eTime,currentTime,speed,totalTime,leftTime,percent
        If rBytes = 0 Then
            startTime = Timer
            Set objDom = Server.CreateObject("Microsoft.XMLDOM")
            objDom.load(xmlPath)
            objDom.selectsinglenode("//totalbytes").text=tBytes
            objDom.save(xmlPath)
        Else
          speed = 0.0001
          currentTime = Timer
        eTime = currentTime - startTime
            If eTime>0 Then speed = rBytes / eTime
            totalTime = tBytes / speed
            leftTime = (tBytes - rBytes) / speed
            percent = Round(rBytes *100 / tBytes)
            'objDom.selectsinglenode("//uploadbytes").text = rBytes
            'objDom.selectsinglenode("//uploadspeed").text = speed
            'objDom.selectsinglenode("//totaltime").text = totalTime
            'objDom.selectsinglenode("//lefttime").text = leftTime
            objDom.selectsinglenode("//uploadbytes").text = FormatFileSize(rBytes) & " / " & FormatFileSize(tBytes)
            objDom.selectsinglenode("//uploadpercent").text = percent
            objDom.selectsinglenode("//uploadspeed").text = FormatFileSize(speed) & "/sec"
            objDom.selectsinglenode("//totaltime").text = SecToTime(totalTime)
            objDom.selectsinglenode("//lefttime").text = SecToTime(leftTime)
            objDom.save(xmlPath)        
        End If
    End Sub
    private Function SecToTime(sec)
        Dim h:h = "0"
        Dim m:m = "0"
        Dim s:s = "0"
        h = round(sec / 3600)
        m = round( (sec mod 3600) / 60)
        s = round(sec mod 60)
        If LEN(h)=1 Then h = "0" & h
        If LEN(m)=1 Then m = "0" & m
        If LEN(s)=1 Then s = "0" & s
        SecToTime = (h & ":" & m & ":" & s)
    End Function
        
    private Function FormatFileSize(fsize)
        Dim radio,k,m,g,unitTMP
        k = 1024
        m = 1024*1024
        g = 1024*1024*1024
        radio = 1
        If Fix(fsize / g) > 0.0 Then
            unitTMP = "GB"
            radio = g
        ElseIf Fix(fsize / m) > 0 Then
            unitTMP = "MB"
            radio = m
        ElseIf Fix(fsize / k) > 0 Then
            unitTMP = "KB"
            radio = k
        Else
            unitTMP = "B"
            radio = 1
        End If
        If radio = 1 Then
            FormatFileSize = fsize & "&nbsp;" & unitTMP
        Else
            FormatFileSize = FormatNumber(fsize/radio,3) & unitTMP
        End If
    End Function
    Private Sub Class_Terminate  
      Set objDom = Nothing
    End Sub
End Class
%>