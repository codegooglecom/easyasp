<%
'######################################################################
'## easp.cache.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp Cache Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com) & SunYu
'## Update Date :   2010/09/30 00:03:09
'## Description :   Save and Get Cache With EasyAsp
'##
'######################################################################
Class EasyAsp_Cache
	Public Items, CountEnabled, Expires, FileType
	Private s_path, b_fsoOn
	Private Sub Class_Initialize
		Set Items = Server.CreateObject("Scripting.Dictionary")
		s_path = Server.MapPath("/_cache") & "\"
		CountEnabled = True
		Expires = 5
		FileType = ".cache"
		Easp.Error(91) = "当前对象不允许缓存到内存缓存"
		Easp.Error(92) = "缓存文件不存在"
		Easp.Error(93) = "当前内容不允许缓存到文件缓存"
		If TypeName(Easp.Fso) = "EasyASP_Fso" Then
			b_fsoOn = True
		Else
			Easp.Use "Fso"
			b_fsoOn = False
		End If
	End Sub
	Private Sub Class_Terminate
		If Not b_fsoOn Then
			Set Easp.Fso = Nothing
			Set Easp.Fso = New EasyAsp_obj
		End If
		Set Items = Nothing
	End Sub
	Public Function [New]()
		Set [New] = New EasyAsp_Cache
	End Function
	Public Property Get Count
		Count = Easp.IIF(CountEnabled,Easp_Cache_Count,-1)
	End Property
	Public Property Let Item(ByVal p, ByVal v)
		If IsNull(p) Then p = ""
		If Not IsObject(Items(p)) Then
			Set Items(p) = New Easp_Cache_Info
			Items(p).CountEnabled = CountEnabled
			Items(p).Expires = Expires
			Items(p).FileType = FileType
		End If
		Items(p).Name = p
		Items(p).Value = v
		Items(p).SavePath = s_path
	End Property
	Public Default Property Get Item(ByVal p)
		If Not IsObject(Items(p)) Then
			Set Items(p) = New Easp_Cache_Info
			Items(p).Name = p
			Items(p).SavePath = s_path
			Items(p).CountEnabled = CountEnabled
			Items(p).Expires = Expires
			Items(p).FileType = FileType
		End If
		set Item = Items(p)
	End Property
  Public Property Let SavePath(ByVal s)
		If Not Instr(s,":") = 2 Then s = Server.MapPath(s)
		If Right(s,1) <> "\" Then s = s & "\"
		s_path = s
  End Property
	Public Property Get SavePath()
		SavePath = s_path
  End Property
	Public Sub SaveAll
		Dim f
		For Each f In Items
			Items(f).Save
		Next
	End Sub
	Public Sub SaveAppAll  
		Dim f 
		For Each f In Items
			Items(f).SaveApp
		Next
	End Sub
	Public Sub RemoveAll
		Dim f
		For Each f In Items
			Items(f).Remove
		Next
	End Sub
	Public Sub RemoveAppAll  
		Dim f 
		For Each f In Items
			Items(f).RemoveApp
		Next
	End Sub
	Public Sub [Clear]
		RemoveAll
		RemoveAppAll
		Easp.RemoveApp "Easp_Cache_Count"
	End Sub
End Class
Function Easp_Cache_Count()
	Easp_Cache_Count = 0
	Dim n : n = Easp.GetApp("Easp_Cache_Count")
	If IsArray(n) Then
		If Ubound(n) = 1 Then Easp_Cache_Count = n(0)
	End If
End Function
Function Easp_CacheCount_Change(ByVal a, ByVal t)
	Dim n : n = Easp.GetApp("Easp_Cache_Count")
	If isArray(n) Then
		If Ubound(n) = 1 Then
			If TypeName(n(1)) = "Dictionary" Then
				If t = 1 Then n(1)(a) = a
				If t = -1 Then
					If n(1).Exists(a) Then n(1).Remove(a)
				End If
				Easp.SetApp "Easp_Cache_Count", Array(n(1).Count,n(1))
			End If
		End If
	Else
		Dim dic : Set dic = Server.CreateObject("Scripting.Dictionary")
		If t = 1 Then dic(a) = a
		Easp.SetApp "Easp_Cache_Count", Array(Easp.IIF(t=1,1,0),dic)
	End If
End Function
class Easp_Cache_Info
	Public SavePath, [Name], CountEnabled, FileType
	Private i_exp, d_exp, o_value
	Private Sub Class_Initialize
		i_exp = 5
		d_exp = ""
	End Sub
	Private Sub Class_Terminate
		If IsObject(o_value) Then Set o_value = Nothing
	End Sub
	Public Property Let Expires(ByVal i)
		If isDate(i) Then
			d_exp = CDate(i)
		ElseIf isNumeric(i) Then
			If i>0 Then
				i_exp = i
			ElseIf i=0 Then
				i_exp = 60*24*365*99
			End If
		End If
  End Property
	Public Property Get Expires()
		Expires = Easp.IfHas(d_exp, i_exp)
	End Property
	Public Property Let [Value](ByVal s)
		If IsObject(s) Then
			Select Case TypeName(s)
				Case "Recordset"
					Set o_value = s.Clone
				Case Else
					Set o_value = s
			End Select
		Else
			o_value = s
		End If
  End Property
	Public Default Property Get [Value]()
		Dim app : app = Easp.GetApp(Me.Name)
		If IsArray(app) Then
			If UBound(app) = 1 Then
				If IsDate(app(0)) Then
					If IsObject(app(1)) Then
						Set [Value] = app(1)
						Exit Property
					Else
						[Value] = app(1)
						If Easp.Has([Value]) Then Exit Property
					End If
				End If
			End If
		End If
		If Easp.Fso.IsFile(FilePath) Then
			On Error Resume Next
			Dim rs
			set rs = Server.CreateObject("Adodb.Recordset")
			rs.Open FilePath
			If Err.Number <> 0 Then
				Err.Clear
				[Value] = Easp.Fso.Read(FilePath)
			Else
				Set [Value] = rs
			End If
		Else
			Easp.Error.Msg = "("""&Easp.HtmlEncode(Me.Name)&""")" : Easp.Error.Raise 92
		End If
	End Property
	Public Sub SaveApp
		Dim appArr(1) : appArr(0) = Now()
		If IsObject(o_value) Then
			Select Case TypeName(o_value)
				Case "Dictionary"
					Set appArr(1) = o_value
				Case "Recordset"
					appArr(1) = o_value.GetRows(-1)
				Case Else
					Easp.Error.Msg = "("""&Easp.HtmlEncode(Me.Name)&" &gt; "&TypeName(o_value)&""")" : Easp.Error.Raise 91
			End Select
		Else
			appArr(1) = o_value
		End If
		Easp.SetApp Me.Name, appArr
		If CountEnabled Then Easp_CacheCount_Change Me.Name, 1
	End Sub
	Public Sub Save
		Select Case TypeName(o_value)
			Case "Recordset"
				Easp.Fso.CreateFile FilePath, "rs"
				Easp.Fso.DelFile FilePath
				o_value.Save FilePath, adPersistXML
				If CountEnabled Then Easp_CacheCount_Change Me.Name, 1
			Case "String"
				Easp.Fso.CreateFile FilePath, o_value
				If CountEnabled Then Easp_CacheCount_Change Me.Name, 1
			Case Else
				Easp.Error.Msg = "("""&Easp.HtmlEncode(Me.Name)&""")" : Easp.Error.Raise 93
		End Select
	End Sub
	Public Sub Remove
		If Not Easp.Test(DelPath,"[*?]") Then
			If Easp.Fso.IsExists(DelPath) Then Easp.Fso.Del DelPath
			If CountEnabled Then Easp_CacheCount_Change Me.Name, -1
		Else
			Easp.Fso.DelFile left(DelPath,len(DelPath)-Len(FileType))
			Easp.Fso.DelFolder left(DelPath,len(DelPath)-Len(FileType))
			If CountEnabled Then Easp_CacheCount_Change Me.Name, -1
		End If
	End Sub
	Public Sub RemoveApp
		If Easp.Has(Me.Name) Then Easp.RemoveApp Me.Name
		If CountEnabled Then Easp_CacheCount_Change Me.Name, -1
	End Sub
	Public Property Get FilePath()
		FilePath = TransPath("[\\:""*?<>|\f\n\r\t\v\s]")
	End Property
	Private Function DelPath()
		DelPath = TransPath("[\\:""<>|\f\n\r\t\v\s]")
	End Function
	Private Function TransPath(ByVal fe)
		Dim s_p : s_p = ""
		Dim parr : parr = split(Me.Name,"/")
		Dim i
		for i = 0 to UBound(parr)
			If Easp.Test(parr(i),fe) Then parr(i)=Server.URLEncode(parr(i))
			s_p = s_p & "_" & parr(i)
			If i < UBound(parr) Then
				s_p = s_p & "\"
			End If
		next
		If s_p="" Then s_p="_"
		TransPath = SavePath & s_p & FileType
	End Function	
	Public Function Ready()
		Dim app : app = Easp.GetApp(Me.Name)
		Ready = False
		If IsArray(app) Then
			If UBound(app) = 1 Then
				If IsDate(app(0)) Then
					Ready = isValid(app(0))
					If Ready Then Exit Function
				End If
			End If
		End If
		If Easp.Fso.IsFile(FilePath) Then
			Ready = isValid(Easp.Fso.GetAttr(FilePath,1))
		End If
	End Function
	Private Function isValid(ByVal t)
		If IsDate(t) Then
			If Easp.Has(d_exp) Then
				isValid = (DateDiff("s",Now,d_exp) > 0)
			Else
				isValid = (DateDiff("s",t,Now) < i_exp*60)
			End If
		End If
	End Function
End Class
%>