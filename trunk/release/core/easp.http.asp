<%
'######################################################################
'## easp.http.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp XMLHTTP Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/09/30 00:03:09
'## Description :   Request XMLHttp Data in EasyASP
'## http://msdn.microsoft.com/en-us/library/ms535874(VS.85).aspx
'######################################################################
Class EasyAsp_Http
	Public Url, Method, CharSet, Async, User, Password, Html, Headers, Body, Text, SaveRandom
	Public ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
	Private s_data, s_url, s_ohtml, o_rh', a_rh()
	
	Private Sub Class_Initialize
		CharSet = ""
		Async = False
		User = ""
		Password = ""
		s_data = ""
		s_url = ""
		Html = ""
		Headers = ""
		Body = Empty
		Text = Empty
		SaveRandom = True
		ResolveTimeout = 20000
		ConnectTimeout = 20000
		SendTimeout = 300000
		ReceiveTimeout = 60000
		Easp.Error(46) = "远程服务器没有响应"
		Easp.Error(47) = "服务器不支持XMLHTTP组件"
		Easp.Error(48) = "要获取的页面地址不能为空"
		Easp.Use "List"
		Set o_rh = Easp.List.New
	End Sub
	
	Private Sub Class_Terminate
		Set o_rh = Nothing
	End Sub
	Public Function [New]()
		Set [New] = New EasyAsp_Http
	End Function
	Public Property Let Data(ByVal s)
		s_data = s
	End Property
	Public Sub SetHeader(ByVal a)
		Dim i,n,v
		If isArray(a) Then
			For i = 0 To Ubound(a)
				n = Replace(Easp.Cleft(a(i),":"),"-","_")
				v = Easp.CRight(a(i),":")
				o_rh(n) = v
			Next
		Else
			n = Replace(Easp.Cleft(a,":"),"-","_")
			v = Easp.CRight(a,":")
			o_rh(n) = v
		End If
	End Sub
	Public Property Let RequestHeader(ByVal n, ByVal v)
		n = Replace(n,"-","_")
		o_rh(n) = v
	End Property
	Public Property Get RequestHeader(ByVal n)
		If Easp.Has(n) Then
			RequestHeader = o_rh(n)
		Else
			RequestHeader = Join(o_rh.Hash,vbCrLf)
		End If
	End Property
	Private Sub SetHeaderTo(ByRef o)
		Dim maps,key
		Set maps = o_rh.Maps
		For Each key In maps
			If Not isNumeric(key) Then
				o.setRequestHeader Replace(key,"_","-"), o_rh(key)
			End If
		Next
		Set maps = Nothing
	End Sub
	Public Function [Open]
		[Open] = GetData(Url, Method, Async, s_data, User, Password)
	End Function
	Public Function [Get](ByVal uri)
		[Get] = GetData(uri, "GET", Async, s_data, User, Password)
	End Function
	Public Function Post(ByVal uri)
		Post = GetData(uri, "POST", Async, s_data, User, Password)
	End Function
	Public Function GetData(ByVal uri, ByVal m, ByVal async, ByVal data, ByVal u, ByVal p)
		Dim o,chru
		If Easp.isInstall("MSXML2.serverXMLHTTP") Then
			Set o = Server.CreateObject("MSXML2.serverXMLHTTP")
		ElseIf Easp.isInstall("MSXML2.XMLHTTP") Then
			Set o = Server.CreateObject("MSXML2.XMLHTTP")
		ElseIf Easp.isInstall("Microsoft.XMLHTTP") Then
			Set o = Server.CreateObject("Microsoft.XMLHTTP")
		Else
			Easp.Error.Raise 47
			Exit Function
		End If
		o.SetTimeOuts ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
		If Easp.IsN(uri) Then Easp.Error.Raise 48 : Exit Function
		If Easp.Test(uri,"^[\w\d-]+>https?://") Then
			CharSet = Easp.CLeft(uri,">")
			uri = Easp.CRight(uri,">")
		End If
		s_url = uri
		m = Easp.IIF(Easp.Has(m),UCase(m),"GET")
		If Easp.IsN(async) Then async = False
		If m = "GET" And Easp.Has(data) Then uri = uri & Easp.IIF(Instr(uri,"?")>0, "&", "?") & Serialize__(data)
		If Easp.Has(u) Then
			o.open m, uri, async, u, p
		Else
			o.open m, uri, async
		End If
		If m = "POST" Then
			If Not o_rh.HasIndex("Content_Type") Then
				o_rh("Content_Type") = "application/x-www-form-urlencoded"
			End If
			SetHeaderTo o
			o.send Serialize__(data)
		Else
			SetHeaderTo o
			o.send
		End If
		If o.readyState <> 4 Then
			GetData = "error:server is down"
			Set o = Nothing
			Easp.Error.Raise 46
			Exit Function
		ElseIf o.Status = 200 Then
			Headers = o.getAllResponseHeaders()
			Body = o.responseBody
			Text = o.responseText
			If Easp.IsN(CharSet) Then
				If Easp.Test(Headers,"charset=([\w-]+)") Then
					CharSet = Easp.RegReplace(Headers,"([\s\S]+)charset=([\w-]+)([\s\S]+)","$2")
				ElseIf Easp.Test(Headers,"Content-Type: ?text/xml") Then
					CharSet = Easp.RegReplace(Text,"^<\?xml\s+[^>]+encoding\s*=\s*""([^""]+)""[^>]*\?>([\s\S]+)","$1")
				ElseIf Easp.Test(Text,"<meta\s+http-equiv\s*=\s*[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>") Then
					CharSet = Easp.RegReplace(Text,"([\s\S]+)<meta\s+http-equiv\s*=\s*[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>([\s\S]+)","$2")
				End If
				If Easp.IsN(CharSet) Then CharSet = Easp.CharSet
			End If
			GetData = Bytes2Bstr__(Body, CharSet)
		Else
			GetData = "error:" & o.Status & " " & o.StatusText
		End If
		Set o = Nothing
		s_ohtml = GetData
		Html = s_ohtml
	End Function
	Public Function Find(ByVal rule)
		Find = Find_(s_ohtml, rule)
	End Function
	Public Function Find_(ByVal s, ByVal rule)
		If Easp.Test(s,rule) Then Find_ = Easp.RegReplace(s,"([\s\S]*)("&rule&")([\s\S]*)","$2")
	End Function
	Public Function [Select](ByVal rule, ByVal part)
		[Select] = Select_(s_ohtml, rule, part)
	End Function
	Public Function Select_(ByVal s, ByVal rule, ByVal part)
		If Easp.Test(s,rule) Then
			part = Replace(part,"$0",Find_(s,rule))
			Select_ = Easp.RegReplace(s,"(?:[\s\S]*)(?:"&rule&")(?:[\s\S]*)",part)
		End If
	End Function
	Public Function Search(ByVal rule)
		Search = Search_(s_ohtml, rule)
	End Function
	Public Function Search_(ByVal s, ByVal rule)
		Dim matches,match,arr(),i : i = 0
		Set matches = Easp.RegMatch(s,rule)
		ReDim arr(matches.Count-1)
		For Each match In matches
			arr(i) = match.Value
			i = i + 1
		Next
		Set matches = Nothing
		Search_ = arr
	End Function
	Public Function SubStr(ByVal tagStart, ByVal tagEnd, ByVal tagSelf)
		SubStr = SubStr_(s_ohtml,tagStart,tagEnd,tagSelf)
	End Function
	Public Function SubStr_(ByVal s, ByVal tagStart, ByVal tagEnd, ByVal tagSelf)
		Dim posA, posB, first, between
		posA = instr(1,s,tagStart,1)
		If posA=0 Then SubStr_ = "源代码中不包括此开始标签" : Exit Function
		posB = instr(PosA+Len(tagStart),s,tagEnd,1) 
		If posB=0 Then SubStr_ = "源代码中不包括此结束标签" : Exit Function
		Select Case tagSelf
			Case 1
				first = posA
				between = posB+len(tagEnd)-first
			Case 2
				first = posA
				between = posB-first
			Case 3
				first = posA+len(tagStart)
				between = posB+len(tagEnd)-first
			Case Else
				first = posA+len(tagStart)
				between = posB-first
		End Select
		SubStr_ = Mid(s,first,between)
	End Function
	Public Function SaveImgTo(ByVal p)
		SaveImgTo = SaveImgTo_(s_ohtml,p)
	End Function
	Public Function SaveImgTo_(ByVal s, ByVal p)
		Dim a,b, i, img, ht, tmp, src
		a = Easp.GetImg(s)
		b = Easp.GetImgTag(s)
		If Easp.Has(a) Then
			Easp.Use "Fso"
			For i = 0 To Ubound(a)
				If SaveRandom Then
					img = Easp.DateTime(Now,"ymmddhhiiss"&Easp.RandStr("5:0123456789")) & Mid(a(i),InstrRev(a(i),"."))
				Else
					img = Mid(a(i),InstrRev(a(i),"/")+1)
				End If
				Set ht = Easp.Http.New
				ht.Get TransPath(s_url, a(i))
				tmp = Easp.Fso.SaveAs(p & img, ht.Body)
				Set ht = Nothing
				If tmp Then
					src = Easp.RegReplace(b(i),"(<img\s[^>]*src\s*=\s*([""|']?))("&a(i)&")(\2[^>]*>)","$1"&p&img&"$4")
					s = Replace(s,b(i),src)
				End If
			Next
		End If
		SaveImgTo_ = s
	End Function
	Public Sub AjaxAgent()
		Easp.NoCache()
		Dim u, qs, qskey, qf, qfkey, m
		u = Easp.Get("easpurl")
		If Easp.IsN(u) Then Easp.WE "error:Invalid URL"
		If Instr(u,"?")>0 Then
			qs = "&" & Easp.CRight(u,"?")
			u = Easp.CLeft(u,"?")
		End If
		If Request.QueryString()<>"" Then
			For Each qskey In Request.QueryString
				If qskey<>"easpurl" Then qs = qs & "&" & qskey & "=" & Request.QueryString(qskey)
			Next
		End If
		u = u & Easp.IfThen(Easp.Has(qs),"?" & Mid(qs,2))
		m = Request.ServerVariables("REQUEST_METHOD")
		If m = "POST" Then
			If Request.Form()<>"" Then
				For Each qfkey In Request.Form
					qf = qf & "&" & qfkey & "=" & Request.Form(qfkey)
				Next
				Data = Mid(qf,2)
			End If
			Easp.WE Post(u)
		Else
			Easp.WE [Get](u)
		End If
	End Sub
	Function TransPath(ByVal u, ByVal p)
		If Left(p,7)="http://" Or Left(p,8)="https://" Then TransPath = p : Exit Function
		Dim tmp,ser, fol
		tmp = Easp.CLeft(u,"?")
		If Left(u,7)<>"http://" And Left(u,8)<>"https://" Then
			ser = ""
		Else
			ser = Easp.RegReplace(tmp,"^(https?://[a-zA-Z0-9-.]+)/(.+)$","$1")
		End If
		fol = Mid(tmp,1,InstrRev(tmp,"/"))
		TransPath = Easp.IIF(Left(p,1) = "/", ser, fol) & p
	End Function
	Private Function Serialize__(ByVal a)
		Dim tmp, i, n, v : tmp = ""
		If Easp.IsN(a) Then Exit Function
		If isArray(a) Then
			For i = 0 To Ubound(a)
				n = Easp.CLeft(a(i),":")
				v = Easp.CRight(a(i),":")
				tmp = tmp & "&" & n & "=" & Server.URLEncode(v)
			Next
			If Len(tmp)>1 Then tmp = Mid(tmp,2)
			Serialize__ = tmp
		Else
			Serialize__ = a
		End If
	End Function
	Private Function Bytes2Bstr__(ByVal s, ByVal char) 
		dim oStrm
		set oStrm = Server.CreateObject("Adodb.Stream")
		With oStrm
			.Type = 1
			.Mode =3
			.Open
			.Write s
			.Position = 0
			.Type = 2
			.Charset = CharSet
			Bytes2Bstr__ = .ReadText
			.Close
		End With
		set oStrm = nothing
	End Function
End Class
%>