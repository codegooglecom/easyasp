<%
'######################################################################
'## easp.http.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp XMLHTTP Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/03/23 23:24:30
'## Description :   Request XMLHttp Data in EasyASP
'## http://msdn.microsoft.com/en-us/library/ms535874(VS.85).aspx
'######################################################################
Class EasyAsp_Http
	Public Url, Method, CharSet, Async, User, Password, Html, Headers
	Public ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
	Private s_data, s_url, s_ohtml
	
	Private Sub Class_Initialize
		'编码默认为空，将自动获取编码
		CharSet = ""
		'异步模式关闭
		Async = False
		User = ""
		Password = ""
		s_data = ""
		s_url = ""
		Html = ""
		Headers = ""
		'服务器解析超时
		ResolveTimeout = 20000
		'服务器连接超时
		ConnectTimeout = 20000
		'发送数据超时
		SendTimeout = 300000
		'接受数据超时
		ReceiveTimeout = 60000
		Easp.Error(46) = "远程服务器没有响应"
		Easp.Error(47) = "服务器不支持XMLHTTP组件"
		Easp.Error(48) = "要获取的页面地址不能为空"

	End Sub
	
	Private Sub Class_Terminate
		
	End Sub

	'建新实例
	Public Function [New]()
		Set [New] = New EasyAsp_Http
	End Function
	
	'提交的数据
	Public Property Let Data(ByVal s)
		s_data = s
	End Property
	
	Public Function [Open]
		[Open] = GetData(Url, Method, Async, s_data, User, Password)
	End Function
	
	'Get取远程页
	Public Function [Get](ByVal uri)
		[Get] = GetData(uri, "GET", Async, s_data, User, Password)
	End Function
	
	'Get取远程页
	Public Function Post(ByVal uri)
		Post = GetData(uri, "POST", Async, s_data, User, Password)
	End Function
	
	'XMLHTTP原始方法
	Public Function GetData(ByVal uri, ByVal m, ByVal async, ByVal data, ByVal u, ByVal p)
		Dim o,chru
		'建立XMLHttp对象
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
		'设置超时时间
		o.SetTimeOuts ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
		'抓取地址
		If Easp.IsN(uri) Then Easp.Error.Raise 48 : Exit Function
		s_url = uri
		'方法：POST或GET
		m = Easp.IIF(Easp.Has(m),UCase(m),"GET")
		'异步
		If Easp.IsN(async) Then async = False
		'构造Get传数据的URL
		If m = "GET" And Easp.Has(data) Then uri = uri & Easp.IIF(Instr(uri,"?")>0, "&", "?") & Serialize__(data)
		'打开远程页
		If Easp.Has(u) Then
			'如果有用户名和密码
			o.open m, uri, async, u, p
		Else
			'匿名
			o.open m, uri, async
		End If
		If m = "POST" Then
			o.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			'有发送的数据
			o.send Serialize__(data)
		Else
			o.send
		End If
		'检测返回数据
		If o.readyState <> 4 Then
			GetData = "error:server is down"
			Set o = Nothing
			Easp.Error.Raise 46
			Exit Function
		ElseIf o.Status = 200 Then
			Headers = o.getAllResponseHeaders()
			If Easp.IsN(CharSet) Then
				If Easp.Test(Headers,"charset=([\w-]+)") Then
					CharSet = Easp.RegReplace(Headers,"([\s\S]+)charset=([\w-]+)([\s\S]+)","$2")
				ElseIf Easp.Test(o.responseText,"<meta\s+http-equiv\s*=[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>") Then
					CharSet = Easp.RegReplace(o.responseText,"([\s\S]+)<meta\s+http-equiv\s*=[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>([\s\S]+)","$2")
				Else
					CharSet = Easp.CharSet
				End If
			End If
			GetData = Bytes2Bstr__(o.responseBody, CharSet)
		Else
			GetData = "error:" & o.Status & " " & o.StatusText
		End If
		Set o = Nothing
		s_ohtml = GetData
		Html = s_ohtml
	End Function
	
	'按正则查找符合的第一个字符串
	Public Function Find(ByVal rule)
		Find = Find_(s_ohtml, rule)
	End Function
	Public Function Find_(ByVal s, ByVal rule)
		If Easp.Test(s,rule) Then Find_ = Easp.RegReplace(s,"([\s\S]*)("&rule&")([\s\S]*)","$2")
	End Function
	
	'按正则查找符合的第一个字符串，可按正则编组选择其中的一部分
	Public Function [Select](ByVal rule, ByVal part)
		[Select] = Select_(s_ohtml, rule, part)
	End Function
	Public Function Select_(ByVal s, ByVal rule, ByVal part)
		If Easp.Test(s,rule) Then
			'$0匹配字符串本身
			part = Replace(part,"$0",Find_(s,rule))
			'按正则编组分别替换
			Select_ = Easp.RegReplace(s,"(?:[\s\S]*)(?:"&rule&")(?:[\s\S]*)",part)
		End If
	End Function
	
	'按正则查找符合的字符串组，返回数组
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
	
	'按标签查找字符串(SubStr)
	'tagStart - 要截取的部分的开头
	'tagEnd   - 要截取的部分的结尾
	'tagSelf  - 结果是否包括tagStart和tagEnd
	'           (0或空:不包括,1:包括,2:只包括tagStart,3:只包括tagEnd)
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
	
	'url参数化
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
	
	'编码转换
	Private Function Bytes2Bstr__(ByVal s, ByVal char) 
		dim oStrm
		set oStrm = Server.CreateObject("Adodb.Stream")
		oStrm.Type = 1
		oStrm.Mode =3
		oStrm.Open
		oStrm.Write s
		oStrm.Position = 0
		oStrm.Type = 2
		oStrm.Charset = CharSet
		Bytes2Bstr__ = oStrm.ReadText
		oStrm.Close
		set oStrm = nothing
	End Function

End Class
%>