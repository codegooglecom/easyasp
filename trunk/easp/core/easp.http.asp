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
	Public CharSet 
	'Private
	
	Private Sub Class_Initialize
		CharSet = Easp.CharSet
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
	
	'Get取远程页
	Public Function [Get](ByVal url)
		[Get] = GetData(url, "GET", False, "", "", "")
	End Function
	
	'Get取远程页
	Public Function Post(ByVal url)
		Post = GetData(url, "POST", False, "", "", "")
	End Function
	
	'XMLHTTP原始方法
	Public Function GetData(ByVal url, ByVal method, ByVal async, ByVal data, ByVal user, ByVal pass)
		Dim o
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
		'抓取地址
		If Easp.IsN(url) Then Easp.Error.Raise 48 : Exit Function
		'方法：POST或GET
		method = Easp.IIF(Easp.Has(method),UCase(method),"GET")
		'异步
		If Easp.IsN(async) Then async = False
		'构造获取条件(Post还是Get)
		If method = "POST" Then
			o.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		Else
			If Easp.Has(data) Then url = url & Easp.IIF(Instr(url,"?")>0, "&", "?") & Serialize(data)
		End If
		'打开远程页
		If Easp.Has(user) And Easp.Has(pass) Then
			'如果有用户名和密码
			o.open method, url, async, user, pass
		Else
			'匿名
			o.open method, url, async
		End If
		If method = "POST" Then
			'有发送的数据
			o.send data
		Else
			o.send
		End If
		'检测返回数据
		If o.readyState <> 4 Then
			GetData = "error:server is down"
			Easp.Error.Raise 46
			Set o = Nothing
			Exit Function
		ElseIf o.Status = 200 Then
			GetData = Bytes2Bstr(o.responseBody, CharSet)
		Else
			Select Case o.Status
				Case 400 GetData = "error:400 bad request"
				Case 403 GetData = "error:403 forbidden"
				Case 404 GetData = "error:404 not found"
				Case 500 GetData = "error:500 server error"
			End Select
		End If
		Set o = Nothing
	End Function
	
	'url参数化
	Private Function Serialize(ByVal a)
		Dim tmp, i, n, v : tmp = ""
		If Easp.IsN(a) Then Exit Function
		If isArray(a) Then
			For i = 0 To Ubound(a)
				n = Easp.CLeft(a(i),":")
				v = Easp.CRight(a(i),":")
				tmp = tmp & "&" & n & "=" & Server.URLEncode(v)
			Next
			If Len(tmp)>1 Then tmp = Mid(tmp,2)
			Serialize = tmp
		Else
			Serialize = a
		End If
	End Function
	
	'编码转换
	Private Function Bytes2Bstr(ByVal s, ByVal char) 
		dim oStrm
		set oStrm = Server.CreateObject("Adodb.Stream")
		oStrm.Type = 1
		oStrm.Mode =3
		oStrm.Open
		oStrm.Write s
		oStrm.Position = 0
		oStrm.Type = 2
		oStrm.Charset = CharSet
		Bytes2Bstr = oStrm.ReadText
		oStrm.Close
		set oStrm = nothing
	End Function

End Class
%>