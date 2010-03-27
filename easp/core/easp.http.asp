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
	
	Public Default Function GetData(ByVal url, ByVal method, ByVal async, ByVal user, ByVal pass, ByVal data)
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
		If Easp.IsN(method) Then method = "GET"
		'异步
		If Easp.IsN(async) Then async = False
		If Easp.Has(user) And Easp.Has(pass) Then
			'如果有用户名和密码
			o.open method, url, async, user, pass
		Else
			'匿名
			o.open method, url, async
		End If
		If method = "POST" Then
			o.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		End If
		If Easp.Has(data) Then
			'有发送的数据
			o.send data
		Else
			'没有数据发送
			o.send
		End If
		If o.readyState <> 4 Then
			Easp.Error.Raise 46
			Set o = Nothing
			Exit Function
		End If
		GetData = Bytes2Bstr(o.responseBody, CharSet)
		Set o = Nothing
	End Function

	'编码转换
	Function Bytes2Bstr(ByVal s, ByVal char) 
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