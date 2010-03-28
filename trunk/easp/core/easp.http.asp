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
	Private o_xmlhttp
	
	Private Sub Class_Initialize
		CharSet = Easp.CharSet
		Easp.Error(46) = "远程服务器没有响应"
		Easp.Error(47) = "服务器不支持XMLHTTP组件"
		Easp.Error(48) = "要获取的页面地址不能为空"
		'建立XMLHttp对象
		If Easp.isInstall("MSXML2.serverXMLHTTP") Then
			Set o_xmlhttp = Server.CreateObject("MSXML2.serverXMLHTTP")
		ElseIf Easp.isInstall("MSXML2.XMLHTTP") Then
			Set o_xmlhttp = Server.CreateObject("MSXML2.XMLHTTP")
		ElseIf Easp.isInstall("Microsoft.XMLHTTP") Then
			Set o_xmlhttp = Server.CreateObject("Microsoft.XMLHTTP")
		Else
			Easp.Error.Raise 47
		End If
	End Sub
	
	Private Sub Class_Terminate
		If isObject(o_xmlhttp) Then Set o_xmlhttp = Nothing
	End Sub

	'建新实例
	Public Function [New]()
		Set [New] = New EasyAsp_Http
	End Function
	
	Public Default Function GetData(ByVal url, ByVal method, ByVal async, ByVal data, ByVal user, ByVal pass)
		'抓取地址
		If Easp.IsN(url) Then Easp.Error.Raise 48 : Exit Function
		'方法：POST或GET
		If Easp.IsN(method) Then method = "GET"
		'异步
		If Easp.IsN(async) Then async = False
		If Easp.Has(user) And Easp.Has(pass) Then
			'如果有用户名和密码
			o_xmlhttp.open method, url, async, user, pass
		Else
			'匿名
			o_xmlhttp.open method, url, async
		End If
		If method = "POST" Then
			o_xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		End If
		If Easp.Has(data) Then
		'有发送的数据
'			If UCase(method) = "GET" Then
'				If Instr(url,"?")>0 Then
'			ElseIf UCase(method) = "POST" Then
'				
'			End If
			o_xmlhttp.send data
		Else
		'没有数据发送
			o_xmlhttp.send
		End If
		If o_xmlhttp.readyState <> 4 Then
			Easp.Error.Raise 46
			Exit Function
		End If
		GetData = Bytes2Bstr(o_xmlhttp.responseBody, CharSet)
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