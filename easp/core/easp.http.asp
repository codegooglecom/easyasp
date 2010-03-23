<%
'######################################################################
'## easp.http.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp XMLHTTP Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/03/23 23:24:30
'## Description :   Request XMLHttp Data in EasyASP
'##
'######################################################################
Class EasyAsp_Http
	Public CharSet 
	'Private
	
	Private Sub Class_Initialize
		CharSet = Easp.CharSet
		Easp.Error(46) = "远程服务器没有响应"
		Easp.Error(47) = ""
	End Sub
	
	Private Sub Class_Terminate
		
	End Sub
	
	Function getHTML(ByVal url, ByVal char)
		Dim oSend, tmp
		If Easp.isInstall("MSXML2.serverXMLHTTP") Then
			Set oSend = Server.CreateObject("MSXML2.serverXMLHTTP")
		ElseIf Easp.isInstall("MSXML2.XMLHTTP") Then
			Set oSend = Server.CreateObject("MSXML2.XMLHTTP")
		ElseIf Easp.isInstall("Microsoft.XMLHTTP") Then
			Set oSend = Server.CreateObject("Microsoft.XMLHTTP")
		End If
		oSend.Open "GET",url,False
		oSend.send
		If oSend.readyState <> 4 Then
			getHTML = "远程服务器没有响应"
			Set oSend = Nothing
			Exit Function
		End If
		If char = "" Then char = "GB2312"
		tmp = Bytes2Bstr(oSend.responseBody,char)
		Set oSend = Nothing
		getHTML = tmp
	End Function

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