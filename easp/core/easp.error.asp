<%
'######################################################################
'## easp.error.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp Exception Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2009/12/15 15:48
'## Description :   EasyAsp�쳣����
'##
'######################################################################
Class EasyAsp_Error
	Private b_debug, b_redirect
	Private i_errNum, i_delay
	Private s_errStr, s_title, s_url, s_css, s_msg
	Private o_err
	Private Sub Class_Initialize
		i_errNum    = ""
		i_delay     = 3000
		s_title     = "����������"
		b_debug     = Easp.Debug
		b_redirect  = True
		s_url       = "javascript:history.go(-1)"
		Set o_err   = Server.CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate
		Set o_err = Nothing
	End Sub
	'�Ƿ�������״̬�������󷵻ؿ����ߴ�����Ϣ��
	Public Property Get [Debug]
		[Debug] = b_debug
	End Property
	Public Property Let [Debug](ByVal b)
		b_debug = b
	End Property
	'ȡ�Ѷ���Ĵ�����Ϣ
	Public Default Property Get E(ByVal n)
		If o_err.Exists(n) Then
			E = o_err(n)
		Else
			E = "δ֪����"
		End If
	End Property
	'�Զ���������
	Public Property Let E(ByVal n, ByVal s)
		If Easp.Has(n) And Easp.Has(s) Then
			If n > "" Then
				o_err(n) = s
			End If
		End If
	End Property
	'ȡ�����������Ĵ���
	Public Property Get LastError
		LastError = i_errNum
	End Property
	'������Ϣ����
	Public Property Get Title
		Title = s_title
	End Property
	Public Property Let Title(ByVal s)
		s_title = s
	End Property
	'�Զ���Ĵ�����Ϣ
	Public Property Get Msg
		Msg = s_msg
	End Property
	Public Property Let Msg(ByVal s)
		s_msg = s
	End Property
	'�Ƿ��Զ�ת��
	Public Property Get [Redirect]
		[Redirect] = b_redirect
	End Property
	Public Property Let [Redirect](ByVal b)
		b_redirect = b
	End Property
	'�Զ�����תҳ
	Public Property Get Url
		Url = s_url
	End Property
	Public Property Let Url(ByVal s)
		s_url = s
	End Property
	'�Զ���תʱ�䣨�룩
	Public Property Get Delay
		Delay = i_delay / 1000
	End Property
	Public Property Let Delay(ByVal i)
		i_delay = i * 1000
	End Property
	'�Զ�����ʽ����
	Public Property Get ClassName
		ClassName = s_css
	End Property
	Public Property Let ClassName(ByVal s)
		s_css = s
	End Property
	'���ɴ���
	Public Sub Raise(ByVal n)
		If Easp.isN(n) Then Exit Sub
		i_errNum = n
		If b_debug Then
			Easp.WE ShowMsg(o_err(n) & s_msg, 1)
		End If
		s_msg = ""
	End Sub
	'�׳�������Ϣ
	Public Sub Throw(ByVal msg)
		If Left(msg,1) = ":" Then
			If o_err.Exists(Mid(msg,2)) Then msg = o_err(Mid(msg,2))
		End If
		Easp.W ShowMsg(msg,0)
	End Sub
	'��ʾ�Ѷ�������д�����뼰��Ϣ
	Public Sub Defined()
		Dim key
		If Easp.Has(o_err) Then
			For Each key In o_err
				Easp.Wn key & " : " & o_err(key)
			Next
		End If
	End Sub
	'��ʾ������Ϣ��
	Private Function ShowMsg(ByVal msg, ByVal t)
		Dim s,x
		s = "<fieldset id=""easpError""" & Easp.IfThen(Easp.Has(s_css)," class=""" & s_css & """") & ">" & vbCrLf
		s = s & "	<legend>" & s_title & "</legend>" & vbCrLf
		s = s & "	<p class=""msg"">" & msg & "</p>" & vbCrLf
		x = Easp.IIF(s_url = "javascript:history.go(-1)", "����", "����")
		If t = 1 Then
			If Err.Number<>0 Then
				s = s & "	<ul class=""dev"">" & vbCrLf
				s = s & "		<li class=""info"">������Ϣ��Կ����ߣ�</li>" & vbCrLf
				s = s & "		<li>������룺0x" & Hex(Err.Number) & "</li>" & vbCrLf
				s = s & "		<li>����������" & Err.Description & "</li>" & vbCrLf
				s = s & "		<li>������Դ��" & Err.Source & "</li>" & vbCrLf
				s = s & "	</ul>" & vbCrLf
			End If
		Else
			If b_redirect Then
				s = s & "	<p class=""back"">ҳ�潫��" & i_delay/1000 & "���Ӻ���ת����������û��������ת��<a href=""" & s_url & """>�����˴�" & x & "</a></p>" & vbCrLf
				s_url = Easp.IIF(Left(s_url,11) = "javascript:", Mid(s_url,12), "location.href='" & s_url & "';")
				s = s & Easp.JsCode("setTimeout(function(){" & s_url & "}," & i_delay & ");")
			Else
				s = s & "	<p class=""back""><a href=""" & s_url & """>�����˴�" & x & "</a></p>" & vbCrLf
			End If
		End If
		s = s & "</fieldset>" & vbCrLf
		ShowMsg = s
	End Function
End Class
%>