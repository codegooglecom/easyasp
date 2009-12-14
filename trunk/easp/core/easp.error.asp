<%
Class EasyAsp_Error
	Private b_debug
	Private i_errNum
	Private s_errStr, s_title
	Private o_err
	Private Sub Class_Initialize
		i_errNum = 0
		b_debug = True
		s_title = "������"
		Set o_err = Server.CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate
		Set o_err = Nothing
	End Sub
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
			If n <> "0" Then
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

	Public Sub Raise(ByVal n)
		If Easp.isN(n) Then Exit Sub
		Dim s : i_errNum = n
		If b_debug Then
			Easp.WE ShowMsg(Me.E(n), 1, "", 0)
		End If
	End Sub

	Public Sub Throw(ByVal msg)
		'Easp.W ShowMsg()
	End Sub
	Private Function ShowMsg(ByVal msg, ByVal t, ByVal url, ByVal relay)
		Dim s, isBack
		If Easp.isN(title) Then title = Me.Title
		If Easp.Has(url) Then
			isBack = True
			If isNumeric(relay) Then
				relay = relay / 1000
			Else
				relay = 3000
			End If
		End If
		s = "<fieldset id=""easpError"" ><legend>" & title & "</legend>" & vbCrLf
		s = s & "	<p class=""msg"">" & msg & "</p>" & vbCrLf
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
			If isBack Then
				s = s & "	<p class=""back"">ҳ�潫��" & relay*1000 & "���Ӻ���ת����������û��������ת��<a href=""" & Easp.IIF(url=":back","javascript:history.go(-1)",url) & """>�����˴�</a>��</p>" & vbCrLf
			End If
		End If
		s = s & "</fieldset>" & vbCrLf
		ShowMsg = s
	End Function
	Public Sub Trace()
		Dim key
		If Easp.Has(o_err) Then
			For Each key In o_err
				Easp.Wn key & ":" & o_err(key)
			Next
		End If
	End Sub
End Class
%>