<%
Class EasyAsp_Error
	Private b_debug
	Private i_errNum
	Private s_errStr
	Private o_err
	Private Sub Class_Initialize
		i_errNum = 0
		b_debug = True
		Set o_err = Server.CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate
		Set o_err = Nothing
	End Sub
	'取已定义的错误信息
	Public Default Property Get E(ByVal n)
		If o_err.Exists(n) Then
			E = o_err(n)
		Else
			E = "Unknown Easp Error"
		End If
	End Property
	'自定义错误代码
	Public Property Let E(ByVal n, ByVal s)
		If Easp.Has(n) And Easp.Has(s) Then
			If n <> "0" Then
				o_err(n) = s
			End If
		End If
	End Property
	'取最近发生错误的代码
	Public Property Get LastError
		LastError = i_errNum
	End Property

	Public Sub Raise(ByVal n)
		If Easp.isN(n) Then Exit Sub
		i_errNum = n
		If b_debug Then
			
		End If
	End Sub

	Public Function ThrowError()
		Dim o : Set o = Server.GetLastError
		
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