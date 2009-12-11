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
		If Not Easp_isN(n) And Not Easp_isN(s) Then
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
		If Easp_isN(n) Then Exit Sub
		Dim s : i_errNum = n
		If b_debug Then
			s = "<fieldset id=""easpErrorField"" ><legend>出错啦</legend><ul><li>" & n & Me.E(n) & "</li>"
			If Err.Number<>0 Then
				s = s & "<li>错误描述：(" & Hex(Err.Number) & ")" & Err.Description & "</li>"
			End If
			s = s & "</ul></fieldset>"
			Response.Write s
		End If
	End Sub

	Public Function ThrowError()
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