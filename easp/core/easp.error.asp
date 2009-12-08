<%
Class EasyAsp_Error
	Private i_errNum
	Private s_errStr
	Private o_err
	Private Sub Class_Initialize
		Set o_err = Server.CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate
		Set o_err = Nothing
	End Sub

	Public Default Property Get E(ByVal n)
		If o_err.Exists(n) Then
			E = o_err(n)
		Else
			E = "ЮДжЊДэЮѓ"
		End If
	End Property
	Public Property Let E(ByVal n, ByVal s)
		If Easp.Has(n) And Easp.Has(s) Then
			o_err(n) = s
		End If
	End Property
	
	Public Sub Setted()
		Dim key
		If Easp.Has(o_err) Then
			For Each key In o_err
				Easp.Wn key & ":" & o_err(key)
			Next
		End If
	End Sub
End Class
%>