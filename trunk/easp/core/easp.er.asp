<%
Class EasyAsp_Er
	Private i_errNum
	Private s_errStr
	Private o_err
	Private Sub Class_Initialize
		Set o_err = Server.CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate
		Set o_err = Nothing
	End Sub
	
	Public Default Property Let [Set](ByVal n, ByVal s)
		
'		Dim n
'		n = Ubound(a_err,1)
'		ReDim Preserve a_err(1,n+1)
'		a_err(0,n+1) = 
	End Property
	
	Public Property Let Tmp(ByVal m)
		
	End Property
End Class
%>