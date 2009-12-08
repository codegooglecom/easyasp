<%
Class EasyAsp_tpl
	Private s_html, s_m, s_ms, s_me, o_loop, o_data
	
	Private Sub Class_Initialize
		s_html = ""
		s_m = "{{*}}"
		getMaskSE s_m
		Set o_loop = CreateObject("Scripting.Dictionary")
		Set o_data = CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate
		Set o_loop = Nothing
		Set o_data = Nothing
	End Sub
	
	Public Property Let [File](ByVal f)
		s_html = Easp.getInclude(f)
	End Property
	
	Public Property Let TagMask(ByVal m)
		s_m = m
		getMaskSE s_m
	End Property
	
	Private Sub getMaskSE(ByVal m)
		s_ms = Easp.CLeft(m,"*")
		s_me = Easp.CRight(m,"*")
	End Sub
	
	Function getLoopBlock(ByVal n)
		Dim reg,rule,m
		rule = "<!--[\s]*?Easp:Loop:"&n&"[\s]*?-->([\s\S]+)<!--[\s]*?Easp:End[ +]Loop:"&n&"[\s]*?-->"
		Set reg = Easp_Match(s_html,rule)
		For Each m In reg
			getLoopBlock = Array(m,m.SubMatches(0))
		Next
		Set reg = Nothing
	End Function
	
	Public Sub Load(ByVal f)
		s_html = Easp.getInclude(f)
	End Sub
	
	Public Default Function Shift(ByVal t, ByVal s)
		Dim b,f,m
		If Instr(t,".")>0 Then
			f = Easp.CLeft(t,".")
			m = Easp.CRight(t,".")
			If Not o_loop.Exists(f) Then
				b = getLoopBlock(f)
				o_loop.Add f&"__b", b(0)
				o_loop.Add f&"__s", b(1)
			End If
		Else
			o_data.Add t, cStr(s)
		End If
	End Function
	
	Public Function MakeTag(ByVal t, ByVal f)
		Dim s,e,a,i
		Select Case Lcase(t)
			Case "css"
				s = "<link href="""
				e = """ rel=""stylesheet"" type=""text/css"" />"
			Case "js"
				s = "<script type=""text/javascript"" src="""
				e = """></scr"&"ipt>"
		End Select
		a = Split(f,"|")
		For i = 0 To Ubound(a)
			a(i) = s & Trim(a(i)) & e
		Next
		MakeTag = Join(a,vbCrLf)
	End Function
	
	Public Sub Show()
		Dim k
		If o_data.Count > 0 Then
			For Each k In o_data
				'Easp.WN k & " - " & Easp.HtmlEncode(o_data(k))
				s_html = Replace(s_html,s_ms&k&s_me,o_data(k))
			Next
		End If
		Response.Write(s_html)
	End Sub
End Class
%>