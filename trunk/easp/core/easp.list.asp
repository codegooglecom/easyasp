<%
'######################################################################
'## easp.list.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp List(Array) Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/03/09 16:08:30
'## Description :   A super Array class in EasyAsp
'##
'######################################################################
Class EasyAsp_List
	Public Size
	Private o_list
	Private a_list
	Private i_count
	Private Sub Class_Initialize
		Set o_list   = Server.CreateObject("Scripting.Dictionary")
		Size = 0
	End Sub
	Private Sub Class_Terminate
		Set o_list = Nothing
	End Sub
	'建新实例
	Public Function [New]
		Set [New] = New EasyAsp_List
	End Function
	'取某一项值
	Public Default Property Get At(ByVal n)
		If isNumeric(n) Then
			If n < Size Then
				At = a_list(n)
			Else
				At = "下标越界"
			End If
		End If
	End Property
	'设置某一项值
	Public Property Let At(ByVal n, ByVal v)
		If isNumeric(n) Then
			If n >= Size Then Redim Preserve a_list(n)
			a_list(n) = v
			Size = n + 1
		End If
	End Property
	'数组的源数据
	Public Property Let Source(ByVal a)
		a_list = a
		Size = Ubound(a) + 1
	End Property
	Public Property Get Length
		Length = Size
	End Property
	'排序
	Public Sub Sort
		Dim a
		If Size = 0 Then Exit Sub
		
	End Sub
	'取一部分项目
	Public Function [Get](ByVal s)
		Dim a,i
		a = Split(s, ",")
		For i = 0 To Ubound(a)
			a(i) = a_list(Trim(a(i)))
		Next
		[Get] = a
	End Function
	'转换为普通数组
	Public Function toArray
		toArray = a_list
	End Function
	'联连字符串
	Public Function J(ByVal s)
		J = Join(a_list, s)
	End Function
	Public Function ToString()
		ToString = J(",")
	End Function
End Class
%>