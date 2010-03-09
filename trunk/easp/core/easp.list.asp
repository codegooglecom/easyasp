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
'Filter 函数可利用

Class EasyAsp_List
	Public Size
	Private o_list
	Private a_list
	Private i_count
	
	Private Sub Class_Initialize
		'Set o_list   = Server.CreateObject("Scripting.Dictionary")
		Size = 0
		Easp.Error(41) = "下标越界"
	End Sub
	
	Private Sub Class_Terminate
		'Set o_list = Nothing
	End Sub
	
	'建新实例
	Public Function [New]()
		Set [New] = New EasyAsp_List
	End Function
	
	'取某一项值
	Public Default Property Get At(ByVal n)
		If isNumeric(n) Then
			If n < Size Then
				At = a_list(n)
			Else
				Easp.Error.Msg = "(当前下标 "&n&" 超过了最大下标 "&Size-1&" )"
				Easp.Error.Raise 41
			End If
		End If
	End Property
	
	'设置某一项值
	Public Property Let At(ByVal n, ByVal v)
		If Size = 0 Then ReDim a_list(0)
		If isNumeric(n) Then
			If n >= Size Then ReDim Preserve a_list(n)
			a_list(n) = v
			Size = n + 1
		End If
	End Property
	
	'源数据
	Public Property Let Data(ByVal a)
		a_list = a
		Size = Ubound(a) + 1
	End Property
	'取出为普通数组
	Public Property Get Data
		Data = a_list
	End Property
	
	'长度
	Public Property Get Length
		Length = Size
	End Property
	
	'数组有效长度（非空值）
	Public Property Get Count
		Dim i,j : j = 0
		For i = 0 To Size-1
			If Easp.Has(At(i)) Then j = j + 1
		Next
		Count = j
	End Property
	
	'按下标取List的一部分元素返回一个新的List对象
	Public Function [Get](ByVal s)
		Dim a,i,j,k,x,y,arr(),e
		a = Split(s, ",")
		ReDim arr(Ubound(a))
		k = 0
		For i = 0 To Ubound(a)
			ReDim Preserve arr(k)
			If Instr(a(i),"-")>0 Then
				x = Int(Easp.CLeft(a(i),"-"))
				y = Int(Easp.CRight(a(i),"-"))
				For j = x To y
					ReDim Preserve arr(k)
					arr(k) = At(j)
					k = k + 1
				Next
			Else
				arr(k) = At(Int(Trim(a(i))))
				k = k + 1
			End If
		Next
		Set [Get] = Me.New
		[Get].Data = arr
	End Function
	
	'排序
	Public Sub Sort
		SortArray a_list, Lbound(a_list), Ubound(a_list)
	End Sub
	Private Sub SortArray(ByRef arr, ByRef low, ByRef high)
		If Not IsArray(arr) Then Exit Sub
		If Easp.IsN(arr) Then Exit Sub
		Dim l, h, m, v, x
		l = low : h = high
		m = (low + high) \ 2 : v = arr(m)
		Do While (l <= h)
			Do While (arr(l) < v And l < high)
				l = l + 1
			Loop
			Do While (v < arr(h) And h > low)
				h = h - 1
			Loop
			If l <= h Then
				x = arr(l) : arr(l) = arr(h) : arr(h) = x   
				l = l + 1 : h = h - 1         
			End If
		Loop
		If (low < h) Then SortArray arr, low, h
		If (l < high) Then SortArray arr,l, high
	End Sub
	
	'删除空元素并返回新的List对象
	Public Function Compact
		Dim arr(), i, j : j = 0
		For i = 0 To Size - 1
			If Easp.Has(At(i)) Then
				Easp.WN "j:" & j & ", At("&i&"):" & At(i)
				ReDim Preserve arr(j)
				arr(j) = At(i)
				j = j + 1
			End If
		Next
		Me.Data = arr
		'Set Compact = Me.New
		'Compact.Data = arr
	End Function
	
	'联连字符串
	Public Function J(ByVal s)
		J = Join(a_list, s)
	End Function
	
	'转换成字符串（,号隔开）
	Public Function ToString()
		ToString = J(",")
	End Function
End Class
%>