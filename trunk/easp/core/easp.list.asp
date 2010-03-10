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
		a_list = Array()
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
		'If Size = 0 Then ReDim a_list(0)
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
	
	'最大下标
	Public Property Get [End]
		[End] = Size - 1
	End Property
	
	'数组有效长度（非空值）
	Public Property Get Count
		Dim i,j : j = 0
		For i = 0 To Size-1
			If Easp.Has(At(i)) Then j = j + 1
		Next
		Count = j
	End Property
	
	'获取第一个元素
	Public Property Get First
		First = At(0)
	End Property
	
	'获取最后一个元素
	Public Property Get Last
		Last = At([End])
	End Property
	
	'获取最大元素
	Public Property Get Max
		Max = MaxMin(0)
	End Property
	
	'获取最小元素
	Public Property Get Min
		Min = MaxMin(1)
	End Property
	
	Private Function MaxMin(ByVal t)
		Dim tmp : Set tmp = Me.Clone
		tmp.Compact
		tmp.Sort
		MaxMin = Easp.IIF(t=0,tmp.Last,tmp.First)
		Set tmp = Nothing
	End Function
	
	'添加一个元素到开头
	Public Sub UnShift(ByVal v)
		
	End Sub
	
	'删除第一个元素
	Public Sub Shift
		
	End Sub
	
	'添加一个元素到结尾
	Public Sub Push(ByVal v)
		At([End]+1) = v
	End Sub
	
	'删除最后一个元素
	Public Sub Pop
		ReDim Preserve a_list([End]-1)
		Size = Size - 1
	End Sub
	
	'插入一个或多个元素
	'Insert 2, "data"
	'Insert "3,12", Array("data3", "data12")
	'Insert "2-4", Array("data2","data3", "data4")
	Public Sub Insert(ByVal i, ByVal v)
		
	End Sub
	
	'删除一个或多个元素
	Public Sub [Delete](ByVal i)
		
	End Sub

	'移除重复元素只保留一个
	Public Sub Uniq
		
	End Sub

	'搜索包含指定字符串的元素
	Public Sub Search(ByVal s)
		
	End Sub

	'随机排序
	Public Sub Rand
		
	End Sub
	
	'反向排列数组
	Public Sub Reverse
		Dim arr(),i,j : j = 0
		ReDim arr([End])
		For i = [End] To 0 Step -1
			arr(j) = At(i)
			j = j + 1
		Next
		Me.Data = arr
	End Sub
	
	'删除空元素
	Public Sub Compact
		Dim arr(), i, j : j = 0
		For i = 0 To [End]
			If Easp.Has(At(i)) Then
				ReDim Preserve arr(j)
				arr(j) = At(i)
				j = j + 1
			End If
		Next
		Me.Data = arr
	End Sub
	
	'清空
	Public Sub Clear
		a_list = Array()
		Size = 0
	End Sub
	
	'排序
	Public Sub Sort
		Me.Data = SortArray(a_list, 0, [End])
	End Sub
	'快速排序法
	Private Function SortArray(ByRef arr, ByRef low, ByRef high)
		If Not IsArray(arr) Then Exit Function
		If Easp.IsN(arr) Then Exit Function
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
		If (low < h) Then arr = SortArray(arr, low, h)
		If (l < high) Then arr = SortArray(arr,l, high)
		SortArray = arr
	End Function
	
	'按下标取List的一部分元素
	Public Sub Slice(ByVal s)
		Me.Data = GetPart(s)
	End Sub
	'按下标取List的一部分元素返回一个新的List对象
	Public Function [Get](ByVal s)
		Set [Get] = Me.New
		[Get].Data = GetPart(s)
	End Function
	Private Function GetPart(ByVal s)
		Dim a,i,j,k,x,y,arr
		a = Split(s, ",")
		arr = Array() : k = 0
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
		GetPart = arr
	End Function
	
	'复制List对象
	Public Function Clone
		Set Clone = Me.New
		Clone.Data = a_list
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