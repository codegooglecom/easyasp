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
	Private o_hash, o_map
	Private a_list
	Private i_count, i_comp
	
	Private Sub Class_Initialize
		Set o_hash = Server.CreateObject("Scripting.Dictionary")
		Set o_map  = Server.CreateObject("Scripting.Dictionary")
		a_list = Array()
		Size = 0
		Easp.Error(41) = "下标越界"
		Easp.Error(42) = "下标不能为空"
		Easp.Error(43) = "下标只能是数字、字母、下划线(_)、点(.)和斜杠(/)组成"
		i_comp = 1
	End Sub
	
	Private Sub Class_Terminate
		Set o_map  = Nothing
		Set o_hash = Nothing
	End Sub
	
	'建新实例
	Public Function [New]()
		Set [New] = New EasyAsp_List
	End Function
	
	'是否忽略大小写
	Public Property Let IgnoreCase(ByVal b)
		i_comp = Easp.IIF(b, 1, 0)
	End Property
	Public Property Get IgnoreCase
		IgnoreCase = (i_comp = 1)
	End Property
	
	'设置某一项值
	Public Property Let At(ByVal n, ByVal v)
		If Easp.IsN(n) Then Easp.Error.Raise 42 : Exit Property
		If Easp.Test(n,"^\d+$") Then
			'如果是数字就直接添加到数组下标
			If n > [End] Then
				ReDim Preserve a_list(n)
				Size = n + 1
			End If
			a_list(n) = v
		ElseIf Easp.Test(n,"^[\w\./]+$") Then
		'如果是字符串
			If Not o_map.Exists(n) Then
			'如果散列中没有此项，添加映射关系及数组值
				o_map(n) = Size
				o_map(Size) = n
				Push v
			Else
			'如果已有该项，更新数组值
				a_list(o_map(n)) = v
			End If
		Else
			Easp.Error.Raise 43
		End If
	End Property
	
	'取某一项值
	Public Default Property Get At(ByVal n)
		If Easp.Test(n,"^\d+$") Then
			If n < Size Then
				At = a_list(n)
			Else
				At = Null
				Easp.Error.Msg = "(当前下标 " & n & " 超过了最大下标 " & [End] & " )"
				Easp.Error.Raise 41
			End If
		ElseIf Easp.Test(n,"^[\w-\./]+$") Then
			If o_map.Exists(n) Then
				At = a_list(o_map(n))
			Else
				At = Null
				Easp.Error.Msg = "(当前列 " & n & " 不在数组列中)"
				Easp.Error.Raise 41
			End If
		End If
	End Property
	
	'源数据
	Public Property Let Data(ByVal a)
		If isArray(a) Then
			a_list = a
		Else
			a_list = Split(a, " ")
		End If
		Size = Ubound(a_list) + 1
	End Property
	'取出为普通数组
	Public Property Get Data
		Data = a_list
	End Property
	
	'映射关系
	Public Property Let Maps(ByVal d)
		If TypeName(d) = "Dictionary" Then CloneDic o_map, d
	End Property
	Public Property Get Maps
		Set Maps = o_map
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
		Dim i, v
		v = At(0)
		If Size > 1 Then
			For i = 1 To [End]
				If StrComp(At(i),v,i_comp) = 1 Then v = At(i)
			Next
		End If
		Max = v
	End Property
	
	'获取最小元素
	Public Property Get Min
		Dim i, v
		v = At(0)
		If Size > 1 Then
			For i = 1 To [End]
				If StrComp(At(i),v,i_comp) = -1 Then v = At(i)
			Next
		End If
		Min = v
	End Property
	
	'添加一个元素到开头
	Public Sub UnShift(ByVal v)
		Insert 0, v
	End Sub
	
	'删除第一个元素
	Public Sub Shift
		[Delete] 0
	End Sub
	
	'添加一个元素到结尾
	Public Sub Push(ByVal v)
		ReDim Preserve a_list(Size)
		a_list(Size) = v
		Size = Size + 1
	End Sub
	
	'删除最后一个元素
	Public Sub Pop
		RemoveMap [End]
		ReDim Preserve a_list([End]-1)
		Size = Size - 1
	End Sub
	Private Sub RemoveMap(ByVal i)
		If o_map.Exists(i) Then
			o_map.Remove o_map(i)
			'Easp.WN "=Delete==mapRemove:" & o_map(i)
			o_map.Remove i
			'Easp.WN "=Delete==mapRemove:" & i
		End If
	End Sub
	Private Sub UpFrom(ByVal n, ByVal i)
		If n = i Then Exit Sub
		If o_map.Exists(i) Then
			o_map(o_map(i)) = n
			o_map(n) = o_map(i)
			o_map.Remove i
			'Easp.WN "=Delete==UpFromRemove:" & i & "  o_map(count):" & o_map.count
		End If
		At(n) = At(i)
	End Sub
	
	'在指定下标插入一个元素
	Public Sub Insert(ByVal n, ByVal v)
		Dim i,j
		If n > [End] Then
		'如果下标大于最大下标
			If isArray(v) Then
			'如果插入一个数组，逐个赋值
				For i = 0 To UBound(v)
					At(n+i) = v(i)
				Next
			Else
			'是字符串直接赋值
				At(n) = v
			End If
		Else
		'如果插入到数组中间
			'如果插入一个数组
			For i = Size To (n+1) Step -1
			'将原数组插入点之后的值移动到新位置（腾出位置）
				If isArray(v) Then
				'如果是数组，要腾出数组的长度个位置
					UpFrom i+UBound(v), i-1
					'Easp.WN "把 " &i-1& "的值修改到 " &i+UBound(v)& " 上"
				Else
				'否则只腾出一个位置
					UpFrom i, i-1
				End If
			Next
			'把新值插入腾出的位置上
			If isArray(v) Then
				For i = 0 To UBound(v)
					At(n+i) = v(i)
				Next
			Else
				At(n) = v
			End If
		End If
	End Sub
	'检测是否包含某元素
	Public Function Has(ByVal v)
		Has = (indexOf__(a_list, v) > -1)
	End Function
	
	'检测元素在数组中的下标
	Public Function IndexOf(ByVal v)
		IndexOf = indexOf__(a_list, v)
	End Function
	Public Function IndexOfHash(ByVal v)
		Dim i : i = indexOf__(a_list, v)
		If i = -1 Then IndexOfHash = -1 : Exit Function
		If o_map.Exists(i) Then
			IndexOfHash = o_map(i)
		Else
			IndexOfHash = 0
		End If
	End Function
	Private Function indexOf__(ByVal arr, ByVal v)
		Dim i
		indexOf__ = -1
		For i = 0 To UBound(arr)
			If StrComp(arr(i),v,i_comp) = 0 Then
				indexOf__ = i
				Exit For
			End If
		Next
	End Function
	
	'删除一个或多个元素
	Public Sub [Delete](ByVal n)
		Dim arr(),tmp,a,x,y,i
		If Instr(n, ",")>0 Or Instr(n,"-")>0 Then
		'如果是删除多个元素
			n = Replace(n,"\s","0")
			n = Replace(n,"\e",[End])
			a = Split(n, ",")
			a = SortArray(a,0,UBound(a))
			tmp = "0-"
			For i = 0 To Ubound(a)
				If Instr(a(i),"-")>0 Then
					x = Easp.CLeft(a(i),"-")
					y = Easp.CRight(a(i),"-")
					'Easp.WN a(i)
					tmp = tmp & x-1 & ","
					tmp = tmp & y+1 & "-"
				Else
					tmp = tmp & a(i)-1 & "," & a(i)+1 & "-"
				End If
			Next
			tmp = tmp & [End]
			'Easp.WN tmp
			Slice tmp
		Else
		'只删除一项
			If Not isNumeric(n) And o_map.Exists(n) Then
				n = o_map(n)
				RemoveMap n
			End If
			For i = n+1 To [End]
				UpFrom i-1, i
			Next
			Pop
		End If
	End Sub

	'移除重复元素只保留一个
	Public Sub Uniq
		Dim arr(),i,j : j = 0
		ReDim arr(0)
		o_hash.RemoveAll
		For i = 0 To [End]
			'如果新数组中没有该值
			If indexOf__(arr, At(i)) = -1 Then
				ReDim Preserve arr(j)
				arr(j) = At(i)
				'Easp.WN "把元素" & i & "存入了新数组的 " & j
				If o_map.Exists(i) Then
					o_hash.Add j, o_map(i)
					o_hash.Add o_map(i), j
					'Easp.WN "把Hash中的第("&i&")项" &o_map(i)& "存入新Hash的" & j
				End If
				j = j + 1
			End If
		Next
		Data = arr
		CloneDic o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Private Sub CloneDic(ByRef map, ByRef hash)
		Dim key
		map.RemoveAll
		For Each key In hash
			map(key) = hash(key)
		Next
	End Sub
	
	'随机排序(洗牌)
	Public Sub Rand
		Dim i, j, tmp, Ei, Ej, Ti, Tj
		For i = 0 To [End]
			j = Easp.Rand(0,[End])
			'检测是否为Hash，如果是Hash就把值存起来
			Ei = o_map.Exists(i)
			Ej = o_map.Exists(j)
			If Ei Then Ti = o_map(i)
			If Ej Then Tj = o_map(j)
			'数组值互换
			tmp = At(j)
			At(j) = At(i)
			At(i) = tmp
			'Hash值互换
			If Ei Then
				o_map(j) = Ti
				o_map(Ti) = j
			End If
			If Ej Then
				o_map(i) = Tj
				o_map(Tj) = i
			End If
			'如果其中至少一个为空，则删除在Hash中的此下标
			If Not (Ei And Ej) Then
				If Ei Then o_map.Remove i
				If Ej then o_map.Remove j
			End If
		Next
	End Sub
	
	'反向排列数组
	Public Sub Reverse
		Dim arr(),i,j : j = 0
		ReDim arr([End])
		o_hash.RemoveAll
		For i = [End] To 0 Step -1
			arr(j) = At(i)
			If o_map.Exists(i) Then
				o_hash.Add j, o_map(i)
				o_hash.Add o_map(i), j
			End If
			j = j + 1
		Next
		Data = arr
		CloneDic o_map, o_hash
		o_hash.RemoveAll
	End Sub

	'搜索包含指定字符串的元素
	Public Sub Search(ByVal s)
		Data = Filter(a_list, s, True, i_comp)
	End Sub

	'搜索不包含指定字符串的元素
	Public Sub SearchNot(ByVal s)
		Data = Filter(a_list, s, False, i_comp)
	End Sub
	
	'删除空元素
	Public Sub Compact
		Dim arr(), i, j : j = 0
		o_hash.RemoveAll
		For i = 0 To [End]
			If Easp.Has(At(i)) Then
				ReDim Preserve arr(j)
				arr(j) = At(i)
				If o_map.Exists(i) Then
					o_hash.Add j, o_map(i)
					o_hash.Add o_map(i), j
				End If
				j = j + 1
			End If
		Next
		Data = arr
		CloneDic o_map, o_hash
		o_hash.RemoveAll
	End Sub
	
	'清空
	Public Sub Clear
		a_list = Array()
		Size = 0
	End Sub
	
	'排序
	Public Sub Sort
		Data = SortArray(a_list, 0, [End])
	End Sub
	Private Function SortArray(ByRef arr, ByRef low, ByRef high)
		If Not IsArray(arr) Then Exit Function
		If Easp.IsN(arr) Then Exit Function
		Dim l, h, m, v, x
		l = low : h = high
		m = (low + high) \ 2 : v = arr(m)
		Do While (l <= h)
			Do While (StrComp(arr(l),v,i_comp) = -1 And l < high)
				l = l + 1
			Loop
			Do While (StrComp(v,arr(h),i_comp) = -1 And h > low)
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
		Data = Slice__(s)
	End Sub
	'按下标取List的一部分元素返回一个新的List对象
	Public Function [Get](ByVal s)
		Set [Get] = Me.New
		[Get].Data = Slice__(s)
	End Function
	Private Function Slice__(ByVal s)
		Dim a,i,j,k,x,y,arr
		s = Replace(s,"\s",0)
		s = Replace(s,"\e",[End])
		a = Split(s, ",")
		arr = Array() : k = 0
		For i = 0 To Ubound(a)
			ReDim Preserve arr(k)
			'Easp.WN "Big:" & k
			If Instr(a(i),"-")>0 Then
				x = Int(Easp.CLeft(a(i),"-"))
				y = Int(Easp.CRight(a(i),"-"))
				For j = x To y
					ReDim Preserve arr(k)
					'Easp.WN "Small:"&k & "=" & x & "-" & y
					arr(k) = At(j)
					If o_map.Exists(j) Then
						
					End If
					If j < y Then k = k + 1
				Next
			Else
				arr(k) = At(Int(Trim(a(i))))
				k = k + 1
			End If
		Next
		Slice__ = arr
	End Function
	
	'复制List对象
	Public Function Clone
		Set Clone = Me.New
		Clone.Data = a_list
		If o_map.Count>0 Then Clone.Maps = o_map
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