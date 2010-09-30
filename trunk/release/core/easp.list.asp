<%
'######################################################################
'## easp.list.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp List(Array) Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/09/30 00:03:10
'## Description :   A super Array class in EasyAsp
'##                 Support Array and Hash (like Dictionary Object)
'##
'######################################################################
Class EasyAsp_List
	Public Size, OverError
	Private o_hash, o_map
	Private a_list
	Private i_count, i_comp
	
	Private Sub Class_Initialize
		Set o_hash = Server.CreateObject("Scripting.Dictionary")
		Set o_map  = Server.CreateObject("Scripting.Dictionary")
		a_list = Array()
		Size = 0
		OverError = True
		Easp.Error(41) = "下标越界"
		Easp.Error(42) = "下标不能为空"
		Easp.Error(43) = "下标只能是数字、字母、下划线(_)、点(.)和斜杠(/)组成"
		Easp.Error(44) = "参数必须是数组或者List对象"
		i_comp = 1
	End Sub
	
	Private Sub Class_Terminate
		Set o_map  = Nothing
		Set o_hash = Nothing
	End Sub
	Public Function [New]()
		Set [New] = New EasyAsp_List
		[New].IgnoreCase = Me.IgnoreCase
	End Function
	Public Function NewArray(ByVal a)
		Set NewArray = New EasyAsp_List
		NewArray.IgnoreCase = Me.IgnoreCase
		NewArray.Data = a
	End Function
	Public Function NewHash(ByVal a)
		Set NewHash = New EasyAsp_List
		NewHash.IgnoreCase = Me.IgnoreCase
		NewHash.Hash = a
	End Function
	Public Property Let IgnoreCase(ByVal b)
		i_comp = Easp.IIF(b, 1, 0)
	End Property
	Public Property Get IgnoreCase
		IgnoreCase = (i_comp = 1)
	End Property
	Public Property Let At(ByVal n, ByVal v)
		If Easp.IsN(n) Then Easp.Error.Raise 42 : Exit Property
		If Easp.Test(n,"^\d+$") Then
			If n > [End] Then
				ReDim Preserve a_list(n)
				Size = n + 1
			End If
			a_list(n) = v
		ElseIf Easp.Test(n,"^[\w\./]+$") Then
			If Not o_map.Exists(n) Then
				o_map(n) = Size
				o_map(Size) = n
				Push v
			Else
				a_list(o_map(n)) = v
			End If
		Else
			Easp.Error.Raise 43
		End If
	End Property
	Public Default Property Get At(ByVal n)
		If Easp.Test(n,"^\d+$") Then
			If n < Size Then
				At = a_list(n)
			Else
				At = Null
				If OverError Then
					Easp.Error.Msg = "(当前下标 " & n & " 超过了最大下标 " & [End] & " )"
					Easp.Error.Raise 41
				End If
			End If
		ElseIf Easp.Test(n,"^[\w-\./]+$") Then
			If o_map.Exists(n) Then
				At = a_list(o_map(n))
			Else
				At = Null
				If OverError Then
					Easp.Error.Msg = "(当前列 " & n & " 不在数组Hash列中)"
					Easp.Error.Raise 41
				End If
			End If
		End If
	End Property
	Public Property Let Data(ByVal a)
		Data__ a, 0
	End Property
	Public Property Get Data
		Data = a_list
	End Property
	Public Property Let Hash(ByVal a)
		Data__ a, 1
	End Property
	Public Property Get Hash
		Dim arr, i
		arr = a_list
		For i = 0 To [End]
			If o_map.Exists(i) Then
				arr(i) = o_map(i) & ":" & arr(i)
			End If
		Next
		Hash = arr
	End Property
	Public Sub Data__(ByVal a, ByVal t)
		Dim arr, i, j
		If isArray(a) Then
			a_list = a
			Size = Ubound(a_list) + 1
			If t = 0 Then Exit Sub
			For i = 0 To Ubound(a)
				If Instr(a(i),":")>0 Then
					j = Easp.CLeft(a(i),":")
					If Not o_map.Exists(j) Then
						o_map.Add i, j
						o_map.Add j, i
					End If
					a(i) = Easp.CRight(a(i),":")
				End If
			Next
			a_list = a
		Else
			arr = Split(a, " ")
			a_list = arr
			Size = Ubound(a_list) + 1
			If t = 0 Then Exit Sub
			If Instr(a, ":")>0 Then
				For i = 0 To Ubound(arr)
					If Instr(arr(i),":")>0 Then
						j = Easp.CLeft(arr(i),":")
						If Not o_map.Exists(j) Then
							o_map.Add i, j
							o_map.Add j, i
						End If
						arr(i) = Easp.CRight(arr(i),":")
					End If
				Next
			End If
			a_list = arr
		End If
	End Sub
	Public Property Let Maps(ByVal d)
		If TypeName(d) = "Dictionary" Then CloneDic__ o_map, d
	End Property
	Public Property Get Maps
		Set Maps = o_map
	End Property
	Public Property Get Length
		Length = Size
	End Property
	Public Property Get [End]
		[End] = Size - 1
	End Property
	Public Property Get Count
		Dim i,j : j = 0
		For i = 0 To Size-1
			If Easp.Has(At(i)) Then j = j + 1
		Next
		Count = j
	End Property
	Public Property Get First
		First = At(0)
	End Property
	Public Property Get Last
		Last = At([End])
	End Property
	Public Property Get Max
		Dim i, v
		v = At(0)
		If Size > 1 Then
			For i = 1 To [End]
				If Compare__("gt", At(i), v) Then v = At(i)
			Next
		End If
		Max = v
	End Property
	Public Property Get Min
		Dim i, v
		v = At(0)
		If Size > 1 Then
			For i = 1 To [End]
				If Compare__("lt", At(i), v) Then v = At(i)
			Next
		End If
		Min = v
	End Property
	Public Property Get Serialize
		Dim tmp, i : tmp = ""
		For i = 0 To [End]
			If o_map.Exists(i) Then
				tmp = tmp & "&" & o_map(i) & "=" & Server.URLEncode(At(i))
			End If
		Next
		If Len(tmp)>1 Then tmp = Mid(tmp,2)
		Serialize = tmp
	End Property
	Public Function HasIndex(ByVal i)
		HasIndex = Index(i) >= 0
	End Function
	Public Function Index(ByVal i)
		If isNumeric(i) Then
			Index = Easp.IIF(i >= 0 And i <= [End], i, -1)
		Else
			If o_map.Exists(i) Then
				Index = o_map(i)
			Else
				Index = -1
			End If
		End If
	End Function
	Public Function IndexHash(ByVal i)
		If isNumeric(i) Then
			IndexHash = Easp.IfThen(o_map.Exists(i), o_map(i))
		Else
			IndexHash = Easp.IfThen(o_map.Exists(i), i)
		End If
	End Function
	Private Function Compare__(ByVal t, ByVal a, ByVal b)
		Dim isStr : isStr = False
		If VarType(a) = 8 Or VarType(b) = 8 Then
			isStr = True
			If IsNumeric(a) And IsNumeric(b) Then isStr = False
			If IsDate(a) And IsDate(b) Then isStr = False
		End If
		If isStr Then
			Select Case LCase(t)
				Case "lt" Compare__ = (StrComp(a,b,i_comp) = -1)
				Case "gt" Compare__ = (StrComp(a,b,i_comp) = 1)
				Case "eq" Compare__ = (StrComp(a,b,i_comp) = 0)
				Case "lte" Compare__ = (StrComp(a,b,i_comp) = -1 Or StrComp(a,b,i_comp) = 0)
				Case "gte" Compare__ = (StrComp(a,b,i_comp) = 1 Or StrComp(a,b,i_comp) = 0)
			End Select
		Else
			Select Case LCase(t)
				Case "lt" Compare__ = (a < b)
				Case "gt" Compare__ = (a > b)
				Case "eq" Compare__ = (a = b)
				Case "lte" Compare__ = (a <= b)
				Case "gte" Compare__ = (a >= b)
			End Select
		End If
	End Function
	Public Sub UnShift(ByVal v)
		Insert 0, v
	End Sub
	Public Function UnShift_(ByVal v)
		Set UnShift_ = Me.Clone
		UnShift_.UnShift v
	End Function
	Public Sub Shift
		[Delete] 0
	End Sub
	Public Function Shift_
		Set Shift_ = Me.Clone
		Shift_.Shift
	End Function
	Public Sub Push(ByVal v)
		ReDim Preserve a_list(Size)
		a_list(Size) = v
		Size = Size + 1
	End Sub
	Public Function Push_(ByVal v)
		Set Push_ = Me.Clone
		Push_.Push v
	End Function
	Public Sub Pop
		RemoveMap__ [End]
		ReDim Preserve a_list([End]-1)
		Size = Size - 1
	End Sub
	Public Function Pop_
		Set Pop_ = Me.Clone
		Pop_.Pop
	End Function
	Private Sub RemoveMap__(ByVal i)
		If o_map.Exists(i) Then
			o_map.Remove o_map(i)
			o_map.Remove i
		End If
	End Sub
	Private Sub UpFrom__(ByVal n, ByVal i)
		If n = i Then Exit Sub
		If o_map.Exists(i) Then
			o_map(o_map(i)) = n
			o_map(n) = o_map(i)
			o_map.Remove i
		End If
		At(n) = At(i)
	End Sub
	Public Sub Insert(ByVal n, ByVal v)
		Dim i,j
		If n > [End] Then
			If isArray(v) Then
				For i = 0 To UBound(v)
					At(n+i) = v(i)
				Next
			Else
				At(n) = v
			End If
		Else
			For i = Size To (n+1) Step -1
				If isArray(v) Then
					UpFrom__ i+UBound(v), i-1
				Else
					UpFrom__ i, i-1
				End If
			Next
			If isArray(v) Then
				For i = 0 To UBound(v)
					At(n+i) = v(i)
				Next
			Else
				At(n) = v
			End If
		End If
	End Sub
	Public Function Insert_(ByVal n, ByVal v)
		Set Insert_ = Me.Clone
		Insert_.Insert n, v
	End Function
	Public Function Has(ByVal v)
		Has = (indexOf__(a_list, v) > -1)
	End Function
	Public Function IndexOf(ByVal v)
		IndexOf = indexOf__(a_list, v)
	End Function
	Public Function IndexOfHash(ByVal v)
		Dim i : i = indexOf__(a_list, v)
		If i = -1 Then IndexOfHash = Empty : Exit Function
		If o_map.Exists(i) Then
			IndexOfHash = o_map(i)
		Else
			IndexOfHash = Empty
		End If
	End Function
	Private Function indexOf__(ByVal arr, ByVal v)
		Dim i
		indexOf__ = -1
		For i = 0 To UBound(arr)
			If Compare__("eq", arr(i),v) Then
				indexOf__ = i
				Exit For
			End If
		Next
	End Function
	Public Sub [Delete](ByVal n)
		Dim tmp,a,x,y,i
		If Instr(n, ",")>0 Or Instr(n,"-")>0 Then
			n = Replace(n,"\s","0")
			n = Replace(n,"\e",[End])
			a = Split(n, ",")
			For i = 0 To Ubound(a)
				If i>0 Then tmp = tmp & ","
				If Instr(a(i),"-")>0 Then
					x = Trim(Easp.CLeft(a(i),"-"))
					y = Trim(Easp.CRight(a(i),"-"))
					If Not Isnumeric(x) And o_map.Exists(x) Then x = o_map(x)
					If Not Isnumeric(y) And o_map.Exists(y) Then y = o_map(y)
					tmp = tmp & x & "-" & y
				Else
					x = Trim(a(i))
					If Not Isnumeric(x) And o_map.Exists(x) Then x = o_map(x)
					tmp = tmp & x
				End If
			Next
			a = Split(tmp,",")
			a = SortArray(a,0,UBound(a))
			tmp = "0-"
			For i = 0 To Ubound(a)
				If Instr(a(i),"-")>0 Then
					x = Easp.CLeft(a(i),"-")
					y = Easp.CRight(a(i),"-")
					tmp = tmp & x-1 & ","
					tmp = tmp & y+1 & "-"
				Else
					tmp = tmp & a(i)-1 & "," & a(i)+1 & "-"
				End If
			Next
			tmp = tmp & [End]
			Slice tmp
		Else
			If Not isNumeric(n) And o_map.Exists(n) Then
				n = o_map(n)
				RemoveMap__ n
			End If
			For i = n+1 To [End]
				UpFrom__ i-1, i
			Next
			Pop
		End If
	End Sub
	Public Function Delete_(ByVal n)
		Set Delete_ = Me.Clone
		Delete_.Delete n
	End Function
	Public Sub Uniq()
		Dim arr(),i,j : j = 0
		ReDim arr(-1)
		If o_hash.Count>0 Then o_hash.RemoveAll
		For i = 0 To [End]
			If indexOf__(arr, At(i)) = -1 Then
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
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Function Uniq_()
		Set Uniq_ = Me.Clone
		Uniq_.Uniq
	End Function
	Private Sub CloneDic__(ByRef map, ByRef hash)
		Dim key
		If map.Count > 0 Then map.RemoveAll
		For Each key In hash
			map(key) = hash(key)
		Next
	End Sub
	Public Sub Rand
		Dim i, j, tmp, Ei, Ej, Ti, Tj
		For i = 0 To [End]
			j = Easp.Rand(0,[End])
			Ei = o_map.Exists(i)
			Ej = o_map.Exists(j)
			If Ei Then Ti = o_map(i)
			If Ej Then Tj = o_map(j)
			tmp = At(j)
			At(j) = At(i)
			At(i) = tmp
			If Ei Then
				o_map(j) = Ti
				o_map(Ti) = j
			End If
			If Ej Then
				o_map(i) = Tj
				o_map(Tj) = i
			End If
			If Not (Ei And Ej) Then
				If Ei Then o_map.Remove i
				If Ej then o_map.Remove j
			End If
		Next
	End Sub
	Public Function Rand_()
		Set Rand_ = Me.Clone
		Rand_.Rand
	End Function
	Public Sub Reverse
		Dim arr(),i,j : j = 0
		ReDim arr([End])
		If o_hash.Count>0 Then o_hash.RemoveAll
		For i = [End] To 0 Step -1
			arr(j) = At(i)
			If o_map.Exists(i) Then
				o_hash.Add j, o_map(i)
				o_hash.Add o_map(i), j
			End If
			j = j + 1
		Next
		Data = arr
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Function Reverse_()
		Set Reverse_ = Me.Clone
		Reverse_.Reverse
	End Function
	Public Sub Search(ByVal s)
		Search__ s, True
	End Sub
	Public Function Search_(ByVal s)
		Set Search_ = Me.Clone
		Search_.Search s
	End Function
	Public Sub SearchNot(ByVal s)
		Search__ s, False
	End Sub
	Public Function SearchNot_(ByVal s)
		Set SearchNot_ = Me.Clone
		SearchNot_.SearchNot s
	End Function
	
	Private Sub Search__(ByVal s, ByVal keep)
		Dim arr,i,tmp
		arr = Filter(a_list, s, keep, i_comp)
		If o_map.Count = 0 Then
			Data = arr
		Else
			AddHash__ arr
		End If
	End Sub
	Public Sub Compact
		Dim arr(), i, j : j = 0
		If o_hash.Count>0 Then o_hash.RemoveAll
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
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Function Compact_()
		Set Compact_ = Me.Clone
		Compact_.Compact
	End Function
	Public Sub Clear
		a_list = Array()
		If o_map.Count>0 Then o_map.RemoveAll
		Size = 0
	End Sub
	Public Sub Sort
		Dim arr
		arr = a_list
		arr = SortArray(arr, 0, [End])
		If o_map.Count = 0 Then
			Data = arr
		Else
			AddHash__ arr
		End If
	End Sub
	Public Function Sort_()
		Set Sort_ = Me.Clone
		Sort_.Sort
	End Function
	Private Function SortArray(ByRef arr, ByRef low, ByRef high)
		If Not IsArray(arr) Then Exit Function
		If Easp.IsN(arr) Then Exit Function
		Dim l, h, m, v, x
		l = low : h = high
		m = (low + high) \ 2 : v = arr(m)
		Do While (l <= h)
			Do While (Compare__("lt",arr(l),v) And l < high)
				l = l + 1
			Loop
			Do While (Compare__("lt",v,arr(h)) And h > low)
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
	Private Sub AddHash__(ByVal arr)
		Dim tmp
		If o_hash.Count > 0 Then o_hash.RemoveAll
		For i = 0 To Ubound(arr)
			If IndexOfHash(arr(i))>"" Then
				tmp = IndexOfHash(arr(i))
				If Not o_hash.Exists(tmp) Then 
					o_hash.Add i, tmp
					o_hash.Add tmp, i
				End If
			End If
		Next
		Data = arr
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Sub Slice(ByVal s)
		Dim a,i,j,k,x,y,arr
		If o_hash.Count>0 Then o_hash.RemoveAll
		s = Replace(s,"\s",0)
		s = Replace(s,"\e",[End])
		a = Split(s, ",")
		arr = Array() : k = 0
		For i = 0 To Ubound(a)
			If Instr(a(i),"-")>0 Then
				x = Trim(Easp.CLeft(a(i),"-"))
				y = Trim(Easp.CRight(a(i),"-"))
				If Not Isnumeric(x) And o_map.Exists(x) Then x = o_map(x)
				If Not Isnumeric(y) And o_map.Exists(y) Then y = o_map(y)
				x = Int(x) : y = Int(y)
				For j = x To y
					ReDim Preserve arr(k)
					arr(k) = At(j)
					If o_map.Exists(j) Then
						If Not o_hash.Exists(o_map(j)) Then
							o_hash.Add k, o_map(j)
							o_hash.Add o_map(j), k
						End If
					End If
					k = k + 1
				Next
			Else
				ReDim Preserve arr(k)
				x = Trim(a(i))
				If Not Isnumeric(x) And o_map.Exists(x) Then x = o_map(x)
				x = Int(x)
				If o_map.Exists(x) Then
					If Not o_hash.Exists(o_map(x)) Then
						o_hash.Add k, o_map(x)
						o_hash.Add o_map(x), k
					End If
				End If
				arr(k) = At(x)
				k = k + 1
			End If
		Next
		Data = arr
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Function Slice_(ByVal s)
		Set Slice_ = Me.Clone
		Slice_.Slice s
	End Function
	Public Function [Get](ByVal s)
		Set [Get] = Slice_(s)
	End Function
	Public Function J(ByVal s)
		J = Join(a_list, s)
	End Function
	Public Function ToString()
		ToString = J(",")
	End Function
	Public Function ToArray
		ToArray = a_list
	End Function
	Public Function Clone
		Set Clone = Me.New
		Clone.Data = a_list
		If o_map.Count>0 Then Clone.Maps = o_map
	End Function
	Public Sub Map(ByVal f)
		Map__ f, 0
	End Sub
	Public Function Map_(ByVal f)
		Set Map_ = Me.Clone
		Map_.Map f
	End Function
	Public Sub [Each](ByVal f)
		Map__ f, 1
	End Sub
	Private Sub Map__(ByVal f, ByVal t)
		Dim i, tmp
		For i = 0 To [End]
			tmp = Value__(At(i))
			If t = 0 Then
				At(i) = Eval(f & "("& tmp &")")
			ElseIf t = 1 Then
				ExecuteGlobal f & "("& tmp &")"
			End If
		Next
	End Sub
	Private Function Value__(ByVal s)
		Dim tmp
		Select Case VarType(s)
			Case 7,8 tmp = """" & s & """"
			Case Else tmp = s
		End Select
		Value__ = tmp
	End Function
	Public Function Find(ByVal f)
		Dim i, k, tmp
		k = "i"
		If Easp.Test(f,"[a-zA-Z]+:(.+)") Then
			k = Easp.CLeft(f,":")
			f = Easp.CRight(f,":")
		End If
		k = "%" & k
		For i = 0 To [End]
			tmp = Replace(Trim(f), k, Value__(At(i)))
			If Eval(tmp) Then
				Find = At(i) : Exit Function
			End If
		Next
		Find = Empty
	End Function
	Public Sub [Select](ByVal f)
		Select__ f, 0
	End Sub
	Public Function Select_(ByVal f)
		Set Select_ = Me.Clone
		Select_.Select f
	End Function
	Private Sub Select__(ByVal f, ByVal t)
		Dim i, j, k, tmp, arr
		arr = Array() : j = 0
		If o_hash.Count>0 Then o_hash.RemoveAll
		k = "i"
		If Easp.Test(f,"[a-zA-Z]+:(.+)") Then
			k = Easp.CLeft(f,":")
			f = Easp.CRight(f,":")
		End If
		k = "%" & k
		For i = 0 To [End]
			tmp = Replace(Trim(f), k, Value__(At(i)))
			If t = 0 Then
				tmp = Eval(tmp)
			ElseIf t = 1 Then
				tmp = (Not Eval(tmp))
			End If
			If tmp Then
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
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Sub Reject(ByVal f)
		Select__ f, 1
	End Sub
	Public Function Reject_(ByVal f)
		Set Reject_ = Me.Clone
		Reject_.Reject f
	End Function
	Public Sub Grep(ByVal g)
		Dim i,j,arr
		arr = Array() : j = 0
		If o_hash.Count>0 Then o_hash.RemoveAll
		For i = 0 To [End]
			If Easp.Test(At(i),g) Then
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
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Function Grep_(ByVal g)
		Set Grep_ = Me.Clone
		Grep_.Grep g
	End Function
	Public Sub SortBy(ByVal f)
		Map f : Sort
	End Sub
	Public Function SortBy_(ByVal f)
		Set SortBy_ = Me.Clone
		SortBy_.SortBy f
	End Function
	Public Sub Times(ByVal t)
		Dim i, arr
		arr = a_list
		For i = 1 To t
			Insert Size, arr
		Next
	End Sub
	Public Function Times_(ByVal t)
		Set Times_ = Me.Clone
		Times_.Times t
	End Function
	Private Function IsList(ByVal o)
		IsList = (Lcase(TypeName(o)) = "easyasp_list")
	End Function
	Public Sub Splice(ByVal o)
		If Not isArray(o) And Not isList(o) Then Easp.Error.Raise 44 : Exit Sub
		Dim omap,dic,i
		If isArray(o) Then
			Insert Size, o
		ElseIf IsList(o) Then
				Set omap = o.Maps
				If omap.Count > 0 Then
					For i = 0 To o.End
						If omap.Exists(i) And (Not o_map.Exists(omap(i))) Then
							o_map.Add Size + i, omap(i)
							o_map.Add omap(i), Size + i
						End If
					Next
				End If
				Insert Size, o.Data
		End If
	End Sub
	Public Function Splice_(ByVal o)
		Set Splice_ = Me.Clone
		Splice_.Splice o
	End Function
	Public Sub Merge(ByVal o)
		Splice o
		Uniq
	End Sub
	Public Function Merge_(ByVal o)
		Set Merge_ = Me.Clone
		Merge_.Merge o
	End Function
	Public Sub Inter(ByVal o)
		If Not isArray(o) And Not isList(o) Then Easp.Error.Raise 44 : Exit Sub
		Dim i,j,k,omap,arr
		arr = Array() : j = 0
		If o_hash.Count>0 Then o_hash.RemoveAll
		If isArray(o) Then
			For i = 0 To Ubound(o)
				If Has(o(i)) Then
					ReDim Preserve arr(j)
					arr(j) = o(i)
					k = IndexOf(o(i))
					If o_map.Exists(k) Then
						o_hash.Add j, o_map(k)
						o_hash.Add o_map(k), j
					End If
					j = j + 1
				End If
			Next
		ElseIf IsList(o) Then
			Set omap = o.Maps
			For i = 0 To o.End
				If Has(o(i)) Then
					ReDim Preserve arr(j)
					arr(j) = o(i)
					k = IndexOf(o(i))
					If o_map.Exists(k) Then
						o_hash.Add j, o_map(k)
						o_hash.Add o_map(k), j
					ElseIf omap.Exists(i) Then
						o_hash.Add j, omap(i)
						o_hash.Add omap(i), j
					End If
					j = j + 1
				End If
			Next
		End If
		Data = arr
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Function Inter_(ByVal o)
		Set Inter_ = Me.Clone
		Inter_.Inter o
	End Function
	Public Sub Diff(ByVal o)
		If Not isArray(o) And Not isList(o) Then Easp.Error.Raise 44 : Exit Sub
		Dim i,j,arr,a
		arr = Array() : j = 0
		If o_hash.Count>0 Then o_hash.RemoveAll
		If isArray(o) Then
			a = o
			Set o = Me.New
			o.Data = a
		End If
		For i = 0 To [End]
			If Not o.Has(At(i)) Then
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
		CloneDic__ o_map, o_hash
		o_hash.RemoveAll
	End Sub
	Public Function Diff_(ByVal o)
		Set Diff_ = Me.Clone
		Diff_.Diff o
	End Function
	Public Function Eq(ByVal o)
		If Not isArray(o) And Not isList(o) Then Easp.Error.Raise 44 : Exit Function
		Dim a, e, m, i, j
		If isArray(o) Then
			a = o
			e = Ubound(a)
		ElseIf isList(o) Then
			a = o.Data
			e = o.End
		End If
		m = Easp.IIF([End] < e, [End], e)
		For i = 0 To m
			If Compare__("gt", At(i), a(i)) Then
				Eq = 1 : Exit Function
			ElseIf Compare__("lt", At(i), a(i)) Then
				Eq = -1 : Exit Function
			End If
		Next
		If [End] > e Then
			Eq = 1
		ElseIf [End] < e Then
			Eq = -1
		Else
			Eq = 0
		End If
	End Function
	Public Function Son(ByVal o)
		If Not isArray(o) And Not isList(o) Then Easp.Error.Raise 44 : Exit Function
		Son = True
		Dim i
		If isList(o) Then o = o.Data
		For i = 0 To Ubound(o)
			If Not Has(o(i)) Then Son = False : Exit Function
		Next
	End Function
End Class
%>