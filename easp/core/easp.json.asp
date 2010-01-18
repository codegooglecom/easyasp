<%
'######################################################################
'## easp.json.asp
'## -------------------------------------------------------------------
'## Feature     :   JSON For ASP
'## Version     :   v2.2 alpha
'## Author      :   Tu?ul Topuz @ 2009 [VBS JSON 2.0.3]
'## Update      :   Coldstone(coldstone[at]qq.com) & Mr.Zhang
'## Update Date :   2010/1/18 22:28
'## Description :   Create JSON strings in EasyASP
'##
'######################################################################
Class EasyAsp_JSON
	Public Collection
	Public Count
	Public QuotedVars
	Public Kind  ' 0 = object, 1 = array

	Private Sub Class_Initialize
		Set Collection = CreateObject("Scripting.Dictionary")
		QuotedVars = False
		Count = 0
		Easp.Error(10) = "不是有效的Easp JSON对象！"
	End Sub

	Private Sub Class_Terminate
		Set Collection = Nothing
	End Sub
	'JSON Kind
'	Public Property Let Kind(k)
'		Select Case LCase(k)
'			Case "0", "object" iKind = 0
'			Case "1", "array"  iKind = 1
'		End Select
'	End Property

	' counter
	Private Property Get Counter 
		Counter = Count
		Count = Count + 1
	End Property

	' - data maluplation
	' -- pair
	Public Property Let Pair(p, v)
		If IsNull(p) Then p = Counter
		Collection(p) = v
	End Property

	Public Property Set Pair(p, v)
		If IsNull(p) Then p = Counter
		If TypeName(v) <> "EasyAsp_JSON" Then
			Easp.Error.Msg = "( Now Type: " & TypeName(v) & " )"
			Easp.Error.Raise 10
		End If
		Set Collection(p) = v
	End Property

	Public Default Property Get Pair(p)
		If IsNull(p) Then p = Count - 1
		If IsObject(Collection(p)) Then
			Set Pair = Collection(p)
		Else
			Pair = Collection(p)
		End If
	End Property

	Public Function [New](ByVal k)
		Set [New] = New EasyASP_JSON
		Select Case LCase(k)
			Case "0", "object" [New].Kind = 0
			Case "1", "array"  [New].Kind = 1
		End Select
	End Function
	' -- pair
	Public Sub Clean
		Collection.RemoveAll
	End Sub

	Public Sub Remove(vProp)
		Collection.Remove vProp
	End Sub
	' data maluplation

	' converting
	Public Function toJSON(vPair)
		Select Case VarType(vPair)
			Case 1	' Null
				toJSON = "null"
			Case 7	' Date
				' yaz saati problemi var
				' jsValue = "new Date(" & Round((vVal - #01/01/1970 02:00#) * 86400000) & ")"
				toJSON = "'" & CStr(vPair) & "'"
			Case 8	' String
				toJSON = "'" & Easp.JSEncode(vPair) & "'"
			Case 9	' Object
				Dim bFI,i 
				bFI = True
				If vPair.Kind Then toJSON = toJSON & "[" Else toJSON = toJSON & "{"
				For Each i In vPair.Collection
					If bFI Then bFI = False Else toJSON = toJSON & ","

					If vPair.Kind Then 
						toJSON = toJSON & toJSON(vPair(i))
					Else
						If QuotedVars Then
							toJSON = toJSON & "'" & i & "':" & toJSON(vPair(i))
						Else
							toJSON = toJSON & i & ":" & toJSON(vPair(i))
						End If
					End If
				Next
				If vPair.Kind Then toJSON = toJSON & "]" Else toJSON = toJSON & "}"
			Case 11
				If vPair Then toJSON = "true" Else toJSON = "false"
			Case 12, 8192, 8204
				toJSON = RenderArray(vPair, 1, "")
			Case Else
				toJSON = Replace(vPair, ",", ".")
		End select
	End Function

	Function RenderArray(arr, depth, parent)
		Dim first : first = LBound(arr, depth)
		Dim last : last = UBound(arr, depth)

		Dim index, rendered
		Dim limiter : limiter = ","

		RenderArray = "["
		For index = first To last
			If index = last Then
				limiter = ""
			End If 

			On Error Resume Next
			rendered = RenderArray(arr, depth + 1, parent & index & "," )

			If Err = 9 Then
				On Error GoTo 0
				RenderArray = RenderArray & toJSON(Eval("arr(" & parent & index & ")")) & limiter
			Else
				RenderArray = RenderArray & rendered & "" & limiter
			End If
		Next
		RenderArray = RenderArray & "]"
	End Function

	Public Property Get jsString
		jsString = toJSON(Me)
	End Property

	Sub Flush
		Easp.W jsString
	End Sub

	Public Function Clone
		Set Clone = ColClone(Me)
	End Function

	Private Function ColClone(core)
		Dim jsc, i
		Set jsc = new EasyAsp_JSON
		jsc.Kind = core.Kind
		For Each i In core.Collection
			If IsObject(core(i)) Then
				Set jsc(i) = ColClone(core(i))
			Else
				jsc(i) = core(i)
			End If
		Next
		Set ColClone = jsc
	End Function
End Class
%>