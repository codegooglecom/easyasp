<%
'######################################################################
'## easp.json.asp
'## -------------------------------------------------------------------
'## Feature     :   JSON For ASP
'## Version     :   v2.2 alpha
'## Author      :   Tu?ul Topuz @ 2009 [VBS JSON 2.0.3]
'## Update      :   Coldstone(coldstone[at]qq.com) & Mr.Zhang & Liaoyizhi
'## Update Date :   2010/01/26 16:08:30
'## Description :   Create JSON strings in EasyASP
'##
'######################################################################
Class EasyAsp_JSON
	Public Collection, Count, QuotedVars, Kind
	'Kind : 0 = object, 1 = array
	Private Sub Class_Initialize
		Set Collection = CreateObject("Scripting.Dictionary")
		'�����Ƿ�������
		If TypeName(Easp.Json) = "EasyAsp_JSON" Then
			QuotedVars = Easp.Json.QuotedVars
		Else
			QuotedVars = True
		End If
		Count = 0
		Easp.Error(10) = "������Ч��Easp JSON����"
	End Sub

	Private Sub Class_Terminate
		Set Collection = Nothing
	End Sub
	'����ʵ��
	Public Function [New](ByVal k)
		Set [New] = New EasyASP_JSON
		Select Case LCase(k)
			Case "0", "object" [New].Kind = 0
			Case "1", "array"  [New].Kind = 1
		End Select
	End Function

	Private Property Get Counter 
		Counter = Count
		Count = Count + 1
	End Property
	'��ֵ��ֵ������Easp��Json����
	Public Property Let Pair(p, v)
		If IsNull(p) Then p = Counter
		If TypeName(v) = "EasyAsp_JSON" Then
			Set Collection(p) = v
		Else
			Collection(p) = v
		End If
	End Property
	Public Default Property Get Pair(p)
		If IsNull(p) Then p = Count - 1
		If IsObject(Collection(p)) Then
			Set Pair = Collection(p)
		Else
			Pair = Collection(p)
		End If
	End Property
	'���������
	Public Sub Clean
		Collection.RemoveAll
	End Sub
	'ɾ��ĳһֵ
	Public Sub Remove(vProp)
		Collection.Remove vProp
	End Sub
	'��Ŀ��ת��Json�ַ���
	Public Function toJSON(vPair)
		Select Case VarType(vPair)
			Case 1
				toJSON = "null"
			Case 7
				toJSON = """" & CStr(vPair) & """"
			Case 8
				toJSON = """" & Easp.JSEncode(vPair) & """"
			Case 9
				Dim bFI,i 
				bFI = True
				toJSON = toJSON & Easp.IIF(vPair.Kind, "[", "{")
				For Each i In vPair.Collection
					If bFI Then bFI = False Else toJSON = toJSON & ","
					toJSON = toJSON & Easp.IfThen(vPair.Kind=0, Easp.IIF(QuotedVars, """" & i & """", i) & ":") & toJSON(vPair(i))
				Next
				toJSON = toJSON & Easp.IIF(vPair.Kind, "]", "}")
			Case 11
				toJSON = Easp.IIF(vPair, "true", "false")
			Case 12, 8192, 8204
				toJSON = RenderArray(vPair, 1, "")
			Case Else
				toJSON = Replace(vPair, ",", ".")
		End select
	End Function
	'�ݹ���������Json�ַ���
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
	'����Json�ַ���
	Public Property Get jsString
		jsString = toJSON(Me)
	End Property
	'��Json�ַ������
	Sub Flush
		Easp.W jsString
	End Sub
	'����Json����
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