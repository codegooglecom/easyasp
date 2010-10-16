<%
'######################################################################
'## easp.tpl.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp Templates Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/10/17 00:42:30
'## Description :   Use Templates with EasyAsp
'##
'######################################################################
Class EasyAsp_Tpl
	Private s_html, s_unknown, s_dict, s_path, s_m, s_ms, s_me
	Private o_tag, o_blockdata, o_block, o_blocktag, o_blocks, o_attr
	Private b_asp

	Private Sub class_Initialize
		s_path = ""
		s_unknown = "keep"
		s_dict = "Scripting.Dictionary"
		Set o_tag = Server.CreateObject(s_dict) : o_tag.CompareMode = 1
		Set o_blockdata = Server.CreateObject(s_dict) : o_blockdata.CompareMode = 1
		Set o_block = Server.CreateObject(s_dict) : o_block.CompareMode = 1
		Set o_blocktag = Server.CreateObject(s_dict) : o_blocktag.CompareMode = 1
		Set o_blocks = Server.CreateObject(s_dict) : o_blocks.CompareMode = 1
		Set o_attr = Server.CreateObject(s_dict) : o_attr.CompareMode = 1
		s_m = "{*}"
		getMaskSE s_m
		b_asp = False
		s_html = ""
	End Sub
	Private Sub Class_Terminate
		Set o_tag = Nothing
		Set o_blockdata = Nothing
		Set o_block = Nothing
		Set o_blockTag = Nothing
		Set o_blocks = Nothing
		Set o_attr = Nothing
	End Sub

	'模板路径
	Public Property Get FilePath
		FilePath = s_path
	End Property
	Public Property Let FilePath(ByVal f)
		If Right(f,1)<>"/" Then f = f & "/"
		s_path = f
	End Property
	'加载模板方法一
	Public Property Let [File](ByVal f)
		Load(f)
	End Property
	'通过文本加载模板
	Public Property Let [Source](ByVal s)
		LoadStr(s)
	End Property
	'标签的标识符
	Public Property Get TagMask
		TagMask = s_m
	End Property
	Public Property Let TagMask(ByVal m)
		s_m = m
		getMaskSE s_m
	End Property
	'模板中是否可以执行ASP代码
	Public Property Get AspEnable
		AspEnable = b_asp
	End Property
	Public Property Let AspEnable(ByVal b)
		b_asp = b
	End Property
	'如何处理未定义的标签
	Public Property Get TagUnknown
		TagUnknown = s_unknown
	End Property
	Public Property Let TagUnknown(ByVal s)
		Select Case LCase(s)
			Case "1", "remove"
				s_unknown = "remove"
			Case "2", "comment"
				s_unknown = "comment"
			Case Else
				s_unknown = "keep"
		End Select
	End Property
	'建新实例
	Public Function [New]()
		Set [New] = New EasyASP_Tpl
	End Function
	'取循环块的属性
	Public Function Attr(ByVal s)
		If Not o_attr.Exists(s) Then Exit Function
		Attr = o_attr.Item(s)
	End Function

	'加载模板方法二
	Public Sub Load(ByVal f)
		s_html = LoadInc(s_path & f,"")
		SetBlocks()
	End Sub
	'从文本加载模板
	Public Sub LoadStr(ByVal s)
		s_html = s
		SetBlocks()
	End Sub
	'加载附加模板
	Public Sub TagFile(ByVal tag, ByVal f)
		LoadToTag tag,0,f
	End Sub
	'从文本加载附加模板
	Public Sub TagStr(ByVal tag, ByVal s)
		LoadToTag tag,1,s
	End Sub
	'加载附加模板原型
	Private Sub LoadToTag(ByVal tag, ByVal t, ByVal f)
		Dim s
		s = Easp.IIF(t = 0, LoadInc(s_path & f,""), f)
		If Easp.Has(tag) Then
			s_html = Easp.regReplace(s_html, s_ms & tag & s_me, s)
		Else
			s_html = s_html & s
		End If
		SetBlocks()
	End Sub
	'替换标签(默认方法)
	Public Default Sub Tag(ByVal s, ByVal v)
		Dim i,f
		If TypeName(v) = "Recordset" Then
			If Easp.Has(v) Then
				For i = 0 To v.Fields.Count - 1
					Tag s & "(" & v.Fields(i).Name & ")", v.Fields(i).Value
				Next
			End If
		Else
			If Easp.IsN(v) Then v = ""
			If o_tag.Exists(s) Then o_tag.Remove s
			o_tag.Add s, Cstr(v)
		End If
	End Sub
	'在已替换标签后添加新内容
	Public Sub Append(ByVal s, ByVal v)
		If Easp.IsN(v) Then v = ""
		Dim tmp
		If o_tag.Exists(s) Then
			tmp = o_tag.Item(s) & Cstr(v)
			o_tag.Remove s
			o_tag.Add s, Cstr(tmp)
		Else
			o_tag.Add s, Cstr(v)
		End If
	End Sub
	'更新循环块数据
	Public Sub [Update](ByVal b)
		Dim Matches, Match, tmp, s, rule, data
		s = BlockData(b)
		rule = Chr(0) & "(\w+?)" & Chr(0)
		Set Matches = Easp.regMatch(s, rule)
		Set Match = Matches
		For Each Match In Matches
			data = Match.SubMatches(0)
			If o_blocktag.Exists(data) Then
				s = Replace(s, Match.Value, o_blocktag.Item(data))
				o_blocktag.Remove(data)
			End If
		Next
		If o_blocktag.Exists(b) Then
			tmp = o_blocktag.Item(b) & s
			o_blocktag.Remove b
			o_blocktag.Add b, Cstr(tmp)
		Else
			o_blocktag.Add b, Cstr(s)
		End If
		Set Matches = Easp.regMatch(s_html, Chr(0) & b & Chr(0))
		Set Match = Matches
		For Each Match In Matches
			s = BlockTag(b)
			s_html = Replace(s_html, Match.Value, s & Match.Value)
		Next
		If o_block.Exists(b) Then o_block.Remove b
	End Sub
	'获取最终html
	Public Function GetHtml()
		Dim Matches, Match, n, b
		'替换标签
		Set Matches = Easp.RegMatch(s_html, s_ms & "(.+?)" & s_me)
		'Easp.WN "rule:" & s_ms & "(.+?)" & s_me
		For Each Match In Matches
			n = Match.SubMatches(0)
			'Easp.WN "match:" & Match.Value
			If o_tag.Exists(n) Then
				s_html = Replace(s_html, Match.Value, o_tag.Item(n))
				'Easp.WN "match_tag:" & Match.Value
				'Easp.WN "match_dic:" & o_tag.Item(n)
			End If
		Next
		'替换未处理循环块
		Set Matches = Easp.regMatch(s_html, Chr(0) & "(\w+?)" & Chr(0))
		For Each Match In Matches
			b = Match.SubMatches(0)
			If o_block.Exists(b) Then [Update](b)
			s_html = Replace(s_html, Match.Value, "")
		Next
		'替换未处理标签
		Set Matches = Easp.RegMatch(s_html, s_ms & "(.+?)" & s_me)
		select case s_unknown
			case "keep"
				'Do Nothing
			case "remove"
				For Each Match In Matches
					s_html = Replace(s_html, Match.Value, "")
				Next
			case "comment"
				For Each Match In Matches
					s_html = Replace(s_html, Match.Value, "<!-- Unknown Tag '" & Match.Submatches(0) & "' -->")
				Next
		End select
		GetHtml = s_html
	End Function
	'输出模板内容
	Public Sub Show()
		Easp.W GetHtml
	End Sub
	'生成html标签
	Public Function MakeTag(ByVal t, ByVal f)
		Dim s,e,a,i,m
		If Instr(t,":")>0 Then
			m = Easp.CRight(t,":")
			t = Easp.CLeft(t,":")
			m = Easp.DateTime(Now,m)
		End If
		Select Case Lcase(t)
			Case "css"
				s = "<link href="""
				e = """ rel=""stylesheet"" type=""text/css"" />"
			Case "js"
				s = "<scr"&"ipt type=""text/javascript"" src="""
				e = """></scr"&"ipt>"
			Case "author", "keywords", "description", "copyright", "generator", "revised", "others"
				MakeTag = MakeTagMeta("name",t,f)
				Exit Function
			Case "content-type", "expires", "refresh", "set-cookie"
				MakeTag = MakeTagMeta("http-equiv",t,f)
				Exit Function
		End Select
		a = Split(f,"|")
		For i = 0 To Ubound(a)
			a(i) = s & Trim(a(i)) & Easp.IfThen(Easp.Has(m),"?" & m) & e
		Next
		MakeTag = Join(a,vbCrLf)
	End Function

	'生成Meta标签
	Private Function MakeTagMeta(ByVal m, ByVal t, ByVal s)
		MakeTagMeta = "<meta " & m & "=""" & t & """ content=""" & s & """ />"
	End Function
	'获取Tag标识
	Private Sub getMaskSE(ByVal m)
		s_ms = Easp.RegEncode(Easp.CLeft(m,"*"))
		s_me = Easp.RegEncode(Easp.CRight(m,"*"))
	End Sub
	'载入模板文件并将无限级include模板载入
	Private Function LoadInc(ByVal f, ByVal p)
		Dim h,pa,rule,inc,Match,incFile,incStr
		pa = Easp.IIF(Left(f,1)="/","",p)
		If b_asp Then
			h = Easp.GetInclude( pa & f )
		Else
			h = Easp.Read( pa & f )
		End If
		rule = "(<!--[\s]*)?" & s_ms & "#include:(.+?)" & s_me & "([\s]*-->)?"
		If Easp.Test(h,rule) Then
			If Easp.isN(p) Then
				If Instr(f,"/")>0 Then p = Left(f,InstrRev(f,"/"))
			Else
				If Instr(f,"/")>0 Then p = pa & Left(f,InstrRev(f,"/"))
			End If
			Set inc = Easp.regMatch(h,rule)
			For Each Match In inc
				incFile = Match.SubMatches(1)
				incStr = LoadInc(incFile, p)
				h = Replace(h,Match,incStr)
			Next
			Set inc = Nothing
		End If
		LoadInc = h
	End Function
	'读取循环块标签
	Private Sub SetBlocks()
		Dim Matches, Match, rule, n, i, j
		i = 0
		rule = "(<!--[\s]*)?" & s_ms & "#:(.+?)" & s_me
		If Not Easp.Test(s_html, rule) Then Exit Sub
		Set Matches = Easp.regMatch(s_html,rule)
		'找到循环块
		For Each Match In Matches
			n = Match.SubMatches(1)
			'Easp.WN "block:" & n
			'把循环块标签加入字典
			If o_blocks.Exists(i) Then o_blocks.Remove i
			o_blocks.Add i, n
			i = i + 1
		Next
		'从最后一层开始初始化循环块，实现无限层嵌套
		For j = i-1 To 0 Step -1
			Begin o_blocks.item(j)
		Next
	End Sub
	'初始化循环块
	Private Sub Begin(ByVal b)
		Dim Matches, Match, rule, data, attrs, attr, att, aname, avalue, atag
		rule = "(<!--[\s]*)?(" & s_ms & ")#:(" & b & ")(" & s_me & ")([\s]*-->)?([\s\S]+?)(<!--[\s]*)?\2/#:\3\4([\s]*-->)?"
		'如果循环块有属性则取出属性
		If Instr(b," ")>0 Then
			attrs = Easp.CRight(b, " ")
			b = Easp.CLeft(b, " ")
		  rule = "(<!--[\s]*)?(" & s_ms & ")#:(" & b & " " & Easp.RegEncode(attrs) & ")(" & s_me & ")([\s]*-->)?([\s\S]+?)(<!--[\s]*)?\2/#:" & b & "\4([\s]*-->)?"
		End If
		Set Matches = Easp.regMatch(s_html, rule)
		Set Match = Matches
		For Each Match In Matches
			'Easp.WN "block_tag:" & b
			'Easp.WN "block_attr:" & attrs
			'取循环块内容
			data = Match.SubMatches(5)
			'把循环块内容存入标签名对应的字典
			If o_blockdata.Exists(b) Then
				o_blockdata.Remove(b)
				o_block.Remove(b)
			End If
			o_blockdata.Add b, Cstr(data)
			o_block.Add b, Cstr(b)
			If Easp.Has(attrs) Then
			'如果有属性则取出每个属性
				Set attr = Easp.RegMatch(attrs,"((\w+)=(['""])(.+?)\3)|((\w+)=([^\s]+))")
				For Each att In attr
					aname = Easp.CLeft(att.Value, "=")
					avalue = Easp.RegReplace(att.Value, "\w+=(['""]?)(.+?)\1", "$2")
					atag = b & "." & aname
					'Easp.WN "attr '" & aname & "' = ["&avalue&"]"
					If o_attr.Exists(atag) Then o_attr.Remove(atag)
					o_attr.Add atag, avalue
				Next
			End If
			'把原始内容中的循环块作临时替换
			s_html = Easp.regReplace(s_html, rule, Chr(0) & b & Chr(0))
		Next
	End Sub
	'取循环块原始模板数据
	Private Function BlockData(ByVal b)
		Dim tmp, s
		If o_blockdata.Exists(b) Then
			tmp = o_blockdata.Item(b)
			'替换已定义标签
			s = UpdateBlockTag(tmp)
			BlockData = s
		Else
			BlockData = "<!--" & Chr(0) & b & Chr(0) & "-->"
		End If
	End Function
	'取循环块临时数据
	Private Function BlockTag(ByVal b)
		Dim tmp, s
		If o_blockdata.Exists(b) Then
			tmp = o_blocktag.Item(b)
			'替换已定义标签
			s = UpdateBlockTag(tmp)
			BlockTag = s
			'删除循环块临时数据
			o_blocktag.Remove(b)
		Else
			BlockTag = "<!--" & Chr(0) & b & Chr(0) & "-->"
		End If
	End Function
	'更新循环块标签
	Private Function UpdateBlockTag(ByVal s)
		Dim Matches, Match, data, rule
		Set Matches = Easp.RegMatch(s, s_ms & "(.+?)" & s_me)
		For Each Match In Matches
			'取标签名
			data = Match.SubMatches(0)
			'如果此标签有替换值
			If o_tag.Exists(data) Then
				rule = Match.Value
				'替换标签为相应的值
				If Easp.isN(o_tag.Item(data)) Then
					s = Replace(s, rule, "")
				Else
					s = Replace(s, rule, o_tag.Item(data))
				End If
			End If
		Next
		UpdateBlockTag = s
	End Function
End Class
%>