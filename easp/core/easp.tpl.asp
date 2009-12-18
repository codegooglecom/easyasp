<%
Class EasyAsp_tpl
	Private s_html, s_m, s_ms, s_me, s_dic,s_block
	Private b_asp
	private o_block, o_data, o_blockTag, o_blockdata
	
	Private Sub Class_Initialize
		s_html = ""
		s_block = ""
		s_m = "{*}"
		b_asp = False
		getMaskSE s_m
		s_dic = "Scripting.Dictionary"
		Set o_blockTag = CreateObject(s_dic)
		Set o_block = CreateObject(s_dic)
		Set o_blockdata = CreateObject(s_dic)
		Set o_data = CreateObject(s_dic)
	End Sub
	
	Private Sub Class_Terminate
		Set o_data = Nothing
		Set o_blockdata = Nothing
		Set o_block = Nothing
		Set o_blockTag = Nothing
	End Sub
	
	Public Property Let [File](ByVal f)
		Load(f)
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
	'获取Tag标识
	Private Sub getMaskSE(ByVal m)
		s_ms = Easp.CLeft(m,"*")
		s_me = Easp.CRight(m,"*")
	End Sub
	'正则表达式特殊字符转义
	Private Function FixRegStr(ByVal s)
		Dim re,i
		re = Split("$,(,),*,+,.,[,?,\,^,{,|",",")
		For i = 0 To Ubound(re)
			s = Replace(s,re(i),"\"&re(i))
		Next
		FixRegStr = s
	End Function
	'分析循环元素
	private Sub GetBlock(ByVal s)
		Dim rule,Matches,Match,i,t
		Dim b,ruleblock
		rule = "(<!--[\s]*)?" & FixRegStr(s_ms) & "#:(.+?)" & FixRegStr(s_me) & "([\s]*-->)?"
		If Not Easp_Test(s,rule) Then Exit Sub
		'取循环标签名t
		Set Matches = Easp_Match(s,rule)
		i = 1
		For Each Match In Matches
			t = Match.SubMatches(1)
			ruleblock = "(<!--[\s]*)?" & FixRegStr(s_ms) & "#:" & t & "" & FixRegStr(s_me) & "([\s]*-->)?([\s\S]+?)(<!--[\s]*)?" & FixRegStr(s_ms) & "/#:" & t & "" & FixRegStr(s_me) & "([\s]*-->)?"
			'取循环块
			If Easp_Test(s,ruleblock) Then
				o_blockTag(i) = t
				Set b = Easp_Match(s,ruleblock)(0)
				o_block(t & "__b") = ruleblock
				o_block(t & "__s") = b.SubMatches(2)
				o_block(t) = b.SubMatches(2)
				Set b = Nothing
				i = i + 1
			End If
		Next
		Set Matches = Nothing
	End Sub
	
	Public Sub Load(ByVal f)
		s_html = LoadInc(f,"")
		Getblock(s_html)
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
		rule = "(<!--[\s]*)?" & FixRegStr(s_ms) & "#include:(.+?)" & FixRegStr(s_me) & "([\s]*-->)?"
		If Easp_Test(h,rule) Then
			If Easp.isN(p) Then
				If Instr(f,"/")>0 Then p = Left(f,InstrRev(f,"/"))
			Else
				If Instr(f,"/")>0 Then p = pa & Left(f,InstrRev(f,"/"))
			End If
			Set inc = Easp_Match(h,rule)
			For Each Match In inc
				incFile = Match.SubMatches(1)
				incStr = LoadInc(incFile, p)
				h = Replace(h,Match,incStr)
			Next
			Set inc = Nothing
		End If
		LoadInc = h
	End Function
	'加载附加模板
	Public Sub Include(ByVal tag, ByVal f)
		Dim s
		s = LoadInc(f)
		Getblock(s)
		s_html = Replace(s_html, s_ms & tag & s_me, s)
	End Sub
	
	Public Default Sub Tag(ByVal t, ByVal s)
		Dim tag,curtag,i
		If Instr(t,".")>0 Then
			tag = Split(t,".")
			curtag = tag(Ubound(tag)-1)
			If o_block.Exists(curtag) Then
'				For i = 0 To Ubound(tag)-1
'					If If Easp.isN(o_block(tag(i))) Then o_block(tag(i)) = o_block(tag(i) & "__s")
'				Next
				o_blockdata.Add t, s
				Easp.Trace(o_blockdata)
			End If
		Else
			o_data.Add t, cStr(s)
		End If
	End Sub

	Public Sub [Update](ByVal t)
		Dim i,tmp,tag,curtag,ftag
		tmp = o_block(t&"__s")
		For Each i In o_blockdata
			tag = Split(i,".")
			curtag = tag(Ubound(tag)-1)
			If curtag = t Then
				tmp = Replace(tmp, s_ms & i & s_me, o_blockdata(i))
				o_blockdata.Remove i
				If Ubound(tag)>1 Then ftag = tag(Ubound(tag)-2)
			End If
		Next
		o_block(t) = Easp.IIF(o_block(t)=o_block(t&"__s") ,tmp ,o_block(t) & tmp)
		If Easp.Has(ftag) Then
			o_block(ftag) = Easp.RegReplace(o_block(ftag),o_block(t & "__b"),o_block(t))
		End If
	End Sub
	
	Public Function MakeTag(ByVal t, ByVal f)
		Dim s,e,a,i
		Select Case Lcase(t)
			Case "css"
				s = "<link href="""
				e = """ rel=""stylesheet"" type=""text/css"" />"
			Case "js"
				s = "<scr"&"ipt type=""text/javascript"" src="""
				e = """></scr"&"ipt>"
			Case "author", "keywords", "description", "copyright"
				MakeTag = MakeTagMeta(t,f)
				Exit Function
		End Select
		a = Split(f,"|")
		For i = 0 To Ubound(a)
			a(i) = s & Trim(a(i)) & e
		Next
		MakeTag = Join(a,vbCrLf)
	End Function
	Private Function MakeTagMeta(ByVal t, ByVal s)
		MakeTagMeta = "<meta name=""" & t & """ content=""" & s & """ />"
	End Function
	
	Public Sub Show()
		Dim i,k
		If o_blockTag.Count>0 Then
			For Each i In o_blockTag
				s_html = Easp.regReplace(s_html,o_block(o_blockTag(i)&"__b"),o_block(o_blockTag(i)))
			Next
		End If
		If o_data.Count > 0 Then
			For Each k In o_data
				'Easp.WN k & " - " & Easp.HtmlEncode(o_data(k))
				s_html = Replace(s_html,s_ms&k&s_me,o_data(k))
			Next
		End If
		Easp.W s_html
	End Sub
	
	Public Sub Trace()
		Easp.wn "<br />========================"
		Easp.Trace(o_blockdata)
		Easp.wn "<br />========================"
		Easp.Trace(o_block)
		Easp.wn "<br />========================"
	End Sub
End Class
%>