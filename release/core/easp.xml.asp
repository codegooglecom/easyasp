<%
'######################################################################
'## easp.xml.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp XML Document Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/09/30 00:03:10
'## Description :   Read and write the XML documents
'##
'######################################################################
Class EasyAsp_Xml
	Public Dom, Doc, IsOpen
	Private s_filePath, s_xsltPath
	Private Sub Class_Initialize()
		s_filePath = ""
		s_xsltPath = ""
		IsOpen = False
		Easp.Error(96) = "XML文件操作出错"
		Easp.Error(97) = "对象不支持此属性或方法"
		Easp.Error(98) = "未找到目标对象"
		Easp.Error(99) = "保存XML文档出错"
	End Sub
	Private Sub Class_Terminate()
		If IsObject(Doc) Then Set Doc = Nothing
		If IsObject(Dom) Then Set Dom = Nothing
	End Sub
	Private Function NewDom()
		Dim o
		If Easp.IsInstall("MSXML2.DOMDocument") Then
			Set o = Server.CreateObject("MSXML2.DOMDocument")
		ElseIf Easp.IsInstall("Microsoft.XMLDOM") Then
			Set o = Server.CreateObject("Microsoft.XMLDOM")
		End If
		o.PreserveWhiteSpace = True
		o.Async = False
		o.SetProperty "SelectionLanguage", "XPath"
		Set NewDom = o
		Set o = Nothing
	End Function
	Public Function Open(byVal f)
		Open = False
		If Easp.IsN(f) Then Exit Function
		Set Dom = NewDom()
		f = absPath(f)
		Dom.load f
		s_filePath = f
		If Not IsErr Then
			Set Doc = NewNode(Dom.documentElement)
			Open = True
			IsOpen = True
		Else
			Set Dom = Nothing
		End If
	End Function
	Private Function absPath(ByVal p)
		If Easp.IsN(p) Then absPath = "" : Exit Function
		If Mid(p,2,1)<>":" Then p = Server.MapPath(p)
		absPath = p
	End Function
	Public Sub Load(ByVal s)
		If Easp.IsN(s) Then Exit Sub
		Dim str
		If Easp.Test(s,"^([\w\d-]+>)?https?://") Then
			Easp.Use "Http"
			Dim h : Set h = Easp.Http.New
			str = h.Get(s)
			Set h = Nothing
		Else
			str = s
		End If
		Set Dom = NewDom()
		Dom.loadXML(str)
		If Not IsErr Then
			Set Doc = NewNode(Dom.documentElement)
		Else
			Set Dom = Nothing
		End If
	End Sub
	Public Sub Close()
		Set Doc = Nothing
		Set Dom = Nothing
		s_filePath = ""
		IsOpen = False
	End Sub
	Public Sub [Save]()
		If IsOpen Then
			Dom.Save(s_filePath)
		Else
			Easp.Error.Msg = "（文档未处于打开状态）"
			Easp.Error.Raise 99
		End If
	End Sub
	Public Sub SaveAs(ByVal p)
		Dim ch,cha,pi
		If Instr(p,">")>0 Then
			ch = Easp.CRight(p,">")
			p = Easp.CLeft(p,">")
		End If
		cha = Easp.IfHas(ch,Easp.CharSet)
		p = absPath(p)
		Set pi = Dom.CreateProcessingInstruction("xml", "version=""1.0"" encoding=""" & cha & """")
		If Dom.FirstChild.BaseName<>"xml" Then
			Dom.InsertBefore pi, Dom.FirstChild
		Else
			If Easp.Has(ch) Then Dom.ReplaceChild pi, Dom.FirstChild
		End If
		Dom.Save(p)
		Set pi = Nothing
	End Sub
	Public Sub SaveAsXHTML(ByVal p, ByVal xsl)
		Dim x,f : Set x = [New]
		If Easp.Test(xsl,"^([\w\d-]+>)?https?://") Then
			x.Load xsl
		Else
			x.Open xsl
		End If
		f = Dom.TransformNode(x.Dom)
		Easp.Use "Fso"
		Easp.Fso.CreateFile p, f
		Set x = Nothing
	End Sub
	Public Function NewNode(ByVal o)
		Set NewNode = New Easp_Xml_Node
		NewNode.Dom = o
	End Function
	Public Property Get Root
		Set Root = NewNode(Dom)
	End Property
	Public Function [New]()
		Set [New] = New EasyAsp_Xml
	End Function
	Public Default Function Find(ByVal t)
		Dim o,s
		If Easp.Test(t,"^<[\s\S]+>$") Then
		Else
			If Easp.Test(t, "[, >\[@:]") Then
				Set o = Dom.selectNodes(Easp_Xml_TransToXpath(t))
			Else
				Set o = Dom.GetElementsByTagName(t)
			End If
		End If
		If o.Length = 0 Then
			Easp.Error.Msg = "("&t&")"
			Easp.Error.Raise 98
		ElseIf o.Length = 1 Then
			Set Find = NewNode(o(0))
		Else
			Set Find = NewNode(o)
		End If
	End Function
	Public Function [Select](ByVal p)
		Set [Select] = NewNode(Dom.selectNodes(p))
	End Function
	Public Function Sel(ByVal p)
		Set Sel = NewNode(Dom.selectSingleNode(p))
	End Function
	Public Function Create(ByVal n, ByVal v)
		Dim o,p,cd
		If Instr(n," ")>0 Then
			cd = LCase(Easp.CRight(n," "))
			n = Easp.CLeft(n," ")
		End If
		If cd="comment" Then
			Set o = Dom.createComment(v)
		Else
			Set o = Dom.CreateElement(n)
			If cd = "cdata" Then
				Set p = Dom.CreateCDATASection(v)
			Else
				Set p = Dom.CreateTextNode(v)
			End If
			o.AppendChild(p)
		End If
		Set Create = NewNode(o)
		Set o = Nothing
		Set p = Nothing
	End Function
  Private Function IsErr()
		Dim s
		IsErr = False
    If Dom.ParseError.Errorcode<>0 Then
			With Dom.ParseError
				s = s & "	<ul class=""dev"">" & vbCrLf
				s = s & "		<li class=""info"">以下信息针对开发者：</li>" & vbCrLf
				s = s & "		<li>错误代码：0x" & Hex(.Errorcode) & "</li>" & vbCrLf
				If Easp.Has(.Reason) Then s = s & "		<li>错误原因：" & .Reason & "</li>" & vbCrLf
				If Easp.Has(.Url) Then s = s & "		<li>错误来源：" & .Url & "</li>" & vbCrLf
				If Easp.Has(.Line) And .Line<>0 Then s = s & "		<li>错误行号：" & .Line & "</li>" & vbCrLf
				If Easp.Has(.Filepos) And .Filepos<>0 Then s = s & "		<li>错误位置：" & .Filepos & "</li>" & vbCrLf
				If Easp.Has(.SrcText) Then s = s & "		<li>源 文 本：" & .SrcText & "</li>" & vbCrLf
				s = s & "	</ul>" & vbCrLf
			End With
			IsErr = True
			Easp.Error.Msg = s
			Easp.Error.Raise 96
    End If
  End Function
End Class
Class Easp_Xml_Node
	Private o_node
	Private Sub Class_Terminate()
		Set o_node = Nothing
	End Sub
	Private Function [New](ByVal o)
		Set [New] = New Easp_Xml_Node
		[New].Dom = o
	End Function
	Public Property Let Dom(ByVal o)
		If Not o Is Nothing Then
			Set o_node = o
		Else
			Easp.Error.Msg = "(不是有效的XML对象)"
			Easp.Error.Raise 97
		End If
	End Property
	Public Property Get Dom
		Set Dom = o_node
	End Property
	Public Default Property Get Item(ByVal n)
		If IsNodes Then
			Set Item = [New](o_node(n))
		ElseIf IsNode And n = 0 Then
			Set Item = [New](o_node)
		Else
			Easp.Error.Msg = "(不是有效的XML元素集合对象&lt;"&TypeName(o_node)&"&gt;)"
			Easp.Error.Raise 97
		End If
	End Property
	Public Property Get IsNode
		IsNode = TypeName(o_node) = "IXMLDOMElement"
	End Property
	Public Property Get IsNodes
		IsNodes = TypeName(o_node) = "IXMLDOMSelection"
	End Property
	Public Property Let Attr(ByVal s, ByVal v)
		If IsNull(v) Then RemoveAttr s : Exit Property
		If IsNode Then
			o_node.setAttribute s, v
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).setAttribute s, v
			Next
		End If
	End Property
	Public Property Get Attr(ByVal s)
		If Not IsNode Then Exit Property
		Attr = o_node.getAttribute(s)
	End Property
	Public Property Let Text(ByVal v)
		If IsNode Then
			If Easp.Has(v) Then o_node.Text = v
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				If Easp.Has(v) Then o_node(i).Text = v
			Next
		End If
	End Property
	Public Property Get Text
		If IsNode Then
			Text = o_node.Text
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				Text = Text & o_node(i).Text
			Next
		End If
	End Property
	Public Property Let Value(ByVal v)
		If IsNode Then
			o_node.ChildNodes(0).NodeValue = v
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).ChildNodes(0).NodeValue = v
			Next
		End If
	End Property
	Public Property Get Value
		If Not IsNode Then Exit Property
		Value = o_node.ChildNodes(0).NodeValue
	End Property
	Public Property Get Xml
		If IsNode Then
			Xml = o_node.Xml
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				If i>0 Then Xml = Xml & vbCrLf
				Xml = Xml & o_node(i).Xml
			Next
		End If
	End Property
	Public Property Get Name
		If Not IsNode Then Exit Property
		Name = o_node.BaseName
	End Property
	Public Property Get [Type]
		If IsNodes Then
			[Type] = 0
		Else
			[Type] = o_node.NodeType
		End If
	End Property
	Public Property Get TypeString
	If IsNodes Then
			TypeString = "selection"
		Else
			TypeString = o_node.NodeTypeString
		End If
	End Property
	Public Property Get Length
		If IsNode Then 
			Length = o_node.ChildNodes.Length
		Else
			Length = o_node.Length
		End If
	End Property
	Public Property Get Root
		If Not IsNode Then Exit Property
		Set Root = [New](o_node.OwnerDocument)
	End property
	Public Property Get Parent
		If Not IsNode Then Exit Property
		Set Parent = [New](o_node.parentNode)
	End property
	Public Property Get Child(ByVal n)
		If Not IsNode Then Exit Property
		Set Child = [New](o_node.ChildNodes(n))
	End property
	Public Property Get Prev
		If Not IsNode Then Exit Property
		Dim o
		Set o = o_node.PreviousSibling
		Do While True
			If TypeName(o) = "Nothing" Or TypeName(o) = "IXMLDOMElement" Then Exit Do
			Set o = o.PreviousSibling
		Loop
		If TypeName(o) = "IXMLDOMElement" Then
			Set [Prev] = [New](o)
			Set o = Nothing
		Else
			Easp.Error.Msg = "(没有上一同级元素)"
			Easp.Error.Raise 96
		End If
	End property
	Public Property Get [Next]
		If Not IsNode Then Exit Property
		Dim o
		Set o = o_node.NextSibling
		Do While True
			If TypeName(o) = "Nothing" Or TypeName(o) = "IXMLDOMElement" Then Exit Do
			Set o = o.NextSibling
		Loop
		If TypeName(o) = "IXMLDOMElement" Then
			Set [Next] = [New](o)
			Set o = Nothing
		Else
			Easp.Error.Msg = "(没有下一同级元素)"
			Easp.Error.Raise 96
		End If
	End property
	Public Property Get First
		If Not IsNode Then Exit Property
		Set First = [New](o_node.FirstChild)
	End Property
	Public Property Get Last
		If Not IsNode Then Exit Property
		Set Last = [New](o_node.LastChild)
	End Property
	Public Function HasAttr(ByVal s)
		If Not IsNode Then Exit Function
		Dim oattr
		Set oattr = o_node.Attributes.GetNamedItem(s)
		HasAttr = Not oattr Is Nothing
		Set oattr = Nothing
	End Function
	Public Function HasChild()
		If Not IsNode Then Exit Function
		HasChild = o_node.hasChildNodes()
	End Function
	Public Function Find(ByVal t)
		If Not IsNode Then Exit Function
		Dim o
		If Easp.Test(t, "[, >\[@:]") Then
			Set o = o_node.selectNodes(Easp_Xml_TransToXpath(t))
		Else
			Set o = o_node.GetElementsByTagName(t)
		End If
		If o.Length = 0 Then
			Easp.Error.Msg = "("&t&")"
			Easp.Error.Raise 98
		ElseIf o.Length = 1 Then
			Set Find = [New](o(0))
		Else
			Set Find = [New](o)
		End If
	End Function
	Public Function [Select](ByVal p)
		If Not IsNode Then Exit Function
		Set [Select] = [New](o_node.selectNodes(p))
	End Function
	Public Function Sel(ByVal p)
		If Not IsNode Then Exit Function
		Set Sel = [New](o_node.selectSingleNode(p))
	End Function
	Public Function Clone(ByVal b)
		If Not IsNode Then Exit Function
		If Easp.IsN(b) Then b = True
		Set Clone = [New](o_node.CloneNode(b))
	End Function
	Private Function GetNodeDom(ByVal o)
		Select Case TypeName(o)
			Case "IXMLDOMElement" Set GetNodeDom = o
			Case "Easp_Xml_Node" Set GetNodeDom = o.Dom
		End Select
	End Function
	Public Function Append(ByVal o)
		If Not IsNode Then Exit Function
		o_node.AppendChild(GetNodeDom(o))
		Set Append = [New](o_node)
	End Function
	Public Function ReplaceWith(ByVal o)
		If IsNode Then
			Call o_node.ParentNode.replaceChild(GetNodeDom(o), o_node)
		ElseIf IsNodes Then
			Dim i,n
			For i = 0 To Length - 1
				Set n = GetNodeDom(o).CloneNode(True)
				Call o_node(i).ParentNode.replaceChild(n, o_node(i))
			Next
		End If
		Set ReplaceWith = [New](o_node)
	End Function
	Public Function Before(ByVal o)
		If IsNode Then
			Call o_node.ParentNode.InsertBefore(GetNodeDom(o), o_node)
		ElseIf IsNodes Then
			Dim i,n
			For i = 0 To Length - 1
				Set n = GetNodeDom(o).CloneNode(True)
				Call o_node(i).ParentNode.InsertBefore(n, o_node(i))
			Next
		End If
		Set Before = [New](o_node)
	End Function
	Public Function After(ByVal o)
		If IsNode Then
			Call InsertAfter(GetNodeDom(o), o_node)
		ElseIf IsNodes Then
			Dim i,n
			For i = 0 To Length - 1
				Set n = GetNodeDom(o).CloneNode(True)
				Call InsertAfter(n, o_node(i))
			Next
		End If
		Set After = [New](o_node)
	End Function
	Private Sub InsertAfter(ByVal n, Byval o)
		Dim p
		Set p = o.ParentNode
		If p.LastChild Is o Then
			p.AppendChild(n)
		Else
			Call p.InsertBefore(n, o.nextSibling)
		End If
	End Sub
	Public Function RemoveAttr(ByVal s)
		If IsNode Then
			o_node.removeAttribute(s)
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).removeAttribute(s)
			Next
		End If
		Set RemoveAttr = [New](o_node)
	End Function
	Public Function [Empty]
		If IsNode Then
			o_node.Text = ""
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).Text = ""
			Next
		End If
		Set [Empty] = [New](o_node)
	End Function
	Public Function Clear
		If IsNode Then
			o_node.Text = ""
			o_node.removeChild(o_node.FirstChild)
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).Text = ""
				o_node(i).removeChild(o_node(i).FirstChild)
			Next
		End If
		Set Clear = [New](o_node)
	End Function
	Public Function Normalize
		If IsNode Then
			o_node.normalize()
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).normalize()
			Next
		End If
		Set Normalize = [New](o_node)
	End Function
	Public Sub Remove
		If IsNode Then
			o_node.ParentNode.RemoveChild(o_node)
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).ParentNode.RemoveChild(o_node(i))
			Next
		End If
	End Sub
End Class
Function Easp_Xml_TransToXpath(ByVal s)
	s = Easp.RegReplace(s, "\s*,\s*", "|//")
	s = Easp.RegReplace(s, "\s*>\s*", "/")
	s = Easp.RegReplace(s, "\s+", "//")
	s = Easp.RegReplace(s, "(\[)([a-zA-Z]+\])", "$1@$2")
	s = Easp.RegReplace(s, "(\[)([a-zA-Z]+[!]?=[^\]]+\])", "$1@$2")
	s = Easp.RegReplace(s, "(?!\[\d)\]\[", " and ")
	s = Replace(s, "|", " | ")
	Easp_Xml_TransToXpath = "//" & s
End Function
%>