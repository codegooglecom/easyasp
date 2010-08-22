<%
'######################################################################
'## easp.xml.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp XML Document Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/08/22 00:24:30
'## Description :   Read and write the XML documents
'## P:http://msdn.microsoft.com/en-us/library/aa924158.aspx
'## M:http://msdn.microsoft.com/en-us/library/aa926433.aspx
'######################################################################
Class EasyAsp_Xml
	Public Dom, Doc, IsOpen
	Private s_filePath
	
	'构造函数
	Private Sub Class_Initialize()
		If Easp.IsInstall("MSXML2.DOMDocument") Then
		'msxml ver 3
			Set Dom = Server.CreateObject("MSXML2.DOMDocument")
		ElseIf Easp.IsInstall("Microsoft.XMLDOM") Then
		'msxml ver 2
			Set Dom = Server.CreateObject("Microsoft.XMLDOM")
		End If
		'保留空格
		Dom.preserveWhiteSpace = True
		'异步
		Dom.async = False
		Dom.setProperty "SelectionLanguage", "XPath"
		s_filePath = ""
		IsOpen = False
		Easp.Error(96) = "XML文件操作出错"
		Easp.Error(97) = "对象不支持此属性或方法"
		Easp.Error(98) = "未找到目标对象"
	End Sub
	
	'析构函数
	Private Sub Class_Terminate()
		'释放Document
		If IsObject(Doc) Then Set Doc = Nothing
		Set Dom = Nothing
	End Sub
	
	'开打一个已经存在的XML文件,返回打开状态
	Public Function Open(byVal f)
		Open = False
		If Easp.IsN(f) Then Exit Function
		'转换为绝对路径
		f = absPath(f)
		'读取文件
		Dom.load f
		'存路径（用于保存）
		s_filePath = f
		If Not IsErr Then
			'设置根元素
			Set Doc = NewNode(Dom.documentElement)
			Open = True
			IsOpen = True
		End If
	End Function
	'取绝对路径
	Private Function absPath(ByVal p)
		If Easp.IsN(p) Then absPath = "" : Exit Function
		If Mid(p,2,1)<>":" Then p = Server.MapPath(p)
		absPath = p
	End Function
	
	'从文本或者URL载入XML结构数据
	Public Sub Load(ByVal s)
		If Easp.IsN(s) Then Exit Sub
		Dim str
		'如果是外部网址则用Http取回,如要指定编码可加在http前，例：gbk>http://....
		If Easp.Test(s,"^([\w\d-]+>)?https?://") Then
			Easp.Use "Http"
			Dim h : Set h = Easp.Http.New
			str = h.Get(s)
			Set h = Nothing
		Else
			str = s
		End If
		'从文本加载
		Dom.loadXML(str)
		'设置根元素
		If Not IsErr Then Set Doc = NewNode(Dom.documentElement)
	End Sub
	
	'关闭文件
	Public Sub Close()
		Set Doc = Nothing
		s_filePath = ""
		IsOpen = False
	End Sub
	'建立新的Easp Node对象
	Public Function NewNode(ByVal o)
		Set NewNode = New Easp_Xml_Node
		NewNode.Dom = o
	End Function
		
	'TagName取对象
	Public Default Function Find(ByVal t)
		Dim o,s
		If Easp.Test(t, "[, >\[@:]") Then
			'按简单表达式取元素
			Set o = Dom.selectNodes(Easp_Xml_TransToXpath(t))
		Else
			'从标签取元素
			Set o = Dom.GetElementsByTagName(t)
		End If
		'如果没有
		If o.Length = 0 Then
			Easp.Error.Msg = "("&t&")"
			Easp.Error.Raise 98
		'如果只有一个元素
		ElseIf o.Length = 1 Then
			Set Find = NewNode(o(0))
		'如果是元素集合
		Else
			Set Find = NewNode(o)
		End If
	End Function
	'XPath取对象集合
	Public Function [Select](ByVal p)
		Set [Select] = NewNode(Dom.selectNodes(p))
	End Function
	'XPath取单个对象
	Public Function Sel(ByVal p)
		Set Sel = NewNode(Dom.selectSingleNode(p))
	End Function
	
	'新建一个节点
	Public Function Create(ByVal n, ByVal v)
		Dim o,p,cd
		'类型可在名称中用空格隔开，例："mytag cdata", " comment"
		If Instr(n," ")>0 Then
			cd = LCase(Easp.CRight(n," "))
			n = Easp.CLeft(n," ")
		End If
		'创建注释节点
		If cd="comment" Then
			Set o = Dom.createComment(v)
		Else
			'创建节点
			Set o = Dom.CreateElement(n)
			If cd = "cdata" Then
				'创建CDATASection节点
				Set p = Dom.CreateCDATASection(v)
			Else
				'创建文本节点
				Set p = Dom.CreateTextNode(v)
			End If
			'追加到节点
			o.AppendChild(p)
		End If
		'返回新建的Node对象
		Set Create = NewNode(o)
		Set o = Nothing
		Set p = Nothing
	End Function
	
  '检查并打印错误信息
  Private Function IsErr()
		Dim s
		IsErr = False
    If Dom.ParseError.errorcode<>0 Then
			With Dom.ParseError
				s = s & "	<ul class=""dev"">" & vbCrLf
				s = s & "		<li class=""info"">以下信息针对开发者：</li>" & vbCrLf
				s = s & "		<li>错误代码：0x" & Hex(.errorcode) & "</li>" & vbCrLf
				If Easp.Has(.reason) Then s = s & "		<li>错误原因：" & .reason & "</li>" & vbCrLf
				If Easp.Has(.url) Then s = s & "		<li>错误来源：" & .url & "</li>" & vbCrLf
				If Easp.Has(.line) And .line<>0 Then s = s & "		<li>错误行号：" & .line & "</li>" & vbCrLf
				If Easp.Has(.filepos) And .filepos<>0 Then s = s & "		<li>错误位置：" & .filepos & "</li>" & vbCrLf
				If Easp.Has(.srcText) Then s = s & "		<li>源 文 本：" & .srcText & "</li>" & vbCrLf
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
	'析构
	Private Sub Class_Terminate()
		Set o_node = Nothing
	End Sub
	'建立新Node对象
	Private Function [New](ByVal o)
		Set [New] = New Easp_Xml_Node
		[New].Dom = o
	End Function
	
	'源对象
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
	
	'取集合中的某一项
	Public Default Property Get El(ByVal n)
		'如果是集合就取其中下标对应子项
		If IsNodes Then
			Set El = [New](o_node(n))
		'如果是节点且下标为0就取节点本身
		ElseIf IsNode And n = 0 Then
			Set El = [New](o_node)
		Else
			Easp.Error.Msg = "(不是有效的XML元素集合对象&lt;"&TypeName(o_node)&"&gt;)"
			Easp.Error.Raise 97
		End If
	End Property
	
	'=======Xml元素属性（自身属性）======
	'是否是元素节点
	Public Property Get IsNode
		IsNode = TypeName(o_node) = "IXMLDOMElement"
	End Property
	'是否是元素集合
	Public Property Get IsNodes
		IsNodes = TypeName(o_node) = "IXMLDOMSelection"
	End Property
	'(可读可写)
	'属性设置
	Public Property Let Attr(ByVal s, ByVal v)
		'如果值为 Null 相当于删除属性
		If IsNull(v) Then RemoveAttr s : Exit Property
		'如果是节点
		If IsNode Then
			o_node.setAttribute s, v
		'如果是集合则设置每个子节点的属性
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
	
	'文本设置
	Public Property Let Text(ByVal v)
		If IsNode Then
			If Easp.Has(v) Then o_node.Text = v
		'如果是集合则设置每个子节点的文本
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
	
	'文本设置
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
	
	'(只读)
	'获取XML
	Public Property Get Xml
		If IsNode Then
			Xml = o_node.Xml
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				Xml = Xml & o_node(i).Xml
			Next
		End If
	End Property
	'元素名
	Public Property Get Name
		If Not IsNode Then Exit Property
		Name = o_node.BaseName
	End Property
	'元素类型
	Public Property Get [Type]
		If IsNodes Then Exit Property
		[Type] = o_node.NodeType
	End Property
	'元素类型名称
	Public Property Get TypeString
		If IsNodes Then Exit Property
		TypeString = o_node.NodeTypeString
	End Property
	'元素长度
	Public Property Get Length
		If IsNode Then 
			Length = o_node.ChildNodes.Length
		Else
			Length = o_node.Length
		End If
	End Property
	
	'=======Xml元素属性（返回新节点元素）======
	'根元素
	Public Property Get Root
		If Not IsNode Then Exit Property
		Set Root = [New](o_node.OwnerDocument)
	End property
	'父元素
	Public Property Get Parent
		If Not IsNode Then Exit Property
		Set Parent = [New](o_node.parentNode)
	End property
	'子元素
	Public Property Get Child(ByVal n)
		If Not IsNode Then Exit Property
		Set Child = [New](o_node.ChildNodes(n))
	End property
	'上一同级元素
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
	'下一同级元素
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
	'第一个元素
	Public Property Get First
		If Not IsNode Then Exit Property
		Set First = [New](o_node.FirstChild)
	End Property
	'最后一个元素
	Public Property Get Last
		If Not IsNode Then Exit Property
		Set Last = [New](o_node.LastChild)
	End Property
	
	'=======Xml元素方法======
	'(查找)
	'是否有某属性
	Public Function HasAttr(ByVal s)
		If Not IsNode Then Exit Function
		Dim oattr
		Set oattr = o_node.Attributes.GetNamedItem(s)
		HasAttr = Not oattr Is Nothing
		Set oattr = Nothing
	End Function
	'是否有子节点
	Public Function HasChild()
		If Not IsNode Then Exit Function
		HasChild = o_node.hasChildNodes()
	End Function
	'查找子元素
	Public Function Find(ByVal t)
		If Not IsNode Then Exit Function
		Dim o
		If Easp.Test(t, "[, >\[@:]") Then
			'按简单表达式取元素
			Set o = o_node.selectNodes(Easp_Xml_TransToXpath(t))
		Else
			'从标签取元素
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
	'XPath取对象集合
	Public Function [Select](ByVal p)
		If Not IsNode Then Exit Function
		Set [Select] = [New](o_node.selectNodes(p))
	End Function
	'XPath取单个对象
	Public Function Sel(ByVal p)
		If Not IsNode Then Exit Function
		Set Sel = [New](o_node.selectSingleNode(p))
	End Function
	
	'(建立)
	'克隆节点
	Public Function Clone(ByVal b)
		If Not IsNode Then Exit Function
		If Easp.IsN(b) Then b = True
		Set Clone = [New](o_node.CloneNode(b))
	End Function
	'统一对象为Dom节点
	Private Function GetNodeDom(ByVal o)
		Select Case TypeName(o)
			Case "IXMLDOMElement" Set GetNodeDom = o
			Case "Easp_Xml_Node" Set GetNodeDom = o.Dom
		End Select
	End Function
	'添加子节点
	Public Function Append(ByVal o)
		If Not IsNode Then Exit Function
		o_node.AppendChild(GetNodeDom(o))
		Set Append = [New](o_node)
	End Function
	'替换节点
	Public Function ReplaceWith(ByVal o)
		If IsNode Then
			'如果是节点则直接替换（是Dom内节点会直接移动），返回被替换的节点
			Call o_node.ParentNode.replaceChild(GetNodeDom(o), o_node)
		ElseIf IsNodes Then
			'如果是集合则依次替换，是Dom内的节点不会移动而是复制
			Dim i,n
			For i = 0 To Length - 1
				Set n = GetNodeDom(o).CloneNode(True)
				Call o_node(i).ParentNode.replaceChild(n, o_node(i))
			Next
		End If
		Set ReplaceWith = [New](o_node)
	End Function
	'在节点前加入另一个节点
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
	'在节点后加入另一个节点
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
	
	'(删除)
	'删除某属性
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
	'清空所有子节点
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
	'清除所有子节点，包括空文本节点
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
	'合并相邻的Text节点并删除空的Text节点
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
	'删除自身
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
'	s = Easp.RegReplace(s, ":([a-z]+)", "[$1()]")
	s = Easp.RegReplace(s, "(?!\[\d)\]\[", " and ")
	s = Replace(s, "|", " | ")
	Easp_Xml_TransToXpath = "//" & s
End Function
%>