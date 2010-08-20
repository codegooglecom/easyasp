<%
'######################################################################
'## easp.xml.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp XML Document Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/08/20 13:24:30
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
		f = absPath(f)
		'读取文件
		Dom.load f
		s_filePath = f
		If Not IsErr Then
			Set Doc = Dom.documentElement
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
		If Easp.Test(s,"^([\w\d-]+>)?https?://") Then
			Easp.Use "Http"
			str = Easp.Http.Get(s)
		Else
			str = s
		End If
		Dom.loadXML(str)
		If Not IsErr Then Set Doc = Dom.documentElement
	End Sub
	
	'关闭文件
	Public Sub Close()
		Set Doc = Nothing
		s_filePath = ""
		IsOpen = False
	End Sub
		
	'TagName取对象
	Public Default Function Find(ByVal t)
		Dim o
		Set Find = New Easp_Xml_Node
		Set o = Dom.GetElementsByTagName(t)
		If o.Length = 0 Then
			Easp.Error.Msg = "("&t&")"
			Easp.Error.Raise 98
		ElseIf o.Length = 1 Then
			Find.Dom = o(0)
		Else
			Find.Dom = o
		End If
	End Function
	'XPath取对象
	Public Function [Select](ByVal p)
		Set [Select] = New Easp_Xml_Node
		[Select].Dom = Dom.selectSingleNode(p)
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
	
	Public Function [New](ByVal o)
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
		If IsNodes Then
			Set El = [New](o_node(n))
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
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		o_node.setAttribute s, v
	End Property
	Public Property Get Attr(ByVal s)
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Attr = o_node.getAttribute(s)
	End Property
	
	'文本设置
	Public Property Let Text(ByVal v)
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		o_node.Text = v
	End Property
	Public Property Get Text
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Text = o_node.Text
	End Property
	
	'文本设置
	Public Property Let Value(ByVal v)
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		o_node.ChildNodes(0).NodeValue = v
	End Property
	Public Property Get Value
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Value = o_node.ChildNodes(0).NodeValue
	End Property
	
	'(只读)
	'获取XML
	Public Property Get Xml
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Xml = o_node.Xml
	End Property
	'元素名
	Public Property Get Name
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Name = o_node.BaseName
	End Property
	'元素名
	Public Property Get [Type]
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		[Type] = o_node.NodeType
	End Property
	'元素长度
	Public Property Get Length
		If IsNodes Then
			Length = o_node.Length
		ElseIf IsNode Then 
			Length = o_node.ChildNodes.Length
		End If
	End Property
	
	'=======Xml元素属性（返回新节点元素）======
	'根元素
	Public Property Get Root
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Set Root = [New](o_node.OwnerDocument)
	End property
	'父元素
	Public Property Get Parent
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Set Parent = [New](o_node.parentNode)
	End property
	'上一同级元素
	Public Property Get Prev
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
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
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
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
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Set First = [New](o_node.FirstChild)
	End Property
	'最后一个元素
	Public Property Get Last
		If Not IsNode Then Easp.Error.Raise 97 : Exit Property
		Set Last = [New](o_node.LastChild)
	End Property
	
	'=======Xml元素方法======
	'(查找)
	'是否有某属性
	Public Function HasAttr(ByVal s)
		If Not IsNode Then Easp.Error.Raise 97 : Exit Function
		Dim oattr
		Set oattr = o_node.Attributes.GetNamedItem(s)
		HasAttr = Not oattr Is Nothing
		Set oattr = Nothing
	End Function
	'是否有子节点
	Public Function HasChild()
		If Not IsNode Then Easp.Error.Raise 97 : Exit Function
		HasChild = o_node.hasChildNodes()
	End Function
	'查找子元素
	Public Function Find(ByVal t)
		If Not IsNode Then Easp.Error.Raise 97 : Exit Function
		Dim o
		Set o = o_node.GetElementsByTagName(t)
		If o.Length = 0 Then
			Easp.Error.Msg = "("&t&")"
			Easp.Error.Raise 98
		ElseIf o.Length = 1 Then
			Set Find = [New](o(0))
		Else
			Set Find = [New](o)
		End If
	End Function
	
	'(建立)
	'克隆节点
	Public Function Clone(ByVal b)
		If Not IsNode Then Easp.Error.Raise 97 : Exit Function
		If Easp.IsN(b) Then b = True
		Set Clone = [New](o_node.CloneNode(b))
	End Function
	'添加子节点
	Public Sub Append(ByVal o)
		
	End Sub
	
	'(删除)
	'删除某属性
	Public Sub RemoveAttr(ByVal s)
		If Not IsNode Then Easp.Error.Raise 97 : Exit Sub
		o_node.removeAttribute(s)
	End Sub
	'清除所有子节点
	Public Sub Clear
		If Not IsNode Then Easp.Error.Raise 97 : Exit Sub
		o_node.Text = ""
		o_node.removeChild(o_node.FirstChild)
	End Sub
	'合并相邻的Text节点并删除空的Text节点
	Public Sub Normalize
		If Not IsNode Then Easp.Error.Raise 97 : Exit Sub
		o_node.normalize()
	End Sub
	'删除自身
	Public Sub Remove
		If Not IsNode Then Easp.Error.Raise 97 : Exit Sub
		o_node.ParentNode.RemoveChild(o_node)
	End Sub
End Class
%>