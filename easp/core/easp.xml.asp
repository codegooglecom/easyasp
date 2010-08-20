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
		Easp.Error(97) = "目标不是有效的XML元素"
		Easp.Error(98) = "目标不是有效的XML元素集合"
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
	
	'从文本载入XML结构数据
	Public Sub Load(ByVal s)
		Dom.loadXML(s)
		If Not IsErr Then Set Doc = Dom.documentElement
	End Sub
	
	'关闭文件
	Public Sub Close()
		Set Doc = Nothing
		s_filePath = ""
		IsOpen = False
	End Sub
		
	'TagName取对象
	Public Default Function Find(ByVal p)
		Set Find = New Easp_Xml_Node
		Find.Dom = Dom.GetElementsByTagName(p)
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
			Easp.Error.Raise 97
		End If
	End Property
	Public Property Get Dom
		Set Dom = o_node
	End Property
	
	'取集合中的某一项
	Public Default Property Get El(ByVal n)
		Set El = [New](o_node(n))
	End Property
	'=======Xml元素属性（自身属性）======
	'(可读可写)
	'属性设置
	Public Property Let Attr(ByVal s, ByVal v)
		o_node.setAttribute s, v
	End Property
	Public Property Get Attr(ByVal s)
		Attr = o_node.getAttribute(s)
	End Property
	
	'文本设置
	Public Property Let Text(ByVal v)
		o_node.Text = v
	End Property
	Public Property Get Text
		Text = o_node.Text
	End Property
	
	'文本设置
	Public Property Let Value(ByVal v)
		o_node.ChildNodes(0).NodeValue = v
	End Property
	Public Property Get Value
		Value = o_node.ChildNodes(0).NodeValue
	End Property
	
	'(只读)
	'获取XML
	Public Property Get Xml
		Xml = o_node.Xml
	End Property
	'元素名
	Public Property Get Name
		Name = o_node.BaseName
	End Property
	'元素名
	Public Property Get [Type]
		[Type] = o_node.NodeType
	End Property
	'元素长度
	Public Property Get Length
		If TypeName(o_node) = "IXMLDOMSelection" Then
			Length = o_node.Length
		Else
			Length = o_node.ChildNodes.Length
		End If
	End Property
	
	'=======Xml元素属性（返回新节点元素）======
	'(只读)
	'根元素
	Public Property Get Root
		Set Root = [New](o_node.OwnerDocument)
	End property
	'父元素
	Public Property Get Parent
		Set Parent = [New](o_node.parentNode)
	End property
	'下一同级元素
	Public Property Get [Next]
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
		Set First = [New](o_node.FirstChild)
	End Property
	'最后一个元素
	Public Property Get Last
		Set Last = [New](o_node.LastChild)
	End Property
End Class
%>