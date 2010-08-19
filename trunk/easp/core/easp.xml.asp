<%
'######################################################################
'## easp.xml.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp XML Document Class
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/08/18 15:24:30
'## Description :   Read and write the XML documents
'## 暂时从网上下载了一个XML类看看有哪些通用方法，Easp的还在构造中
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
	
	Public Default Property Get Item(ByVal el)
		
		'Set [Select] = 
	End Property
	
	'XPath取对象
	Public Function [Select](ByVal p)
		Set [Select] = New Easp_Xml_Node
		[Select].Node = Dom.selectSingleNode(p)
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
	
	Public Property Let Node(ByVal o)
		If Not o Is Nothing Then
			Set o_node = o
		Else
			Easp.Error.Raise 97
		End If
	End Property
	Public Property Get Node
		Set Node = o_node
	End Property
	
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
	
	'XML获取
	Public Property Get Xml
		Xml = o_node.Xml
	End Property
End Class
%>