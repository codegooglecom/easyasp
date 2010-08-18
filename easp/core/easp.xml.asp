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
	Private fNode,fANode
	Private fErrInfo,fFileName,fOpen
	Public Dom
	'构造函数
	Private Sub Class_Initialize()
		Set Dom = CreateObject("Microsoft.XMLDOM")
		Dom.PreserveWhiteSpace = True
		Set fNode = Nothing
		Set fANode = Nothing
		fErrInfo = ""
		fFileName = ""
		fopen = False
	End Sub
	
	'析构函数
	Private Sub Class_Terminate()
		Set fNode = Nothing
		Set fANode = Nothing 
		Set Dom = Nothing
		fopen = False
	End Sub
	'取绝对路径
	Private Function absPath(ByVal p)
		If Easp.IsN(p) Then absPath = "" : Exit Function
		If Mid(p,2,1)<>":" Then p = Server.MapPath(p)
		absPath = p
	End Function
	
	'开打一个已经存在的XML文件,返回打开状态
	Function Open(byVal f)
		Open = False
		f = absPath(f)
		If Easp.IsN(f) Then Exit Function
		Dom.Async = False
		Dom.Load f
		fFileName = f
		If Not IsError Then
			Open = True
			fopen = True
		End If
	End Function
	
	'关闭
	Sub Close()
		Set fNode = Nothing
		Set fANode = Nothing
		fErrInfo = ""
		fFileName = ""
		fopen = False
	End Sub
	
'  '返回节点的缩进字串
'  Private Property Get TabStr(byVal Node)
'    TabStr = ""
'    If Node Is Nothing Then Exit Property
'    If not Node.parentNode Is Nothing Then TabStr = "  "&TabStr(Node.parentNode)
'  End Property
'  
'  '返回一个子节点对象,ElementOBJ为父节点,ChildNodeObj要查找的节点,IsAttributeNode指出是否为属性对象
'  Public Property Get ChildNode(byVal ElementOBJ,byVal ChildNodeObj,byVal IsAttributeNode)
'    Dim Element
'    Set ChildNode = Nothing
'        
'    If IsNull(ChildNodeObj) Then
'      If IsAttributeNode = False Then
'        Set ChildNode = fNode
'      Else
'        Set ChildNode = fANode
'      End If
'      Exit Property
'    ElseIf IsObject(ChildNodeObj) Then
'      Set ChildNode = ChildNodeObj
'      Exit Property
'    End If
'    
'    Set Element = Nothing
'    If LCase(TypeName(ChildNodeObj)) = "string" and Trim(ChildNodeObj)<>"" Then
'      If IsNull(ElementOBJ) Then 
'         Set Element = fNode
'      ElseIf LCase(TypeName(ElementOBJ)) = "string" Then
'         If Trim(ElementOBJ)<>"" Then
'           Set Element = Dom.selectSingleNode("//"&Trim(ElementOBJ))
'           If Lcase(Element.nodeTypeString) = "attribute" Then Set Element = Element.selectSingleNode("..")
'         End If
'      ElseIf IsObject(ElementOBJ) Then
'         Set Element = ElementOBJ
'      End If
'      
'      If Element Is Nothing Then
'        Set ChildNode = Dom.selectSingleNode("//"&Trim(ChildNodeObj))
'      ElseIf IsAttributeNode = True Then
'        Set ChildNode = Element.selectSingleNode("./@"&Trim(ChildNodeObj))
'      Else
'        Set ChildNode = Element.selectSingleNode("./"&Trim(ChildNodeObj))
'      End If
'    End If
'  End Property
'  
'  '读取最后的错误信息
'  Public Property Get ErrInfo
'    ErrInfo = fErrInfo
'  End Property
'
'  '给xml内容
'  Public Property Get xmlText(byVal ElementOBJ)
'    xmlText = ""
'    If fopen = False Then Exit Property
'    
'    Set ElementOBJ = ChildNode(Dom,ElementOBJ,False)
'    If ElementOBJ Is Nothing Then Set ElementOBJ = Dom
'
'    xmlText = ElementOBJ.xml
'  End Property
'  
'  '=================================================================
'  
'  '=====================================================================
'  '建立一个XML文件，RootElementName：根结点名。XSLURL：使用XSL样式地址
'  '返回根结点
'  Function Create(byVal RootElementName)',byVal XslUrl)
'    Dim PINode,RootElement
'    Set Create = Nothing
'    If (Dom Is Nothing) Or (fopen = True) Then Exit Function
'    
'    If Trim(RootElementName) = "" Then RootElementName = "Root"
'    
'    Set PINode = Dom.CreateProcessingInstruction("xml", "version=""1.0""  encoding=""UTF-8""")
'    Dom.appendChild PINode
'
'    'Set PINode = Dom.CreateProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href="""&XslUrl&"""")
'    'Dom.appendChild PINode
'
'    Set RootElement = Dom.createElement(Trim(RootElementName))
'    Dom.appendChild RootElement
'    
'    Set Create = RootElement
'    
'    fopen = True
'    set fNode = RootElement
'  End Function
'  
'  
'  '读取一个NodeOBJ的节点Text的值
'  'NodeOBJ可以是节点对象或节点名，为null就取当前默认fNode
'  Function getNodeText(byVal NodeOBJ)
'    getNodeText = ""
'    If fopen = False Then Exit Function
'    
'    Set NodeOBJ = ChildNode(null,NodeOBJ,False)
'    If NodeOBJ Is Nothing Then Exit Function
'
'    If Lcase(NodeOBJ.nodeTypeString) = "element" Then
'      set fNode = NodeOBJ
'    Else
'      set fANode = NodeOBJ
'    End If
'    getNodeText = NodeOBJ.text
'  End function
'  
'  '插入在BefelementOBJ下面一个名为ElementName，Value为ElementText的子节点。
'  'IsFirst：是否插在第一个位置；IsCDATA：说明节点的值是否属于CDATA类型
'  '插入成功就返回新插入这个节点
'  'BefelementOBJ可以是对象也可以是节点名，为null就取当前默认对象
'  Function InsertElement(byVal BefelementOBJ,byVal ElementName,byVal ElementText,byVal IsFirst,byVal IsCDATA)
'    Dim Element,TextSection,SpaceStr
'    Set InsertElement = Nothing
'    
'    If not fopen Then Exit Function
'
'    Set BefelementOBJ = ChildNode(Dom,BefelementOBJ,False)
'    If BefelementOBJ Is Nothing Then Exit Function
'    
'    Set Element = Dom.CreateElement(Trim(ElementName))
'    
'    'SpaceStr = vbCrLf&TabStr(BefelementOBJ)
'    'Set STabStr = Dom.CreateTextNode(SpaceStr)
'    
'    'If Len(SpaceStr)>2 Then  SpaceStr = Left(SpaceStr,Len(SpaceStr)-2)
'    'Set ETabStr = Dom.CreateTextNode(SpaceStr)
'    
'    If IsFirst = True Then 
'      'BefelementOBJ.InsertBefore ETabStr,BefelementOBJ.firstchild
'      BefelementOBJ.InsertBefore Element,BefelementOBJ.firstchild
'      'BefelementOBJ.InsertBefore STabStr,BefelementOBJ.firstchild
'    Else
'      'BefelementOBJ.appendChild STabStr
'      BefelementOBJ.appendChild Element
'      'BefelementOBJ.appendChild ETabStr
'    End If
'
'    If IsCDATA = True Then 
'      set TextSection = Dom.createCDATASection(ElementText)
'      Element.appendChild TextSection
'    ElseIf ElementText<>"" Then
'      Element.Text = ElementText
'    End If
'
'    Set InsertElement = Element
'    Set fNode = Element
'  End Function
'  
'  '在ElementOBJ节点上插入或修改名为AttributeName，值为：AttributeText的属性
'  '如果已经存在名为AttributeName的属性对象，就进行修改。
'  '返回插入或修改属性的Node
'  'ElementOBJ可以是Element对象或名，为null就取当前默认对象
'  Function setAttributeNode(byVal ElementOBJ,byVal AttributeName,byVal AttributeText)
'    Dim AttributeNode
'    Set setAttributeNode = Nothing
'
'    If not fopen Then Exit Function
'   
'    Set ElementOBJ = ChildNode(Dom,ElementOBJ,False)
'    If ElementOBJ Is Nothing Then Exit Function 
'   
'    Set AttributeNode = ElementOBJ.attributes.getNamedItem(AttributeName)
'    If AttributeNode Is Nothing Then 
'       Set AttributeNode = Dom.CreateAttribute(AttributeName)
'       ElementOBJ.setAttributeNode AttributeNode
'    End If
'    AttributeNode.text = AttributeText
'    
'    set fNode = ElementOBJ
'    set fANode = AttributeNode
'    Set setAttributeNode = AttributeNode
'  End Function
'  
'  '修改ElementOBJ节点的Text值，并返回这个节点
'  'ElementOBJ可以对象或对象名，为null就取当前默认对象
'  Function UpdateNodeText(byVal ElementOBJ,byVal NewElementText,byVal IsCDATA)
'    Dim TextSection
'
'    set UpdateNodeText = Nothing
'    If not fopen Then Exit Function
'    
'    Set ElementOBJ = ChildNode(Dom,ElementOBJ,False)
'    If ElementOBJ Is Nothing Then Exit Function 
'
'    If IsCDATA = True Then 
'      set TextSection = Dom.createCDATASection(NewElementText)
'      If ElementOBJ.firstchild Is Nothing Then 
'        ElementOBJ.appendChild TextSection
'      ElseIf LCase(ElementOBJ.firstchild.nodeTypeString) = "cdatasection" Then
'        ElementOBJ.replaceChild TextSection,ElementOBJ.firstchild
'      End If
'    Else
'      ElementOBJ.Text = NewElementText
'    End If
'    
'    set fNode = ElementOBJ
'    Set UpdateNodeText = ElementOBJ
'  End Function
'  
'  '返回符合testValue条件的第一个ElementNode，为null就取当前默认对象
'  Function getElementNode(byVal ElementName,byVal testValue)
'    Dim Element,regEx,baseName
'    
'    Set getElementNode = Nothing
'    If not fopen Then Exit Function
'
'    testValue = Trim(testValue)
'    Set regEx = New RegExp
'    regEx.Pattern = "^[A-Za-z]+"
'    regEx.IgnoreCase = True
'    If regEx.Test(testValue) Then testValue = "/"&testValue
'    Set regEx = Nothing
'    
'    baseName = LCase(Right(ElementName,Len(ElementName)-InStrRev(ElementName,"/",-1)))
'
'    Set Element = Dom.SelectSingleNode("//"&ElementName&testValue)
'
'    If Element Is Nothing Then
'      'Response.write ElementName&testValue
'      Set getElementNode = Nothing
'      Exit Function
'    End If
'
'    Do While LCase(Element.baseName)<>baseName
'      Set Element = Element.selectSingleNode("..")
'      If Element Is Nothing Then Exit Do
'    Loop
'        
'    If LCase(Element.baseName)<>baseName Then 
'      Set getElementNode = Nothing
'    Else
'      Set getElementNode = Element
'      If Lcase(Element.nodeTypeString) = "element" Then 
'        Set fNode = Element
'      Else
'        Set fANode = Element
'      End If
'    End If
'  End Function
'  
'  '删除一个子节点
'  Function removeChild(byVal ElementOBJ)
'    removeChild = False
'    If not fopen Then Exit Function
'
'    Set ElementOBJ = ChildNode(null,ElementOBJ,False)
'    If ElementOBJ Is Nothing Then Exit Function 
'    
'    'response.write ElementOBJ.baseName
'
'    If Lcase(ElementOBJ.nodeTypeString) = "element" Then
'      If ElementOBJ Is fNode Then set fNode = Nothing
'      If ElementOBJ.parentNode Is Nothing Then
'        Dom.removeChild(ElementOBJ)
'      Else
'        ElementOBJ.parentNode.removeChild(ElementOBJ)
'      End If
'      removeChild = True
'    End If
'  End Function
'  
'  '清空一个节点所有子节点
'  Function ClearNode(byVal ElementOBJ)
'     set ClearNode = Nothing
'     If not fopen Then Exit Function
'    
'     Set ElementOBJ = ChildNode(null,ElementOBJ,False)
'     If ElementOBJ Is Nothing Then Exit Function 
'     
'     ElementOBJ.text = ""
'     ElementOBJ.removeChild(ElementOBJ.firstchild)
'     
'     Set ClearNode = ElementOBJ
'     Set fNode = ElementOBJ
'  End Function
'
'  '删除子节点的一个属性
'  Function removeAttributeNode(byVal ElementOBJ,byVal AttributeOBJ)
'    removeAttributeNode = False
'    If not fopen Then Exit Function
'    
'    Set ElementOBJ = ChildNode(Dom,ElementOBJ,False)
'    If ElementOBJ Is Nothing Then Exit Function 
'    
'    Set AttributeOBJ = ChildNode(ElementOBJ,AttributeOBJ,True)
'    If not AttributeOBJ Is Nothing Then
'      ElementOBJ.removeAttributeNode(AttributeOBJ)
'      removeAttributeNode = True
'    End If
'  End Function
'
'  '保存打开过的文件，只要保证FileName不为空就可以实现保存
'  Function Save()
'    'On Error Resume Next
'    Save = False
'    If (not fopen) or (fFileName = "") Then Exit Function
'    
'    Dom.Save fFileName
'    Save=(not IsError)
'    If Err.number<>0 then
'      Err.clear
'      Save = False
'    End If
'  End Function
'
'  '另存为XML文件，只要保证FileName不为空就可以实现保存
'  Function SaveAs(SaveFileName)
'    'On Error Resume Next
'    SaveAs = False
'    If (not fopen) or SaveFileName = "" Then Exit Function
'    Dom.Save SaveFileName
'    SaveAs=(not IsError)
'    If Err.number<>0 then
'      Err.clear
'      SaveAs = False
'    End If
'  End Function
'
  '检查并打印错误信息
  Private Function IsError()
    If Dom.ParseError.errorcode<>0 Then
       fErrInfo = "<h1>Error"&Dom.ParseError.errorcode&"</h1>"
       fErrInfo = fErrInfo&"<B>Reason :</B>"&Dom.ParseError.reason&"<br>"
       fErrInfo = fErrInfo&"<B>URL &nbsp; &nbsp;:</B>"&Dom.ParseError.url&"<br>"
       fErrInfo = fErrInfo&"<B>Line &nbsp; :</B>"&Dom.ParseError.line&"<br>"
       fErrInfo = fErrInfo&"<B>FilePos:</B>"&Dom.ParseError.filepos&"<br>"
       fErrInfo = fErrInfo&"<B>srcText:</B>"&Dom.ParseError.srcText&"<br>"
       IsError = True
    Else
      IsError = False
    End If
  End Function
End Class
%>