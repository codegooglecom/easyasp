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
		Set Dom = CreateObject("Microsoft.XMLDOM")
		Dom.PreserveWhiteSpace = True
		Dom.Async = False
		s_filePath = ""
		IsOpen = False
		Easp.Error(96) = "XML文件操作出错"
	End Sub
	
	'析构函数
	Private Sub Class_Terminate()
		If IsObject(Doc) Then Set Doc = Nothing
		Set Dom = Nothing
	End Sub
	
	'开打一个已经存在的XML文件,返回打开状态
	Function Open(byVal f)
		Open = False
		If Easp.IsN(f) Then Exit Function
		f = absPath(f)
		Dom.Load f
		s_filePath = f
		If Not IsError Then
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
	
	'关闭文件
	Sub Close()
		Set Doc = Nothing
		s_filePath = ""
		IsOpen = False
	End Sub
	
  '检查并打印错误信息
  Private Function IsError()
		IsError = False
    If Dom.ParseError.Errorcode<>0 Then
       s = "<h4>Error" & Dom.ParseError.Errorcode & "</h4>"
       s = s & "<B>Reason :</B>" & Dom.ParseError.Reason & "<br />"
       s = s & "<B>URL &nbsp; &nbsp;:</B>" & Dom.ParseError.Url & "<br />"
       s = s & "<B>Line &nbsp; :</B>" & Dom.ParseError.Line & "<br />"
       s = s & "<B>FilePos:</B>" & Dom.ParseError.Filepos & "<br />"
       s = s & "<B>srcText:</B>" & Dom.ParseError.SrcText & "<br />"
       IsError = True
			 Easp.Error.Msg = s
			 Easp.Error.Raise 96
    End If
  End Function
End Class
%>