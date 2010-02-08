<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include file="../../easp/easp.asp" --><%
Easp.Debug = True
Dim txt,i,j,html
txt = "这篇文档是Easp的模板类的测试文档和示例文件"

'加载tpl核心
Easp.Use "tpl"

'允许在模板文件中使用ASP代码
'Easp.Tpl.AspEnable = True

'模板文件所在文件夹，支持绝对路径和相对路径
Easp.Tpl.FilePath = "../tpl/html/"
'模板文件中可以用{#include}标签包含无限层次的子模板，也都支持相对路径和绝对路径，请参考html文件夹内的模板文件

'如何处理未替换的标签,"keep"-保留，"remove"-移除，"comment"-转成注释
'Easp.Tpl.TagUnknown = "comment"

'模板标签的样式，默认为"{*}"，*号为标签名
'Easp.Tpl.TagMask = "{$*$}"

'加载模板
Easp.Tpl.Load "tpl.html"

'也可以用下面这种方式加载模板
'Easp.Tpl.File = "tpl.html"

'开始解析标签，MakeTag可以快速生成html标签
Easp.Tpl "author", Easp.Tpl.MakeTag("author","Coldstone, TainRay")
Easp.Tpl "keywords", Easp.Tpl.MakeTag("keywords", "EasyAsp, Easp, Version 2.2")
Easp.Tpl "description", Easp.Tpl.MakeTag("description","This is a EasyAsp TPL Sample.")

'将标签替换为副模板
Easp.Tpl.TagFile "style", "inc/style.html"

Easp.Tpl "jsfile", Easp.Tpl.MakeTag("js","html/inc.js")
Easp.Tpl "cssfile", Easp.Tpl.MakeTag("css","html/style.css")
Easp.Tpl "title", "EasyAsp 模板类测试页"
Easp.Tpl "subtitle", txt
Easp.Tpl "color", "#F60"

If Hour(Now)>=10 Then
	'追加标签内容：
	Easp.Tpl.Append "subtitle", " <small>[10点之后显示]</small>"
End If

'开始循环
For i = 1 to 3
	Easp.Tpl "A.title", "A标题" & i
	Easp.Tpl "A.addtime", Easp.DateTime(Now(),"y/mm/dd")
	'更新本次循环数据，每次循环后必须调用此方法
	Easp.Tpl.Update "A"
Next

'嵌套循环演示，嵌套可以无限层的，这是父循环
For i = 1 to 4
	Easp.Tpl "B.title", "B标题" & i & " | "
	'Demo中的 B. 这个前缀不是必须的，只是为了代码方便阅读
	Easp.Tpl "id", i+10
	Easp.Tpl "addtime", Easp.DateTime(Now()-5,"y/m/d/ h:i:s")
	'这是子循环
	For j = 20 to 23
		'替换标签
		Easp.Tpl "page.list", " "&i&">"&j
		'更新本次循环数据
		Easp.Tpl.Update "page"
	Next
	'更新本次循环数据
	Easp.Tpl.Update "B"
Next

'将替换完毕的html输出至浏览器
Easp.Tpl.Show

'或者，也可以生成静态页
'得到替换完毕的html代码
html = Easp.Tpl.GetHtml
'生成静态页
'Easp.Use "Fso"
'Call Easp.Fso.CreateFile("demo.tpl.html",html)

'释放Easp对象，Easp已载入的各核心类资源也会同时自动释放
Set Easp = Nothing
%>