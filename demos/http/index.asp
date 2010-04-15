<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Use "Http"
Dim http, tmp
Set http = Easp.Http.New()
'直接获取页面源码
tmp = http.Get("http://easp.lengshi.com/docs/easp.db._conn.html")

'通过属性配置
Set http = Easp.Http.New()
http.Url = ""
http.Method = ""
http.Async = False
http.Charset = "gb2312"
http.Data = ""
http.Data = Array("a:one", "b:two")
http.User = ""
http.Password = ""

Easp.W tmp

Set http = Nothing
Easp.WN ""
Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>