<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Use "Http"
'Easp.Http.CharSet = "GBK"
'Easp.WN Easp.Http.GetData("http://easp.ambox.com/demos/http/post1.asp","POST",False,"key=这是一定的",Null,Null)

Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>