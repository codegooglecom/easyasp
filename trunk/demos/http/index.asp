<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Use "Http"
'Easp.Http.CharSet = "GBK"
Easp.W Easp.Http.GetData("http://easp.ambox.com/demos/http/post.asp","POST",Null,Null,Null,"key=这是一定的")
Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>