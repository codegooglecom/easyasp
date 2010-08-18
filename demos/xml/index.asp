<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Use "Xml"

Easp.Xml.Open "cd_catalog.xml"
'Easp.WN Easp.Xml.getNodeText("CD")

Easp.WN ""
Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>