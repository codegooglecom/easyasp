<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Debug = True
Easp.Use "Xml"

Easp.Xml.Open "cd_catalog.xml"
'Easp.WN Easp.Xml.Select("//Author").Attr("Sex")
Easp.Xml.Select("//CD[1]/YEAR").Attr("year") = 2009
Easp.WN Easp.HtmlEncode(Easp.Xml.Select("//CD[1]/YEAR").Xml)

Easp.WN ""
Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>