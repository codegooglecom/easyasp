<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Debug = True
Easp.Use "Xml"
Dim str,n,i
'Easp.Xml.Open "cd_catalog.xml"
'Easp.WN Easp.Xml.Select("//Author").Attr("Sex")
'Set n = Easp.Xml.Select("//site")
'n.Attr("year") = 2009
'Easp.WN "n.Name => " & n.Name
'Easp.WNH "n.Xml => " & n.Xml
'Easp.WNH "n.Text => " & n.Text
'Easp.WN "n.Length => " & n.Length
'Easp.WN TypeName(n.Last.Dom.childNodes)
'Easp.WNH n.Last.Length

str = 			"<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
str = str & "<microblog>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Tencent"">腾讯微博</name>" & vbCrLf
str = str & "		<url>http://t.qq.com</url>" & vbCrLf
str = str & "		<account>@lengshi</account>" & vbCrLf
str = str & "		<last><![CDATA[今天我们这里下<em>大雨</em>啦！]]></last>" & vbCrLf
str = str & "	</site>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Sina"">新浪微博</name>" & vbCrLf
str = str & "		<url>http://t.sina.com.cn</url>" & vbCrLf
str = str & "		<account>@tainray</account>" & vbCrLf
str = str & "		<last><![CDATA[是不是<font color=""red"">这样</font>的噢，我也不知道哈。<img src=""http://bbs.lengshi.com/max-assets/icon-emoticon/12.gif"" />]]></last>" & vbCrLf
str = str & "	</site>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Twitter"">Twitter</name>" & vbCrLf
str = str & "		<url>http://twitter.com</url>" & vbCrLf
str = str & "		<account>@lengshi</account>" & vbCrLf
str = str & "		<last><![CDATA[I don't need this feature <strong>(>_<)</strong> any more.]]></last>" & vbCrLf
str = str & "	</site>" & vbCrLf
str = str & "</microblog>"
Easp.Xml.Load str
'Easp.WN TypeName(Easp.Xml.Dom.GetElementsByTagName("site"))
'Easp.WN Easp.Xml("last")(2).Value
Set n = Easp.Xml("last")
For i = 0 To n.Length-1
	Easp.WN n(i).Type
	Easp.WN n(i).Value
Next
Easp.WN n(1).Root.Type
Easp.WN n(2).Parent.Name
Easp.WN TypeName(n(1).Dom)
Easp.WN n(0).Next.Name
Set n = Nothing
Easp.WN ""
Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>