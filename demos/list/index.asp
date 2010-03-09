<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Dim list, arr

Easp.Use "List"

Set list = Easp.List.New()

list.Source = Array(42,56,53,34,839,"aaa",75,"bbb",9023,844,322,65,775)

list(15) = 2548

Easp.WN list.J(",")

arr = list.Get("1,5,9,10,11")
Easp.WN Join(arr,", ")

Easp.wn "---现在的List---"
For i = 0 To list.Size - 1
	Easp.WN list(i)
Next
Easp.wn "---转成原始数组后---"
arr = list.toArray
For i = 0 To Ubound(arr)
	Easp.WN arr(i)
Next

Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set Easp = Nothing
%>