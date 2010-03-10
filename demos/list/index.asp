<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Dim list, arr, arr1
'Easp.Error.Defined()
Easp.Debug = True
Easp.Use "List"
'创建一个List对象
Set list = Easp.List.New()
'把数组存入List对象管理
list.Data = Array(42,56,53,34,839,"aaa",75,"bbb",9023,"",844,322,65,775)
Easp.WN "初始数组为：" & list.ToString
'添加一个元素到下标15
list(15) = 2548
Easp.WN "添加一个元素到下标15后的数组为：" & list.ToString
Easp.WN "数组的长度（元素个数）是：" & list.Length
Easp.WN "数组的有效长度（非空值）是：" & list.Count
Easp.WN "数组的最大值是：" & list.Max
Easp.WN "数组的最小值是：" & list.Min
Easp.WN "=========="
'取得数组的其中一部分（按下标），返回一个新的List对象
Set arr = list.Get("3,12,4-7")
Easp.WN "取得下标为3,12,4-7的新数组为：" & arr.ToString
Easp.WN "新数组的长度是：" & arr.Length
arr.Reverse
Easp.WN "新数组反向排列结果是：" & arr.ToString
'数组排序
arr.Sort
Easp.WN "排序后的数据用 | 符号连接起是：" & arr.J("|")
arr.Pop
arr.Push "thisisP"
Easp.WN arr.Last
Set arr = Nothing
Easp.WN "=========="
'复制List对象
Set arr1 = list.Clone
'去除所有空元素
arr1.Compact
Easp.WN "去除所有空元素的数组：" & arr1.ToString
Set arr1 = Nothing


Easp.wn "---现在的List---"
For i = 0 To list.Size - 1
	Easp.WN list(i)
Next
Easp.wn "---取出为普通数组后---"
arr = list.Data
For i = 0 To Ubound(arr)
	Easp.WN arr(i)
Next

Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set list = Nothing
Set Easp = Nothing
%>