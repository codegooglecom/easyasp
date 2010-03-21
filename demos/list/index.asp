<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
'Easp.Debug = False
Function testmy(ByVal s)
	testmy = "U : " & UCase(s)
End Function

Sub WN(ByVal s)
	Easp.WN s & "----" & s
End Sub

'先构造一个随机数组
Dim arrayA(19),i
For i = 0 To 19
	arrayA(i) = Easp.RandStr(Easp.Rand(0,6)&":abcdeABCDE1234567890")
Next
Dim list, Alist, list1, list2
'加载List核心
Easp.Use "List"
'创建一个List对象
Set list = Easp.List.New
'忽略大小写(去重复项、搜索、取索引值、排序、比较时)
'list.IgnoreCase = False
'-----------------------------------------------
'把数组存入List对象管理(可以用2种方法接受共4种形式的数据)
'-------------
'第1种：简单数组
'list.Data = arrayA
'第2种：用空格隔开的字符串，每个字符串会解析为数组的一个元素
'list.Data = "aa a ee ddd A AA aa Ab ab a bb b  bb c ccc  ddd b d"
'-------------
'第3种：带下标的数组，如果数组元素中包含 : 号，则会解析为Hash表值对， : 号前的字符串为Hash的下标
list.Hash = Array("test:34", "name:coldstone", 344.89, "birth:81/01/01", "Others", "btime:81/01/32", "addtime:"&True)
'第4种：用空格隔开的字符串，字符串中包含 : 号，也会把带 : 号的字符串解析为Hash表值对
'list.Hash = "aa a ee se:ddd A AA aa my:Ab ab a bb b  bb c la:ccc  ddd b d"
'-----------------------------------------------
Easp.WN "初始数组为：" & list.ToString

list("one") = "OneNumber"
list("two") = "2222"
list("six") = "SSSSix"
list.Push "wobu"
list.Push "zhidao"
list.Push -349.89
list.Push 80
list.Push "ssssix"
list.Insert 22, Array("seven","eight","nine")
Easp.WN "添加一些元素后为：" & list.ToString
list.Pop
list.Delete 4
list.Delete "two"
Easp.WN "删除一些元素后为：" & list.ToString
Easp.WN "现在数组的长度是：" & list.Size
Easp.WN "数组的有效值个数（非空值）是：" & list.Count
'去除重复元素
list.Uniq
Easp.WN "去除重复元素后为：" & list.ToString
list.Compact
Easp.WN "去除空元素后为：" & list.ToString
Easp.WN "数组的最大值是：" & list.Max
Easp.WN "数组的最小值是：" & list.Min
Easp.WN "数组的第一个元素是：" & list.First
Easp.WN "数组的最后一个元素是：" & list.Last
list.Sort
Easp.WN "排序后为：" & list.ToString
list.Reverse
Easp.WN "倒序后为：" & list.ToString
list.Rand
Easp.WN "打乱顺序后为：" & list.ToString
Easp.WN "执行迭代处理(不影响原数组)：" & list.Map_("testmy").ToString
'list.Each("WN")
Easp.WN "第一个是数字的值是：" & list.Find("isNumeric(%i)")
Easp.WN "选择所有非数字的值(不影响原数组)：" & list.Select_("Not isNumeric(%i)").ToString
Easp.WN "选择所有以数字开头的值(不影响原数组)：" & list.Grep_("^\d.+").ToString
Easp.WN "执行迭代处理后排序(不影响原数组)：" & list.SortBy_("testmy").ToString
'数组重复
'list.Times 2

Set Alist = Easp.List.New
'Alist.Hash = "aaa:ssssix b:wefewr c:sfwef one:weioid six:yesterday ee"
Alist.Data = Array("ssssix","OneNumber","zhidao",234.234,35235,3534.345)
'附加数组
'list.Splice Alist
'合并数组
'list.Merge Alist
'数组交集
'list.Inter Alist
'数组差集
'list.Diff Alist

Easp.WN "=========="
Easp.wn "---遍历现在的List---"
For i = 0 To list.End
	Easp.WN "list("&i&") 的值是：" & list(i)
Next
Easp.WN "=========="
Easp.wn "---遍历现在的List中的散列对值---"
Dim Maps,key,x,y
Set Maps = list.Maps
For Each key In Maps
	If Not isNumeric(key) Then
		Easp.WN "list(""" & key & """) = list(" & Maps(key) &  ") = " & list(key)
	End If
Next
Set Maps = Nothing
Easp.WN "=========="



'Easp.WN "是否包含字符串 bb ：" & list.Has("bb") & "，在数组中第1次出现的下标是：" & list.IndexOf("bb")
''取值
'Easp.WN "下标为3的元素为：" & list.At(3)
''赋值
'list.At(3) = "three"
''list.At(n)可以简写为 list(n)
'Easp.WN "下标为5的元素为：" & list(5)
'list(5) = "five"
''可以向超过当前最大下标的元素赋值，相当于添加元素
'list(22) = "this22"
'Easp.WN "添加和更改元素后的数组为：" & list.ToString
'Easp.WN "数组的长度（元素个数）是：" & list.Size
'Easp.WN "数组的有效长度（非空值）是：" & list.Count
''去除空元素
'list.Compact
'Easp.WN "去除所有空元素的结果是：" & list.ToString
'Easp.WN "数组的最大值是：" & list.Max
'Easp.WN "数组的最小值是：" & list.Min
'Easp.WN "数组的第一个元素是：" & list.First
'Easp.WN "数组的最后一个元素是：" & list.Last
''排序
'list.Sort
'Easp.WN "将数组排序后的结果是：" & list.ToString
''反向
'list.Reverse
'Easp.WN "将数组反向排列结果是：" & list.ToString
''随机排序
'list.Rand
'Easp.WN "将数随机排序结果是：" & list.ToString
''删除指定下标项
'list.Delete 12
'Easp.WN "删除下标为12的元素后结果是：" & list.ToString
'Set list1 = list.Clone
''因为Search方法会更改当前的List对象，所以这里复制为新list对象操作
'list1.Search("a")
'Easp.WN "数组中包含字符串 a 的元素：" & list1.ToString
'Set list1 = list.Clone
'list1.SearchNot("a")
'Easp.WN "数组中不包含字符串 a 的元素：" & list1.ToString
'Easp.C(list1)
'Easp.WN "=========="
''取得数组的其中一部分（按下标），返回一个新的List对象
''可取多个元素，用逗号隔开，可以用 - 表示范围（如2-5表示第2到第5下标，\s表示开头，\e表示结尾）
'Set arr = list.Get("1,3,7-\e")
'Easp.WN "取得下标为1,3,7-\e的新数组为：" & arr.ToString
'Easp.WN "新数组的长度是：" & arr.Length
'Easp.WN "用 | 符号连接起是：" & arr.J("|")
''删除第一个元素
'arr.Shift
''添加一个元素到开头
'arr.UnShift "first"
''删除最后一个元素
'arr.Pop
''添加一个元素到最后
'arr.Push "last"
''插入一个元素到指定下标
'arr.Insert 4, "four"
'Easp.WN "删除和添加元素后是：" & arr.ToString
''删除多个元素，用逗号隔开，可以用 - 表示范围（如2-5表示第2到第5下标，\s表示开头，\e表示结尾）
'arr.Delete "\s-2,5-\e"
'Easp.WN "删除\s-2,5-\e后是：" & arr.ToString
'Easp.C(arr)
'Easp.WN "=========="
'
'Easp.wn "---遍历现在的List---"
'For i = 0 To list.End
'	Easp.WN "list("&i&") 的值是：" & list(i)
'Next

Easp.WN "=========="
Easp.wn "---取出为普通数组(如果是Hash表就把Hash名称转换为前缀)后再遍历---"
arr = list.Hash
For i = 0 To Ubound(arr)
	Easp.WN "arr("&i&") 的值是：" & arr(i)
Next
Easp.WN "=========="
Easp.wn "---取出为普通数组后再遍历---"
arr = list.Data
For i = 0 To Ubound(arr)
	Easp.WN "arr("&i&") 的值是：" & arr(i)
Next

Easp.wn "------------------------------------"
Easp.w "页面执行时间： " & Easp.ScriptTime & " 秒"
Set list = Nothing
Set Easp = Nothing
%>