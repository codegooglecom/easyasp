<%
Option Explicit
'######################################################################
'## easp.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyAsp Class
'## Version     :   v2.2 alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2009/12/1 0:02:32
'## Description :   EasyAsp Class
'##
'## Update Info	:
'    1. 修改Easp.CutString为Easp.CutStr，Easp.GetCookie为Easp.Cookie；
'    2. 增加Easp.Str和Easp.WStr输出字符串；
'    3. 增加Easp.JsCode方法，返回生成的javascript代码字符串；
'    4. 增加Easp.Rewrite和Easp.RewriteRule方法，用于伪Rewrite的实现；
'    5. 增加Easp.Get和Easp.Post方法，可全面取代Easp.R系列函数，更加安全；
'    6. 增加Easp.Use方法，用于引用Easp的官方类库，如Easp.Aes、Easp.Fso、
'       Easp.Upload等，此方法为动态加载，可多次调用但只引用一次文件；
'    7. 增加Easp.MD5和Easp.MD5_16方法，用于Md5加密，此方法为动态加载文件；
'    8. 增加Easp.CLeft和Easp.CRight方法，用于取特殊字符隔开的左右字符串；
'    9. 修改Easp.IfThen方法，现在只有两个参数，用于条件为真的赋值；
'   10. 增加Easp.Ext方法，用于动态载入和使用Easp的插件；
'   11. 优化Easp.isN方法，增加了判断Recordset和Dictionary是否为空；
'   12. 增加Easp.Has方法，用于判断对象是否不为空，与Easp.isN刚好相反；
'   13. 增加Easp.Aes类，用于对中英文字符串的AES算法加密，可使用中文密码(钥)；
'   14. 优化Easp.Cookie/SetCookie，可对cookie按AES算法加密，防伪造；
'       同时方法参数有所变化，原来的分隔符:更改为>，且支持Easp.Get的参数方式；
'   15. 新增Easp.Fso类，用于FSO文件操作，功能非常全面和易于使用；
'   16. 优化Easp.GetUrlWith方法，可以将参数带到其它页面；
'   17. 优化Easp.CheckForm方法，rule规则如果以:开头，并用||隔开，则可以验
'       证多个表示"或"关系的规则项，符合其中任意一个规则则验证通过；
'   18. 优化Easp.JsEncode方法，会对双字节字符进行编码，更加严谨且无乱码问题；
'######################################################################
Dim Easp_Timer : Easp_Timer = Timer()
Dim Easp : Set Easp = New EasyASP
Dim EasyAsp_s_html
%>
<!--#include file="easp.config.asp"-->
<%
Class EasyAsp
	Public db,fso,upload,tpl,aes,[error]
	Private s_path, s_plugin, s_fsoName, s_dicName, s_charset,s_rq
	Private o_md5, o_rwt, o_ext
	Private b_cooen, i_rule
	Private Sub Class_Initialize()
		s_path		= "/easp/"
		s_plugin	= "/easp/plugin/"
		s_fsoName	= "Scripting.FileSystemObject"
		s_dicName	= "Scripting.Dictionary"
		s_charset	= "GB2312"
		s_rq		= Request.QueryString()
		i_rule		= 1
		b_cooen		= True
		Set o_rwt 	= Server.CreateObject(s_dicName)
		Set o_ext 	= Server.CreateObject(s_dicName)
		Set [error]	= New EasyAsp_Error
		Set db		= New EasyAsp_db
		Set o_md5	= New EasyAsp_obj
		Set fso		= New EasyAsp_obj
		Set upload	= New EasyAsp_obj
		Set tpl		= New EasyAsp_obj
		Set aes		= New EasyAsp_obj
	End Sub
	Private Sub Class_Terminate()
		Set aes		= Nothing
		Set tpl		= Nothing
		Set upload	= Nothing
		Set fso		= Nothing
		Set o_md5	= Nothing
		Set db 		= Nothing
		Set [error]	= Nothing
		ClearExt() : Set o_ext	= Nothing
		Set o_rwt	= Nothing
	End Sub
	Public Property Let basePath(ByVal p)
		p = IIF(Left(p,1)= "/", p, "/" & p)
		p = IIF(Right(p,1)="/", p, p & "/")
		s_path = p
	End Property
	Public Property Get basePath()
		basePath = s_path
	End Property
	Public Property Let pluginPath(ByVal p)
		s_plugin = p
	End Property
	Public Property Get pluginPath()
		pluginPath = s_plugin
	End Property
	Public Property Let fsoName(ByVal s)
		s_fsoName = s
	End Property
	Public Property Get fsoName()
		fsoName = s_fsoName
	End Property
	Public Property Let [CharSet](ByVal s)
		s_charset = Ucase(s)
	End Property
	Public Property Get [CharSet]()
		[CharSet] = s_charset
	End Property
	Public Property Let CookieEncode(ByVal b)
		b_cooen = b
	End Property
	Public Property Get CookieEncode()
		CookieEncode = b_cooen
	End Property

	Private Function rqsv(ByVal s)
		rqsv = Request.ServerVariables(s)
	End Function
	
	'输出字符串(简易断点调试)
	Sub W(ByVal s)
		Response.Write(s)
	End Sub
	Sub WC(ByVal s)
		W(s & VbCrLf)
	End Sub
	Sub WN(ByVal s)
		W(s & "<br />" & VbCrLf)
	End Sub
	Sub WE(ByVal s)
		W(s)
		Response.End()
	End Sub
	'生成动态字符串
	Function Str(ByVal s, ByVal v)
		Dim i
		s = Replace(s,"\\",Chr(0))
		s = Replace(s,"\{",Chr(1))
		If isArray(v) Then
			For i = 0 To Ubound(v)
				s = Replace(s,"{"&(i+1)&"}",v(i))
			Next
		Else
			s = Replace(s,"{1}",v)
		End If
		s = Replace(s,Chr(1),"{")
		Str = Replace(s,Chr(0),"\")
	End Function
	'输出动态字符串
	Sub WStr(ByVal s, ByVal v)
		W Str(s,v)
	End Sub
	'服务器端跳转
	Sub RR(ByVal u)
		Response.Redirect(u)
	End Sub
	'判断是否为空值
	Function isN(ByVal s)
		isN = Easp_isN(s)
	End Function
	Function Has(ByVal s)
		Has = Not Easp_isN(s)
	End Function
	'判断三元表达式
	Function IIF(ByVal Cn, ByVal T, ByVal F)
		IIF = Easp_IIF(Cn,T,F)
	End Function
	Function IfThen(ByVal Cn, ByVal T)
		IfThen = Easp_IIF(Cn,T,"")
	End Function
	'服务器端输出javascript
	Sub Js(ByVal s)
		W JsCode(s)
	End Sub
	Function JsCode(ByVal s)
		JsCode = Str("<{1} type=""text/java{1}"">{2}{3}{4}{2}</{1}>{2}", Array("sc"&"ript",vbCrLf,vbTab,s))
	End Function
	'服务器端输出javascript弹出消息框并返回前页
	Sub Alert(ByVal s)
		WE JsCode(Str("alert('{1}');history.go(-1);",JsEncode(s)))
	End Sub
	'服务器端输出javascript弹出消息框并转到URL
	Sub AlertUrl(ByVal s, ByVal u)
		WE JsCode(Str("alert('{1}');location.href='{2}';",Array(JsEncode(s),u)))
	End Sub
	'服务器端输出javascript确认消息框并根据选择转到URL
	Sub ConfirmUrl(ByVal s, ByVal tu, ByVal fu)
		WE JsCode(Str("if(confirm('{1}')){{4}='{2}';}else{{4}='{3}';}",Array(JsEncode(s),tu,fu,"location.href")))
	End Sub
	'处理字符串中的Javascript特殊字符
	Function JsEncode(ByVal s)
		JsEncode = Easp_JsEncode(s)
	End Function
	'特殊字符编码
	Function Escape(ByVal s)
		Escape = Easp_Escape(s)
	End Function
	'特殊字符解码
	Function UnEscape(ByVal s)
		UnEscape = Easp_UnEscape(s)
	End Function
	'格式化日期时间
	Function DateTime(ByVal iTime, ByVal iFormat)
		If Not IsDate(iTime) Then DateTime = "Date Error" : Exit Function
		If Instr(",0,1,2,3,4,",","&iFormat&",")>0 Then DateTime = FormatDateTime(iTime,iFormat) : Exit Function
		Dim diffs,diffd,diffw,diffm,diffy,dire,before,pastTime
		Dim iYear, iMonth, iDay, iHour, iMinute, iSecond,iWeek,tWeek
		Dim iiYear, iiMonth, iiDay, iiHour, iiMinute, iiSecond,iiWeek
		Dim iiiWeek, iiiMonth, iiiiMonth
		Dim SpecialText, SpecialTextRe,i,t
		iYear = right(Year(iTime),2) : iMonth = Month(iTime) : iDay = Day(iTime)
		iHour = Hour(iTime) : iMinute = Minute(iTime) : iSecond = Second(iTime)
		iiYear = Year(iTime) : iiMonth = right("0"&Month(iTime),2)
		iiDay = right("0"&Day(iTime),2) : iiHour = right("0"&Hour(iTime),2)
		iiMinute = right("0"&Minute(iTime),2) : iiSecond = right("0"&Second(iTime),2)
		tWeek = Weekday(iTime)-1 : iWeek = Array("日","一","二","三","四","五","六")
		If isDate(iFormat) or isN(iFormat) Then
			If isN(iFormat) Then : iFormat = Now() : pastTime = true : End If
			dire = "后" : If DateDiff("s",iFormat,iTime)<0 Then : dire = "前" : before = True : End If
			diffs = Abs(DateDiff("s",iFormat,iTime))
			diffd = Abs(DateDiff("d",iFormat,iTime))
			diffw = Abs(DateDiff("ww",iFormat,iTime))
			diffm = Abs(DateDiff("m",iFormat,iTime))
			diffy = Abs(DateDiff("yyyy",iFormat,iTime))
			If diffs < 60 Then DateTime = "刚刚" : Exit Function
			If diffs < 1800 Then DateTime = Int(diffs\60) & "分钟" & dire : Exit Function
			If diffs < 2400 Then DateTime = "半小时"  & dire : Exit Function
			If diffs < 3600 Then DateTime = Int(diffs\60) & "分钟" & dire : Exit Function
			If diffs < 259200 Then
				If diffd = 3 Then DateTime = "3天" & dire & " " & iiHour & ":" & iiMinute : Exit Function
				If diffd = 2 Then DateTime = IIF(before,"前天 ","后天 ") & iiHour & ":" & iiMinute : Exit Function
				If diffd = 1 Then DateTime = IIF(before,"昨天 ","明天 ") & iiHour & ":" & iiMinute : Exit Function
				DateTime = Int(diffs\3600) & "小时" & dire : Exit Function
			End If
			If diffd < 7 Then DateTime = diffd & "天" & dire & " " & iiHour & ":" & iiMinute : Exit Function
			If diffd < 14 Then
				If diffw = 1 Then DateTime = IIF(before,"上星期","下星期") & iWeek(tWeek) & " " & iiHour & ":" & iiMinute : Exit Function
				If Not pastTime Then DateTime = diffd & "天" & dire : Exit Function
			End If
			If Not pastTime Then
				If diffd < 31 Then
					If diffm = 2 Then DateTime = "2个月" & dire : Exit Function
					If diffm = 1 Then DateTime = IIF(before,"上个月","下个月") & iDay & "日" : Exit Function
					DateTime = diffw & "星期" & dire : Exit Function
				End If
				If diffm < 36 Then
					If diffy = 3 Then DateTime = "3年" & dire : Exit Function
					If diffy = 2 Then DateTime = IIF(before,"前年","后年") & iMonth & "月" : Exit Function
					If diffy = 1 Then DateTime = IIF(before,"去年","明年") & iMonth & "月" : Exit Function
					DateTime = diffm & "个月" & dire : Exit Function
				End If
				DateTime = diffy & "年" & dire : Exit Function
			Else
				iFormat = "yyyy-mm-dd hh:ii"
			End If
		End If
		iiWeek = Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
		iiiWeek = Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
		iiiMonth = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
		iiiiMonth = Array("January","February","March","April","May","June","July","August","September","October","November","December")
		SpecialText = Array("y","m","d","h","i","s","w")
		SpecialTextRe = Array(Chr(0),Chr(1),Chr(2),Chr(3),Chr(4),Chr(5),Chr(6))
		For i = 0 To 6 : iFormat = Replace(iFormat,"\"&SpecialText(i), SpecialTextRe(i)) : Next
		t = Replace(iFormat,"yyyy", iiYear) : t = Replace(t, "yyy", iiYear)
		t = Replace(t, "yy", iYear) : t = Replace(t, "y", iiYear)
		t = Replace(t, "mmmm", iiiiMonth(iMonth-1)) : t = Replace(t, "mmm", iiiMonth(iMonth-1))
		t = Replace(t, "mm", iiMonth) : t = Replace(t, "m", iMonth)
		t = Replace(t, "dd", iiDay) : t = Replace(t, "d", iDay)
		t = Replace(t, "hh", iiHour) : t = Replace(t, "h", iHour)
		t = Replace(t, "ii", iiMinute) : t = Replace(t, "i", iMinute)
		t = Replace(t, "ss", iiSecond) : t = Replace(t, "s", iSecond)
		t = Replace(t, "www", iiiWeek(tWeek)) : t = Replace(t, "ww", iiWeek(tWeek))
		t = Replace(t, "w", iWeek(tWeek))
		For i = 0 To 6 : t = Replace(t, SpecialTextRe(i),SpecialText(i)) : Next
		DateTime = t
	End Function
	Sub RewriteRule(ByVal s, ByVal u)
		If (Left(s,3)<>"^\/" And Left(s,2)<>"\/") Or Left(u,1)<>"/" Then Exit Sub
		o_rwt.Add ("rule" & i_rule), Array(s,u)
		i_rule = i_rule + 1
	End Sub
	Sub Rewrite(ByVal p, ByVal s, Byval u)
		Dim rp
		If Left(s,1) = "^" Or Right(s,1) = "$" Then Exit Sub
		If Left(p,1) <> "/" Then Exit Sub
		rp = Replace(p,"/","\/")
		s = "^" & rp & "\?" & s & "$"
		u = p & "?" & u
		o_rwt.Add ("rule" & i_rule), Array(s,u)
		i_rule = i_rule + 1
	End Sub
	'获取QueryString值，支持取Rewrite值
	Function [Get](Byval s)
		Dim tmp, isRwt, url, rule, i, qs, arrQs, t
		isRwt = False : url = GetUrl(1)
		If Instr(s,":")>0 Then
		'如果有类型参数，则取出为t
			t = CRight(s,":") : s = CLeft(s,":")
		End If
		If Has(o_rwt) Then
		'如果有Rewrite规则，则检测匹配否
			For Each i In o_rwt
				rule = o_rwt(i)(0)
				If Easp_Test(url,rule) Then
					qs = CRight(o_rwt(i)(1),"?")
					isRwt = True
					Exit For
				End If
			Next
		End If
		If isRwt Then
		'如果是Rewrite的页面地址
			arrQs = Split(qs,"&")
			For i = 0 To Ubound(arrQs)
				If s = CLeft(arrQs(i),"=") Then
					tmp = RegReplace(url,rule,CRight(arrQs(i),"="))
					Exit For
				End If
			Next
		Else
		'否则直接取QueryString
			tmp = Request.QueryString(s)
		End If
		[Get] = Safe(tmp,t)
	End Function
	'取Form值
	Function Post(ByVal s)
		Dim t,tmp
		If Instr(s,":")>0 Then
			t = CRight(s,":") : s = CLeft(s,":")
		End If
		tmp = Request.Form(s)
		Post = Safe(tmp,t)
	End Function
	'安全获取值新版
	Function Safe(ByVal s, ByVal t)
		Dim spl,d,l,li,i,tmp,arr() : l = False
		'如果类型中有默认值
		If Instr(t,":")>0 Then
			d = CRight(t,":") : t = CLeft(t,":")
		End If
		If Instr(",sa,da,na,", "," & Left(LCase(t),2) & ",")>0 Then
			'如果有分隔符且要警告
			If Len(t)>2 Then
				spl = Mid(t,3) : t = LCase(Left(t,2)) : l = True
			End If
		ElseIf Instr("sdn", Left(LCase(t),1))>0 Then
			'如果有分隔符且不警告
			If Len(t)>1 Then
				spl = Mid(t,2) : t = LCase(Left(t,1)) : l = True
			End If
		ElseIf Has(t) Then
			'仅有分隔符无类型
			spl = t : t = "" : l = True
		End If
		li = Split(s,spl)
		If l Then Redim arr(Ubound(li))
		For i = 0 To Ubound(li)
			If i<>0 Then tmp = tmp & spl
			Select Case t
				Case "s","sa"
				'字符串类型
					If isN(li(i)) Then li(i) = d
					tmp = tmp & Replace(li(i),"'","''")
					If l Then arr(i) = Replace(li(i),"'","''")
				Case "d","da"
				'日期类型
					If t = "da" Then
						If Not isDate(li(i)) And Has(li(i)) Then Alert("不正确的日期值！")
					End If
					tmp = IIF(isDate(li(i)), tmp & li(i), tmp & d)
					If l Then arr(i) = IIF(isDate(li(i)), li(i), d)
				Case "n","na"
				'数字类型
					If t = "na" Then
						If Not isNumeric(li(i)) And Has(li(i)) Then Alert("不正确的数值！")
					End If
					tmp = IIF(isNumeric(li(i)), tmp & li(i), tmp & d)
					If l Then arr(i) = IIF(isNumeric(li(i)), li(i), d)
				Case Else
				'未指定类型则不处理
					tmp = IIF(isN(li(i)), tmp & d, tmp & li(i))
					If l Then arr(i) = IIF(isN(li(i)), d, li(i))
			End Select
		Next
		Safe = IIF(l,arr,tmp)
	End Function
	'检查提交数据来源
	Function CheckDataFrom()
		Dim v1, v2
		CheckDataFrom = False
		v1 = Cstr(rqsv("HTTP_REFERER"))
		v2 = Cstr(rqsv("SERVER_NAME"))
		If Mid(v1,8,Len(v2)) = v2 Then
			CheckDataFrom = True
		End If
	end Function
	Sub CheckDataFromA()
		If Not CheckDataFrom Then alert "禁止从站点外部提交数据！"
	end Sub
	'截取长字符串左边部分并以特殊符号代替
	Function CutStr(ByVal s, ByVal strlen)
		Dim l,t,i,j,d,f,n
		s = Replace(s,vbCrLf,"")
		l = len(s) : t = 0 : d = "…" : f = Easp_Param(strlen)
		If Has(f(1)) Then
			strlen = Int(f(0)) : d = f(1) : f = ""
		End If
		For j = 1 To Len(d)
			n = IIF(Abs(Ascw(Mid(d,j,1)))>255, n+2, n+1)
		Next
		strlen = strlen - n
		For i = 1 to l
			t = IIF(Abs(Ascw(Mid(s,i,1)))>255, t+2, t+1)
			If t >= strlen Then
				f = Left(s,i) & d
				Exit For
			Else
				f = s
			End If
		Next
		CutStr = f
	End Function
	'取字符隔开的左段
	Function CLeft(ByVal s, ByVal m)
		CLeft = Easp_LR(s,m,0)
	End Function
	'取字符隔开的右段
	Function CRight(ByVal s, ByVal m)
		CRight = Easp_LR(s,m,1)
	End Function
	'获取当前文件的地址
	Function GetUrl(param)
		Dim script_name,url,dir
		Dim out,qitem,qtemp,i,hasQS,qstring
		script_name = rqsv("SCRIPT_NAME")
		url = script_name
		dir  = Left(script_name,InstrRev(script_name,"/"))
		If isN(param) or param = "-1" Then
			Dim ustart,uport
			If rqsv("HTTPS")="on" Then
				ustart = "https://"
				uport = IIF(Int(rqsv("SERVER_PORT"))=443,"",":"&rqsv("SERVER_PORT"))
			Else
				ustart = "http://"
				uport = IIF(Int(rqsv("SERVER_PORT"))=80,"",":"&rqsv("SERVER_PORT"))
			End If
			url = ustart & rqsv("SERVER_NAME") & uport
			If isN(param) Then
				url = url & script_name
			Else
				GetUrl = url : Exit Function
			End If
			If Has(s_rq) Then url = url & "?" & s_rq
			GetUrl = url : Exit Function
		End If
		If param = "0" Then : GetUrl = url : Exit Function
		If param = "2" Then : GetUrl = dir : Exit Function
		If InStr(param,":")>0 Then
			url = dir
			out = Mid(param,2)
			hasQS = IIF(isN(out),0,1)
		Else
			out = param : hasQS = 1
		End If
		If Has(s_rq) Then
			If param="1" Or hasQS = 0 Then
				url = url & "?" & s_rq
			Else
				qtemp = "" : i = 0 : out = ","&out&","
				qstring = IIF(InStr(out,"-")>0,"Not InStr(out,"",-""&qitem&"","")>0","InStr(out,"",""&qitem&"","")>0")
				For Each qitem In Request.QueryString()
					If Eval(qstring) Then
						If i<>0 Then qtemp = qtemp & "&"
						qtemp = qtemp & qitem & "=" & Request.QueryString(qitem)
						i = i + 1
					End If
				Next
				If Has(qtemp) Then url = url & "?" & qtemp
			End If
		End If
		GetUrl = url
	End Function
	'获取本页URL地址并带上新的URL参数
	Function GetUrlWith(ByVal p, ByVal v)
		Dim u,s,n
		s = IIF(p=-1,GetUrl(-1)&"/","")
		s = IIF(IsN(p),GetUrl(""),GetUrl(0))
		If Instr(p,":")>0 Then
			If Has(CLeft(p,":")) Then
				n = Cleft(p,":") : p = CRight(p,":")
			End If
		End If
		u = GetUrl(p)
		If Left(p,1)=":" Then s = Left(u,InstrRev(u,"/"))
		u = u & IfThen(Has(v),IIF(isN(Mid(u,len(s)+1)),"?","&") & v)
		If Has(n) Then
			If Instr(u,"?")>0 Then
				u = n & Mid(u,Instr(u,"?"))
			Else
				u = n
			End If
		End If
		GetUrlWith = u
	End Function
	'获取用户IP地址
	Function GetIP()
		Dim addr, x, y
		x = rqsv("HTTP_X_FORWARDED_FOR")
		y = rqsv("REMOTE_ADDR")
		addr = IIF(isN(x) or lCase(x)="unknown",y,x)
		If InStr(addr,".")=0 Then addr = "0.0.0.0"
		GetIP = addr
	End Function
	'仅格式化HTML文本（可带HTML标签）
	Function HtmlFormat(ByVal s)
		If Has(s) Then
			Dim m : Set m = RegMatch(s, "<([^>]+)>")
			For Each Match In m
				 s = Replace(s, Match.SubMatches(0), regReplace(Match.SubMatches(0), "\s+", Chr(0)))
			Next
			Set m = Nothing
			s = Replace(s, Chr(32), "&nbsp;")
			s = Replace(s, Chr(9), "&nbsp;&nbsp; &nbsp;")
			s = Replace(s, Chr(0), " ")
			s = regReplace(s, "(<[^>]+>)\s+", "$1")
			s = Replace(s, vbCrLf, "<br />")
		End If
		HtmlFormat = s
	End Function
	'HTML加码函数
	Function HtmlEncode(ByVal s)
		If Has(s) Then
			s = Replace(s, Chr(38), "&#38;")
			s = Replace(s, "<", "&lt;")
			s = Replace(s, ">", "&gt;")
			s = Replace(s, Chr(39), "&#39;")
			s = Replace(s, Chr(32), "&nbsp;")
			s = Replace(s, Chr(34), "&quot;")
			s = Replace(s, Chr(9), "&nbsp;&nbsp; &nbsp;")
			s = Replace(s, vbCrLf, "<br />")
		End If
		HtmlEncode = s
	End Function
	'HTML解码函数
	Function HtmlDecode(ByVal s)
		If Has(s) Then
			s = regReplace(s, "<br\s*/?\s*>", vbCrLf)
			s = Replace(s, "&nbsp;&nbsp; &nbsp;", Chr(9))
			s = Replace(s, "&quot;", Chr(34))
			s = Replace(s, "&nbsp;", Chr(32))
			s = Replace(s, "&#39;", Chr(39))
			s = Replace(s, "&apos;", Chr(39))
			s = Replace(s, "&gt;", ">")
			s = Replace(s, "&lt;", "<")
			s = Replace(s, "&amp;", Chr(38))
			s = Replace(s, "&#38;", Chr(38))
			HtmlDecode = s
		End If
	End Function
	'过滤HTML标签
	Function HtmlFilter(ByVal s)
		s = regReplace(s,"<[^>]+>","")
		s = Replace(s, ">", "&gt;")
		HtmlFilter = Replace(s, "<", "&lt;")
	End Function
	'精确到毫秒的脚本执行时间
	Function GetScriptTime(t)
		If t = "" Or t = "0" Then t = Easp_Timer
		GetScriptTime = FormatNumber((Timer()-t)*1000, 2, -1)
	End Function
	'取指定长度的随机字符串
	Function RandStr(ByVal f)
		RandStr = Easp_RandStr(f)
	End Function
	'取一个随机数
	Function Rand(ByVal n, ByVal m)
		Rand = Easp_Rand(n,m)
	End Function
	'格式化数字
	Function toNumber(ByVal n, ByVal d)
		toNumber = FormatNumber(n,d,-1)
	End Function
	'将数字转换为货币格式
	Function toPrice(ByVal n)
		toPrice = FormatCurrency(n,2,-1,0,-1)
	End Function
	'将数字转换为百分比格式
	Function toPercent(ByVal n)
		toPercent = FormatPercent(n,2,-1)
	End Function
	'关闭对象并释放资源
	Sub C(ByRef o)
		On Error Resume Next
		o.Close() : Set o = Nothing
		Err.Clear()
	End Sub
	'不缓存页面信息
	Sub noCache()
		Response.Buffer = True
		Response.Expires = 0
		Response.ExpiresAbsolute = Now() - 1
		Response.CacheControl = "no-cache"
		Response.AddHeader "Expires",Date()
		Response.AddHeader "Pragma","no-cache"
		Response.AddHeader "Cache-Control","private, no-cache, must-revalidate"
	End Sub
	'设置一个Cookies值
	Sub SetCookie(ByVal cooName, ByVal cooValue, ByVal cooCfg)
		Dim n,i,cExp,cDomain,cPath,cSecure
		If isArray(cooCfg) Then
			For i = 0 To Ubound(cooCfg)
				If isDate(cooCfg(i)) Then
					cExp = cDate(cooCfg(i))
				ElseIf Test(cooCfg(i),"int") Then
					If cooCfg(i)<>0 Then cExp = Now()+Int(cooCfg(i))/60/24
				ElseIf Test(cooCfg(i),"domain") or Test(cooCfg(i),"ip") Then
					cDomain = cooCfg(i)
				ElseIf Instr(cooCfg(i),"/")>0 Then
					cPath = cooCfg(i)
				ElseIf cooCfg(i)="True" or cooCfg(i)="False" Then
					cSecure = cooCfg(i)
				End If
			Next
		Else
			If isDate(cooCfg) Then
				cExp = cDate(cooCfg)
			ElseIf Test(cooCfg,"int") Then
				If cooCfg<>0 Then cExp = Now()+Int(cooCfg)/60/24
			ElseIf Test(cooCfg,"domain") or Test(cooCfg,"ip") Then
				cDomain = cooCfg
			ElseIf Instr(cooCfg,"/")>0 Then
				cPath = cooCfg
			ElseIf cooCfg = "True" or cooCfg = "False" Then
				cSecure = cooCfg
			End If
		End If
		If Has(cooValue) Then
			If b_cooen Then
				Use("Aes") : cooValue = Aes.Encode(cooValue)
			End If
		End If
		If Instr(cooName,">")>0 Then
			n = CRight(cooName,">")
			cooName = CLeft(cooName,">")
			Response.Cookies(cooName)(n) = cooValue
		Else
			Response.Cookies(cooName) = cooValue
		End If
		If Has(cExp) Then Response.Cookies(cooName).Expires = cExp
		If Has(cDomain) Then Response.Cookies(cooName).Domain = cDomain
		If Has(cPath) Then Response.Cookies(cooName).Path = cPath
		If Has(cSecure) Then Response.Cookies(cooName).Secure = cSecure
	End Sub
	'获取一个Cookies值
	Function Cookie(ByVal s)
		Dim p,t,coo
		If Instr(s,">") > 0 Then
			p = CLeft(s,">")
			s = CRight(s,">")
		End If
		If Instr(s,":")>0 Then
			t = CRight(s,":")
			s = CLeft(s,":")
		End If
		If Has(p) And Has(s) Then
			If Response.Cookies(p).HasKeys Then
				coo = Request.Cookies(p)(s)
			End If
		ElseIf Has(s) Then
			coo = Request.Cookies(s)
		Else
			Cookie = "" : Exit Function
		End If
		If IsN(coo) Then Cookie = "": Exit Function
		If  b_cooen Then
			Use("Aes") : coo = Aes.Decode(coo)
		End If
		Cookie = Safe(coo,t)
	End Function
	'删除一个Cookies值
	Sub RemoveCookie(ByVal cooName)
		Dim n : n = Easp_Param(cooName)
		If Response.Cookies(n(0)).HasKeys And Has(n(1)) Then
			Response.Cookies(n(0))(n(1)) = Empty
		Else
			Response.Cookies(n(0)) = Empty
			Response.Cookies(n(0)).Expires = Now()
		End If
	End Sub
	'设置缓存记录
	Sub SetApp(AppName,AppData)
		Application.Lock
		Application.Contents.Item(AppName) = AppData
		Application.UnLock
	End Sub
	'获取一个缓存记录
	Function GetApp(AppName)
		If IsN(AppName) Then GetApp = "" : Exit Function
		GetApp = Application.Contents.Item(AppName)
	End Function
	'删除一个缓存记录
	Sub RemoveApp(AppName)
		Application.Lock
		Application.Contents.Remove(AppName)
		Application.UnLock
	End Sub
	'验证身份证号码
	Private Function isIDCard(ByVal s)
		Dim Ai, BirthDay, arrVerifyCode, Wi, i, AiPlusWi, modValue, strVerifyCode
		isIDCard = False
		If Len(s) <> 15 And Len(s) <> 18 Then Exit Function
		Ai = IIF(Len(s) = 18,Mid(s, 1, 17),Left(s, 6) & "19" & Mid(s, 7, 9))
		If Not IsNumeric(Ai) Then Exit Function
		If Not Test(Left(Ai,6),"^(1[1-5]|2[1-3]|3[1-7]|4[1-6]|5[0-4]|6[1-5]|8[12]|91)\d{2}[01238]\d{1}$") Then Exit Function
		BirthDay = Mid(Ai, 7, 4) & "-" & Mid(Ai, 11, 2) & "-" & Mid(Ai, 13, 2)
		If IsDate(BirthDay) Then
			If cDate(BirthDay) > Date() Or cDate(BirthDay) < cDate("1870-1-1") Then  Exit Function
		Else
			Exit Function
		End If
		arrVerifyCode = Split("1,0,x,9,8,7,6,5,4,3,2", ",")
		Wi = Split("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2", ",")
		For i = 0 To 16
			AiPlusWi = AiPlusWi + CInt(Mid(Ai, i + 1, 1)) * Wi(i)
		Next
		modValue = AiPlusWi Mod 11
		strVerifyCode = arrVerifyCode(modValue)
		Ai = Ai & strVerifyCode
		If Len(s) = 18 And LCase(s) <> Ai Then Exit Function
		isIDCard = True
	End Function
	'简易的服务端检查表单
	Function CheckForm(ByVal s, ByVal Rule, ByVal Require, ByVal ErrMsg)
		Dim tmpMsg, Msg, i
		tmpMsg = Replace(ErrMsg,"\:",chr(0))
		Msg = IIF(Instr(tmpMsg,":")>0,Split(tmpMsg,":"),Array("有项目不能为空",tmpMsg))
		If Require = 1 And IsN(s) Then
			If Instr(tmpMsg,":")>0 Then
				Alert Replace(Msg(0),chr(0),":") : Exit Function
			Else
				Alert Replace(tmpMsg,chr(0),":") : Exit Function
			End If
		End If
		If Not (Require = 0 And isN(s)) Then
			If Left(Rule,1)=":" Then
				pass = False
				arrRule = Split(Mid(Rule,2),"||")
				For i = 0 To Ubound(arrRule)
					If Test(s,arrRule(i)) Then pass = True : Exit For
				Next
				If Not pass Then Alert(Replace(Msg(1),chr(0),":")) : Exit Function
			Else
				If Not Test(s,Rule) Then Alert(Replace(Msg(1),chr(0),":")) : Exit Function
			End If
		End If
		CheckForm = s
	End Function
	'返回正则验证结果
	Function [Test](ByVal s, ByVal p)
		Dim Pa
		Select Case Lcase(p)
			Case "date"		Test = IIF(isDate(s),True,False) : Exit Function
			Case "idcard"	Test = IIF(isIDCard(s),True,False) : Exit Function
			Case "english"	Pa = "^[A-Za-z]+$"
			Case "chinese"	Pa = "^[\u0391-\uFFE5]+$"
			Case "username"	Pa = "^[a-zA-Z]\w{2,19}$"
			Case "email"	Pa = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
			Case "int"		Pa = "^[-\+]?\d+$"
			Case "number"	Pa = "^\d+$"
			Case "double"	Pa = "^[-\+]?\d+(\.\d+)?$"
			Case "price"	Pa = "^\d+(\.\d+)?$"
			Case "zip"		Pa = "^[1-9]\d{5}$"
			Case "qq"		Pa = "^[1-9]\d{4,9}$"
			Case "phone"	Pa = "^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$"
			Case "mobile"	Pa = "^((\(\d{2,3}\))|(\d{3}\-))?(1[35][0-9]|189)\d{8}$"
			Case "url"		Pa = "^(http|https|ftp):\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\""])*$"
			Case "domain"	Pa = "^[A-Za-z0-9\-\.]+\.([A-Za-z]{2,4}|[A-Za-z]{2,4}\.[A-Za-z]{2})$"
			Case "ip"		Pa = "^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$"
			Case Else Pa = p
		End Select
		[Test] = Easp_Test(CStr(s),Pa)
	End Function
	'正则替换
	Function regReplace(ByVal s, ByVal rule, Byval Result)
		regReplace = Easp_Replace(s,rule,Result,0)
	End Function
	'正则替换多行模式
	Function regReplaceM(ByVal s, ByVal rule, Byval Result)
		regReplaceM = Easp_Replace(s,rule,Result,1)
	End Function
	'正则匹配捕获
	Function regMatch(ByVal s, ByVal rule)
		Set regMatch =  Easp_Match(s,rule)
	End Function
	'检测组件是否安装
	Function isInstall(Byval s)
		On Error Resume Next : Err.Clear()
		isInstall = False
		Dim obj : Set obj = Server.CreateObject(s)
		If Err.Number = 0 Then isInstall = True
		Set obj = Nothing : Err.Clear()
	End Function
	'动态载入文件
	Sub Include(ByVal filePath)
		ExecuteGlobal GetIncCode(filePath,0)
	End Sub
	Function getInclude(ByVal filePath)
		ExecuteGlobal GetIncCode(filePath,1)
		getInclude = EasyAsp_s_html
	End Function
	'读取文件内容
	Private Function Read(ByVal filePath)
		Dim Fso, p, f, tmpStr,o_strm
		p = filePath
		If Not (Mid(filePath,2,1)=":") Then p = Server.MapPath(filePath)
		Set Fso = Server.CreateObject(s_fsoName)
		If Fso.FileExists(p) Then
			If s_charset = "GB2312" Then
				Set f = Fso.OpenTextFile(p)
				tmpStr = f.ReadAll
				f.Close()
				Set f = Nothing
			Else
				Set o_strm = Server.CreateObject("ADODB.Stream")
				With o_strm
					.Type = 2
					.Mode = 3
					.Open
					.LoadFromFile p
					.Charset = s_charset
					.Position = 2
					tmpStr = .ReadText
					.Close
				End With
				Set o_strm = Nothing
			End If
		Else
			tmpStr = "文件未找到:" & filePath
		End If
		Set Fso = Nothing
		Read = tmpStr
	End Function
	'读取包含文件内容（无限级）
	Private Function IncRead(ByVal filePath)
		Dim content, rule, inc, incFile, incStr
		content = Read(filePath)
		If isN(content) Then Exit Function
		content = regReplace(content,"<% *?@.*?%"&">","")
		content = regReplace(content,"(<%[^>]+?)(option +?explicit)([^>]*?%"&">)","$1'$2$3")
		rule = "<!-- *?#include +?(file|virtual) *?= *?""??([^"":?*\f\n\r\t\v]+?)""?? *?-->"
		If Easp_Test(content,rule) Then
			Set inc = regMatch(content,rule)
			For Each Match In inc
				If LCase(Match.SubMatches(0))="virtual" Then
					incFile = Match.SubMatches(1)
				Else
					incFile = Mid(filePath,1,InstrRev(filePath,IIF(Instr(filePath,":")>0,"\","/"))) & Match.SubMatches(1)
				End If
				incStr = IncRead(incFile)
				content = Replace(content,Match,incStr)
			Next
			Set inc = Nothing
		End If
		IncRead = content
	End Function
	'将包含文件转换为ASP代码
	Private Function GetIncCode(ByVal filePath, ByVal getHtml)
		Dim content,tmpStr,code,tmpCode,s_code,st,en
		content = IncRead(filePath)
		code = "" : st = 1 : en = Instr(content,"<%") + 2
		s_code = IIF(getHtml=1,"EasyAsp_s_html = EasyAsp_s_html & ","Response.Write ")
		While en > st + 1
			tmpStr = Mid(content,st,en-st-2)
			st = Instr(en,content,"%"&">") + 2
			If Has(tmpStr) Then
				tmpStr = Replace(tmpStr,"""","""""")
				tmpStr = Replace(tmpStr,vbCrLf&vbCrLf,vbCrLf)
				tmpStr = Replace(tmpStr,vbCrLf,"""&vbCrLf&""")
				code = code & s_code & """" & tmpStr & """" & vbCrLf
			End If
			tmpStr = Mid(content,en,st-en-2)
			tmpCode = regReplace(tmpStr,"^\s*=\s*",s_code) & vbCrLf
			If getHtml = 1 Then
				tmpCode = regReplaceM(tmpCode,"^(\s*)response\.write","$1" & s_code) & vbCrLf
				tmpCode = regReplaceM(tmpCode,"^(\s*)Easp\.(W|WC|WN)","$1" & s_code) & vbCrLf
			End If
			code = code & Replace(tmpCode,vbCrLf&vbCrLf,vbCrLf)
			en = Instr(st,content,"<%") + 2
		Wend
		tmpStr = Mid(content,st)
		If Has(tmpStr) Then
			tmpStr = Replace(tmpStr,"""","""""")
			tmpStr = Replace(tmpStr,vbCrLf&vbCrLf,vbCrLf)
			tmpStr = Replace(tmpStr,vbcrlf,"""&vbCrLf&""")
			code = code & s_code & """" & tmpStr & """" & vbCrLf
		End If
		If getHtml = 1 Then code = "EasyAsp_s_html = """" " & vbCrLf & code
		GetIncCode = Replace(code,vbCrLf&vbCrLf,vbCrLf)
	End Function
	'加载引用EasyAsp库类
	Sub Use(ByVal f)
		Dim p, o, t : o = f
		p = "easp." & Lcase(o) & ".asp"
		If LCase(o) = "md5" Then o = "o_md5"
		t = Eval("LCase(TypeName("&o&"))")
		If t = "easyasp_obj" Then
			Include(s_path & "core/" & p)
			Execute("Set "&o&" = New EasyAsp_"&f)
			Select Case Lcase(f)
				Case "fso"
					fso.fsoName = s_fsoName
					fso.CharSet = s_charset
				Case "upload"
					upload.CharSet = s_charset
				Case "tpl"
					tpl.CharSet = s_charset
			End Select
		End If
	End Sub
	'加载插件
	Function Ext(ByVal f)
		Dim loaded
		f = Lcase(f) : loaded = True
		If Not o_ext.Exists(f) Then
			loaded = False
		Else
			If LCase(TypeName(o_ext(f))) <> "easyasp_" & f Then loaded = False
		End If
		If Not loaded Then
			Include(s_plugin & "easp." & f & ".asp")
			Execute("Set o_ext(""" & f & """) = New EasyAsp_" & f)
		End If
		Set Ext = o_ext(f)
	End Function
	'清除加载插件
	Private Sub ClearExt()
		Dim i
		If Has(o_ext) Then
			For Each i In o_ext
				Set o_ext(i) = Nothing
			Next
			o_ext.RemoveAll
		End If
	End Sub
	'Md5加密字符串
	Function MD5(ByVal s)
		Use("Md5") : MD5 = o_md5(s)
	End Function
	Function MD5_16(ByVal s)
		Use("Md5") : MD5_16 = o_md5.To16(s)
	End Function
End Class
Class EasyAsp_obj : End Class

'EasyASP及子类通用函数部分
Function Easp_IIF(ByVal Cn, ByVal T, ByVal F)
	If Cn Then
		Easp_IIF = T
	Else
		Easp_IIF = F
	End If
End Function
Function Easp_Param(ByVal s)
	Dim arr(1),t : t = Instr(s,":")
	If t > 0 Then
		arr(0) = Left(s,t-1) : arr(1) = Mid(s,t+1)
	Else
		arr(0) = s : arr(1) = ""
	End If
	Easp_Param = arr
End Function
Function Easp_isN(ByVal s)
	Easp_isN = False
	Select Case VarType(s)
		Case vbEmpty, vbNull
			Easp_isN = True : Exit Function
		Case vbString
			If s="" Then Easp_isN = True : Exit Function
		Case vbObject
			Select Case TypeName(s)
				Case "Nothing","Empty"
					Easp_isN = True : Exit Function
				Case "Recordset"
					If s.State = 0 Then Easp_isN = True : Exit Function
					If s.Bof And s.Eof Then Easp_isN = True : Exit Function
				Case "Dictionary"
					If s.Count = 0 Then Easp_isN = True : Exit Function
			End Select
		Case vbArray,8194,8204,8209
			If Ubound(s)=-1 Then Easp_isN = True : Exit Function
	End Select
End Function
Function Easp_JsEncode(ByVal s)
	Dim i, j, aL1, aL2, c, p, t
	aL1 = Array(&h22, &h5C, &h2F, &h08, &h0C, &h0A, &h0D, &h09)
	aL2 = Array(&h22, &h5C, &h2F, &h62, &h66, &h6E, &h72, &h74)
	For i = 1 To Len(s)
		p = True
		c = Mid(s, i, 1)
		For j = 0 To 7
			If c = Chr(aL1(j)) Then
				t = t & "\" & Chr(aL2(j))
				p = False
				Exit For
			End If
		Next
		If p Then 
			Dim a
			a = AscW(c)
			If a > 31 And a < 127 Then
				t = t & c
			ElseIf a > -1 Or a < 65535 Then
				t = t & "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
			End If 
		End If
	Next
	Easp_JsEncode = t
End Function
Function Easp_Escape(ByVal ss)
	Dim i,c,a,s : s = ""
	If Easp_isN(ss) Then Easp_Escape = "" : Exit Function
	For i = 1 To Len(ss)
		c = Mid(ss,i,1)
		a = ASCW(c)
		If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
			s = s & c
		ElseIf InStr("@*_+-./",c)>0 Then
			s = s & c
		ElseIf a>0 and a<16 Then
			s = s & "%0" & Hex(a)
		ElseIf a>=16 and a<256 Then
			s = s & "%" & Hex(a)
		Else
			s = s & "%u" & Hex(a)
		End If
	Next
	Easp_Escape = s
End Function
Function Easp_UnEscape(ByVal ss)
	Dim x, s
	x = InStr(ss,"%")
	s = ""
	Do While x>0
		s = s & Mid(ss,1,x-1)
		If LCase(Mid(ss,x+1,1))="u" Then
			s = s & ChrW(CLng("&H"&Mid(ss,x+2,4)))
			ss = Mid(ss,x+6)
		Else
			s = s & Chr(CLng("&H"&Mid(ss,x+1,2)))
			ss = Mid(ss,x+3)
		End If
		x=InStr(ss,"%")
	Loop
	Easp_UnEscape = s & ss
End Function
Function Easp_RandStr(ByVal cfg)
	Dim a, p, l, t, reg, m, mi, ma
	cfg = Replace(Replace(Replace(cfg,"\<",Chr(0)),"\>",Chr(1)),"\:",Chr(2))
	a = ""
	If Easp_Test(cfg, "(<\d+>|<\d+-\d+>)") Then
		t = cfg
		p = Easp_Param(cfg)
		If Not Easp_isN(p(1)) Then
			a = p(1) : t = p(0) : p = ""
		End If
		Set reg = Easp_Match(cfg, "(<\d+>|<\d+-\d+>)")
		For Each m In reg
			p = m.SubMatches(0)
			l = Mid(p,2,Len(p)-2)
			If Easp_Test(l,"^\d+$") Then
				t = Replace(t,p,Easp_RandString(l,a),1,1)
			Else
				mi = Easp_LR(l,"-",0)
				ma = Easp_LR(l,"-",1)
				t =  Replace(t,p,Easp_Rand(mi, ma),1,1)
			End If
		Next
		Set reg = Nothing
	ElseIf Easp_Test(cfg,"^\d+-\d+$") Then
		mi = Easp_LR(cfg,"-",0)
		ma = Easp_LR(cfg,"-",1)
		t = Easp_Rand(mi, ma)
	ElseIf Easp_Test(cfg, "^(\d+)|(\d+:.)$") Then
		l = cfg : p = Easp_Param(cfg)
		If Not Easp_isN(p(1)) Then
			a = p(1) : l = p(0) : p = ""
		End If
		t = Easp_RandString(l, a)
	Else
		t = cfg
	End If
	Easp_RandStr = Replace(Replace(Replace(t,Chr(0),"<"),Chr(1),">"),Chr(2),":")
End Function
Function Easp_RandString(ByVal length, ByVal allowStr)
	Dim i
	If Easp_IsN(allowStr) Then allowStr = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	For i = 1 To length
		Randomize() : Easp_RandString = Easp_RandString & Mid(allowStr, Int(Len(allowStr) * Rnd + 1), 1)
	Next
End Function
Function Easp_Rand(ByVal min, ByVal max)
    Randomize() : Easp_Rand = Int((max - min + 1) * Rnd + min)
End Function
Function Easp_Test(ByVal s, ByVal p)
	If Easp_IsN(s) Then Easp_Test = False : Exit Function
	Dim Reg
	Set Reg = New RegExp
	Reg.Global = True
	Reg.Pattern = p
	Easp_Test = Reg.Test(CStr(s))
	Set Reg = Nothing
End Function
Function Easp_Replace(ByVal s, ByVal rule, Byval Result, ByVal isM)
	Dim tmpStr,Reg : tmpStr = s
	If Not Easp_isN(s) Then
		Set Reg = New Regexp
		Reg.Global = True
		Reg.IgnoreCase = True
		If isM = 1 Then Reg.Multiline = True
		Reg.Pattern = rule
		tmpStr = Reg.Replace(tmpStr,Result)
		Set Reg = Nothing
	End If
	Easp_Replace = tmpStr
End Function
'正则匹配
Function Easp_Match(ByVal s, ByVal rule)
	Dim Reg
	Set Reg = New Regexp
	Reg.Global = True
	Reg.IgnoreCase = True
	Reg.Pattern = rule
	Set Easp_Match = Reg.Execute(s)
	Set Reg = Nothing
End Function
'取字符串的两头
Function Easp_LR(ByVal s, ByVal m, ByVal t)
	Dim n : n = Instr(s,m)
	If n>0 Then
		If t = 0 Then
			Easp_LR = Left(s,n-1)
		ElseIf t = 1 Then
			Easp_LR = Mid(s,n+1)
		End If
	Else
		Easp_LR = s
	End If
End Function
%>
<!--#include file="core/easp.error.asp"-->
<!--#include file="core/easp.db.asp"-->