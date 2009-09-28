<%
Class EasyAsp_db
	Private idbConn, idbType, idebug, idbErr, iQueryType
	Private iPageParam, iPageIndex, iPageSize, iPageSpName, iPageCount, iRecordCount, iPageDic

	Private Sub Class_Initialize()
		On Error Resume Next
		idbType = ""
		idebug = False
		idbErr = ""
		iQueryType = 0
		If TypeName(Conn) = "Connection" Then
			Set idbConn = Conn : idbType = GetDataType(Conn)
		End If
		iPageParam = "page"
		iPageSize = 20
		iPageSpName = "easp_sp_pager"
		Set iPageDic = Server.CreateObject("Scripting.Dictionary")

		iPageDic("default_html") = "<div class=""pager"">{first}{prev}{liststart}{list}{listend}{next}{last} 跳转到{jump}页</div>"
		iPageDic("default_config") = ""
	End Sub
	Private Sub Class_Terminate()
		If TypeName(idbConn) = "Connection" Then
			If idbConn.State = 1 Then idbConn.Close()
			Set idbConn = Nothing
		End If
		Set iPageDic = Nothing
	End Sub
	'属性：定义数据库连接
	Public Property Let dbConn(ByVal pdbConn)
		If TypeName(pdbConn) = "Connection" Then
			Set idbConn = pdbConn
			idbType = GetDataType(pdbConn)
		Else
			ErrMsg "无效的数据库连接", Err.Description
		End If
	End Property
	Public Property Get dbConn()
		Set dbConn = idbConn
	End Property
	'属性：当前数据库类型
	Public Property Get DatabaseType()
		DatabaseType = idbType
	End Property
	'属性：设置是否开启调试模式
	Public Property Let Debug(ByVal bool)
		idebug = bool
	End Property
	Public Property Get Debug()
		Debug = idebug
	End Property
	'属性：返回错误信息
	Public Property Get dbErr()
		dbErr = idbErr
	End Property
	'属性：设置获取记录集的方式
	Public Property Let QueryType(ByVal str)
		str = Lcase(str)
		If str = "1" or str = "command" Then
			iQueryType = 1
		Else
			iQueryType = 0
		End If
	End Property
	'属性：设置分页数量
	Public Property Let PageSize(ByVal num)
		iPageSize = num
	End Property
	'属性：返回分页数量
	Public Property Get PageSize()
		PageSize = iPageSize
	End Property
	'属性：返回总页数
	Public Property Get PageCount()
		PageCount = iPageCount
	End Property
	'属性：返回当前页码
	Public Property Get PageIndex()
		PageIndex = Easp_IIF(Easp_isN(iPageIndex),GetCurrentPage,iPageIndex)
	End Property
	'属性：返回总记录数
	Public Property Get PageRecordCount()
		PageRecordCount = iRecordCount
	End Property
	'属性：设置获取分页参数
	Public Property Let PageParam(ByVal str)
		iPageParam = str
	End Property
	'属性：设置分页存储过程名
	Public Property Let PageSpName(ByVal str)
		iPageSpName = str
	End Property
	Private Sub ErrMsg(e,d)
		idbErr = "<div id=""easp_db_err"">" & e
		If d<>"" Then idbErr = idbErr & "<br/>错误信息：" & d
		idbErr = idbErr & "</div>"
		If idebug Then
			Response.Write idbErr
			Response.End()
		End If
	End Sub
	'生成数据库连接字符串
	Public Function OpenConn(ByVal dbType, ByVal strDB, ByVal strServer)
		Dim TempStr, objConn, s, u, p, port
		s = "" : u = "" : p = "" : port = ""
		If Instr(strServer,"@")>0 Then
			s = Trim(Mid(strServer,InstrRev(strServer,"@")+1))
			u = Trim(Left(strServer,InstrRev(strServer,"@")-1))
			If Instr(s,":")>0 Then : port = Trim(Mid(s,Instr(s,":")+1)) : s = Trim(Left(s,Instr(s,":")-1))
			If Instr(u,":")>0 Then : p = Trim(Mid(u,Instr(u,":")+1)) : u = Trim(Left(u,Instr(u,":")-1))
		Else
			If Instr(strServer,":")>0 Then
				u = Trim(Left(strServer,Instr(strServer,":")-1))
				p = Trim(Mid(strServer,Instr(strServer,":")+1))
			Else
				p = Trim(strServer)
			End If
		End If
		idbType = UCase(Cstr(dbType))
		Select Case idbType
			Case "0","MSSQL"
				If port = "" Then port = "1433"
				TempStr = "Provider=sqloledb;Data Source="&s&","&port&";Initial Catalog="&strDB&";User Id="&u&";Password="&p&";"
			Case "1","ACCESS"
				Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
				TempStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&p&";"
			Case "2","MYSQL"
				If port = "" Then port = "3306"
				TempStr = "Driver={mySQL};Server="&s&";Port="&port&";Option=131072;Stmt=;Database="&strDB&";Uid="&u&";Pwd="&p&";"
			Case "3","ORACLE"
				TempStr = "Provider=msdaora;Data Source="&s&";User Id="&u&";Password="&p&";"
		End Select
		Set OpenConn = CreatConn(TempStr)
	End Function
	'建立数据库连接对象
	Public Function CreatConn(ByVal ConnStr)
		On Error Resume Next
		Dim objConn : Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open ConnStr
		If Err.number <> 0 Then
			ErrMsg "数据库服务器端连接错误，请检查数据库连接。", Err.Description
			objConn.Close
			Set objConn = Nothing
		End If
		Set CreatConn = objConn
	End Function
	Private Function GetDataType(ByVal connObj)
		Dim str,i : str = UCase(connObj.Provider)
		Dim MSSQL, ACCESS, MYSQL, ORACLE
		MSSQL = Split("SQLNCLI10, SQLXMLOLEDB, SQLNCLI, SQLOLEDB, MSDASQL",", ")
		ACCESS = Split("MICROSOFT.ACE.OLEDB.12.0, MICROSOFT.JET.OLEDB.4.0",", ")
		MYSQL = "MYSQLPROV"
		ORACLE = Split("MSDAORA, OLEDB.ORACLE",", ")
		For i = 0 To Ubound(MSSQL)
			If Instr(str,MSSQL(i))>0 Then
				GetDataType = "MSSQL" : Exit Function
			End If
		Next
		For i = 0 To Ubound(ACCESS)
			If Instr(str,ACCESS(i))>0 Then
				GetDataType = "ACCESS" : Exit Function
			End If
		Next
		If Instr(str,MYSQL)>0 Then
			GetDataType = "MYSQL" : Exit Function
		End If
		For i = 0 To Ubound(ORACLE)
			If Instr(str,ORACLE(i))>0 Then
				GetDataType = "ORACLE" : Exit Function
			End If
		Next
	End Function
	'自动获取唯一序列号（自动编号）
	Public Function AutoID(ByVal TableName)
		On Error Resume Next
		Dim rs, tmp, fID, tmpID : fID = "" : tmpID = 0
		tmp = Easp_Param(TableName)
		If Not Easp_isN(tmp(1)) Then : TableName = tmp(0) : fID = tmp(1) : tmp = "" : End If
		Set rs = GRS("Select " & Easp_IIF(fID<>"", "Max("&fID&")", "Top 1 *") & " From ["&TableName&"]")
		If rs.eof Then
			AutoID = 1 : Exit Function
		Else
			If fID<>"" Then
				If Easp_isN(rs.Fields.Item(0).Value) Then AutoID = 1 : Exit Function
				AutoID = rs.Fields.Item(0).Value + 1 : Exit Function
			Else
				Dim newRs
				Set newRs = GRS("Select Max("&rs.Fields.Item(0).Name&") From ["&TableName&"]")
				tmpID = newRS.Fields.Item(0).Value + 1
				newRs.Close() : Set newRs = Nothing
			End If
		End If
		If Err.number <> 0 Then ErrMsg "无效的查询条件，无法获取新的ID号！", Err.Description
		rs.Close() : Set rs = Nothing
		AutoID = tmpID
	End Function
	'取得符合条件的纪录列表
	Public Function GetRecord(ByVal TableName,ByVal Condition,ByVal OrderField)
		Set GetRecord = GRS(wGetRecord(TableName,Condition,OrderField))
	End Function
	Public Function wGetRecord(ByVal TableName,ByVal Condition,ByVal OrderField)
		Dim strSelect, FieldsList, ShowN, o, p
		FieldsList = "" : ShowN = 0
		o = Easp_Param(TableName)
		If Not Easp_isN(o(1)) Then
			TableName = Trim(o(0)) : FieldsList = Trim(o(1)) : o = ""
			p = Easp_Param(FieldsList)
			If Not Easp_isN(p(1)) Then
				FieldsList = Trim(p(0)) : ShowN = Int(Trim(p(1))) : p = ""
			Else
				If isNumeric(FieldsList) Then ShowN = Int(FieldsList) : FieldsList = ""
			End If
		End If
		strSelect = "Select "
		If ShowN > 0 Then strSelect = strSelect & "Top " & ShowN & " "
		strSelect = strSelect & Easp_IIF(FieldsList <> "", FieldsList, "* ")
		strSelect = strSelect & " From [" & TableName & "]"
		If isArray(Condition) Then
			strSelect = strSelect & " Where " & ValueToSql(TableName,Condition,1)
		Else
			If Condition <> "" Then strSelect = strSelect & " Where " & Condition
		End If
		If OrderField <> "" Then strSelect = strSelect & " Order By " & OrderField
		wGetRecord = strSelect
	End Function
	Public Function GR(ByVal TableName,ByVal Condition,ByVal OrderField)
		Set GR = GetRecord(TableName, Condition, OrderField)
	End Function
	Public Function wGR(ByVal TableName,ByVal Condition,ByVal OrderField)
		wGR = wGetRecord(TableName, Condition, OrderField)
	End Function
	'根据sql语句返回记录集
	Public Function GetRecordBySQL(ByVal str)
		On Error Resume Next
		If iQueryType = 1 Then
			Dim cmd : Set cmd = Server.CreateObject("ADODB.Command")
			With cmd
				.ActiveConnection = idbConn
				.CommandText = str
				Set GetRecordBySQL = .Execute
			End With
			Set cmd = Nothing
		Else
			Dim rs : Set rs = Server.CreateObject("Adodb.Recordset")
			With rs
				.ActiveConnection = idbConn
				.CursorType = 1
				.LockType = 1
				.Source = str
				.Open
			End With
			Set GetRecordBySQL = rs
		End If
		If Err.number <> 0 Then ErrMsg "无效的查询条件，无法获取记录集！", Err.Description & "<br/>SQL：" & str
		Err.Clear
	End Function
	Public Function GRS(ByVal strSelect)
		Set GRS = GetRecordBySQL(strSelect)
	End Function
	'根据记录集生成Json格式代码
	Public Function Json(ByVal jRs, ByVal jName)
		On Error Resume Next
		Dim tmpStr, rs, fi, i, j, o, isE,tName,tValue : i = 0
		isE = False
		o = Easp_Param(jName)
		If Not Easp_isN(o(1)) Then
			jName = o(0)
			isE = True
		End If
		Set rs = jRs
		tmpStr = "{ """&jName&""" : ["
		rs.MoveFirst()
		If Not rs.bof And Not rs.eof Then
			While Not rs.Eof
				j = 0 : If i<>0 Then tmpStr = tmpStr & ", "
				tmpStr = tmpStr & "{"
				For Each fi In rs.Fields
					If j<>0 Then tmpStr = tmpStr & ", "
					tName = fi.Name : tValue = fi.Value
					If isE Then
						tmpStr = tmpStr & """" & Easp_Escape(tName) & """:""" & Easp_Escape(Easp_jsEncode(tValue)) & """"
					Else
						tmpStr = tmpStr & """" & tName & """:""" & Easp_jsEncode(tValue) & """"
					End If
					j = j + 1
				Next
				tmpStr = tmpStr & "}"
				i = i + 1 : rs.MoveNext()
			Wend
		End If
		tmpStr = tmpStr & "]}"
		If Err.number <> 0 Then ErrMsg "生成Json格式代码出错！", Err.Description
		rs.Close() : Set rs = Nothing
		Json = tmpStr
	End Function
	'生成指定长度的不重复的字符串
	Public Function RandStr(length,TableField)
		On Error Resume Next
		Dim tb, fi, tmpStr, rs
		tb = Easp_Param(TableField)(0)
		fi = Easp_Param(TableField)(1)
		tmpStr = Easp_RandStr(length)
		Do While (True)
			Set rs = GR(tb&":"&fi&":1",fi&"='"&tmpStr&"'","")
			If Not rs.Bof And Not rs.Eof Then
				tmpStr = Easp_RandStr(length)
			Else
				RandStr = tmpStr
				Exit Do
			End If
			C(rs)
		Loop
		If Err.number <> 0 Then ErrMsg "生成不重复的随机字符串出错！", Err.Description
	End Function
	'生成一个不重复的随机数
	Public Function Rand(min,max,TableField)
		On Error Resume Next
		Dim tb, fi, tmpInt, rs
		tb = Easp_Param(TableField)(0)
		fi = Easp_Param(TableField)(1)
		tmpInt = Easp_Rand(min,max)
		Do While (True)
			Set rs = GR(tb&":"&fi&":1",Array(fi&":"&tmpInt),"")
			If Not rs.Bof And Not rs.Eof Then
				tmpInt = Easp_Rand(min,max)
			Else
				Rand = tmpInt
				Exit Do
			End If
			C(rs)
		Loop
		If Err.number <> 0 Then ErrMsg "生成不重复的随机数出错！", Err.Description
	End Function
	'取得某一指定纪录的详细资料
	Public Function GetRecordDetail(ByVal TableName,ByVal Condition)
		Dim strSelect
		strSelect = "Select * From [" & TableName & "] Where " & ValueToSql(TableName,Condition,1)
		Set GetRecordDetail = GRS(strSelect)
	End Function
	Public Function GRD(ByVal TableName,ByVal Condition)
		Set GRD = GetRecordDetail(TableName, Condition)
	End Function
	'取指定数量的随机记录
	Public Function GetRandRecord(ByVal TableName,ByVal Condition)
		Dim sql,o,p,fi,IdField,showN,where
		o = Easp_Param(TableName)
		If Not Easp_isN(o(1)) Then
			TableName = o(0)
			p = Easp_Param(o(1))
			If Easp_isN(p(1)) Then
				ErrMsg "获取随机记录失败！", "请输入要取的记录数量"
				Exit Function
			Else
				fi = p(0) : showN = p(1)
				If Instr(fi,",")>0 Then
					IdField = Trim(Left(fi,Instr(fi,",")-1))
				Else
					IdField = fi : fi = "*"
				End If
			End If
		Else
			ErrMsg "获取随机记录失败！", "请在表名后输入:ID字段的名称"
			Exit Function
		End If
		Condition = Easp_IIF(Easp_isN(Condition),""," Where " & ValueToSql(TableName,Condition,1))
		sql = "Select Top " & showN & " " & fi & " From ["&TableName&"]" & Condition
		Select Case idbType
			Case "ACCESS" : Randomize
				sql = sql & " Order By Rnd(-(" & IdField & "+" & Rnd() & "))"
			Case "MSSQL"
				sql = sql & " Order By newid()"
			Case "MYSQL"
				sql = "Select " & fi & " From ["&TableName&"]" & Condition & " Order By rand() limit " & showN
			Case "ORACLE"
				sql = "Select " & fi & " From (Select " & fi & " From ["&TableName&"] Order By dbms_random.value) " & Easp_IIF(Easp_isN(Condition),"Where",Condition & " And") & " rownum < " & Int(showN)+1
		End Select
		Set GetRandRecord = GRS(sql)
	End Function
	Public Function GRR(ByVal TableName,ByVal Condition)
		Set GRR = GetRandRecord(TableName,Condition)
	End Function
	'添加一个新的纪录
	Public Function AddRecord(ByVal TableName,ByVal ValueList)
		On Error Resume Next
		Dim o : o = Easp_Param(TableName) : If Not Easp_isN(o(1)) Then TableName = o(0)
		DoExecute wAddRecord(TableName,ValueList)
		If Err.number <> 0 Then
			ErrMsg "向数据库添加记录出错！", Err.Description
			AddRecord = 0
			Exit Function
		End If
		If Not Easp_isN(o(1)) Then
			AddRecord = AutoID(o(0)&":"&o(1))-1
		Else
			AddRecord = 1
		End If
	End Function
	Public Function wAddRecord(ByVal TableName,ByVal ValueList)
		Dim TempSQL, TempFiled, TempValue, o
		o = Easp_Param(TableName) : If Not Easp_isN(o(1)) Then TableName = o(0)
		TempFiled = ValueToSql(TableName,ValueList,2)
		TempValue = ValueToSql(TableName,ValueList,3)
		TempSQL = "Insert Into [" & TableName & "] (" & TempFiled & ") Values (" & TempValue & ")"
		wAddRecord = TempSQL
	End Function
	Public Function AR(ByVal TableName,ByVal ValueList)
		AR = AddRecord(TableName,ValueList)
	End Function
	Public Function wAR(ByVal TableName,ByVal ValueList)
		wAR = wAddRecord(TableName,ValueList)
	End Function
	'修改某一纪录
	Public Function UpdateRecord(ByVal TableName,ByVal Condition,ByVal ValueList)
		On Error Resume Next
		DoExecute wUpdateRecord(TableName,Condition,ValueList)
		If Err.number <> 0 Then
			ErrMsg "更新数据库记录出错！", Err.Description
			UpdateRecord = 0
			Exit Function
		End If
		UpdateRecord = 1
	End Function
	Public Function wUpdateRecord(ByVal TableName,ByVal Condition,ByVal ValueList)
		Dim TmpSQL
		TmpSQL = "Update ["&TableName&"] Set "
		TmpSQL = TmpSQL & ValueToSql(TableName,ValueList,0)
		TmpSQL = TmpSQL & " Where " & ValueToSql(TableName,Condition,1)
		wUpdateRecord = TmpSQL
	End Function
	Public Function UR(ByVal TableName,ByVal Condition,ByVal ValueList)
		UR = UpdateRecord(TableName, Condition, ValueList)
	End Function
	Public Function wUR(ByVal TableName,ByVal Condition,ByVal ValueList)
		wUR = wUpdateRecord(TableName, Condition, ValueList)
	End Function
	'删除指定的纪录
	Public Function DeleteRecord(ByVal TableName,ByVal Condition)
		On Error Resume Next
		DoExecute wDeleteRecord(TableName,Condition)
		If Err.number <> 0 Then
			ErrMsg "从数据库删除数据出错！", Err.Description
			DeleteRecord = 0
			Exit Function
		End If
		DeleteRecord = 1
	End Function
	Public Function wDeleteRecord(ByVal TableName,ByVal Condition)
		Dim IDFieldName, IDValues, Sql, p : IDFieldName = "" : IDValues = ""
		If Not isArray(Condition) Then
			p = Easp_Param(Condition)
			If Not Easp_isN(p(1)) Then
				IDFieldName = p(0)
				If Instr(IDFieldName," ")=0 Then
					IDValues = p(1)
				Else
					IDFieldName = ""
				End If
			End If
		End If
		Sql = "Delete From ["&TableName&"] Where " & Easp_IIF(IDFieldName="", ValueToSql(TableName,Condition,1), "["&IDFieldName&"] In (" & IDValues & ")")
		wDeleteRecord = Sql
	End Function
	Public Function DR(ByVal TableName,ByVal Condition)
		DR = DeleteRecord(TableName, Condition)
	End Function
	Public Function wDR(ByVal TableName,ByVal Condition)
		wDR = wDeleteRecord(TableName, Condition)
	End Function
	'从某一表中，根据一个条件获取一条记录的其他字段的值
	Public Function ReadTable(ByVal TableName,ByVal Condition,ByVal GetFieldNames)
		On Error Resume Next
		Dim rs,Sql,arrTemp,arrStr,TempStr,i
		TempStr = "" : arrStr = ""
		Sql = "Select "&GetFieldNames&" From ["&TableName&"] Where " & ValueToSql(TableName,Condition,1)
		Set rs = GRS(Sql)
		If Not rs.Eof Then
			If Instr(GetFieldNames,",") > 0 Then
				arrTemp = Split(GetFieldNames,",")
				For i = 0 To Ubound(arrTemp)
					If i<>0 Then arrStr = arrStr & Chr(0)
					arrStr = arrStr & rs.Fields.Item(i).Value
				Next
				TempStr = Split(arrStr,Chr(0))
			Else
				TempStr = rs.Fields.Item(0).Value
			End If
		End If
		If Err.number <> 0 Then ErrMsg "从数据库获取数据出错！", Err.Description
		rs.close() : Set rs = Nothing : Err.Clear
		ReadTable = TempStr
	End Function
	Public Function RT(ByVal TableName,ByVal Condition,ByVal GetFieldNames)
		RT = ReadTable(TableName, Condition, GetFieldNames)
	End Function
	'调用存储过程
	Public Function doSP(ByVal spName, ByVal spParam)
		On Error Resume Next
		Dim p, spType, cmd, outParam, i, NewRS : spType = ""
		If Not idbType="0" And Not idbType="MSSQL" Then
			MsgErr "仅支持从MS SQL Server数据库调用存储过程！",""
			Exit Function
		End If
		p = Easp_Param(spName)
		If Not Easp_isN(p(1)) Then : spType = UCase(Trim(p(1))) : spName = Trim(p(0)) : p = "" : End If
		Set cmd = Server.CreateObject("ADODB.Command")
			With cmd
				.ActiveConnection = idbConn
				.CommandText = spName
				.CommandType = 4
				.Prepared = true
				.Parameters.append .CreateParameter("return",3,4)
				outParam = "return"
				If Not IsArray(spParam) Then
					If spParam<>"" Then
						spParam = Easp_IIF(Instr(spParam,",")>0, spParam = Split(spParam,","), Array(spParam))
					End If
				End If
				If IsArray(spParam) Then
					For i = 0 To Ubound(spParam)
						Dim pName, pValue
						If (spType = "1" or spType = "OUT" or spType = "3" or spType = "ALL") And Instr(spParam(i),"@@")=1 Then
							.Parameters.append .CreateParameter(spParam(i),200,2,8000)
							outParam = outParam & "," & spParam(i)
						Else
							If Instr(spParam(i),"@")=1 And Instr(spParam(i),":")>2 Then
								pName = Left(spParam(i),Instr(spParam(i),":")-1)
								outParam = outParam & "," & pName
								pValue = Mid(spParam(i),Instr(spParam(i),":")+1)
								If pValue = "" Then pValue = NULL
								.Parameters.append .CreateParameter(pName,200,1,8000,pValue)
							Else
								.Parameters.append .CreateParameter("@param"&(i+1),200,1,8000,spParam(i))
								outParam = outParam & "," & "@param"&(i+1)
							End If
						End If
					Next
				End If
			End With
			outParam = Easp_IIF(Instr(outParam,",")>0, Split(outParam,","), Array(outParam))
			If spType = "1" or spType = "OUT" Then
				cmd.Execute : doSP = cmd
			ElseIf spType = "2" or spType = "RS" Then
				Set doSP = cmd.Execute
			ElseIf spType = "3" or spType = "ALL" Then
				Dim NewOut,pa : Set NewOut = Server.CreateObject("Scripting.Dictionary")
				Set NewRS = cmd.Execute : NewRS.close
				For i = 0 To Ubound(outParam)
					NewOut(Trim(outParam(i))) = cmd(i)
				Next
				NewRs.open : doSP = Array(NewRS,NewOut)
				Set NewOut = Nothing
			Else
				cmd.Execute : doSP = cmd(0)
			End If
		If Err.number <> 0 Then ErrMsg "调用存储过程出错！", Err.Description
		Set cmd = Nothing
		Err.Clear
	End Function
	'释放记录集对象
	Public Function C(ByRef ObjRs)
		On Error Resume Next
		ObjRs.close()
		Set ObjRs = Nothing
	End Function
	'执行指定的SQL语句,可返回记录集
	Public Function Exec(ByVal str)
		On Error Resume Next
		If Lcase(Left(str,6)) = "select" Then
			Dim i : i = iQueryType
			iQueryType = 1
			Set Exec = GRS(str)
			iQueryType = i
		Else
			Exec = 1 : DoExecute(str)
			If Err.number <> 0 Then Exec = 0
		End If
		If Err.number <> 0 Then
			ErrMsg "执行SQL语句出错！", Err.Description
		End If
		Err.Clear
	End Function
	
	Private Function ValueToSql(ByVal TableName, ByVal ValueList, ByVal sType)
		On Error Resume Next
		Dim StrTemp : StrTemp = ValueList
		If IsArray(ValueList) Then
			StrTemp = ""
			Dim rsTemp, CurrentField, CurrentValue, i
			Set rsTemp = GRS("Select Top 1 * From [" & TableName & "] Where 1 = -1")
			For i = 0 to Ubound(ValueList)
				CurrentField = Easp_Param(ValueList(i))(0)
				CurrentValue = Easp_Param(ValueList(i))(1)
				If i <> 0 Then StrTemp = StrTemp & Easp_IIF(sType=1, " And ", ", ")
				If sType = 2 Then
					StrTemp = StrTemp & "[" & CurrentField & "]"
				Else
					Select Case rsTemp.Fields(CurrentField).Type
						Case 7,8,129,130,133,134,135,200,201,202,203
							StrTemp = StrTemp & Easp_IIF(sType=3, "'"&CurrentValue&"'", "[" & CurrentField & "] = '"&CurrentValue&"'")
						Case 11
							Dim tmpTF, tmpTFV : tmpTFV = UCase(cstr(Trim(CurrentValue)))
							tmpTF = Easp_IIF(tmpTFV="TRUE" or tmpTFV = "1", Easp_IIF(idbType="ACCESS","True","1"), Easp_IIF(idbType="ACCESS",Easp_IIF(tmpTFV="","NULL","False"),Easp_IIF(tmpTFV="","NULL","0")))
							StrTemp = StrTemp & Easp_IIF(sType = 3, tmpTF, "[" & CurrentField & "] = " & tmpTF)
						Case Else
							CurrentValue = Easp_IIF(Easp_IsN(CurrentValue),"NULL",CurrentValue)
							StrTemp = StrTemp & Easp_IIF(sType = 3, CurrentValue, "[" & CurrentField & "] = " & CurrentValue)
					End Select
				End If
			Next
			If Err.number <> 0 Then ErrMsg "生成SQL语句出错！", Err.Description
			rsTemp.Close() : Set rsTemp = Nothing : Err.Clear
		End If
		ValueToSql = StrTemp
	End Function
	Private Function DoExecute(ByVal sql)
		Dim ExecuteCmd : Set ExecuteCmd = Server.CreateObject("ADODB.Command")
		With ExecuteCmd
			.ActiveConnection = idbConn
			.CommandText = sql
			.Execute
		End With
		Set ExecuteCmd = Nothing
	End Function
	'以下是分页程序部分
	'获取分页后的记录集
	Public Function GetPageRecord(ByVal PageSetup, ByVal Condition)
		On Error Resume Next
		Dim pType,spResult,rs,o,p,Sql,n,i,spReturn
		o = Easp_Param(Cstr(PageSetup))
		pType = o(0)
		If Not Easp_isN(o(1)) Then
			p = Easp_Param(o(1))
			If Not Easp_isN(p(1)) Then
				iPageParam = Lcase(p(0))
				iPageSize = Int(p(1))
			Else
				If isNumeric(o(1)) Then
					iPageSize = Int(o(1))
				Else
					iPageParam = Lcase(o(1))
				End If
			End If
		End If
		iPageIndex = GetCurrentPage()
		Select Case Lcase(pType)
			Case "array","0"
				If isArray(Condition) Then
					Dim Table,Fi,Where
					o = Easp_Param(Condition(0))
					If Not Easp_isN(o(1)) Then
						Table = o(0) : Fi = o(1)
					Else
						Table = Condition(0) : Fi = "*"
					End If
					If isArray(Condition(1)) Then
						Where = ValueToSql(Table,Condition(1),1)
					Else
						Where = Condition(1)
					End If
					iRecordCount = Int(RT(Table, Easp_IIF(Easp_isN(Where),"1=1",Where), "Count(0)"))
					n = iRecordCount / iPageSize
					iPageCount = Easp_IIF(n=Int(n), n, Int(n)+1)
					iPageIndex = Easp_IIF(iPageIndex > iPageCount, iPageCount, iPageIndex)
					If idbType = "1" or idbType = "ACCESS" Then
						Set rs = GR(Table&":"&Fi,Where,Condition(2))
						rs.PageSize = iPageSize
						If iRecordCount>0 Then rs.AbsolutePage = iPageIndex
						Set GetPageRecord = rs : Exit Function
					ElseIf idbType = "2" or idbType = "MYSQL" Then
						Sql = "Select "& fi & " From [" & Table & "]"
						If Not Easp_isN(Where) Then Sql = Sql & " Where " & Where
						If Not Easp_isN(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
						Sql = Sql & " Limit " & iPageSize*(iPageIndex-1) & ", " & iPageSize
					Else
						If Ubound(Condition)<>3 Then ErrMsg "获取分页数据出错！", "数组必须是4个元素（必须提供数据库表的主键）！"
						Sql = "Select Top " & iPageSize & " " & fi
						Sql = Sql & " From [" & Table & "]"
						If Not Easp_isN(Where) Then Sql = Sql & " Where " & Where
						If iPageIndex > 1 Then
							Sql = Sql & " " & Easp_IIF(Easp_isN(Where), "Where", "And") & " " & Condition(3) & " Not In ("
							Sql = Sql & "Select Top " & iPageSize * (iPageIndex-1) & " " & Condition(3) & " From [" & Table & "]"
							If Not Easp_isN(Where) Then Sql = Sql & " Where " & Where
							If Not Easp_isN(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
							Sql = Sql & ") "
						End If
						If Not Easp_isN(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
					End If
					Set GetPageRecord = GRS(Sql)
				Else
					ErrMsg "获取分页数据出错！", "使用数组条件获取分页数据时条件参数必须为数组！"
				End If
			Case "sql","1" Set rs = GRS(Condition)
			Case "rs","2" Set rs = Condition
			Case Else
				If isArray(Condition) Then
					If pType = "" Then pType = iPageSpName
					Select Case pType
						Case "easp_sp_pager"	'使用自带分页存储过程分页
							If Ubound(Condition)<>5 Then ErrMsg "获取分页数据出错！", "使用自带分页存储过程时条件数组参数必须为6个元素！"
							spResult = doSP("easp_sp_pager:3",Array("@TableName:"&Condition(0),"@FieldList:"&Condition(1),"@Where:"&Condition(2),"@Order:"&Condition(3),"@PrimaryKey:"&Condition(4),"@SortType:"&Condition(5),"@RecorderCount:0","@pageSize:"&iPageSize,"@PageIndex:"&iPageIndex,"@@RecordCount","@@PageCount"))
						Case Else	'使用自定义分页存储过程
							spReturn = Array(False,False)
							For i = 0 To Ubound(Condition)
								If LCase(Condition(i)) = "@@recordcount" Then spReturn(0) = True
								If LCase(Condition(i)) = "@@pagecount" Then spReturn(1) = True
								If spReturn(0) And spReturn(1) Then Exit For
							Next
							If spReturn(0) And spReturn(1) Then
								spResult = doSP(pType&":3",Condition)
							Else
								ErrMsg "获取分页数据出错！", "使用自定义分页存储过程时必须包含@@RecordCount和@@PageCount输出参数！"
							End If
					End Select
					Set GetPageRecord = spResult(0)
					iRecordCount = int(spResult(1)("@@RecordCount"))
					iPageCount = int(spResult(1)("@@PageCount"))
					iPageIndex = Easp_IIF(iPageIndex > iPageCount, iPageCount, iPageIndex)
				Else
					ErrMsg "获取分页数据出错！", "使用存储过程获取分页数据时条件参数必须为数组！"
				End If
		End Select
		If Instr(",sql,rs,1,2,", "," & pType & ",")>0 Then
			iRecordCount = rs.RecordCount
			rs.PageSize = iPageSize
			iPageCount = rs.PageCount
			iPageIndex = Easp_IIF(iPageIndex > iPageCount, iPageCount, iPageIndex)
			If iRecordCount>0 Then rs.AbsolutePage = iPageIndex
			Set GetPageRecord = rs
		End If
	End Function
	Public Function GPR(ByVal PageSetup, ByVal Condition)
		Set GPR = GetPageRecord(PageSetup, Condition)
	End Function
	'生成分页导航链接
	Public Function Pager(ByVal PagerHtml, ByRef PagerConfig)
		On Error Resume Next
		Dim pList, pListStart, pListEnd, pFirst, pPrev, pNext, pLast
		Dim pJump, pJumpLong, pJumpStart, pJumpEnd, pJumpValue
		Dim i, j, tmpStr, pStart, pEnd, cfg, pcfg(1)
		tmpStr = Easp_IIF(PagerHtml="",iPageDic("default_html"),PagerHtml)
		Set cfg = Server.CreateObject("Scripting.Dictionary")
		cfg("recordcount")	= iRecordCount
		cfg("pageindex")	= iPageIndex
		cfg("pagecount")	= iPageCount
		cfg("pagesize")		= iPageSize
		cfg("listlong")		= 9
		cfg("listsidelong")	= 2
		cfg("list")			= "*"
		cfg("currentclass")	= "current"
		cfg("link")			= GetRQ(0) & "*"
		cfg("first")		= "&laquo;"
		cfg("prev")			= "&#8249;"
		cfg("next")			= "&#8250;"
		cfg("last")			= "&raquo;"
		cfg("more")			= "..."
		cfg("disabledclass")= "disabled"
		cfg("jump")			= "input"
		cfg("jumpplus")		= ""
		cfg("jumpaction")	= ""
		cfg("jumplong")		= 50
		PagerConfig = Easp_IIF(isArray(PagerConfig),PagerConfig, Easp_IIF(Easp_isN(PagerConfig),iPageDic("default_config"),Array(PagerConfig,"pagerconfig:1")))
		If isArray(PagerConfig) Then
			Dim ConfigName, ConfigValue
			For i = 0 To Ubound(PagerConfig)
				ConfigName = LCase(Left(PagerConfig(i),Instr(PagerConfig(i),":")-1))
				ConfigValue = Mid(PagerConfig(i),Instr(PagerConfig(i),":")+1)
				If Instr(",recordcount,pageindex,pagecount,pagesize,listlong,listsidelong,jumplong,", ","&ConfigName&",") > 0 Then
					cfg(ConfigName) = Int(ConfigValue)
				Else
					cfg(ConfigName) = ConfigValue
				End If
			Next
		End If
		pStart = cfg("pageindex") - ((cfg("listlong") \ 2) + (cfg("listlong") Mod 2)) + 1
		pEnd = cfg("pageindex") + (cfg("listlong") \ 2)
		If pStart < 1 Then
			pStart = 1 : pEnd = cfg("listlong")
		End If
		If pEnd > cfg("pagecount") Then
			pStart = cfg("pagecount") - cfg("listlong") + 1 : pEnd = cfg("pagecount")
		End If
		If pStart < 1 Then pStart = 1
		For i = pStart To pEnd
			If i = cfg("pageindex") Then
				pList = pList & " <span class="""&cfg("currentclass")&""">" & Replace(cfg("list"),"*",i) & "</span> "
			Else
				pList = pList & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
			End If
		Next
		If cfg("listsidelong")>0 Then
			If cfg("listsidelong") < pStart Then
				For i = 1 To cfg("listsidelong")
					pListStart = pListStart & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
				pListStart = pListStart & Easp_IIF(cfg("listsidelong")+1=pStart,"",cfg("more") & " ")
			ElseIf cfg("listsidelong") >= pStart And pStart > 1 Then
				For i = 1 To (pStart - 1)
					pListStart = pListStart & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			End If
			If (cfg("pagecount") - cfg("listsidelong")) > pEnd Then
				pListEnd = " " & cfg("more") & pListEnd
				For i = ((cfg("pagecount") - cfg("listsidelong"))+1) To cfg("pagecount")
					pListEnd = pListEnd & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			ElseIf (cfg("pagecount") - cfg("listsidelong")) <= pEnd And pEnd < cfg("pagecount") Then
				For i = (pEnd+1) To cfg("pagecount")
					pListEnd = pListEnd & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			End If
		End If
		If cfg("pageindex") > 1 Then
			pFirst = " <a href="""&Replace(cfg("link"),"*","1")&""">" & cfg("first") & "</a> "
			pPrev = " <a href="""&Replace(cfg("link"),"*",cfg("pageindex")-1)&""">" & cfg("prev") & "</a> "
		Else
			pFirst = " <span class="""&cfg("disabledclass")&""">" & cfg("first") & "</span> "
			pPrev = " <span class="""&cfg("disabledclass")&""">" & cfg("prev") & "</span> "
		End If
		If cfg("pageindex") < cfg("pagecount") Then
			pLast = " <a href="""&Replace(cfg("link"),"*",cfg("pagecount"))&""">" & cfg("last") & "</a> "
			pNext = " <a href="""&Replace(cfg("link"),"*",cfg("pageindex")+1)&""">" & cfg("next") & "</a> "
		Else
			pLast = " <span class="""&cfg("disabledclass")&""">" & cfg("last") & "</span> "
			pNext = " <span class="""&cfg("disabledclass")&""">" & cfg("next") & "</span> "
		End If
		Select Case LCase(cfg("jump"))
			Case "input"
				pJumpValue = "this.value"
				pJump = "<input type=""text"" size=""3"" title=""请输入要跳转到的页数并回车""" & Easp_IIF(cfg("jumpplus")="",""," "&cfg("jumpplus"))
				pJump = pJump & " onkeydown=""javascript:if(event.charCode==13||event.keyCode==13){if(!isNaN(" & pJumpValue & ")){"
				pJump = pJump & Easp_IIF(cfg("jumpaction")="",Easp_IIF(Lcase(Left(cfg("link"),11))="javascript:",Replace(Mid(cfg("link"),12),"*",pJumpValue),"document.location.href='" & Replace(cfg("link"),"*","'+" & pJumpValue & "+'") & "';"),Replace(cfg("jumpaction"),"*", pJumpValue))
				pJump = pJump & "}return false;}"" />"
			Case "select"
				pJumpValue = "this.options[this.selectedIndex].value"
				pJump = "<select" & Easp_IIF(cfg("jumpplus")="",""," "&cfg("jumpplus")) & " onchange=""javascript:"
				pJump = pJump & Easp_IIF(cfg("jumpaction")="",Easp_IIF(Lcase(Left(cfg("link"),11))="javascript:",Replace(Mid(cfg("link"),12),"*",pJumpValue),"document.location.href='" & Replace(cfg("link"),"*","'+" & pJumpValue & "+'") & "';"),Replace(cfg("jumpaction"),"*",pJumpValue))
				pJump = pJump & """ title=""请选择要跳转到的页数""> "
				If cfg("jumplong")=0 Then
					For i = 1 To cfg("pagecount")
						pJump = pJump & "<option value=""" & i & """" & Easp_IIF(i=cfg("pageindex")," selected=""selected""","") & ">" & i & "</option> "
					Next
				Else
					pJumpLong = Int(cfg("jumplong") / 2)
					pJumpStart = Easp_IIF(cfg("pageindex")-pJumpLong<1, 1, cfg("pageindex")-pJumpLong)
					pJumpStart = Easp_IIF(cfg("pagecount")-cfg("pageindex")<pJumpLong, pJumpStart-(pJumpLong-(cfg("pagecount")-cfg("pageindex")))+1, pJumpStart)
					pJumpStart = Easp_IIF(pJumpStart<1,1,pJumpStart)
					j = 1
					For i = pJumpStart To cfg("pageindex")
						pJump = pJump & "<option value=""" & i & """" & Easp_IIF(i=cfg("pageindex")," selected=""selected""","") & ">" & i & "</option> "
						j = j + 1
					Next
					pJumpLong = Easp_IIF(cfg("pagecount")-cfg("pageindex")<pJumpLong, pJumpLong, pJumpLong + (pJumpLong-j)+1)
					pJumpEnd = Easp_IIF(cfg("pageindex")+pJumpLong>cfg("pagecount"), cfg("pagecount"), cfg("pageindex")+pJumpLong)
					For i = cfg("pageindex")+1 To pJumpEnd
						pJump = pJump & "<option value=""" & i & """>" & i & "</option> "
					Next
				End If
				pJump = pJump & "</select>"
		End Select
		tmpStr = Replace(tmpStr,"{recordcount}",cfg("recordcount"))
		tmpStr = Replace(tmpStr,"{pagecount}",cfg("pagecount"))
		tmpStr = Replace(tmpStr,"{pageindex}",cfg("pageindex"))
		tmpStr = Replace(tmpStr,"{pagesize}",cfg("pagesize"))
		tmpStr = Replace(tmpStr,"{list}",pList)
		tmpStr = Replace(tmpStr,"{liststart}",pListStart)
		tmpStr = Replace(tmpStr,"{listend}",pListEnd)
		tmpStr = Replace(tmpStr,"{first}",pFirst)
		tmpStr = Replace(tmpStr,"{prev}",pPrev)
		tmpStr = Replace(tmpStr,"{next}",pNext)
		tmpStr = Replace(tmpStr,"{last}",pLast)
		tmpStr = Replace(tmpStr,"{jump}",pJump)
		Set cfg = Nothing
		Pager = vbCrLf & tmpStr & vbCrLf
	End Function
	'配置分页样式
	Public Sub SetPager(ByVal PagerName, ByVal PagerHtml, ByRef PagerConfig)
		If PagerName = "" Then PagerName = "default"
		If Not Easp_isN(PagerHtml) Then iPageDic.item(PagerName&"_html") = PagerHtml
		If Not Easp_isN(PagerConfig) Then iPageDic.item(PagerName&"_config") = PagerConfig
	End Sub
	'调用分页样式
	Public Function GetPager(ByVal PagerName)
		If PagerName = "" Then PagerName = "default"
		GetPager = Pager(iPageDic(PagerName&"_html"),iPageDic(PagerName&"_config"))
	End Function
	'取得当前页码
	Private Function GetCurrentPage()
		Dim rqParam, thisPage : thisPage = 1
		rqParam = Request.QueryString(iPageParam)
		If isNumeric(rqParam) Then
			If Int(rqParam) > 0 Then thisPage = Int(rqParam)
		End If
		GetCurrentPage = thisPage
	End Function
	'返回除去页码的当前URL参数
	Private Function GetRQ(pageNumer)
		Dim tmpStr,rq : tmpStr = ""
		For Each rq In Request.QueryString()
			If rq<>iPageParam Then tmpStr = tmpStr & "&" & rq & "=" & Server.UrlEncode(Request.QueryString(rq))
		Next
		GetRQ = Request.ServerVariables("SCRIPT_NAME") & "?" & Easp_IIF(tmpStr="","",Mid(tmpStr,2)&"&") & iPageParam & "=" & Easp_IIF(pageNumer=0,"",pageNumer)
	End Function
End Class
%>