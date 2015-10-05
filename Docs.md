### EasyASP V2.0 方法列表 ###

#### 1.公共方法类(此类包含ASP(VBScript)中常用的公共函数集) ####
```
====字符串处理====
Easp.W|WC|WN|WE   输出字符串及简单断点调试
Easp.IsN          判断多种类型目标是否为空
Easp.JS           输出JavaScript代码 
Easp.Alert        弹出js消息框并返回上页 
Easp.AlertUrl     弹出js消息框并跳转到新页 
Easp.JsEncode     转换字符串为安全的JavaScript字符串 
Easp.CutString    截取字符串并以自定义符号代替被截部分 
Easp.HtmlEncode   HTML加码函数 
Easp.HtmlDecode   HTML解码函数 
Easp.HtmlFilter   过滤HTML标签 
Easp.Escape       用Unicode编码特殊字符 
Easp.UnEscape     把用Unicode编码的特殊字符解码为普通字符 
Easp.RandStr      生成指定长度和范围的随机字符串 
Easp.R            安全获取传入值并转换为SQL安全字符串 
Easp.Ra           安全获取传入值并在错误时弹出js消息框 
Easp.RF           安全获取传入的表单值 
Easp.RFa          安全获取传入的表单值(有警告) 
Easp.RQ           安全获取传入的URL参数值 
Easp.RQa          安全获取传入的URL参数值(有警告) 
Easp.Test         根据正则表达式验证数据合法性 
Easp.RegReplace   根据正则表达式替换文本内容 
```

```
====时间日期处理====
Easp.DateTime      按指定格式输出日期和时间 
Easp.DiffDay       返回一个日期时间变量和现在相比相差的小时数 
Easp.DiffHour      返回一个日期时间变量和现在相比相差的天数 
Easp.GetScriptTime 根据时间戳返回精确到毫秒的脚本执行时间 
```

```
====数值处理====
Easp.Rand         生成一个随机数 
Easp.toNumber     转换数字为指定小数位数的格式 
Easp.toPrice      转换数字为货币格式 
Easp.toPercent    转换数字为百分比格式 
```

#### 2.数据库操作类(EasyASP中简化对数据库的操作的数据库操作类) ####
```
====属性====
Easp.db.dbConn          (读写) 设置和获取当前数据库连接对象 
Easp.db.DatabaseType    (只读) 查询当前使用的数据库类型 
Easp.db.Debug           (读写) 设置和查询错误调试开关 
Easp.db.dbErr           (只读) 查询错误信息 
Easp.db.PageCount       (只读) 查询分页记录集总页数 
Easp.db.PageIndex       (只读) 查询分页记录集当前页码 
Easp.db.PageParam       (只写) 设置默认分页页码的URL参数名 
Easp.db.PageRecordCount (只读) 查询分页记录集总记录数 
Easp.db.PageSize        (读写) 设置和查询默认分页每页记录数 
Easp.db.PageSpName      (只写) 设置默认分页存储过程名称 
```

```
====方法====
Easp.db.OpenConn         根据模板建立数据库连接对象 
Easp.db.CreatConn        根据自定义字符串建立数据库连接对象 
Easp.db.AutoId           根据表名获取自动编号 
Easp.db.GetRecord        根据条件获取记录集 
Easp.db.wGetRecord       输出获取记录集的SQL语句 
Easp.db.GetRecordBySql   根据SQL语句获取记录集
Easp.db.GetRecordDetail  根据条件获取指定记录的详细数据 
Easp.db.AddRecord        添加一条新的记录 
Easp.db.wAddRecord       输出添加新记录的SQL语句 
Easp.db.UpdateRecord     根据条件更新记录 
Easp.db.wUpdateRecord    输出更新记录的SQL语句 
Easp.db.DeleteRecord     根据条件删除记录 
Easp.db.wDeleteRecord    输出删除记录的SQL语句 
Easp.db.ReadTable        根据条件获取指定字段数据 
Easp.db.Json             根据记录集生成Json格式数据 
Easp.db.doSP             调用一个SQL存储过程并返回数据 
Easp.db.GetPageRecord    初始化分页数据并得到记录集 
Easp.db.SetPager         设置分页导航列表样式 
Easp.db.GetPager         获取分页导航列表样式 
Easp.db.Pager            即时生成分页导航列表样式 
Easp.db.Rand             生成一个不重复的随机数 
Easp.db.RandStr          生成一个不重复的随机字符串 
Easp.db.C                关闭记录集并释放资源 
Easp.db.GR               GetRecord 方法的简写 
Easp.db.wGR              wGetRecord 方法的简写 
Easp.db.GRS              GetRecordBySql 方法的简写 
Easp.db.GRD              GetRecordDetail 方法的简写 
Easp.db.AR               AddRecord 方法的简写 
Easp.db.wAR              wAddRecord 方法的简写 
Easp.db.UR               UpdateRecord 方法的简写 
Easp.db.wUR              wUpdateRecord 方法的简写 
Easp.db.DR               DeleteRecord 方法的简写
Easp.db.wDR              wDeleteRecord 方法的简写 
Easp.db.RT               ReadTable 方法的简写 
Easp.db.GPR              GetPageRecord 方法的简写 
```