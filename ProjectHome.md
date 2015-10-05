### EasyASP v2.2 能做什么？ ###
> [![](http://easyasp.googlecode.com/files/EasyASP.v2.2.Map.s.jpg)](http://www.easyasp.cn)

### 关于 EasyASP ###

**EasyASP 是一个简单方便的用来快速开发ASP程序的类库。EasyASP 包含完善的全参数化查询多数据库操作、高效Json数据生成与解析、各种字符串及日期处理函数、功能强大动态数组处理、领先的文件系统处理、远程文件及XML文档处理、内存缓存和文件缓存处理、简单实用的模板引擎等等丰富的功能。而为了解决ASP调试不方便的问题，EasyASP 推出了独创的控制台调试功能以及丰富的异常信息显示，能让你开发 ASP 程序时最大程度的从错误调试的纷繁中解放出来。EasyASP 目前提供下载的是VBScript版本。**

**EasyASP v3 现已推出测试版本，最新消息请查看EasyASP官方网站  http://www.easyasp.cn**

#### EasyASP v2.1 的特点： ####
  * [数据库] 能方便的实现一个或多个数据库的增、删、改等控制操作。
  * [数据库] 对数据库字段进行操作时可以不用考虑字段值数据类型的差别(如文本字段不用加单引号)。
  * [数据库] 自带记录集分页和调用存储过程分页功能，拥有功能丰富的可完全自定义配置及调用。
  * [数据库] 能方便的执行带各种参数的MSSQL存储过程并返回多个值或多个记录集。
  * [数据库] 完善的数据库操作容错功能，能即时输出出错SQL语句方便调试。
  * [数据库] 在使用已经存在的数据库连接对象时能自动判断数据库类型。
  * [数据库] 专为Ajax设计的数据获取方式及输出Json格式数据。
  * [数据库] 能有效防止SQL注入。
  * [ASP](ASP.md) 自带大量的ASP通用过程及方法，简化大部分ASP操作。
  * [ASP](ASP.md) 完美实现ASP文件的动态载入，并支持无限级的ASP原生include。
  * [ASP](ASP.md) 自带数据类型验证及服务器端表单验证功能。
  * [ASP](ASP.md) 能轻松实现页面地址获取并对URL参数进行过滤以及替换。
  * [工具] 具有专为EasyASP开发的适用于Dreamweaver CS3 和 CS4 的代码高亮及代码提示扩展插件。
  * [工具] 具有完善的帮助手册及大量应用实例。
  * ……

### EasyASP V2.1 更新日志 @2009-08-31 by coldstone ###

**新增功能：**
  * 新增Easp.Include方法，完美实现了ASP的动态包含，且支持ASP源码中无限级层次的<!--#include...-->。
  * 新增Easp.GetInclude方法，用于获取ASP文件运行的结果或获取html文件等文本文件的源码。
  * 新增Easp.Charset属性，用于设置Easp.Include方法和Easp.getInclude方法载入文件的编码。
  * 新增Easp.ConfirmUrl方法，用于输出确认信息框并根据选择进行Url跳转。
  * 新增Easp.HtmlFormat方法，用于对html文本进行简单的格式化(仅转换空格和换行为可见)。
  * 新增Easp.RegReplaceM方法，用于正则替换的多行模式。
  * 新增Easp.RegMatch方法，用于正则匹配的编组捕获。
  * 新增Easp.IsInstall方法，用于检测系统是否安装了某个组件。
  * [db](db.md)新增Easp.db.QueryType属性，可设置用ADO的RecordSet还是Command方式获取记录集。
  * [db](db.md)新增Easp.db.GetRandRecord方法，用于取得指定数量的随机记录集。
  * [db](db.md)新增Easp.db.Exec方法，用于执行SQL语句或者返回Command方式查询的记录集。

**其他更新**
  * 优化Easp.DateTime方法，格式化为时间差时的显示更人性化。
  * 优化Easp.RandStr和Easp.db.RandStr方法，提供更强大更人性化的随机字符串和随机数生成功能。
  * 修正Easp.GetUrlWith方法第一个参数为空时生成的新URL出错的Bug。
  * 修正Easp.GetApp方法无法获取缓存数据的Bug。
  * 修正Easp.AlertUrl跳转前会继续执行服务器端代码的Bug。
  * 修正v2.1beta版中Easp.JsEncode和Easp.db.Json方法会报“类型不匹配”错误的Bug。
  * 修正v2.1beta版中Easp.RandStr和Easp.db.RandStr的一个Bug。
  * [db](db.md)优化Easp.db.AddRecord方法，现在仅当指定了ID字段的名称时才返回新增的记录ID号。(影响以前版本)
  * [db](db.md)修正分页下拉框中页面数量小于jumplong配置时出现负数的Bug。

### EasyASP v2.0 (2009-02-07更新) ###

**新增功能：**
  * 新增数据分页功能，可以根据多种方式实现较高性能的记录集分页，还可以对分页导航进行完全自定义的多个配置，并且可在多个配置间自由切换。还可以很方便的生成静态分页导航及Ajax分页导航。
  * 新增在使用已经存在的数据库连接对象时可以自动判断数据库类型的功能。
  * 新增Escape方法和UnEscape方法，用于将特殊字符编码和解码，可解决非UTF-8下的Ajax中文乱码问题。
  * 新增IIF方法，功能同IfThen方法一样，而且更符合大多数人的习惯。
  * 新增GetUrlWith方法，可以在getUrl方法的基础上加上新的Url参数和值，非常实用。
  * 新增regReplace方法，支持按正则表达式替换字符串内容。
  * 新增EasyASP的Dreamweaver代码提示及代码高亮扩展插件。

**其它更新：**
  * 将EasyASP原来的多个asp文件合并为了一个文件easp.asp以方便调用，但如果有需要的话仍然可以把数据库操作类EasyASP\_db单独出来使用。
  * 优化了isN方法，可以检测多种类型的数据是否为空。
  * 优化DateTime方法，增加英文月份及缩写格式，增加输出为如“3个小时前”等格式。
  * 优化RandStr及db.RandStr方法，可以指定获取的随机字符串的字符范围。
  * 将原来的Easp.Close方法更新为Easp.C方法。
  * 优化SetCookie、GetCookie、RemoveCookie方法，现在可以设置Cookie的集合、域、路径及安全。
  * 修正db.OpenConn方法中服务器密码不能包含“:”、“@”等特殊字符的Bug。
  * 再次优化db.AutoId方法，能更好的解决并发量大时的自动编号获取的问题。
  * 重新制作了EasyASP v2.0帮助手册，并添加了代码高亮功能。
  * 对代码进行了简单重构，减少了大量的冗余代码，并修正了其它一些小的Bug。

**更新说明：**
> EasyASP在开发之初就首先定位于Easy，所以在编写分页的时候也主要是考虑如何使用方便和简单。EasyASP的分页功能在性能上做了最大程度的努力，可以根据数组参数(和Easp.db.GetRecord方法相似)、SQL语句、记录集和存储过程来生成分页数据，对于MSSQL来说，根据数组条件和存储过程分页的效率是比较高的，而Access数据库的话则可以使用数组条件、SQL语句方式或记录集方式。在使用的方便性上EasyASP也采用了颠覆性的方式，采用了类似javascript中常用的Json方式的配置方法，而且您可以预先配置多个样式的分页导航，并在不同的地方直接调用事先配置的各种样式，轻松实现在一个页面中包含多种不同样式的分页(包括嵌套分页)，这一点俺非常有自信大家会喜欢这种方式的。 看几个用EasyASP生成的分页样式：
> > ![http://easyasp.googlecode.com/files/pager.jpg](http://easyasp.googlecode.com/files/pager.jpg)

> 在编写分页功能的过程中，偶然看到Jorkin写的《Kin\_Db\_Pager分页类》，受到了很大启发。EasyASP的分页导航多配置部分就是基于他的这个分页类的一些思想实现的，优化后的isN方法也是基本照搬他的isBlank函数，在此特别感谢一下。

> EasyASP中的用MSSQL存储过程分页功能中内置了一个默认的存储过程easp\_sp\_pager，该存储过程是使用了nzperfect编写的《单主键高效通用分页存储过程》，也在此特别感谢一下。

> 另外为了使用方便，这一次在更新 EasyASP v2.0 的时候特别制作了一个用于Dreamweaver CS3 和 CS4 的 EasyASP v2.0 代码提示和代码高亮的扩展插件，大家在用Dreamweaver编写程序的时候应该会非常有用的，就像下面这样：
> > ![http://easyasp.googlecode.com/files/EasyASP_v2_for_Dreamweaver.jpg](http://easyasp.googlecode.com/files/EasyASP_v2_for_Dreamweaver.jpg)

---


### 使用说明 ###

**1、使用方法：**

(1) Easp类的所有功能都已包含在easp.asp中，EasyASP v2.0只有一个easp.asp文件，所以只需要在页首引入该文件，如：

```
<!--#include file="inc/easp/easp.asp" -->
```

或：

```
<!--#include virtual="/inc/easp/easp.asp"-->
```

(2) 该类已经实例化，无需再单独实例化，直接使用Easp.前缀调用即可，如：

```
Easp.wn("Test String")  或  Easp.db.AutoId("Table:ID")
```

(3) 如要同时操作多个数据库，请实例化新的EasyASP\_db对象，如：

```
  Dim db : Set db = New EasyASP_db
  db.dbConn = db.OpenConn(0,dbase,server)
```

**2、参数约定：**

(1) 数组参数：由于VBScript不能使用动态参数，所以，在本类涉及到数据库数据的代码中，使用了Array(数组)来达到这一效果。本类中的部分参数可以使用数组(参数说明中有注明)，但使用数组时应参照以下格式：

```
Array("Field1:Value1", "Field2:True", "Field3:100")
```


> 对，有点像json的格式，如果涉及到变量，那就这样：

```
Array("Field1:" & Value1, "Field2:" & Value2, "Field3:" & Value3)
```

可以这样说，本类中的几乎所有与数据库字段相关的内容都可以用以上的数组格式来设置条件或者是获取内容，包括调用存储过程要传递的参数。而这个类里最大的优点就是在使用时不用去考虑字段的类型，在字段后跟一个冒号，接着跟上相应的值就行了。如果你经常手写ASP程序的话，你很快就会感受到运用这种方式的魅力，除了数据类型不用考虑之外，它也很方便随时添加和删除条件。这里举个例子说明这个用法：

比如添加新记录的方法：

```
Easp.db.AddRecord "Table", Array("FieldString:测试数据","FieldDate:"&Now(),"FieldBoolean:True","FieldInt:5874")
```

参数只有两个，一个是表名，另一个就是这样的数组参数，不用考虑数据类型。而且如果要改变数据库结构，修改上面的程序代码就非常简单了。

(2) 共用参数（用特殊符号分隔）: 也是考虑到要尽量减少参数，如果有些参数在很多时候都可以没有的话，那就没有必要专门为它增加一个参数。在本类里采用了特殊符号如冒号(:)分隔一个参数中的多个值来达到传递多个参数的效果。举几个例子说明一下，同时也可以预览一下采用本类的一些优势：

比如建立MSSQL数据库连接对象的方法：

```
Set Conn = Easp.db.Open(0,"Database","User:Password@ServerAddress")
```

这样应该更符合我们平时描述服务器地址的方式了。另外如果是Access数据库有密码则在上面的第3个参数中输入就行了，不用新增参数。

再比如获取记录集的方法：

```
Set rs = Easp.db.GetRecord("Table:FieldsA,FieldsB,FieldsC:20","ID > 10","ID Desc")
```

其中第1个参数中包含了表名，要取的字段和要取的记录数，因为字段和记录数很多时候是并不需要的，所以俺索性把参数也省略了，这样要记的参数要少很多滴。

再比如本类里有一个GetUrl()的获取本页面地址的方法，很多地方都见过是吧，但是本类里这个方法带一个参数，通过这个参数可以取得很多结果，看例子：

> 比如一个页面的实际地址为：
```
http://www.ambox.cn/public/news/index.asp?type=public&feed=on&page=23
```
> 接下来是使用不同参数返回的结果：
```
    方法                    返回结果
    GetUrl("")              http://www.ambox.cn/public/news/index.asp?type=public&feed=on&page=23
    GetUrl(0)               /public/news/index.asp
    GetUrl(1)               /public/news/index.asp?type=public&feed=on&page=23
    GetUrl(2)               /public/news/
    GetUrl("page")          /public/news/index.asp?page=23
    GetUrl("-page")         /public/news/index.asp?type=public&feed=on
    GetUrl(":")             /public/news/?type=public&feed=on&page=23
    GetUrl(":-feed,-page")  /public/news/?type=public
```

就是这样，可以方便的过滤URL参数。本类中灵活使用共用参数的地方还有很多，这也是EasyASP的一大特色，大家自己下载手册来看吧。


---


### EasyASP v1.5 (2008-10-22更新) ###

**新增功能：**
  * 将数据库控制类(原clsDbCtrl.asp)封装入Easp类，均通过Easp.db调用，也可独立使用。
  * 新增MSSQL存储过程调用方法，可灵活调用存储过程并返回返回值、记录集及出参。
  * 新增db.CreatConn方法，可以根据自定义的连接字符串连接数据库。
  * 新增db.Json方法，可以将数据库记录集按Json格式输出。
  * 新增db.Rand和db.RandStr方法，可以生成一个不重复的随机数或者随机字符串
  * 新增数据库操作各方法的简写方法，更节约书写代码时间。
  * 在Easp类中新增大量的实用方法，如安全获取值、防Sql注入、服务器端表单验证等。

**其它更新：**
  * 优化db.AutoId自动获取编号，效率提高20倍以上，数据量越大越明显。
  * 修改db.OpenConn数据库连接方法，更符合日常描述习惯。
  * 修改db.GetRecord取记录集方法，参数更少。修正条件使用数组报错的Bug。
  * 修改并优化db.DeleteRecord删除记录方法，目前只有两个参数了。
  * 修改了错误调试方法，增加Debug全局属性控制错误显示。

**更新说明：**

> 以前写了一个clsDbCtrl.asp数据库控制类，收到一些反馈，还有朋友发来邮件告诉我一些改进的方法，很感谢他们。而我在原帖的跟帖中看到一条留言说“有记参数的时间，SQL语句早都写完了”，更是直接指出了其中的尴尬，的确，尽管VBS没有arguments属性，但用太多的参数也不是个好主意。所以我花了些时间把这个类的许多代码都重写了一下，在保证功能只能更强不能更弱的前提下，一个方法最多只有3个参数了。另外新增加了一个调用MSSQL存储过程的方法，可以灵活的调用存储过程并根据需要返回一个或多个记录集、输出参数及返回值，当然，吸取教训了，这个方法只有两个参数。现在都封装在这个新的名叫EasyASP的家伙中了，顾名思义，无非是想一切都简单点。另外还有一个更尴尬的，那就是VBScript并不是面向对象的语言，所以这个类其实说穿了也只是一些过程和方法的封装，方便使用而已，所以其中大部分的方法和过程都可以提出来单独使用。当然，如果有需要，也可以把它封装成wsc或者dll组件使用。