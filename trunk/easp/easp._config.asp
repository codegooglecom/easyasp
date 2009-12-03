<%
'配置数据库连接
'Easp.Use("db")
'Easp.db.dbConn = Easp.db.OpenConn(1,"/data/data.mdb","")
Easp.db.dbConn = Easp.db.OpenConn(0,"EduFile","jpzxoa:jpzx_SQL_1860@192.168.133.2")
Easp.db.Debug = True
%>