<%
'######################################################################
'## easp.config.asp
'## -------------------------------------------------------------------
'## EasyASP �����ļ�
'######################################################################

'������ȷ����'easp.asp'�ļ�����վ�е�·������"/"��ͷ:
Easp.BasePath = "/easp/"

'�����ļ����� (ͨ��Ϊ'GBK'����'UTF-8'):
Easp.CharSet = "GBK"

''�򿪿����ߵ���ģʽ��
'Easp.Debug = True

''������Cookies����:
'Easp.CookieEncode = False

''����FSO��������ƣ�������������޸Ĺ���:
'Easp.FsoName = "Scripting.FileSystemObject"

''������δ��������UTF-8�ļ���BOM��Ϣ(keep/remove/add)��
'Easp.FileBOM = "remove"

''�������ݿ����ӣ�
''Access:
'Easp.db.Conn = Easp.db.OpenConn(1,"/data/data.mdb","")
''MS SQL Server:
'Easp.db.Conn = Easp.db.OpenConn(0,"Data","sa:admin@127.0.0.1")
%>