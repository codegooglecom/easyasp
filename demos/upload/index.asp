<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/easp/easp.asp" --><%
Easp.Debug = True
'EasyASP 上传类 Demo，此例子为UTF-8编码
Dim random, jsonFile, i, j, f
'载入核心类
Easp.Use "Upload"
'使用无组件进度条（默认为不使用）
Easp.Upload.UseProgress = True
'保存进度条数据临时文件的目录（默认为/__uptemp）
'Easp.Upload.ProgressPath = "/__uptemp"
'点击上传后执行
If Easp.Get("act") = "upload" Then
	'仅允许上传文件类型(建议先在客户端判断)
	Easp.Upload.Allowed = "exe|jpg|gif|png"
	'禁止上传的文件类型，如果设置了仅允许上传文件类型，则此设置不生效(建议先在客户端判断)
	'Easp.Upload.Denied = "exe|msi|bat|cmd|asp|asa"
	'单个文件最大允许值，单位为KB（如果是图片建议在客户端判断）
	Easp.Upload.FileMaxSize = 10240
	'全部文件最大允许值，单位为KB
	Easp.Upload.TotalMaxSize = 1024*30
	'上传文件保存路径，可用<>带日期标志（参见Easp.DateTime）按日期建立相应文件夹，
	Easp.Upload.SavePath = "uploadfiles/<yyyy>/<mm>/"
	'保存时使用随机文件名
	Easp.Upload.Random = True
	'获取上传的唯一KEY用于生成进度条数据Json文件给js调用
	Easp.Upload.Key = Easp.Get("json")
	'Easp.we ""
	Easp.Upload.StartUpload()
	'保存全部上传文件
	Easp.Upload.SaveAll
	Easp.WC "<ul>"
	For Each i In Easp.Upload.Form
		Easp.WC "<li>表单项 '" & i & "' 的值 : " & Easp.Upload.Form(i) & "</li>"
	Next
	Easp.WC "</ul>"
	Easp.WN "可上传个 " & Easp.Upload.File.Count & " 文件，本次上传成功了 " & Easp.Upload.Count & " 个文件。成功上传的文件信息如下："
	Easp.WN "================="
	For Each j In Easp.Upload.File
		Set f = Easp.Upload.File(j)
		If f.Size>0 Then
			Easp.WN "文件原位置：" & f.Client
			Easp.WN "文件原目录：" & f.OldPath
			Easp.WN "文件大小：" & f.Size
			Easp.WN "文件名称：" & f.Name
			Easp.WN "文件扩展名：" & f.Ext
			Easp.WN "文件类型：" & f.MIME
			Easp.WN "新路径："& f.NewPath
			Easp.WN "新名称："& f.NewName
			Easp.WN "================="
		End If
	Next
End If
'生成本次上传的唯一KEY
random = Easp.Upload.GenKey
'获取给js使用的Json文件的地址
jsonFile = Easp.Upload.ProgressFile(random)
Set Easp = Nothing
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>EasyAsp Upload Demo</title>
<script type="text/javascript" src="../jquery-1.4.1.min.js"></script>
<style type="text/css">
/*.progress {
    position: absolute;
    padding: 4px;
    top: 50;
    left: 400;
    font-family: Verdana, Helvetica, Arial, sans-serif;
    font-size: 12px;
    width: 250px;
    height:100px;
    background: #FFFBD1;
    color: #3D2C05;
    border: 1px solid #715208;
    -moz-border-radius: 5px;
}
.progress table,.progress td{
  font-size:9pt;
}
.Bar{
  width:100%;
    height:15px;
    background-color:#CCCCCC;
    border: 1px inset #666666;
    margin-bottom:4px;
}
.ProgressPercent{
    font-size: 9pt;
    color: #000000;
    height: 15px;
    position: absolute;
    z-index: 20;
    width: 100%;
    text-align: center;
}
.ProgressBar{
  background-color:#91D65C;
    width:1px;
    height:15px;
}*/
/*表单样式*/
.upload {font-size:14px; font-family:Tahoma;width:500px; padding:0 20px;}
.upload p{ padding:0; margin:0 0 10px 0;}
.upload input{ font-size:12px;font-family:Tahoma; padding:4px;}
.upload input.ipt{ width:436px;}
.upload input.ipts{ width:180px;}
.upload .btns{ padding:10px 0;}
/*进度条样式*/
.upload #formUpload{position:relative;}
.upload .progress { font-size:12px;position:absolute;top:80px;left:256px;height:130px;width:240px; background-color:#EEE;}
.upload .progress .txt{ line-height:130px; text-align:center; display:block;}
.upload .progress .info{padding:10px 0;text-align:center; line-height:1.5em;}
</style>
</head>
<body>
<fieldset class="upload"><legend>Easp.Upload上传文件示例</legend>
<form id="formUpload" method="post" enctype="multipart/form-data">
	<p>昵 称：<input type="text" name="nick" class="ipt" /></p>
	<p>密 码：<input type="password" name="pwd" class="ipt" /></p>
	<p>附件1：<input name="file1" type="file" /></p>
	<p>附件2：<input name="file2" type="file" /></p>
	<p>附件3：<input name="file3" type="file" /></p>
	<p>附件3：<input name="file4" type="file" /></p>
	<div class="btns"><input type="submit" id="btnSubmit" value="确认提交"/></div>
    <div id="progress" class="progress">
    	<!--<span class="txt">请选择一个或多个要上传的文件！</span> -->
        <div class="info">
        	<strong>正在上传，请稍候…</strong><br />
            总大小： <span>0 KB</span> / 已上传： <span>0 KB</span><br />
            总共时间： <span>00:00:00</span><br />
            剩余时间： <span>00:00:00</span>
        </div>
    </div>
</form>
</fieldset>
<!--<div id="progress" style="display:none;" class="progress">
    <div class="bar">
        <div id="uploadPercent" class="ProgressPercent">0%</div>
        <div id="uploadProgressBar" class="ProgressBar"></div>
    </div>
    <table border="0" cellspacing="0" cellpadding="2">
        <tr>
            <td>已经上传</td>
            <td>:</td>
            <td id="uploadSize">&nbsp;</td>
        </tr>
        <tr>
            <td>上传速度</td>
            <td>:</td>
            <td id="uploadSpeed">&nbsp;</td>
        </tr>
        <tr>
            <td>共需时间</td>
            <td>:</td>
            <td id="uploadTotalTime">&nbsp;</td>
        </tr>
        <tr>
            <td>剩余时间</td>
            <td>:</td>
            <td id="uploadRemainTime">&nbsp;</td>
        </tr>
    </table>
</div> -->
</body>
<script language="javascript">
$('#formUpload').submit(function(){
	var flag = false;
	$(this).find(':file').each(function(){
		if ($(this).val()!=''){
			flag = true;
			return false;
		}
	});
	if (!flag) {
		alert('请至少上传一个文件！');
		return false;
	} else {
		//在Form的action中加入上传的唯一KEY
		this.action = '?act=upload&json=<%=random%>';
		//显示进度条
		startProgress('<%=jsonFile%>');
		return true;
	}
});

//显示进度条
function startProgress(path){
	//0.5秒后启动进度条
	setTimeout('readProgress("' + path + '")',500);
}
//读取进度条
function readProgress(path){
	var percent = 0;
	try{
		//Ajax读取
		$.get(path,{rnd:Math.floor(Math.random()*10000)},function(d){
			//解析Json
			var progress = eval('('+d+')');
			//已上传大小 和 总大小
			$('#uploadSize').text(progress.uploaded +' / '+ progress.total);
			//上传速度
			$('#uploadSpeed').text(progress.speed);
			//上传总共需要时间（估计值）
			$('#uploadTotalTime').text(progress.totaltime);
			//上传剩余时间
			$('#uploadRemainTime').text(progress.remaintime);
			//上传百分比，更新显示状态
			percent = progress.percent;
			$('#uploadPercent').text(percent+'%');
			$('#uploadProgressBar').width(percent+'%');
		});
	} catch(e){ }
	//上传如未完成继续刷新(时间为0.5秒刷新一次)
	if (percent<100){
		setTimeout('readProgress("'+path+'")',500);
		//显示进度条
		$('#progress').show();
	}
}
</script>
</html>
