<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-13
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'系统信息  (2014,05,27)

function getObjectArrayCofnig()
	dim Sot(53, 2)
	Sot(0, 0) = "Scripting.FileSystemObject"
	  Sot(0, 2) = "文件操作组件"
	Sot(1, 0) = "wscript.shell"
	  Sot(1, 2) = "命令行执行组件"
	Sot(2, 0) = "ADOX.Catalog"
	  Sot(2, 2) = "ACCESS建库组件"
	Sot(3, 0) = "JRO.JetEngine"
	  Sot(3, 2) = "ACCESS压缩组件"
	Sot(4, 0) = "Scripting.Dictionary"
	  Sot(4, 2) = "数据流上传辅助组件"
	Sot(5, 0) = "Adodb.connection"
	  Sot(5, 2) = "数据库连接组件"
	Sot(6, 0) = "Adodb.Stream"
	  Sot(6, 2) = "数据流上传组件"
	Sot(7, 0) = "SoftArtisans.FileUp"
	  Sot(7, 2) = "SA-FileUp 文件上传组件"
	Sot(8, 0) = "LyfUpload.UploadFile"
	  Sot(8, 2) = "刘云峰文件上传组件"
	Sot(9, 0) = "Persits.Upload.1"
	  Sot(9, 2) = "ASPUpload 文件上传组件"
	Sot(10, 0) = "JMail.SmtpMail"
	  Sot(10, 2) = "JMail 邮件收发组件"
	Sot(11, 0) = "CDONTS.NewMail"
	  Sot(11, 2) = "虚拟SMTP发信组件"
	Sot(12, 0) = "SmtpMail.SmtpMail.1"
	  Sot(12, 2) = "SmtpMail发信组件"
	Sot(13, 0) = "JMail.Message"
	  Sot(13, 2) = "JMail.Message发信组件" 
	  
	Sot(14, 0) = "Microsoft.XMLHTTP"
	  Sot(14, 2) = "数据传输组件" 
	Sot(15, 0) = "NetBox.HttpServer"
	  Sot(15, 2) = "NetBox.HttpServer哈哈" 
	Sot(16, 0) = "ADODB.Stream.6.0"
	  Sot(16, 2) = "ADODB.Stream.6.0" 
	  
	Sot(20, 0) = "MSXML2.DOMDocument"
	  Sot(20, 2) = "xml处理" 
	  
	  
	  
	  
	  
	Sot(31, 0) = "yunsun.AspCode"
	  Sot(31, 2) = "云祥孙ASP代码格式化" 
	
	Sot(32, 0) = "yunsun.AspToPhpClass"
	  Sot(32, 2) = "云祥孙ASP转PHP" 
	
	Sot(33, 0) = "yunsun.calldll"
	  Sot(33, 2) = "云祥孙回调" 
	  
	Sot(34, 0) = "yunsun.clsAccessPass"
	  Sot(34, 2) = "云祥孙Access密码破解" 
	
	Sot(35, 0) = "yunsun.dllCommand"
	  Sot(35, 2) = "云祥孙公共" 
	
	Sot(36, 0) = "yunsun.dllSafe"
	  Sot(36, 2) = "云祥孙安全" 
	
	Sot(37, 0) = "yunsun.GPS"
	  Sot(37, 2) = "云祥孙图片处理" 
	  
	Sot(38, 0) = "yunsun.dllFZ"
	  Sot(38, 2) = "云祥孙仿站类" 
	  
	Sot(40, 0) = "WScript.Shell.1"
	  Sot(40, 2) = "WScript.Shell.1" 
	  
	  
	  
	  
    Sot(50, 0) = "yunsun.calldll_NONO"
    Sot(50, 2) = "不存在的组件（只为测试）"
	getObjectArrayCofnig=sot
end function

'显示对象  =showObject()
Function showObject()
	Dim i, t, s, isObj, okC, errC,en,cn,sot
	
    sot=getObjectArrayCofnig()
	
	For i = 0 To ubound(Sot) 
		en=Sot(i, 0)
		cn=Sot(i, 2)
		if en<>"" then
			If checkObject(en)=true Then
				s=" √"
				okC=okC & en & "  " & s & "   " & cn & vbcrlf
			Else
				s=" ×"
				errC=errC & en & "  " & s & "   " & cn & vbcrlf
			End If
		end if
	Next
	showObject = okC & errC
End Function
'显示组件第二种
Function showObject2()
	dim sot,s,c,i,en,cn,okC,errC
    sot=getObjectArrayCofnig()
  c="<br><table width='80%' bgcolor='menu' border='0' cellspacing='1' cellpadding='0' align='center'>"
  c=c&"<tr><td height='20' colspan='3' align='center' bgcolor='menu'>服务器组件信息</td></tr>"
  c=c&"<tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器名</td><td bgcolor='#FFFFFF'>&nbsp;</td><td bgcolor='#FFFFFF'>"&request.serverVariables("SERVER_NAME")&"</td></tr>"
  c=c&"<form method=post action='http://www.ip138.com/index.asp' name='ipform' target='_blank'><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器IP</td><td bgcolor='#FFFFFF'>&nbsp;</td><td bgcolor='#FFFFFF'>"
  c=c&"<input type='text' name='ip' size='15' value='"&Request.ServerVariables("LOCAL_ADDR")&"'> <input type='submit' value='查询'><input type='hidden' name='action' value='2'></td></tr></form>"
  c=c&"<tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器时间</td><td bgcolor='#FFFFFF'>&nbsp;</td><td bgcolor='#FFFFFF'>"&now&"&nbsp;</td></tr>"
  c=c&"<tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器CPU数量</td><td bgcolor='#FFFFFF'>&nbsp;</td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("NUMBER_OF_PROCESSORS")&"</td></tr>"
  c=c&"<tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器操作系统</td><td bgcolor='#FFFFFF'>&nbsp;</td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("OS")&"</td></tr>"
  c=c&"<tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>WEB服务器版本</td><td bgcolor='#FFFFFF'>&nbsp;</td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("SERVER_SOFTWARE")&"</td></tr>"
 
	For i = 0 To ubound(Sot) 
		en=Sot(i, 0)
		cn=Sot(i, 2)
		if en<>"" then
			If checkObject(en)=true Then
				s=" √" 
			okC=okC & "<tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>"& en &"</td><td bgcolor='#FFFFFF'>"& s &"</td><td bgcolor='#FFFFFF' align=left>"& cn &"</td></tr>"
			Else
				s=" ×" 
			errC=errC & "<tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>"& en &"</td><td bgcolor='#FFFFFF'>"& s &"</td><td bgcolor='#FFFFFF' align=left>"& cn &"</td></tr>"
			End If
		end if
	Next
  
	showObject2=c & okC & errC
End Function

'检测对象
function checkObject(objecName)
	dim t
	on error resume next
	Set T=CreateObject(objecName) 
	If -2147221005 <> Err and 429<>Err Then
	  checkObject=true
	Else
		checkObject=false
	  Err.Clear 
	End If
	Set T=Nothing  
end function


'操作系统版本
function operationSystem()
    dim httpAgent, systemVer 
    httpAgent = request.serverVariables("HTTP_USER_AGENT") 
    if inStr(httpAgent, "NT 5.2") > 0 then
        systemVer = "Windows Server 2003" 
    elseIf inStr(httpAgent, "NT 5.1") > 0 then
        systemVer = "Windows XP" 
    elseIf inStr(httpAgent, "NT 5") > 0 then
        systemVer = "Windows 2000" 
    elseIf inStr(httpAgent, "NT 4") > 0 then
        systemVer = "Windows NT4" 
    elseIf inStr(httpAgent, "4.9") > 0 then
        systemVer = "Windows ME" 
    elseIf inStr(httpAgent, "98") > 0 then
        systemVer = "Windows 98" 
    elseIf inStr(httpAgent, "95") > 0 then
        systemVer = "Windows 95" 
    else
        systemVer = httpAgent 
    end if 
    operationSystem = httpAgent 
end function 
'检测是否为手机浏览
function checkMobile()
    checkMobile = false 
	
	dim regex,sagent
	if 1=1 then
	set regex = createObject("VBscript.RegExp")
	regex.ignorecase = true 
	regex.global = true 
	regex.pattern = "mini 9.5|vx1000|lge |m800|e860|u940|ux840|compal|wireless| mobi|ahong|lg380|lgku|lgu900|lg210|lg47|lg920|lg840|lg370|sam-r|mg50|s55|g83|t66|vx400|mk99|d615|d763|el370|sl900|mp500|samu3|samu4|vx10|xda_|samu5|samu6|samu7|samu9|a615|b832|m881|s920|n210|s700|c-810|_h797|mob-x|sk16d|848b|mowser|s580|r800|471x|v120|rim8|c500foma:|160x|x160|480x|x640|t503|w839|i250|sprint|w398samr810|m5252|c7100|mt126|x225|s5330|s820|htil-g1|fly v71|s302|-x113|novarra|k610i|-three|8325rc|8352rc|sanyo|vx54|c888|nx250|n120|mtk |c5588|s710|t880|c5005|i;458x|p404i|s210|c5100|teleca|s940|c500|s590|foma|samsu|vx8|vx9|a1000|_mms|myx|a700|gu1100|bc831|e300|ems100|me701|me702m-three|sd588|s800|8325rc|ac831|mw200|brew |d88|htc\/|htc_touch|355x|m50|km100|d736|p-9521|telco|sl74|ktouch|m4u\/|me702|8325rc|kddi|phone|lg |sonyericsson|samsung|240x|x320vx10|nokia|sony cmd|motorola|up.browser|up.link|mmp|symbian|smartphone|midp|wap|vodafone|o2|pocket|kindle|mobile|psp|treo|iris|3g_t|windows ce|opera mobi|windows ce; smartphone;|windows ce; iemobile|ipod|iphone|android|opera mini|blackberry|palm os|palm|hiptop|avantgo|fennec|plucker|xiino|blazer|elaine|iris|3g_t|windows ce|opera mobi|windows ce; smartphone;|windows ce; iemobile" 
	
	sagent = request.ServerVariables("HTTP_USER_AGENT") & "" 
	if sagent <> "" then  
		if regex.test(sagent) then 
			checkMobile = true 
		else
			checkMobile = false
		end if 
	end if 
	else
	'这是旧版
    if request.serverVariables("HTTP_X_WAP_PROFILE") <> "" then
        checkMobile = true 
    end if
	end if 
end function 

'获得IIS版本号
function getIISVersion()
    getIISVersion = request.serverVariables("SERVER_SOFTWARE") 
end function 

%> 

