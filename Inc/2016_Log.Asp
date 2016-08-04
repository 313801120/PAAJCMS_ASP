<%
'************************************************************
'作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-08-04
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'日志文件

'错误日志
Function errorLog(content)
    If openErrorLog = True Then
        Call rw(content) 
    End If 
End Function 

'写入系统操作日志
Function writeSystemLog(tableName, msgStr)
    Dim logFile, s, url, ip, addDateTime,logDir
	logDir= adminDir & "/data/systemLog/" 
	call createDirFolder(logDir)		'创建文件夹
    logFile =logDir & "/" & format_Time(Now(), 2) & ".txt" 
    url = ADSql(getThisUrlFileParam())
    addDateTime = format_Time(Now(), 1) 
    ip = getIP() 
    If InStr(openWriteSystemLog, "|txt|") > 0 Then
        s = s & "姓名：" & Session("adminusername") & vbCrLf 
        s = s & "表：" & tableName & vbCrLf 
        s = s & "信息：" & msgStr & vbCrLf 
        s = s & "网址：" & url & vbCrLf 
        s = s & "时间：" & addDateTime & vbCrLf 
        s = s & "IP：" & ip & vbCrLf 
        s = s & "------------------------" & vbCrLf 
        Call createAddFile(logFile, s) 
    'call echo(logfile,"log")
    End If 

    If InStr(openWriteSystemLog, "|txt|") > 0 Then
        Call OpenConn()
		'判断表存在
		if instr(getHandleTableList(),"|"&  db_PREFIX & "systemlog" &"|")>0 then
			conn.Execute("insert into " & db_PREFIX & "SystemLog (tablename,msgstr,url,adminname,ip,adddatetime) values('" & tableName & "','" & msgStr & "','" & url & "','" & Session("adminusername") & "','" & ip & "','" & addDateTime & "')")
		end if
    End If 

End Function 

%>
