<%
'************************************************************
'作者：云孙World(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-09-22
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'xml文件操作

'批量导入xls文件，并保存 Call BatchImportXlsFile("E:\E盘\WEB网站\网站案例\", "优秀网址大全2.txt")
function batchImportXlsFile(folderPath, saveXmlFilePath)
    dim i, s, s2, sql, connStr, tableList, splTable, tableName, tableFieldList, splField, fieldName, urlList, url 
    dim xlsFilePath, xlsFileList, splXls 

    xlsFileList = getFileFolderList(folderPath, true, "xls", "", "", "", "") 
    splXls = split(xlsFileList, vbCrLf) 
    for each xlsFilePath in splXls
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xlsFilePath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1';" 
        call CloseConn() 
		set conn = server.createobject("adodb.connection")
		set rs = createObject("Adodb.RecordSet")
		set rsx = createObject("Adodb.RecordSet")
		set rss = createObject("Adodb.RecordSet")
		set rst = createObject("Adodb.RecordSet")
		set rsd = createObject("Adodb.RecordSet")
		set tempRs = createObject("Adodb.RecordSet")
		set tempRs2 = createObject("Adodb.RecordSet")
		set rsTemp = createObject("Adodb.RecordSet")
		'Call Eerr("connStr",connStr)
		conn.open connStr 
		
		tableList = getTableList_Conn(conn) 
		splTable = split(tableList, vbCrLf) 
		for each tableName IN splTable
			if tableName <> "" then
				tableFieldList = getFieldList(tableName) 
				splField = split(tableFieldList, ",") 
				if uBound(splField) <> - 1 then
					'Call Echo(Ubound(SplField),TableFieldList):Doevents
		
					rsx.open "Select * From [" & tableName & "]", conn, 1, 1 
					'Call Echo("记录",Rsx.RecordCount)
					while not rsx.eof
						i = 0 
						for each fieldName in splField
							if fieldName <> "" then
								url = rsx(i) 
								'Call Echo(FieldName & "("& CheckUrl(Rsx(I)) &")", Rsx(I) )
								if checkUrl(url) = true and inStr(vbCrLf & urlList & vbCrLf, vbCrLf & url & vbCrLf) = false then
									urlList = urlList & url & vbCrLf 
									call echo(tableName, url) : doevents 
								end if 
								i = i + 1 
							end if 
						next 
					rsx.moveNext : wend : rsx.close 
				end if 
			end if 
		next 
    next
    urlList = CStr(urlList) 
    call echo("UrlList", urlList) 
    call WriteToFile(saveXmlFilePath, urlList, 1) 
'Call CreateFile("优秀网址大全.txt",UrlList)
end function 
%> 

