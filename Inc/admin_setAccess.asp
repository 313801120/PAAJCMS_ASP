<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-03-13
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'��#jump#��trueΪtrue��Ϊ������ǰ
'��#top#��trueΪtrue��Ϊֹͣȫ��

'����callfile_setAccess�ļ�����
function callfile_setAccess()
    dim sType 
    sType = request("stype") 
    if sType = "backupDatabase" then
        call backupDatabase()                                '�������ݿ�
    elseif sType = "recoveryDatabase" then
        call recoveryDatabase()                              '�ָ����ݿ�
    else
        call eerr("setAccessҳ��û�ж���", request("stype")) 
    end if 
end function

'�ָ����ݿ�
function recoveryDatabase()
    dim backupDir, backupFilePath 
    dim content, s, splStr, tableName 
    call handlePower("�ָ����ݿ�") 
    backupDir = adminDir & "/Data/BackUpDateBases/" 
    backupFilePath = backupDir & "/" & request("databaseName") 
    if checkFile(backupFilePath) = false then
        call eerr("���ݿ��ļ�������", backupFilePath) 
    end if 
    content = getftext(backupFilePath) 
    splStr = split(content, "===============================" & vbCrLf) 
    for each s in splStr
        tableName = newGetStrCut(s, "table") 
        if tableName <> "" then
            conn.execute("delete from " & db_PREFIX & tableName) 
            call echo(tableName, nImportTXTData(s, tableName, "���")) 
        end if 
    next 
    call echo("�ָ����ݿ����", "") 
end function 

'�������ݿ�
function backupDatabase()
    dim isUnifyToFile, tableNameList, databaseTableNameList, fieldConfig, fieldName, fieldType, splField, fieldValue, nLen, isOK 
    dim splStr, splxx, tableName, s, c, backupDir, backupFilePath 
    call handlePower("�������ݿ�") 
    tableNameList = lcase(request("tableNameList"))                                 '�Զ��屸�����ݱ��б�
    isUnifyToFile = IIF(request("isUnifyToFile") = "1", true, false)                'ͳһ�ŵ�һ���ļ���
    databaseTableNameList = lcase(db_PREFIX & "webcolumn" & vbCrLf & getTableList()) '��db_PREFIX����ǰ�棬��Ϊ��������Ҫ�������ȡ
    nLen = len(db_PREFIX) 

    '�����Զ�����б�
    if tableNameList <> "" then
        splStr = split(tableNameList, "|") 
        for each tableName in splStr
            if instr(vbCrLf & databaseTableNameList & vbCrLf, vbCrLf & db_PREFIX & tableName & vbCrLf) > 0 then
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & db_PREFIX & tableName 
            end if 
        next 
        if c = "" then
            call eerr("�Զ��屸�ݱ���ȷ <a href=""javascript:history.go(-1)"">�������</a>", tableNameList) 
        end if 
        databaseTableNameList = c 
    end if 
    splStr = split(databaseTableNameList, vbCrLf) 
    c = "" 
    for each tableName in splStr
        tableName = trim(tableName) 
        isOK = true 
        '�ж�ǰ׺�Ƿ�һ��
        if nLen > 0 then
            if mid(tableName, 1, nLen) <> db_PREFIX then
                isOK = false 
            end if 
        end if 
        if isOK = true then
            fieldConfig = lcase(getFieldConfigList(tableName)) 
            call echo(tableName, fieldConfig) 
			'��@��jsp��ʾ@��try{
            rs.open "select * from " & tableName, conn, 1, 1 
            c = c & "��table��" & mid(tableName, len(db_PREFIX) + 1) & vbCrLf 
            while not rs.eof
                splField = split(fieldConfig, ",") 
                for each s in splField
                    if instr(s, "|") > 0 then
                        splxx = split(s & "|", "|") 
                        fieldName = splxx(0) 
                        fieldType = splxx(1) 
                        fieldValue = rs(fieldName) 
                        if fieldType = "numb" then
                            fieldValue = replace(replace(fieldValue, "True", "1"), "False", "0") 
                        end if 
                        '��̨�˵�
                        if tableName = db_PREFIX & "listmenu" and fieldName = "parentid" then
                            fieldValue = getListMenuName(fieldValue) 
                        '��վ��Ŀ
                        elseif tableName = db_PREFIX & "webcolumn" and fieldName = "parentid" then
                            fieldValue = getColumnName(fieldValue) 
                        end if 
                        if fieldValue <> "" then
                            if instr(fieldValue, vbCrLf) > 0 then
                                fieldValue = fieldValue & "��/" & fieldName & "��" 
                            end if 
                            c = c & "��" & fieldName & "��" & fieldValue & vbCrLf 
                        end if 
                    end if 
                next 
                c = c & "-------------------------------" & vbCrLf 
            rs.movenext : wend : rs.close 
			'��@��jsp��ʾ@��}catch(Exception e){} 
            c = c & "===============================" & vbCrLf 
        end if 
    next 
    backupDir = adminDir & "/Data/BackUpDateBases/" 
    backupFilePath = backupDir & "/" & format_Time(now(), 4) & ".txt" 
    call createDirFolder(backupDir) 
    call deleteFile(backupFilePath)                                                 'ɾ���ɱ����ļ�
    call createfile(backupFilePath, c)                                              '���������ļ�
    call hr() 
    call echo("backupDir", backupDir) 
    call echo("backupFilePath", backupFilePath) 
    call eerr("�������", "<a href='?act=displayLayout&templateFile=layout_manageDatabases.html&lableTitle=���ݿ�'>������� ���ݻָ�����</a>") 
end function 

'�������ݿ�����
sub resetAccessData()
    call handlePower("�ָ�ģ������")                                                '����Ȩ�޴���
    call OpenConn() 
    dim splStr, i, s, columnname, title, nCount, webdataDir 
    webdataDir = request("webdataDir") 
    if webdataDir <> "" then
        if checkFolder(webdataDir) = false then
            call eerr("��վ����Ŀ¼�����ڣ��ָ�Ĭ������δ�ɹ�", webdataDir) 
        end if 
    else
        webdataDir = "/Data/WebData/" 
    end if 

    '�޸���վ����
    call nImportTXTData(getftext(webdataDir & "/website.txt"), "website", "�޸�") 
    call batchImportDirTXTData(webdataDir, db_PREFIX & "WebColumn" & vbCrLf & getTableList()) '��webcolumn����Ϊwebcolumn�����µ������ݣ���Ϊ��̨��������Ҫ������20160711

    call echo("��ʾ", "�ָ��������") 
    call rw("<hr><a href='"& WEB_VIEWURL &"' target='_blank'>������ҳ</a> | <a href=""?"" target='_blank'>�����̨</a>") 



    call writeSystemLog("", "�ָ�Ĭ������" & db_PREFIX)                             'ϵͳ��־
end sub 

'����������Ӧ����Ϣ
function batchImportDirTXTData(webdataDir, tableNameList)
    dim folderPath, tableName, splStr, content, splxx, filePath, fileName, handleTableNameList 
    splStr = split(tableNameList, vbCrLf) 
    for each tableName in splStr
        if tableName <> "" then
            if db_PREFIX <> "" then
                tableName = mid(tableName, len(db_PREFIX) + 1) 
            end if 
            tableName = trim(lcase(tableName)) 
            '�жϱ� ���ظ�����
            if instr("|" & handleTableNameList & "|", "|" & tableName & "|") = false then
                handleTableNameList = handleTableNameList & tableName & "|" 

                folderPath = handlePath(webdataDir & "/" & tableName) 
                if checkFolder(folderPath) = true then
                    conn.execute("delete from " & db_PREFIX & tableName)                            'ɾ����ǰ��ȫ������
                    call echo("tableName", tableName) 
                    content = getDirAllFileList(folderPath, "txt") 
                    splxx = split(content, vbCrLf) 
                    for each filePath in splxx
                        fileName = getFileName(filePath) 
                        if filePath <> "" and inStr("_#", left(fileName, 1)) = false then
                            call echo(tableName, filePath) 
                            call nImportTXTData(getftext(filePath), tableName, "���") 
                            doevents 
                        end if 
                    next 
                end if 
            end if 
        end if 
    next 
end function 

'��������
function nImportTXTData(content, tableName, sType)
    dim fieldConfigList, splList, listStr, splStr, splxx, s, sql, nOK 
    dim fieldName, fieldType, fieldValue, addFieldList, addValueList, updateValueList 
    dim fieldStr,isJump,isStop
    tableName = trim(lcase(tableName))                                              '��
    '��������Ϊ�˴�GitHub����ʱ����vbcrlfת�� chr(10)  20160409
    if instr(content, vbCrLf) = false then
        content = replace(content, chr(10), vbCrLf) 
		content=replace(content,vbcrlf & vbcrlf, vbcrlf)
    end if 
    fieldConfigList = lcase(getFieldConfigList(db_PREFIX & tableName)) 
    splStr = split(fieldConfigList, ",") 
    splList = split(content, vbCrLf & "-------------------------------") 
    nOK = 0 
    for each listStr in splList
        addFieldList = ""                                                               '����ֶ��б����
        addValueList = ""                                                               '����ֶ��б�ֵ
        updateValueList = ""                                                            '�޸��ֶ��б�

        isJump = lcase(trim(newGetStrCut(listStr, "#jump#"))) 
        isStop = lcase(trim(newGetStrCut(listStr, "#stop#")))
 
		'ֹͣ����
		if isStop = "1" or isStop = "true" then
    		nImportTXTData = nOK 
			exit function
		end if
		'�жϲ�����ת��ǰ
        if isJump <> "1" and isJump <> "true" then
            for each fieldStr in splStr
                if fieldStr <> "" then
                    splxx = split(fieldStr & "| ", "|")			'�������| ��jsp�������������֪��Ϊʲô�� ����û��ϵ����Ӱ������ִ�н��� 
                    fieldName = splxx(0) 
                    fieldType = splxx(1) 
                    if instr(listStr, "��" & fieldName & "��") > 0 then
                        listStr = listStr & vbCrLf                                                      '�Ӹ�������Ϊ�������һ����������ӽ�ȥ 20160629
                        if addFieldList <> "" then
                            addFieldList = addFieldList & "," 
                            addValueList = addValueList & "," 
                            updateValueList = updateValueList & "," 
                        end if 
                        addFieldList = addFieldList & fieldName 
'call echo(fieldName,fieldType)
'doevents
                        fieldValue = newGetStrCut(listStr, fieldName) 
                        if fieldType = "textarea" or ( EDITORTYPE="jsp" and fieldType="") then
                            fieldValue = contentTranscoding(fieldValue) 
						'�������Ϊ�����������Ͳ�Ҫ����true �� false ��  sqlserver��  20160803
						elseif fieldType = "yesno" or fieldType = "numb" then
							if lcase(fieldValue)="true" then
								fieldValue="1"
							elseif lcase(fieldValue)="false" then
								fieldValue="0"
							end if		
						 
                        end if 
                        'call echo(tableName,fieldName)
                        '���´���
                        if(tableName = "articledetail" or tableName = "webcolumn") and fieldName = "parentid" then
                            'call echo(tableName,fieldName)
                            'call echo("fieldValue",fieldValue)
                            fieldValue = getColumnId(fieldValue) 
                            'call echo("fieldValue",fieldValue)
						'BBS����20171003
						elseif(tableName = "bbsdetail" or tableName = "bbscolumn") and fieldName = "parentid" then	 
                            fieldValue = handleGetColumnID("bbscolumn", fieldValue)  
						'CAI����20171117
						elseif(tableName = "caidetail" or tableName = "caicolumn") and fieldName = "parentid" then	 
                            fieldValue = handleGetColumnID("caicolumn", fieldValue)  
                        '��̨�˵�
                        elseif tableName = "listmenu" and fieldName = "parentid" then
                            fieldValue = getListMenuId(fieldValue) 
                        end if 
                        if fieldType = "date" and fieldValue = "" then
                            fieldValue = date() 
                        elseif(fieldType = "time" or fieldType = "now") and fieldValue = "" then
                            fieldValue = cstr(now()) 
                        end if 
                        if fieldType <> "yesno" and fieldType <> "numb" then
                            fieldValue = "'" & fieldValue & "'" 
                        'Ĭ����ֵ����Ϊ0
                        elseif fieldValue = "" then
                            fieldValue = "0" 
                        end if
'call echo(fieldName,fieldType & "("& replace(left(fieldValue,22),"<","&lt;") &")" ) :doevents
                        addValueList = addValueList & fieldValue                                        '���ֵ
                        updateValueList = updateValueList & fieldName & "=" & fieldValue                '�޸�ֵ
                    end if 
                end if 
            next 
            '�ֶ��б�Ϊ��
            if addFieldList <> "" then
                if sType = "�޸�" then
                    sql = "update " & db_PREFIX & "" & tableName & " set " & updateValueList 
                else
                    sql = "insert into " & db_PREFIX & "" & tableName & " (" & addFieldList & ") values(" & addValueList & ")" 
                end if 
                '���SQL
                if checksql(sql) = false then
                    call eerr("������ʾ1", "<hr>sql=" & sql & "<br>") 
                end if 
                nOK = nOK + 1 
            else
                nOK = nBatchImportColumnList(content,splStr, listStr, nOK, tableName) 

            end if 
        end if 

    next 
    nImportTXTData = nOK 
end function 
'����������Ŀ�б� 20160716
function nBatchImportColumnList(content,splField, byval listStr, nOK, tableName)
    dim splStr, splxx, isColumn, columnName, s, c, nLen, id, parentIdArray(99), columntypeArray(99), flagsArray(99), nIndex, fieldStr, fieldName, valueStr, nCount, bodycontent
    dim fileName, templatepath, rowC 
    isColumn = false 
    nCount = 0 
    listStr = replace(listStr, vbTab, "    ") 
    splStr = split(listStr, vbCrLf) 
    for each s in splStr
        rowC = "" 
        if left(trim(s), 1) = "#" then
        'Ϊ#����ִ��
        elseif s = "��#sub#��" then
            isColumn = true 
        elseif isColumn = true then
            columnName = s 
            if instr(columnName, "��|��") > 0 then
                columnName = mid(columnName, 1, instr(columnName, "��|��") - 1) 
            end if 
            columnName = rTrim(columnName) 
            nLen = len(columnName) 
            columnName = ltrim(columnName) 
            nLen = nLen - len(columnName) 
            nIndex = cint(nLen / 4) 
            if columnName <> "" then
                parentIdArray(nIndex) = columnName 
                rowC = rowC & "��columnname��" & columnName & vbCrLf 
                for each fieldStr in splField
                    splxx = split(fieldStr & "|", "|") 
                    fieldName = splxx(0) 
                    if fieldName <> "" and fieldName <> "columnname" and instr(s, fieldName & "='") > 0 then
                        valueStr = getStrCut(s, fieldName & "='", "'", 2) 
                        rowC = rowC & "��" & fieldName & "��" & valueStr & vbCrLf 

                        if fieldName = "columntype" then
                            columntypeArray(nIndex) = valueStr 
                        elseif fieldName = "flags" then
                            flagsArray(nIndex) = valueStr 

                        elseif fieldName = "templatepath" then
                            templatepath = valueStr 
                        elseif fieldName = "filename" then
                            fileName = valueStr 

                        end if 
                    end if 
                next 
                'call echo(filename,templatepath)
                if instr(fileName, "[thistemplate]") > 0 then
                    rowC = vbCrLf & "��filename��" & replace(fileName, "[thistemplate]", templatepath) & vbCrLf & rowC 
                end if 
                if nIndex <> 0 then
                    rowC = rowC & "��parentid��" & parentIdArray(nIndex - 1) & vbCrLf 
                    rowC = rowC & "��columntype��" & columntypeArray(nIndex - 1) & vbCrLf 
                    rowC = rowC & "��flags��" & flagsArray(nIndex - 1) & vbCrLf 
                    if columntypeArray(nIndex) = "" then
                        columntypeArray(nIndex) = columntypeArray(nIndex - 1) 
                    end if 
                    if flagsArray(nIndex) = "" then
                        flagsArray(nIndex) = flagsArray(nIndex - 1) 
                    end if 

                else
                    rowC = rowC & "��parentid��-1" & vbCrLf 
                end if 
                rowC = rowC & "��sortrank��" & nCount & vbCrLf
				s=getStrCut(content,"��"& columnName &"��","��/"& columnName &"��",2)
                rowC = rowC & "��bodycontent��" & s & "��/bodycontent��" & vbCrLf
				if s <>""then
					'call echo("s",s)
				
				end if
				 
                nCount = nCount + 1 
				rowC = rowC & getRowConfigData(content,parentIdArray,nIndex) 
                c = c & rowC & "-------------------------------" & vbCrLf 
            end if 
        end if 
    next 
    'call die(createfile("1.txt",c))
    '��������
    if c <> "" then
        call nImportTXTData(c, tableName, "���") 
    end if 
    nBatchImportColumnList = nCount 
end function
'��õ�ǰ����������
function getRowConfigData(content,parentIdArray,nIndex)
	dim i,s,c
	for i=0 to nIndex
		s=parentIdArray(i)
		if c<>"" then
			c=c & "/"
		end if
		c=c & s
	next
	s=getStrCut(content,"��###"& c &"###��","��/###"& c &"###��",2)
	getRowConfigData= vbcrlf & s & vbcrlf
end function

'�µĽ�ȡ�ַ�20160216
function newGetStrCut(content, title)
    dim s 
    '��������Ϊ�˴�GitHub����ʱ����vbcrlfת�� chr(10)  20160409
    if instr(content, vbCrLf) = false then
        content = replace(content, chr(10), vbCrLf) 
    end if 
    if inStr(content, "��/" & title & "��") > 0 then
        s = aDSql(phptrim(getStrCut(content, "��" & title & "��", "��/" & title & "��", 0))) 
    else
        s = aDSql(phptrim(getStrCut(content, "��" & title & "��", vbCrLf, 0))) 
    end if 
    newGetStrCut = s 
end function 

'����ת��
function contentTranscoding(byVal content)
    content = replace(replace(replace(replace(content, "<?", "&lt;?"), "?>", "?&gt;"), "<" & "%", "&lt;%"), "?>", "%&gt;") 
 
    dim splStr, i, s, c, isTranscoding, isBR 
    isTranscoding = false 
    isBR = false 
	
    '��������Ϊ�˴�GitHub����ʱ����vbcrlfת�� chr(10)  20160409
    if instr(content, vbCrLf) = false then
        content = replace(content, chr(10), vbCrLf) 
    end if 
	
    splStr = split(content, vbCrLf) 
    for each s in splStr
        if inStr(s, "[&htmlת��&]") > 0 then
            isTranscoding = true 
        end if 
        if inStr(s, "[&htmlת��end&]") > 0 then
            isTranscoding = false 
        end if 
        if inStr(s, "[&ȫ������&]") > 0 then
            isBR = true 
        end if 
        if inStr(s, "[&ȫ������end&]") > 0 then
            isBR = false 
        end if 

        if isTranscoding = true then
            s = replace(replace(s, "[&htmlת��&]", ""), "<", "&lt;") 
        else
            s = replace(s, "[&htmlת��end&]", "") 
        end if 
        if isBR = true then
            s = replace(s, "[&ȫ������&]", "") 
            if right(trim(s), 8) <> "������/div>" then
                s = s & "<br>" 
            end if 
        else
            s = replace(s, "[&ȫ������end&]", "") 
        end if 
        '��ǩ��ʽ������� 20160628
        if instr(s, "��article_lable��") > 0 then
            s = replace(s, "��article_lable��", "") 
            s = "<div class=""article_lable"">" & s & "</div>" 
        elseif instr(s, "��article_blockquote��") > 0 then
            s = replace(s, "��article_blockquote��", "") 
            s = "<div class=""article_blockquote"">" & s & "</div>" 
        end if 


        if c <> "" then
            c = c & vbCrLf 
        end if 
        c = c & s 
    next 
    c = replace(replace(c, "��b��", "<b>"), "��/b��", "</b>") 
    c = replace(c, "������", "<") 

    c = replace(replace(c, "��strong��", "<strong>"), "��/strong��", "</strong>") 
    contentTranscoding = c 
end function 

 
%>