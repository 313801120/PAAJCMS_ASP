<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-02-27
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'����function2�ļ�����
function callFunction_cai()
    dim sType 
    sType = request("stype") 
    if sType = "cai" then
        call cai()                                                       '�ɼ�
    elseif sType = "newcai" then
        call newCai()                                                       '�°�ɼ�
    elseif sType = "clearAllData" then
        call cai_clearAllData()                                          '���ȫ������
    elseif sType = "importCaiData" then
        call cai_importCaiData()                                         '����ɼ�����
    elseif sType = "callFunction_cai_test" then
        call callFunction_cai_test()                                     '����
    else
        call eerr("callFunction_caiҳ��û�ж���", request("stype")) 
    end if 
end function 
'�°�ɼ�
function newCai()
	dim id,sFileCharSet,httpurl,morepageurl,startPage,endPage,url,content,handleContent,dirPath,filePath,nState
	dim i,nI,s,c,splstr,splxx,arrData(99,99,99)
    dirPath = handlePath("../cache/cai") 
    call createDirFolder(dirPath)  
	id=request("id") 
	rs.open "select * from " & db_PREFIX & "caiColumn where id="&id,conn,1,1
	if not rs.eof then
		sFileCharSet=rs("charset")			'�ļ�����
		httpurl=rs("httpurl")				'��ҳ��ַ
		morepageurl=rs("morepageurl")		'����ҳ��ַ
		startPage=rs("startPage")			'��ʼҳ��
		endPage=rs("endPage")				'����ҳ��
		if startPage<0 then startPage=0
		call echo("��ʼ/����ҳ",rs("startPage") & "/" & rs("endPage"))
		if endPage>=startPage then
			for i=startPage to endPage 
				if i=startPage and httpurl<>"" then
					url =httpurl
				elseif morepageurl<>"" then
					url=morepageurl
				else
					if httpurl<>"" then
						url=httpurl
					else
						url=morepageurl
					end if
				end if
				url=replace(url,"{*}",i)
				filePath=dirPath & "/" & mymd5(url)
				if checkFile(filePath)=false then
					call echo("��ʾ�����������ļ�",url)
					nState=saveRemoteFile(url,filePath)
					if nState<>200 then
						call deleteFile(filePath)
						call echo("��ʾ������ʧ�� ״̬��="& nState &"��ɾ���ļ�",filePath)
					end if
				else
					call echo("��ʾ����ַ�л���",url)
				end if
				'������ 
				if checkFile(filePath)=true then
					content=readFile(filePath,sFileCharSet) 
					call pubFormatObj.handleFormatting(content) 
					call echo("��ʾ�����ļ�����",filePath)
					rsx.open "select * from " & db_PREFIX & "caiColumn where parentId="&id & " order by sortRank",conn,1,1
					while not rsx.eof
						call echo("���ö���",rsx("sType") & " =>> " & rsx("labelNameList")&" =>> " & rsx("htmlValue")  )
						if rsx("stype")<>"" then 
							handleContent=pubFormatObj.handleLabelContent(rsx("labelNameList"),"*",rsx("htmlValue"),rsx("sType")) 
							if rsx("sAction")="�ָ�" then
								splstr=split(handleContent,"$Array$  ")
								nI=-1
								for each s in splstr
									if s<>"" then
										nI=nI+1 
										arrData(nI,0,0)=s
									end if
								next
							end if
						end if
					rsx.movenext:wend:rsx.close
				end if
			'call rwend(handleContent)
			next
			if arrData(0,0,0)<>"" then
				call rwend(arrData(0,0,0))
			end if
			
					
		end if
		
		
		
	end if:rs.close

	call rw(getTimer())
end function

'����ɼ�����
function cai_importCaiData()
    dim id, addTableFieldName, addTableFieldValue, sql, i, articleFieldList, selectSql 
    dim fieldName, nAddCount, parentid 
    id = request("id") 
    sql = "select * from " & db_PREFIX & "caidata" 
    if id <> "" then
        sql = sql & " where id=" & id 
    end if 
    nAddCount = 0 
    articleFieldList = getHandleFieldList(db_PREFIX & "articledetail", "�ֶ������б�") 
    call echo("articleFieldList", articleFieldList) 
    '��@��jsp��ʾ@��try{
    rs.open sql, conn, 1, 1 
    while not rs.eof
        selectSql = "" 
        parentid = getColumnId(rs("columnname")) 
        if parentid <> "" then
            addTableFieldName = "parentid" 
            addTableFieldValue = parentid 
        end if 
        for i = 1 to 6
            fieldName = lcase(phptrim(rs("fieldname" & i))) 
            if fieldName <> "" and instr(articleFieldList, "," & fieldName & "|") > 0 then
                if addTableFieldName <> "" then
                    addTableFieldName = addTableFieldName & "," 
                end if 
                addTableFieldName = addTableFieldName & fieldName 
                if selectSql = "" then
                    selectSql = " where " & fieldName & "=" & handleSqlValue(articleFieldList, fieldName, ADSql(rs("value" & i))) 
                end if 
                if addTableFieldValue <> "" then
                    addTableFieldValue = addTableFieldValue & "," 
                end if 
                addTableFieldValue = addTableFieldValue & handleSqlValue(articleFieldList, fieldName, ADSql(rs("value" & i))) 
            end if 
        next 

        sql = "select * from " & db_PREFIX & "articledetail" & selectSql 
        rsx.open sql, conn, 1, 1 
        if rsx.eof then
            sql = "insert into " & db_PREFIX & "articledetail (" & addTableFieldName & ") values(" & addTableFieldValue & ")" 
            conn.execute(sql) 
            nAddCount = nAddCount + 1 
            'call echo("sql",sql)
            call echo("�������ݳɹ�", nAddCount) 
        end if : rsx.close 
        doevents 
        'call eerr(addTableFieldName,addTableFieldValue)
    rs.movenext : wend : rs.close 
    '��@��jsp��ʾ@��}catch(Exception e){}
    call echo("�������", "<a href='?act=dispalyManageHandle&actionType=CaiData&addsql=order by id desc&lableTitle=�ɼ�����'>OK</a>") 
end function 
'���sql��ֵ �ı�����������ǰ���'����
function handleSqlValue(fieldConfigList, fieldName, valueStr)
    handleSqlValue = "'" & valueStr & "'" 
    if instr(fieldConfigList, "," & fieldName & "|numb|") > 0 then
        handleSqlValue = valueStr 
    end if 
end function 
'����
function callFunction_cai_test()
    call echo("����", "callFunction_cai_test") 
end function 
'���ȫ������
function cai_clearAllData()
    call openconn() 
    conn.execute("delete from " & db_PREFIX & "caidata") 
    call echo("�������", "<a href='?act=dispalyManageHandle&actionType=CaiData&addsql=order by id desc&lableTitle=�ɼ�����'>OK</a>") 
end function 
'�ɼ�
function cai()
    dim dirPath, id, url, i, msg, did 
    dim httpurl, morePageUrl, sCharSet, nThisPage, nCountPage, content, tempContent, listData, htmlPath 
    dim startStr, endStr, startAddStr, endAddStr, splStr, s, s1, c, saction 
    dim value1, value2, value3, value4, value5, value6, value7, value8, sql, addSql, tempAddSql, sqlFields, sqlValues 
    dim spl1, spl2, spl3, spl4, spl5, spl6, j, fieldName, fieldcheck, isAdd, nAddSuccessCount 
    dim defaultSqlFields, defaultSqlValues, valueStr, columnname 
    id = request("id") 
    dirPath = handlePath("../cache/cai") 
    call createDirFolder(dirPath) 
    call echo(dirPath, webDir) 
    call openconn() 
    nAddSuccessCount = 0                                                            '׷�ӳɹ�����
    '��@��jsp��ʾ@��try{
    rs.open "select * from " & db_PREFIX & "caiweb where id=" & id, conn, 1, 1 
    if not rs.EOF then
        did = rs("bigclassname") 
        httpurl = rs("httpurl") 
        morePageUrl = rs("morePageUrl") 
        sCharSet = rs("charset") 
        nThisPage = CInt(rs("thisPage")) 
        nCountPage = CInt(rs("countPage")) 
        columnname = rs("columnname")                                    '��Ŀ����
        for i = nThisPage to nCountPage
            url = getHandleCaiUrl(httpurl, morePageUrl, cstr(i)) 
            htmlPath = dirPath & "/" & setfileName(url) 
            if checkFile(htmlPath) = false then
                content = gethttpurl(url, sCharSet) 
                call createfile(htmlPath, content) 
                msg = "���������ȡ��" 
            else
                content = getftext(htmlPath) 
                msg = "���ӱ��ض�ȡ��" 
            end if 
            tempContent = content 
            call echo(i & "/" & nCountPage & msg, url) 

            rsx.open "select * from " & db_PREFIX & "caiconfig where bigclassname='" & did & "' and isthrough<>0 order by sortRank", conn, 1, 1 
            value1 = "" : value2 = "" : value3 = "" : value4 = "" : value5 = "" : value6 = "" : value7 = "" : value8 = "" '���ֵ
            defaultSqlFields = "bigclassname,isthrough,columnname" 
            defaultSqlValues = "'" & did & "',1,'" & columnname & "'" 
            addSql = "" 

            while not rsx.EOF
                startStr = rsx("startstr") 
                endStr = rsx("endstr") 
                startAddStr = rsx("startAddStr") 
                endAddStr = rsx("endAddStr") 
                saction = rsx("saction") 
                '׷��ǰ��
                if startAddStr = "��startstr��" then
                    startAddStr = startStr 
                end if 
                '׷�Ӻ���
                if endAddStr = "��endstr��" then
                    endAddStr = startStr 
                end if 
                if rsx("stype") = "��ȡ" then
                    content = startAddStr & getStrCut(content, startStr, endStr, 2) & endAddStr 
                elseIf rsx("stype") = "�ָ�" then
                    listData = getArray(content, startStr, endStr, false, false) 
                    listData = startAddStr & replace(listData, "$Array$", endAddStr & "$Array$" & startAddStr) & endAddStr 
                elseIf inStr(rsx("stype"), "����") > 0 then
                    splStr = split(listData, "$Array$") 
                    for each s in splStr
                        if rsx("stype") = "����1" then
                            value1 = setThisFieldValue(saction, value1, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "����2" then
                            value2 = setThisFieldValue(saction, value2, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "����3" then
                            value3 = setThisFieldValue(saction, value3, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "����4" then
                            value4 = setThisFieldValue(saction, value4, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "����5" then
                            value5 = setThisFieldValue(saction, value5, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "����6" then
                            value6 = setThisFieldValue(saction, value6, s, url, startStr, endStr, startAddStr, endAddStr) 
                        end if 
                        fieldName = "fieldname" & replace(rsx("stype"), "����", "") 
                        if fieldName <> "" and instr("," & defaultSqlFields & ",", "," & fieldName & ",") = false then
                            defaultSqlFields = defaultSqlFields & "," & fieldName 
                            defaultSqlValues = defaultSqlValues & ",'" & ADSql(rsx("fieldname")) & "'" 
                        end if 
                        '�ֶμ�ⲻΪ��
                        if rsx("fieldcheck") <> 0 then

                            fieldName = "value" & replace(rsx("stype"), "����", "") 
                            addSql = fieldName & "='��" & rsx("stype") & "��'" 
                        end if 
                    next 
                end if 
                call echo("����", rsx("stype")) 
            rsx.moveNext : wend : rsx.close 
            spl1 = split(value1, "$Array$") 
            spl2 = split(value2, "$Array$") 
            spl3 = split(value3, "$Array$") 
            spl4 = split(value4, "$Array$") 
            spl5 = split(value5, "$Array$") 
            spl6 = split(value6, "$Array$") 
            for j = 0 to uBound(spl1)
                sqlFields = defaultSqlFields 
                sqlValues = defaultSqlValues 
                tempAddSql = addSql 
                valueStr = ADSql(spl1(j)) 
                sqlFields = sqlFields & ",value1" 
                sqlValues = sqlValues & ",'" & valueStr & "'" 
                tempAddSql = replace(tempAddSql, "������1��", valueStr) 
                if uBound(spl2) >= j then
                    valueStr = ADSql(spl2(j)) 
                    sqlFields = sqlFields & ",value2" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "������2��", valueStr) 
                end if 
                if uBound(spl3) >= j then
                    valueStr = ADSql(spl3(j)) 
                    sqlFields = sqlFields & ",value3" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "������3��", valueStr) 
                end if 
                if uBound(spl4) >= j then
                    valueStr = ADSql(spl4(j)) 
                    sqlFields = sqlFields & ",value4" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "������4��", valueStr) 
                end if 
                if uBound(spl5) >= j then
                    valueStr = ADSql(spl5(j)) 
                    sqlFields = sqlFields & ",value5" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "������5��", valueStr) 
                end if 
                if uBound(spl6) >= j then
                    valueStr = ADSql(spl6(j)) 
                    sqlFields = sqlFields & ",value6" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "������6��", valueStr) 
                end if 


                isAdd = true 
                if tempAddSql <> "" then
                    sql = "select * from " & db_PREFIX & "caidata where " & tempAddSql 
                    rsx.open sql, conn, 1, 1 
                    if not rsx.eof then
                        isAdd = false 
                    end if : rsx.close 
                end if 

                'Ϊ������ӵ����ݿ�
                if isAdd = true then
                    sql = "insert into " & db_PREFIX & "caiData (" & sqlFields & ") values(" & sqlValues & ")" 
                    conn.execute(sql) 
                    nAddSuccessCount = nAddSuccessCount + 1 
                    call echo("��ӳɹ�", nAddSuccessCount) 
                end if 
            next 
            doevents 
        'die(value1)
        next 
    end if : rs.close 
    '��@��jsp��ʾ@��}catch(Exception e){}

end function 
'��ָ�����ݸ�ֵ
function setThisFieldValue(saction, byVal valueName, byref s, url, startStr, endStr, startAddStr, endAddStr)
    dim s1 
    if valueName <> "" then
        valueName = valueName & "$Array$" 
    end if 
    s1 = startAddStr & getStrCut(s, startStr, endStr, 2) & endAddStr 
    if saction = "�����������ַ" then
        s1 = fullHttpUrl(url, s1) 
    end if 
    s = s1 
    setThisFieldValue = valueName & s1 
end function 
'��ô����������ַ
function getHandleCaiUrl(httpurl, morePageUrl, id)
    dim url 
    id = CStr(id) 
    if id = "1" then
        if httpurl <> "" then
            url = httpurl 
        else
            url = morePageUrl 
        end if 
    else
        if morePageUrl = "" then
            url = httpurl 
        else
            url = morePageUrl 
        end if 
    end if 
    url = replace(replace(replace(replace(url, "{*}", id), "[*]", id), "{id}", id), "[id]", id) 
    getHandleCaiUrl = url 
end function 
%>  

