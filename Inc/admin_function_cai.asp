<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-02-27
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'调用function2文件函数
function callFunction_cai()
    dim sType 
    sType = request("stype") 
    if sType = "cai" then
        call cai()                                                       '采集
    elseif sType = "newcai" then
        call newCai()                                                       '新版采集
    elseif sType = "clearAllData" then
        call cai_clearAllData()                                          '清除全部数据
    elseif sType = "importCaiData" then
        call cai_importCaiData()                                         '导入采集数据
    elseif sType = "callFunction_cai_test" then
        call callFunction_cai_test()                                     '测试
    else
        call eerr("callFunction_cai页里没有动作", request("stype")) 
    end if 
end function 
'新版采集
function newCai()
	dim id,sFileCharSet,httpurl,morepageurl,startPage,endPage,url,content,handleContent,dirPath,filePath,nState
	dim i,nI,s,c,splstr,splxx,arrData(99,99,99)
    dirPath = handlePath("../cache/cai") 
    call createDirFolder(dirPath)  
	id=request("id") 
	rs.open "select * from " & db_PREFIX & "caiColumn where id="&id,conn,1,1
	if not rs.eof then
		sFileCharSet=rs("charset")			'文件编码
		httpurl=rs("httpurl")				'首页网址
		morepageurl=rs("morepageurl")		'更多页网址
		startPage=rs("startPage")			'开始页数
		endPage=rs("endPage")				'结束页数
		if startPage<0 then startPage=0
		call echo("开始/结束页",rs("startPage") & "/" & rs("endPage"))
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
					call echo("提示，正在下载文件",url)
					nState=saveRemoteFile(url,filePath)
					if nState<>200 then
						call deleteFile(filePath)
						call echo("提示，下载失败 状态码="& nState &"，删除文件",filePath)
					end if
				else
					call echo("提示，网址有缓冲",url)
				end if
				'读内容 
				if checkFile(filePath)=true then
					content=readFile(filePath,sFileCharSet) 
					call pubFormatObj.handleFormatting(content) 
					call echo("提示，读文件内容",filePath)
					rsx.open "select * from " & db_PREFIX & "caiColumn where parentId="&id & " order by sortRank",conn,1,1
					while not rsx.eof
						call echo("配置动作",rsx("sType") & " =>> " & rsx("labelNameList")&" =>> " & rsx("htmlValue")  )
						if rsx("stype")<>"" then 
							handleContent=pubFormatObj.handleLabelContent(rsx("labelNameList"),"*",rsx("htmlValue"),rsx("sType")) 
							if rsx("sAction")="分割" then
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

'导入采集数据
function cai_importCaiData()
    dim id, addTableFieldName, addTableFieldValue, sql, i, articleFieldList, selectSql 
    dim fieldName, nAddCount, parentid 
    id = request("id") 
    sql = "select * from " & db_PREFIX & "caidata" 
    if id <> "" then
        sql = sql & " where id=" & id 
    end if 
    nAddCount = 0 
    articleFieldList = getHandleFieldList(db_PREFIX & "articledetail", "字段配置列表") 
    call echo("articleFieldList", articleFieldList) 
    '【@是jsp显示@】try{
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
            call echo("导入数据成功", nAddCount) 
        end if : rsx.close 
        doevents 
        'call eerr(addTableFieldName,addTableFieldValue)
    rs.movenext : wend : rs.close 
    '【@是jsp显示@】}catch(Exception e){}
    call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=CaiData&addsql=order by id desc&lableTitle=采集数据'>OK</a>") 
end function 
'获得sql赋值 文本与数字类型前后加'符号
function handleSqlValue(fieldConfigList, fieldName, valueStr)
    handleSqlValue = "'" & valueStr & "'" 
    if instr(fieldConfigList, "," & fieldName & "|numb|") > 0 then
        handleSqlValue = valueStr 
    end if 
end function 
'测试
function callFunction_cai_test()
    call echo("测试", "callFunction_cai_test") 
end function 
'清除全部数据
function cai_clearAllData()
    call openconn() 
    conn.execute("delete from " & db_PREFIX & "caidata") 
    call echo("操作完成", "<a href='?act=dispalyManageHandle&actionType=CaiData&addsql=order by id desc&lableTitle=采集数据'>OK</a>") 
end function 
'采集
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
    nAddSuccessCount = 0                                                            '追加成功总数
    '【@是jsp显示@】try{
    rs.open "select * from " & db_PREFIX & "caiweb where id=" & id, conn, 1, 1 
    if not rs.EOF then
        did = rs("bigclassname") 
        httpurl = rs("httpurl") 
        morePageUrl = rs("morePageUrl") 
        sCharSet = rs("charset") 
        nThisPage = CInt(rs("thisPage")) 
        nCountPage = CInt(rs("countPage")) 
        columnname = rs("columnname")                                    '栏目名称
        for i = nThisPage to nCountPage
            url = getHandleCaiUrl(httpurl, morePageUrl, cstr(i)) 
            htmlPath = dirPath & "/" & setfileName(url) 
            if checkFile(htmlPath) = false then
                content = gethttpurl(url, sCharSet) 
                call createfile(htmlPath, content) 
                msg = "（从网络读取）" 
            else
                content = getftext(htmlPath) 
                msg = "（从本地读取）" 
            end if 
            tempContent = content 
            call echo(i & "/" & nCountPage & msg, url) 

            rsx.open "select * from " & db_PREFIX & "caiconfig where bigclassname='" & did & "' and isthrough<>0 order by sortRank", conn, 1, 1 
            value1 = "" : value2 = "" : value3 = "" : value4 = "" : value5 = "" : value6 = "" : value7 = "" : value8 = "" '清空值
            defaultSqlFields = "bigclassname,isthrough,columnname" 
            defaultSqlValues = "'" & did & "',1,'" & columnname & "'" 
            addSql = "" 

            while not rsx.EOF
                startStr = rsx("startstr") 
                endStr = rsx("endstr") 
                startAddStr = rsx("startAddStr") 
                endAddStr = rsx("endAddStr") 
                saction = rsx("saction") 
                '追加前面
                if startAddStr = "【startstr】" then
                    startAddStr = startStr 
                end if 
                '追加后面
                if endAddStr = "【endstr】" then
                    endAddStr = startStr 
                end if 
                if rsx("stype") = "截取" then
                    content = startAddStr & getStrCut(content, startStr, endStr, 2) & endAddStr 
                elseIf rsx("stype") = "分割" then
                    listData = getArray(content, startStr, endStr, false, false) 
                    listData = startAddStr & replace(listData, "$Array$", endAddStr & "$Array$" & startAddStr) & endAddStr 
                elseIf inStr(rsx("stype"), "内容") > 0 then
                    splStr = split(listData, "$Array$") 
                    for each s in splStr
                        if rsx("stype") = "内容1" then
                            value1 = setThisFieldValue(saction, value1, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "内容2" then
                            value2 = setThisFieldValue(saction, value2, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "内容3" then
                            value3 = setThisFieldValue(saction, value3, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "内容4" then
                            value4 = setThisFieldValue(saction, value4, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "内容5" then
                            value5 = setThisFieldValue(saction, value5, s, url, startStr, endStr, startAddStr, endAddStr) 
                        elseIf rsx("stype") = "内容6" then
                            value6 = setThisFieldValue(saction, value6, s, url, startStr, endStr, startAddStr, endAddStr) 
                        end if 
                        fieldName = "fieldname" & replace(rsx("stype"), "内容", "") 
                        if fieldName <> "" and instr("," & defaultSqlFields & ",", "," & fieldName & ",") = false then
                            defaultSqlFields = defaultSqlFields & "," & fieldName 
                            defaultSqlValues = defaultSqlValues & ",'" & ADSql(rsx("fieldname")) & "'" 
                        end if 
                        '字段检测不为假
                        if rsx("fieldcheck") <> 0 then

                            fieldName = "value" & replace(rsx("stype"), "内容", "") 
                            addSql = fieldName & "='【" & rsx("stype") & "】'" 
                        end if 
                    next 
                end if 
                call echo("类型", rsx("stype")) 
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
                tempAddSql = replace(tempAddSql, "【内容1】", valueStr) 
                if uBound(spl2) >= j then
                    valueStr = ADSql(spl2(j)) 
                    sqlFields = sqlFields & ",value2" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "【内容2】", valueStr) 
                end if 
                if uBound(spl3) >= j then
                    valueStr = ADSql(spl3(j)) 
                    sqlFields = sqlFields & ",value3" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "【内容3】", valueStr) 
                end if 
                if uBound(spl4) >= j then
                    valueStr = ADSql(spl4(j)) 
                    sqlFields = sqlFields & ",value4" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "【内容4】", valueStr) 
                end if 
                if uBound(spl5) >= j then
                    valueStr = ADSql(spl5(j)) 
                    sqlFields = sqlFields & ",value5" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "【内容5】", valueStr) 
                end if 
                if uBound(spl6) >= j then
                    valueStr = ADSql(spl6(j)) 
                    sqlFields = sqlFields & ",value6" 
                    sqlValues = sqlValues & ",'" & valueStr & "'" 
                    tempAddSql = replace(tempAddSql, "【内容6】", valueStr) 
                end if 


                isAdd = true 
                if tempAddSql <> "" then
                    sql = "select * from " & db_PREFIX & "caidata where " & tempAddSql 
                    rsx.open sql, conn, 1, 1 
                    if not rsx.eof then
                        isAdd = false 
                    end if : rsx.close 
                end if 

                '为真则添加到数据库
                if isAdd = true then
                    sql = "insert into " & db_PREFIX & "caiData (" & sqlFields & ") values(" & sqlValues & ")" 
                    conn.execute(sql) 
                    nAddSuccessCount = nAddSuccessCount + 1 
                    call echo("添加成功", nAddSuccessCount) 
                end if 
            next 
            doevents 
        'die(value1)
        next 
    end if : rs.close 
    '【@是jsp显示@】}catch(Exception e){}

end function 
'给指定内容赋值
function setThisFieldValue(saction, byVal valueName, byref s, url, startStr, endStr, startAddStr, endAddStr)
    dim s1 
    if valueName <> "" then
        valueName = valueName & "$Array$" 
    end if 
    s1 = startAddStr & getStrCut(s, startStr, endStr, 2) & endAddStr 
    if saction = "处理成完整网址" then
        s1 = fullHttpUrl(url, s1) 
    end if 
    s = s1 
    setThisFieldValue = valueName & s1 
end function 
'获得处理后配置网址
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

