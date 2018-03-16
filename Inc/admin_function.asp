<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-16
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'后台操作核心程序 添加 删除 修改 列表

'调用function文件函数
function callFunction()
    dim sType 
    sType = request("stype") 
    if sType = "updateWebsiteStat" then
        updateWebsiteStat()                                  '更新网站统计
    elseif sType = "clearWebsiteStat" then
        call clearWebsiteStat()                              '清空网站统计
    elseif sType = "updateTodayWebStat" then
        call updateTodayWebStat()                            '更新网站今天统计
    elseif sType = "websiteDetail" then
        call websiteDetail()                                 '详细网站统计
    elseif sType = "displayAccessDomain" then
        call displayAccessDomain()                           '显示访问域名
    elseif sType = "delTemplate" then
        call delTemplate()                                   '删除模板
    else
        call eerr("function1页里没有动作", request("stype")) 
    end if 
end function

'显示访问域名
function displayAccessDomain()
    dim visitWebSite, visitWebSiteList, urlList, nOK 
    call handlePower("显示访问域名") 
    call openconn() 
    nOK = 0 
	'【@是jsp显示@】try{	
    rs.open "select * from " & db_PREFIX & "websitestat", conn, 1, 1 
    while not rs.EOF
        visitWebSite = lcase(getWebSite(rs("visiturl"))) 
        'call echo("visitWebSite",visitWebSite)
        if instr(vbCrLf & visitWebSiteList & vbCrLf, vbCrLf & visitWebSite & vbCrLf) = 0 then
            if visitWebSite <> lcase(getWebSite(webDoMain())) then
                visitWebSiteList = visitWebSiteList & visitWebSite & vbCrLf 
                nOK = nOK + 1 
                urlList = urlList & nOK & "、<a href='" & rs("visiturl") & "' target='_blank'>" & rs("visiturl") & "</a><br>" 
            end if 
        end if 
    rs.movenext : wend : rs.close 
	'【@是jsp显示@】}catch(Exception e){} 
    call echo("显示访问域名", "操作完成 <a href='javascript:history.go(-1)'>点击返回</a>") 
    call rwend(visitWebSiteList & "<br><hr><br>" & urlList) 
end function 
'获得处理后表列表 20160313
function getHandleTableList()
    dim s, lableStr 
    lableStr = "表列表[" & request("mdbpath") & "]" 
    if WEB_CACHEContent = "" then
        WEB_CACHEContent = getftext(WEB_CACHEFile) 
    end if 
    s = getConfigContentBlock(WEB_CACHEContent, "#" & lableStr & "#") 
    if s = "" then
        s = LCase(getTableList()) 
        s = "|" & replace(s, vbCrLf, "|") & "|" 
        WEB_CACHEContent = setConfigFileBlock(WEB_CACHEFile, s, "#" & lableStr & "#") 
        if isCacheTip = true then
            call echo("缓冲", lableStr) 
        end if 
    end if 
    getHandleTableList = s 
end function 

'获得处理的字段列表   getHandleFieldList("ArticleDetail","字段列表")
function getHandleFieldList(tableName, sType)
    dim s 
    if WEB_CACHEContent = "" then
        WEB_CACHEContent = getftext(WEB_CACHEFile) 
    end if 
    s = getConfigContentBlock(WEB_CACHEContent, "#" & tableName & sType & "#") 

    if s = "" then
        if sType = "字段配置列表" then
            s = LCase(getFieldConfigList(tableName)) 
        else
            s = LCase(getFieldList(tableName)) 
        end if 
        WEB_CACHEContent = setConfigFileBlock(WEB_CACHEFile, s, "#" & tableName & sType & "#") 
        if isCacheTip = true then
            call echo("缓冲", tableName & sType) 
        end if 
    end if 
    getHandleFieldList = s 
end function 
'读模板内容 20160310
function getTemplateContent(templateFileName)
    call loadWebConfig() 
    '读模板
    dim templateFile, customTemplateFile, c 
    customTemplateFile = ROOT_PATH & "template/" & db_PREFIX & "/" & templateFileName 
    '为手机端
    if checkMobile() = true or request("m") = "mobile" then
        templateFile = ROOT_PATH & "/Template/mobile/" & templateFileName 
    end if 
    '判断手机端文件是否存在20160330
    if checkFile(templateFile) = false then
        if checkFile(customTemplateFile) = true then
            templateFile = customTemplateFile 
        else
            templateFile = ROOT_PATH & templateFileName 
        end if 
    end if 
    c = readFile(templateFile,"") 
    c = replaceLableContent(c) 
    getTemplateContent = c 
end function 
'替换标签内容
function replaceLableContent(content)
    dim s, c, splStr, list 
    content = replace(content, "{$webVersion$}", webVersion)                        '网站版本
    content = replace(content, "{$Web_Title$}", cfg_webTitle)                       '网站标题
    content = replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'ASP与PHP
    content = replace(content, "{$adminDir$}", adminDir)                            '后台目录
    content = replace(content, "{$incDir$}", incDir)                            '后台目录
	

    content = replace(content, "[$adminId$]", getsession("adminId"))              '管理员ID
    content = replace(content, "{$adminusername$}", getsession("adminusername"))       '管理账号名称
    content = replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        '程序类型
    content = replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      '前台
    content = replace(content, "{$webVersion$}", webVersion)                        '版本
    content = replace(content, "{$WebsiteStat$}", getConfigFileBlock(WEB_CACHEFile, "#访客信息#")) '最近访客信息


    content = replace(content, "{$databaseType$}", databaseType)                          '数据为类型
    content = replace(content, "{$DB_PREFIX$}", db_PREFIX)                          '表前缀
    content = replace(content, "{$adminflags$}", IIF(getsession("adminflags") = "|*|", "超级管理员", "普通管理员")) '管理员类型
    content = replace(content, "{$SERVER_SOFTWARE$}", request.serverVariables("SERVER_SOFTWARE")) '服务器版本
    content = replace(content, "{$SERVER_NAME$}", request.serverVariables("SERVER_NAME")) '服务器网址
    content = replace(content, "{$LOCAL_ADDR$}", request.serverVariables("LOCAL_ADDR")) '服务器IP
    content = replace(content, "{$SERVER_PORT$}", request.serverVariables("SERVER_PORT")) '服务器端口
    content = replaceValueParam(content, "mdbpath", request("mdbpath")) 
    content = replaceValueParam(content, "webDir", webDir) 
    content = replaceValueParam(content, "EDITORTYPE", EDITORTYPE)                        'ASP与PHP
 
    '20160628
    if instr(content, "{$backupDatabaseSelectHtml$}") > 0 then
        c = getDirTxtNameList(adminDir & "/Data/BackUpDateBases/") 
        splStr = split(c, vbCrLf) 
        for each s in splStr
            list = list & "<option value=""" & s & """>" & s & "</option>" & vbCrLf 
        next 
        content = replace(content, "{$backupDatabaseSelectHtml$}", list) 
    end if 

    '20160614
    if EDITORTYPE = "php" then
        content = replace(content, "{$EDITORTYPE_PHP$}", "php")                         '给phpinc/用
    end if 
    content = replace(content, "{$EDITORTYPE_PHP$}", "")                            '给phpinc/用

    replaceLableContent = content 
end function 

'文章列表旗
function displayFlags(flags)
    dim c 
    '头条[h]
    if inStr("|" & flags & "|", "|h|") > 0 then
        c = c & "头 " 
    end if 
    '推荐[c]
    if inStr("|" & flags & "|", "|c|") > 0 then
        c = c & "推 " 
    end if 
    '幻灯[f]
    if inStr("|" & flags & "|", "|f|") > 0 then
        c = c & "幻 " 
    end if 
    '特荐[a]
    if inStr("|" & flags & "|", "|a|") > 0 then
        c = c & "特 " 
    end if 
    '滚动[s]
    if inStr("|" & flags & "|", "|s|") > 0 then
        c = c & "滚 " 
    end if 
    '加粗[b]
    if inStr("|" & flags & "|", "|b|") > 0 then
        c = c & "粗 " 
    end if 
    if c <> "" then c = "[<font color=""red"">" & c & "</font>]" 

    displayFlags = c 
end function 


'栏目类别循环配置        showColumnList(parentid, "webcolumn", ,"",0, defaultStr,3,"")   nCount为深度值   thisPId为交点的id
function showColumnList(byVal parentid, byVal tableName, showFieldName, byVal thisPId, nCount, byVal action)
    dim i, s, c, selectcolumnname, selStr, url, isFocus, sql, addSql, listLableStr, topnav,nRecordCount
    dim thisColumnName, sNavheaderStr, sNavfooterStr ,focusRootColumeId
	dim titleFieldName			'标题字段名称
	titleFieldName="title"
	if instr("|webcolumn|bbscolumn|caicolumn|", "|" & lcase(tableName) & "|")>0 then
		titleFieldName="columnname"
	end if
	
    parentid = trim(parentid) 
    listLableStr = "list" 

    topnav = getStrCut(action, "[topnav]", "[/topnav]", 2) 
    focusRootColumeId = getStrCut(action, "[rootcolumeid]", "[/rootcolumeid]", 2) 
    thisColumnName = getColumnName(parentid) 
    'call echo(parentid,topnav)

    if parentid <> topnav then
		'深度20180116
		if instr(action, "[small"& nCount &"-list") > 0 then
            listLableStr = "small"& nCount & "-list"   
        elseif instr(action, "[small-list") > 0 then
            listLableStr = "small-list"
        end if 
    end if 
    'call echo("listLableStr",listLableStr)
    dim rs : set rs = createObject("Adodb.RecordSet")
		'【@是.netc显示@】OleDbDataReader rs=null;				//要不会出错的
        dim fieldNameList, splFieldName, nK, fieldName, replaceStr, startStr, endStr, nTop, nModI, title 
        dim subHeaderStr, subFooterStr, subHeaderStartStr, subHeaderEndStr, subFooterStartStr, subFooterEndStr 


        fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 
        splFieldName = split(fieldNameList, ",") 
        sql = "select * from " & db_PREFIX & tableName & " where parentid=" & parentid 
        'call echo("sql1111111111111",tableName)
        '处理追加SQL
        startStr = "[sql-" & nCount & "]" : endStr = "[/sql-" & nCount & "]" 
        if inStr(action, startStr) = false and inStr(action, endStr) = 0 then
            startStr = "[sql]" : endStr = "[/sql]" 
        end if 
        addSql = getStrCut(action, startStr, endStr, 2) 
        if addSql <> "" then
            sql = getWhereAnd(sql, addSql) 
        end if 
		'call echo(sql,addSql)
        sql = sql & " order by sortrank asc"
		'【@是jsp显示@】try{	
        rs.open sql, conn, 1, 1 
		'【PHP】删除rs
		nRecordCount=rs.recordCount
		'【@是jsp显示@】rs = Conn.executeQuery(handleSqlTop(sql));
        for i = 1 to nRecordCount
            if not rs.EOF then				
            	'【PHP】$rs=mysql_fetch_array($rsObj);                                            //给PHP用，因为在 asptophp转换不完善  特殊
                startStr = "" : endStr = "" 
                selStr = "" 
                isFocus = false 
				'改进
                if CStr(rs("id")) = CStr(thisPId) or (focusRootColumeId<>"" and CStr(rs("id"))=cstr(focusRootColumeId)) then
                    selStr = " selected " 
                    isFocus = true 
                end if 
                '网址判断
                if isFocus = true then
                    startStr = "[" & listLableStr & "-focus]" : endStr = "[/" & listLableStr & "-focus]" 
                else

                    startStr = "[" & listLableStr & "-" & thisColumnName & "]" : endStr = "[/" & listLableStr & "-" & thisColumnName & "]" 

                    if inStr(action, startStr) = 0 and instr(action, endStr) = 0 then
                        startStr = "[" & listLableStr & "-" & i & "]" : endStr = "[/" & listLableStr & "-" & i & "]" 
                    else
                    'call echo(rs("columnname"),startStr)
                    end if 
                end if 

                '在最后时排序当前交点20160202
                if i = nTop and isFocus = false then
                    startStr = "[" & listLableStr & "-end]" : endStr = "[/" & listLableStr & "-end]" 
                end if 
                '例[list-mod2]  [/list-mod2]    20150112
                for nModI = 6 to 2 step - 1
                    if inStr(action, startStr) = false and i mod nModI = 0 then
                        startStr = "[" & listLableStr & "-mod" & nModI & "]" : endStr = "[/" & listLableStr & "-mod" & nModI & "]" 
                        if inStr(action, startStr) > 0 then
                            exit for 
                        end if 
                    end if 
                next 

                '没有则用默认
                if inStr(action, startStr) = 0 and instr(action, endStr) = 0 then
                    startStr = "[" & listLableStr & "]" : endStr = "[/" & listLableStr & "]" 
                end if 
                'call rwend(action)
                'call echo(startStr,endStr)
                if inStr(action, startStr) > 0 and inStr(action, endStr) > 0 then
                    s = strCut(action, startStr, endStr, 2) 

                    s = replaceValueParam(s, "id", rs("id")) 
                    s = replaceValueParam(s, "selected", selStr) 
                    selectcolumnname = rs(showFieldName) : title = selectcolumnname 
                    if nCount >= 1 then
                        selectcolumnname = copystr("&nbsp;&nbsp;", nCount) & "├─" & selectcolumnname 
                    end if 
                    s = replaceValueParam(s, "selectcolumnname", selectcolumnname) 
                    s = replaceValueParam(s, "title", title) 


                    for nK = 0 to uBound(splFieldName)
                        if splFieldName(nK) <> "" then
                            fieldName = splFieldName(nK) 
                            replaceStr = rs(fieldName) & "" 

                            s = replaceValueParam(s, fieldName, replaceStr) 
                        end if 
                    next 

                    'url = WEB_VIEWURL & "?act=nav&columnName=" & rs(showFieldName)             '以栏目名称显示列表
                    url = WEB_VIEWURL & "?act=nav&id=" & rs("id")                               '以栏目ID显示列表



                    '自定义网址
                    if trim(rs("customaurl")) <> "" then
                        url = trim(rs("customaurl")) 
                    end if 
                    s = replace(s, "[$viewWeb$]", url) 
                    s = replaceValueParam(s, "url", url) 
					s = replaceValueParam(s, "i", i)                                                '循环编号
					s = replaceValueParam(s, "编号", i)                                             '循环编号

                    '网站栏目没有page位置处理 追加于20160716 home
                    url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=网站栏目&nPageSize=10&page=&id=" & rs("id") & "&n=" & getRnd(11) 
                    s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span")                    '处理是否添加在线修改管理器


                    if EDITORTYPE = "php" then
                        s = replace(s, "[$phpArray$]", "[]") 
                    else
                        s = replace(s, "[$phpArray$]", "") 
                    end if 

                    's=copystr("",nCount) & rs("columnname") & "<hr>"
                    if rs("parentid") = "-1" and instr(action, "[navheader]") > 0 then
                        sNavheaderStr = getStrCut(action, "[navheader]", "[/navheader]", 2) 
                        sNavfooterStr = getStrCut(action, "[navfooter]", "[/navfooter]", 2) 
						if isFocus = true then
							if instr(action,"[navheader-focus]")>0 then
                        		sNavheaderStr = getStrCut(action, "[navheader-focus]", "[/navheader-focus]", 2) 
							end if
							if instr(action,"[navfooter-focus]")>0 then
                        		sNavfooterStr = getStrCut(action, "[navfooter-focus]", "[/navfooter-focus]", 2) 
							end if
						end if
                    end if
					
					if EDITORTYPE<>"jsp" then 
						c = c & sNavheaderStr & s & vbCrLf 
						s = showColumnList(rs("id"), tableName, showFieldName, thisPId, nCount + 1, action) & sNavfooterStr 
					end if
					

                    subHeaderStartStr = "[subheader-" & rs(titleFieldName) & "]" : subHeaderEndStr = "[/subheader-" & rs(titleFieldName) & "]" 
                    if instr(action, subHeaderStartStr) = 0 and instr(action, subHeaderEndStr) = 0 then
                        subHeaderStartStr = "[subheader]" : subHeaderEndStr = "[/subheader]"  
                    end if
					
					 
                    subFooterStartStr = "[subfooter-" & rs(titleFieldName) & "]" : subFooterEndStr = "[/subfooter-" & rs(titleFieldName) & "]" 
                    if instr(action, subFooterStartStr) = 0 and instr(action, subFooterStartStr) = 0 then
                        subFooterStartStr = "[subfooter]" : subFooterEndStr = "[/subfooter]" 
                    end if
					'在最后20180308
					if i=nRecordCount then
						subFooterStartStr = "[subfooter-end]" : subFooterEndStr = "[/subfooter-end]"
						if instr(action, subFooterStartStr) = 0 and instr(action, subFooterStartStr) = 0 then
							subFooterStartStr = "[subfooter]" : subFooterEndStr = "[/subfooter]" 
						end if
					end if
                    subHeaderStr = getStrCut(action, subHeaderStartStr, subHeaderEndStr, 2) 
                    subFooterStr = getStrCut(action, subFooterStartStr, subFooterEndStr, 2) 
                    'call echo(rs("columnname"),"哈哈")

                    if s <> "" then s = vbCrLf & subHeaderStr & s & subFooterStr 
                    c = c & s 
                end if 
            end if 
        rs.moveNext : next : rs.close 
		'【@是jsp显示@】}catch(Exception e){}
        showColumnList = c 
end function
'msg1  辅助
function getMsg1(msgStr, url)
    dim content 
    content = getFText(ROOT_PATH & "msg.html") 
    msgStr = msgStr & "<br>" & jsTiming(url, 5) 
    content = replace(content, "[$msgStr$]", msgStr) 
    content = replace(content, "[$url$]", url) 


    content = replaceL(content, "提示信息") 
    content = replaceL(content, "如果您的浏览器没有自动跳转，请点击这里") 
    content = replaceL(content, "倒计时") 


    getMsg1 = content 
end function 

'检测权力
function checkPower(powerName)
	dim sql
	checkPower=false
    if getsession("adminId") <> "" then
        call openconn()                                                                 '打开数据库 要不然在php报错，晕
        '这个做会很慢，测试时用
		sql="select * from " & db_PREFIX & "admin where id=" & getsession("adminId")
        
		'【@是jsp显示@】try{	
		rss.open sql, conn, 1, 1 
        if not rss.eof then
            call setsession("adminflags", rss("flags")) 
        end if : rss.close 
		'【@是jsp显示@】}catch(Exception e){}
		
        if inStr("|" & getsession("adminflags") & "|", "|" & powerName & "|") > 0 or inStr("|" & getsession("adminflags") & "|", "|*|") > 0 then
            checkPower = true 
        else
            checkPower = false 
        end if 
    else
        checkPower = true 
    end if 
end function 
'处理后台管理权限
function handlePower(powerName)
    if checkPower(powerName) = false then
        call eerr("提示", "你没有【" & powerName & "】权限，<a href='javascript:history.go(-1);'>点击返回</a>") 
    end if 
end function 
'显示管理列表
function dispalyManage(actionName, lableTitle, byVal nPageSize, addSql)
    call handlePower("显示" & lableTitle)                                           '管理权限处理
    call loadWebConfig() 
    dim content, i, s, c, fieldNameList, sql, action 
    dim nX, url, nCount, nPage 
    dim idInputName 

    dim tableName, j, splxx 
    dim fieldName                                                                   '字段名称
    dim splFieldName                                                                '分割字段
    dim searchfield, keyWord                                                        '搜索字段，搜索关键词
    dim parentid                                                                    '栏目id

    dim replaceStr                                                                  '替换字符
    tableName = LCase(actionName)                                                   '表名称
	
	dim columnTalbeName:columnTalbeName="webColumn"								    '类表名称 
	if instr(lcase("|bbsdetail|"), lcase(tableName))>0 then
		columnTalbeName="bbsColumn"												    '类表名称 
	elseif instr(lcase("|caidetail|"), lcase(tableName))>0 then
		columnTalbeName="caiColumn"
	end if

    searchfield = request("searchfield")                                            '获得搜索字段值
    keyWord = request("keyword")                                                    '获得搜索关键词值
    if request.form("parentid") <> "" then
        parentid = request.form("parentid") 
    else
        parentid = request.queryString("parentid") 
    end if 

    dim id 
    dim focusid                                                                     '是判断传过来的id是否在当前列表中是交点20160715 home
    id = rq("id") 
    focusid = rq("focusid") 

    fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 

    fieldNameList = specialStrReplace(fieldNameList)                                '特殊字符处理
    splFieldName = split(fieldNameList, ",")                                        '字段分割成数组

	'追加于20170702
	dim customTemplatePath,templatePath
	templatePath="manage_" & tableName & ".html"
	if request("template")<>"" then
		customTemplatePath="manage_" & request("template") & ".html"
		if checkFile(customTemplatePath)=true then
			templatePath=customTemplatePath
		end if
	end if
    '读模板
    content = getTemplateContent(templatePath) 

    action = getStrCut(content, "[list]", "[/list]", 2) 
    '网站栏目单独处理      栏目不一样20160301
    if actionName = "WebColumn" or actionName = "BBSColumn" or actionName = "CaiColumn" then
        action = getStrCut(content, "[action]", "[/action]", 1) 
        content = replace(content, action, showColumnList( "-1", actionName, "columnname", "", 0, action))
		
		
    elseIf actionName = "ListMenu" then
        action = getStrCut(content, "[action]", "[/action]", 1) 
        content = replace(content, action, showColumnList( "-1", "listmenu", "title", "", 0, action)) 
    else
        if keyWord <> "" and searchfield <> "" then
			if left(keyWord,2)="==" then
				keyWord=mid(keyWord,3)
				if searchfield<>"id" and instr(getHandleFieldList(db_PREFIX & tableName, "字段配置列表"), ","& searchfield &"|numb|")=false then
					keyWord="'"& keyWord &"'" 
				end if  
				
            	addSql = getWhereAnd(" where " & searchfield & " = " & keyWord & " ", addSql) 
			else
            	addSql = getWhereAnd(" where " & searchfield & " like '%" & keyWord & "%' ", addSql) 
			end if
        end if 
        if parentid <> "" then
            addSql = getWhereAnd(" where parentid=" & parentid & " ", addSql) 
        end if 
        'call echo(tableName,addsql)
        sql = getWhereAnd("select * from " & db_PREFIX & tableName , addSql) 	'改进于20180128
        '检测SQL
        if checksql(sql) = false then
            call errorLog("出错提示5：<br>action=" & action & "<hr>sql=" & sql & "<br>") 
            exit function 
        end if 
		'【@是jsp显示@】try{	
        rs.open sql, conn, 1, 1 
        '【PHP】删除rs
        nCount = rs.recordCount
        s = handleNumber(request("page"))
		if s = "" then
			nPage=0
		else
			nPage=cint(s)
		end if
        content = replace(content, "[$pageInfo$]", webPageControl(nCount, nPageSize, cstr(nPage), url, "")) 
        content = replace(content, "[$accessSql$]", sql) 

        if EDITORTYPE = "asp" then
            nX = getRsPageNumber(rs, nCount, nPageSize, nPage)     		'【@不是asp屏蔽@】
			
		elseif EDITORTYPE = "aspx" then 	
				
			'【@是.netc显示@】int  nCountPage = getCountPage(nCount, nPageSize);
			'【@是.netc显示@】if(nPage<=1){
			'【@是.netc显示@】	nX=nPageSize;
			'【@是.netc显示@】	if(nX>nCount){
			'【@是.netc显示@】		nX=nCount;
			'【@是.netc显示@】	} 
			'【@是.netc显示@】}else{
			'【@是.netc显示@】	for(int nI2=0;nI2<nPageSize*(nPage-1);nI2++){
            '【@是.netc显示@】		rs.Read();
			'【@是.netc显示@】	}
			'【@是.netc显示@】	if(nPage<nCountPage){
			'【@是.netc显示@】		nX=nPageSize;
			'【@是.netc显示@】	}else{
			'【@是.netc显示@】		nX=nCount-nPageSize*(nPage-1);
			'【@是.netc显示@】	}
			'【@是.netc显示@】} 
		elseif EDITORTYPE = "jsp" then
		
			'【@是jsp显示@】int  nCountPage = getCountPage(nCount, nPageSize);
			'【@是jsp显示@】rs = Conn.executeQuery(sql);			
			'【@是jsp显示@】if(nPage<=1){
			'【@是jsp显示@】	nX=nPageSize;
			'【@是jsp显示@】	if(nX>nCount){
			'【@是jsp显示@】		nX=nCount;
			'【@是jsp显示@】	} 
			'【@是jsp显示@】}else{
			'【@是jsp显示@】	for(int nI2=0;nI2<nPageSize*(nPage-1);nI2++){
            '【@是jsp显示@】		rs.next();
			'【@是jsp显示@】	}
			'【@是jsp显示@】	if(nPage<nCountPage){
			'【@是jsp显示@】		nX=nPageSize;
			'【@是jsp显示@】	}else{
			'【@是jsp显示@】		nX=nCount-nPageSize*(nPage-1);
			'【@是jsp显示@】	}
			'【@是jsp显示@】}  

        else									 
            if nPage<>0 then'【@是.netc屏蔽@】'【@是jsp屏蔽@】
                nPage = nPage - 1 '【@是.netc屏蔽@】'【@是jsp屏蔽@】
            end if '【@是.netc屏蔽@】'【@是jsp屏蔽@】
            sql = "select * from " & db_PREFIX & "" & tableName & " " & addSql & " limit " & nPageSize * nPage & "," & nPageSize '【@是.netc屏蔽@】'【@是jsp屏蔽@】
            rs.open sql, conn, 1, 1 '【@是.netc屏蔽@】'【@是jsp屏蔽@】
            '【PHP】删除rs
            nX = rs.recordCount '【@是.netc屏蔽@】'【@是jsp屏蔽@】
        end if
		'待屏蔽
    	content = replaceValueParam(content, "print_sql", sql)     '打印出SQL
		
        for i = 1 to nX
            '【PHP】$rs=mysql_fetch_array($rsObj);                                            //给PHP用，因为在 asptophp转换不完善  特殊
            '【@是.netc显示@】rs.Read();
            '【@是jsp显示@】rs.next();
			s = replace(action, "[$id$]", rs("id")) 
            for j = 0 to uBound(splFieldName)
                if splFieldName(j) <> "" then
                    splxx = split(splFieldName(j) & "|||", "|") 
                    fieldName = splxx(0) 
                    replaceStr = rs(fieldName) & "" 
                    '对文章旗处理
                    if fieldName = "flags" then
                        replaceStr = displayFlags(replaceStr) 
                    end if 
                    'call echo("fieldname",fieldname)
                    's = Replace(s, "[$" & fieldName & "$]", replaceStr)
                    s = replaceValueParam(s, fieldName, replaceStr) 

                end if 
            next 

            idInputName = "id" 
            s = replace(s, "[$selectid$]", "<input type='checkbox' name='" & idInputName & "' id='" & idInputName & "' value='" & rs("id") & "' >") 
            s = replace(s, "[$phpArray$]", "") 
            url = "【NO】" 
            if actionName = "ArticleDetail" then
                url = WEB_VIEWURL & "?act=detail&id=" & rs("id") 
            elseIf actionName = "OnePage" then
                url = WEB_VIEWURL & "?act=onepage&id=" & rs("id") 
            '给评论加预览=文章  20160129
            elseIf actionName = "TableComment" then
                url = WEB_VIEWURL & "?act=detail&id=" & rs("itemid") 
            end if 
            '必需有自定义字段
            if inStr(fieldNameList, "customaurl") > 0 then
                '自定义网址
                if trim(rs("customaurl")) <> "" then
                    url = trim(rs("customaurl")) 
                end if 
            end if 
            s = replace(s, "[$viewWeb$]", url) 
            s = replaceValueParam(s, "cfg_websiteurl", cfg_webSiteUrl) 
            'call echo(focusid & "/" & rs("id"),IIF(focusid=cstr(rs("id")),"true","false"))
            s = replaceValueParam(s, "focusid", focusid) 

            c = c & s  

        rs.moveNext : next : rs.close 
		'【@是jsp显示@】}catch(Exception e){}
        content = replace(content, "[list]" & action & "[/list]", c) 
        '表单提交处理，parentid(栏目ID) searchfield(搜索字段) keyword(关键词) addsql(排序)
        url = "?page=[id]&addsql=" & request("addsql") & "&keyword=" & request("keyword") & "&searchfield=" & request("searchfield") & "&parentid=" & request("parentid") 
        url = getUrlAddToParam(getUrl(), url, "replace") 
        'call echo("url",url)
        content = replace(content, "[list]" & action & "[/list]", c) 

    end if 

    if inStr(content, "[$input_parentid$]") > 0 then  
        action = "[list]<option value=""[$id$]""[$selected$]>[$selectcolumnname$]</option>[/list]" 
        c = "<select name=""parentid"" id=""parentid""><option value="""">≡ 选择栏目 ≡</option>" & showColumnList( "-1", columnTalbeName, "columnname", parentid, 0, action) & vbCrLf & "</select>" 
        content = replace(content, "[$input_parentid$]", c)                        '上级栏目
    end if 

    content = replaceValueParam(content, "searchfield", request("searchfield"))     '搜索字段
    content = replaceValueParam(content, "keyword", request("keyword"))             '搜索关键词
    content = replaceValueParam(content, "nPageSize", request("nPageSize"))         '每页显示条数
    content = replaceValueParam(content, "addsql", request("addsql"))               '追加sql值条数
    content = replaceValueParam(content, "tableName", tableName)                    '表名称
    content = replaceValueParam(content, "actionType", request("actionType"))       '动作类型
    content = replaceValueParam(content, "lableTitle", request("lableTitle"))       '动作标题
    content = replaceValueParam(content, "id", id)                                  'id
    content = replaceValueParam(content, "page", request("page"))                   '页

    content = replaceValueParam(content, "parentid", request("parentid"))           '栏目id
    content = replaceValueParam(content, "focusid", focusid) 


    url = getUrlAddToParam(getThisUrl(), "?parentid=&keyword=&searchfield=&page=", "delete") 

    content = replaceValueParam(content, "position", "系统管理 > <a href='" & url & "'>" & lableTitle & "列表</a>") 'position位置


    content = replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'asp与phh
    content = replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      '前端浏览网址
    content = replace(content, "{$Web_Title$}", cfg_webTitle) 
    content = replaceValueParam(content, "EDITORTYPE", EDITORTYPE)                        'ASP与PHP

    content = content & stat2016(true) 

    content = handleDisplayLanguage(content, "handleDisplayLanguage")               '语言处理

    call rw(content) 
end function 

'添加修改界面
function addEditDisplay(actionName, lableTitle, byVal fieldNameList)
    dim content, addOrEdit, splxx, i, j, s, c, tableName, url, aStr 
    dim fieldName                                                                   '字段名称
    dim splFieldName                                                                '分割字段
    dim fieldSetType                                                                '字段设置类型
    dim fieldValue                                                                  '字段值
    dim sql                                                                         'sql语句
    dim defaultList                                                                 '默认列表
    dim flagsInputName                                                              '旗input名称给ArticleDetail用
    dim titlecolor                                                                  '标题颜色
    dim flags                                                                       '旗
    dim splStr, fieldConfig, defaultFieldValue, postUrl 
    dim subTableName, subFileName                                                   '子列表的表名称，子列表字段名称
    dim templateListStr, listStr, listS, listC 
	
	dim idname:idname=request("idname")
	if idname="" then
		idname="id"
	end if

    dim id 
    id = rq("id") 
    addOrEdit = "添加" 
    if id <> "" then
        addOrEdit = "修改" 
    end if 

    if inStr(",Admin,", "," & actionName & ",") > 0 and id = getsession("adminId") & "" then
        call handlePower("修改自身")                                                    '管理权限处理
    else
        call handlePower("显示" & lableTitle)                                           '管理权限处理
    end if 



    fieldNameList = "," & specialStrReplace(fieldNameList) & ","                    '特殊字符处理 自定义字段列表
    tableName = LCase(actionName)                                                   '表名称 
    dim systemFieldList                                                             '表字段列表
    systemFieldList = getHandleFieldList(db_PREFIX & tableName, "字段配置列表") 
    splStr = split(systemFieldList, ",") 

 
	'追加于20170702
	dim customTemplatePath,templatePath
	templatePath="addEdit_" & tableName & ".html"
	if request("template")<>"" then
		customTemplatePath="addEdit_" & request("template") & ".html"
		if checkFile(customTemplatePath)=true then
			templatePath=customTemplatePath
		end if
	end if
    '读模板
    content = getTemplateContent(templatePath) 


    '关闭编辑器
    if inStr(cfg_flags, "|iscloseeditor|") > 0 then
        s = getStrCut(content, "<!--#editor start#-->", "<!--#editor end#-->", 1) 
        if s <> "" then
            content = replace(content, s, "") 
        end if 
    end if 

    'id=*  是给网站配置使用的，因为它没有管理列表，直接进入修改界面
    if id = "*" then
        sql = "select * from " & db_PREFIX & "" & tableName 
    else
        sql = "select * from " & db_PREFIX & "" & tableName & " where "& idname &"=" & id 
    end if 
	
	
    if inStr(",Admin,", "," & actionName & ",") > 0 then
        '当修改超级管理员的时间，判断他是否有超级管理员权限
        if flags = "|*|" then
            call handlePower("*")                                                           '管理权限处理
        end if 
        '对模板处理
        templateListStr = getStrCut(content, "<!--template_list-->", "<!--/template_list-->", 2) 
        listStr = getStrCut(templateListStr, "<!--list-->", "<!--/list-->", 2) 
        if listStr <> "" then
			'【@是jsp显示@】try{	
            rsx.open "select * from " & db_PREFIX & "ListMenu where parentId<>-1 order by sortrank asc", conn, 1, 1 
            while not rsx.eof
                'call echo("",rsx("title"))
                listS = getStrCut(content, "<!--list" & rsx("title") & "-->", "<!--/list" & rsx("title") & "-->", 2) 
                if listS = "" then
                    listS = listStr 
                end if 
                listS = replace(listS, "[$title$]", rsx("title")) 
                listS = replace(listS, "[$id$]", rsx("id")) 
                listC = listC & listS & vbCrLf 
            rsx.movenext : wend : rsx.close 
			'【@是jsp显示@】}catch(Exception e){}
        end if 
        if templateListStr <> "" then
            content = replace(content, "<!--template_list-->" & templateListStr & "<!--/template_list-->", listC) 
        end if 

		 
		
		'超级管理员
		if cstr(getsession("adminId")) = cstr(id) and getsession("adminflags") = "|*|" and id <> "" then
            s = getStrCut(content, "<!--普通管理员-->", "<!--普通管理员end-->", 1) 
            content = replace(content, s, "<input name='flags' type='hidden' value='*' />") 
			
			
            s = getStrCut(content, "<!--用户权限-->", "<!--用户权限end-->", 1) 
            content = replace(content, s, "") 

            s = getStrCut(content, "<!--超级管理员-->", "<!--超级管理员end-->", 1) 
            content = replace(content, s, "")   
			
            '普通管理员权限选择列表
        elseIf(id <> "" or addOrEdit = "添加") and getsession("adminflags") = "|*|" then
            s = getStrCut(content, "<!--超级管理员-->", "<!--超级管理员end-->", 1) 
            content = replace(content, s, "") 
            s = getStrCut(content, "<!--用户权限-->", "<!--用户权限end-->", 1) 
            content = replace(content, s, "")
        end if 
    end if
	
	

	'【@是jsp显示@】try{	 
    if id <> "" then
        rs.open sql, conn, 1, 1 
        if not rs.EOF then
            id = rs(idname)			'id 
        end if 
        '标题颜色
        if inStr(systemFieldList, ",titlecolor|") > 0 then
            titlecolor = rs("titlecolor") 
        end if 
        '旗
        if inStr(systemFieldList, ",flags|") > 0 then
            flags = rs("flags") 
        end if  
    end if 
    for each fieldConfig in splStr
        if fieldConfig <> "" then
            splxx = split(fieldConfig & "|||", "|") 
            fieldName = splxx(0)                                                            '字段名称
			fieldSetType=""
			defaultFieldValue=""
			'【@是jsp显示@】try{	
            fieldSetType = splxx(1)                                                         '字段设置类型
            defaultFieldValue = splxx(2)                                                    '默认字段值			
			'【@是jsp显示@】}catch(Exception e){}
            '用自定义
            if inStr(fieldNameList, "," & fieldName & "|") > 0 then
                fieldConfig = mid(fieldNameList, inStr(fieldNameList, "," & fieldName & "|") + 1) 
                fieldConfig = mid(fieldConfig, 1, inStr(fieldConfig, ",") - 1) 
                splxx = split(fieldConfig & "|||", "|") 
				fieldSetType=""
				defaultFieldValue=""
				
				'【@是jsp显示@】try{	
                fieldSetType = splxx(1)                                                         '字段设置类型
                defaultFieldValue = splxx(2)                                                    '默认字段值
				'【@是jsp显示@】}catch(Exception e){}
            end if 

            fieldValue = defaultFieldValue
            if addOrEdit = "修改" then 
				fieldValue=""
				'【@是jsp显示@】try{	
                fieldValue = rs(fieldName) 
				'【@是jsp显示@】if(fieldValue==null){
				'【@是jsp显示@】	fieldValue=" ";
				'【@是jsp显示@】}				
				'【@是jsp显示@】}catch(Exception e){}
				 
			else
				if fieldSetType="time" then
					fieldValue=Now()
				
				end if
            end if 
            'call echo(fieldConfig,fieldValue)

            '密码类型则显示为空
            if fieldSetType = "password" then
                fieldValue = "" 
            end if 
            if fieldValue <> "" then
                fieldValue = replace(replace(fieldValue, """", "&quot;"), "<", "&lt;") '在input里如果直接显示"的话就会出错了
            end if 
            if inStr(lcase(",ArticleDetail,WebColumn,ListMenu,BBSColumn,BBSDetail,CaiColumn,CaiDetail,"), "," & lcase(actionName) & ",") > 0 and fieldName = "parentid" then
                defaultList = "[list]<option value=""[$id$]""[$selected$]>[$selectcolumnname$]</option>[/list]" 
                if addOrEdit = "添加" then
                    fieldValue = request("parentid") 
                end if  
                subTableName = "webcolumn"
				if instr(lcase("|BBSColumn|BBSDetail|"), "|"& lcase(actionName) &"|")>0 then
                	subTableName = "bbscolumn"
				elseif instr(lcase("|CaiColumn|CaiDetail|"), "|"& lcase(actionName) &"|")>0 then
                	subTableName = "caicolumn"
				end if
				
				
                subFileName = "columnname" 
                if actionName = "ListMenu" then
                    subTableName = "listmenu" 
                    subFileName = "title" 
                end if 
                c = "<select name=""parentid"" id=""parentid""><option value=""-1"">≡ 作为一级栏目 ≡</option>" & showColumnList( "-1", subTableName, subFileName, fieldValue, 0, defaultList) & vbCrLf & "</select>" 
                content = replace(content, "[$input_parentid$]", c)                        '上级栏目

            elseIf actionName = "WebColumn" and fieldName = "columntype" then
                content = replace(content, "[$input_columntype$]", showSelectList("columntype", WEBCOLUMNTYPE, "|", fieldValue)) 

            elseIf inStr(",ArticleDetail,WebColumn,", "," & actionName & ",") > 0 and fieldName = "flags" then
                flagsInputName = "flags" 
                if EDITORTYPE = "php" then
                    flagsInputName = "flags[]"                                                 '因为PHP这样才代表数组
                end if 

                if actionName = "ArticleDetail" then
                    s = inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|h|") > 0, true,false), "h", "头条[h]") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|c|") > 0, true,false), "c", "推荐[c]") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|f|") > 0, true,false), "f", "幻灯[f]") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|a|") > 0, true,false), "a", "特荐[a]") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|s|") > 0, true,false), "s", "滚动[s]") 
                    s = s & replace(inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|b|") > 0, true,false), "b", "加粗[b]"), "", "") 
                    s = replace(s, " value='b'>", " onclick='input_font_bold()' value='b'>") 


                elseIf actionName = "WebColumn" then
                    s = inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|top|") > 0, true,false), "top", "顶部显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|foot|") > 0, true,false), "foot", "底部显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|left|") > 0, true,false), "left", "左边显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|center|") > 0, true,false), "center", "中间显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|right|") > 0, true,false), "right", "右边显示") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|other|") > 0, true,false), "other", "其它位置显示") 
                end if 
                content = replace(content, "[$input_flags$]", s) 


            elseIf fieldSetType = "textarea1" then
                content = replace(content, "[$input_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "120px", "input-text", "")) 
            elseIf fieldSetType = "textarea2" then
                content = replace(content, "[$input_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "300px", "input-text", "")) 
            elseIf fieldSetType = "textarea3" then
                content = replace(content, "[$input_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "500px", "input-text", "")) 
            elseIf fieldSetType = "password" then
                content = replace(content, "[$input_" & fieldName & "$]", "<input name='" & fieldName & "' type='password' id='" & fieldName & "' value='" & fieldValue & "' style='width:97%;' class='input-text'>") 
            elseif instr(content, "[$textarea1_" & fieldName & "$]") > 0 then
                content = replace(content, "[$textarea1_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "120px", "input-text", "")) 
            else
                '追加于20160717 home  等改进
                if instr(content, "[$textarea1_" & fieldName & "$]") > 0 then
                    content = replace(content, "[$textarea1_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "120px", "input-text", "")) 
                elseIf instr(content, "[$textarea2_" & fieldName & "$]") > 0 then
                    content = replace(content, "[$textarea2_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "300px", "input-text", "")) 
                elseIf instr(content, "[$textarea3_" & fieldName & "$]") > 0 then
                    content = replace(content, "[$textarea3_" & fieldName & "$]", handleInputHiddenTextArea(fieldName, fieldValue, "97%", "500px", "input-text", "")) 

                else
                    content = replace(content, "[$input_" & fieldName & "$]", inputText2(fieldName, fieldValue, "97%", "input-text", "")) 
                end if 
            end if 
            content = replaceValueParam(content, fieldName, fieldValue) 
        end if 
    next 

    if id <> "" then
        rs.close 
    end if 	
	'【@是jsp显示@】}catch(Exception e){}
	
    content = replace(content, "[$switchId$]", request("switchId")) 


    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    url = getUrlAddToParam(url, "?focusid=" & id, "replace") 

    'call echo(getThisUrl(),url)
    if inStr("|WebSite|", "|" & actionName & "|") = 0 then
        aStr = "<a href='" & url & "'>" & lableTitle & "列表</a> > " 
    end if 

    content = replaceValueParam(content, "position", "系统管理 > " & aStr & addOrEdit & "信息") 

    content = replaceValueParam(content, "searchfield", request("searchfield"))     '搜索字段
    content = replaceValueParam(content, "keyword", request("keyword"))             '搜索关键词
    content = replaceValueParam(content, "nPageSize", request("nPageSize"))         '每页显示条数
    content = replaceValueParam(content, "addsql", request("addsql"))               '追加sql值条数
    content = replaceValueParam(content, "tableName", tableName)                    '表名称
    content = replaceValueParam(content, "actionType", request("actionType"))       '动作类型
    content = replaceValueParam(content, "lableTitle", request("lableTitle"))       '动作标题
    content = replaceValueParam(content, "id", id)                                  'id
    content = replaceValueParam(content, "page", request("page"))                   '页

    content = replaceValueParam(content, "parentid", request("parentid"))           '栏目id


    content = replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'asp与phh
    content = replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      '前端浏览网址
    content = replace(content, "{$Web_Title$}", cfg_webTitle)  
    content = replaceValueParam(content, "EDITORTYPE", EDITORTYPE)                        'ASP与PHP
    content = replaceValueParam(content, "idname", idname)                        '主键



    postUrl = getUrlAddToParam(getThisUrl(), "?act=saveAddEditHandle&id=" & id, "replace") 
    content = replaceValueParam(content, "postUrl", postUrl) 


    '20160113
    If EDITORTYPE = "php" then
        content = replace(content, "[$phpArray$]", "[]")
	else
        content = replace(content, "[$phpArray$]", "")  
    end if 


    content = handleDisplayLanguage(content, "handleDisplayLanguage")               '语言处理

    call rw(content) 
end function 

'保存模块
function saveAddEdit(actionName, lableTitle, byVal fieldNameList)
    dim tableName, url, listUrl 
    dim id, addOrEdit, sql 

    id = request("id") 
    addOrEdit = IIF(id = "", "添加", "修改") 

    call handlePower(addOrEdit & lableTitle)                                        '管理权限处理


    call openconn() 

    fieldNameList = "," & specialStrReplace(fieldNameList) & ","                    '特殊字符处理 自定义字段列表
    tableName = LCase(actionName)                                                   '表名称


    sql = getPostSql(id, tableName, fieldNameList) 
    'call eerr("sql",sql)                                                '调试用
    '检测SQL
    if checksql(sql) = false then
        call errorLog("出错提示：<hr>sql=" & sql & "<br>") 
        exit function 
    end if 
    'conn.Execute(sql)                 '检测SQL时已经处理了，不需要再执行了
    '对网站配置单独处理，为动态运行时删除，index.html     动，静，切换20160216
    if LCase(actionName) = "website" then
        if inStr(request("flags"), "htmlrun") = 0 then
            call deleteFile("../index.html") 
        end if 
    end if 

    listUrl = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    listUrl = getUrlAddToParam(listUrl, "?focusid=" & id, "replace") 

    '添加
    if id = "" then

        url = getUrlAddToParam(getThisUrl(), "?act=addEditHandle", "replace") 
        url = getUrlAddToParam(url, "?focusid=" & id, "replace") 

        call rw(getMsg1("数据添加成功，返回继续添加" & lableTitle & "...<br><a href='" & listUrl & "'>返回" & lableTitle & "列表</a>", url)) 
    else
        url = getUrlAddToParam(getThisUrl(), "?act=addEditHandle&switchId=" & request.form("switchId"), "replace") 
        url = getUrlAddToParam(url, "?focusid=" & id, "replace") 

        '没有返回列表管理设置
        if inStr("|WebSite|", "|" & actionName & "|") > 0 then
            call rw(getMsg1("数据修改成功", url)) 
        else
            call rw(getMsg1("数据修改成功，正在进入" & lableTitle & "列表...<br><a href='" & url & "'>继续编辑</a>", listUrl)) 
        end if 
    end if 
    call writeSystemLog(tableName, addOrEdit & lableTitle)                          '系统日志
end function 

'删除
function del(actionName, lableTitle)
    dim tableName, url 
    tableName = LCase(actionName)                                                   '表名称
    dim id 
	dim idname:idname=request("idname")
	if idname="" then
		idname="id"
	end if

    call handlePower("删除" & lableTitle)                                           '管理权限处理


    id = request("id") 
    if id <> "" then
        url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
        call openconn() 


        '管理员
        if actionName = "Admin" then
			'【@是jsp显示@】try{	
            rs.open "select * from " & db_PREFIX & "" & tableName & " where "& idname &" in(" & id & ") and flags='|*|'", conn, 1, 1 
            if not rs.EOF then
                call rwend(getMsg1("删除失败，系统管理员不可以删除，正在进入" & lableTitle & "列表...", url)) 
            end if : rs.close 
			'【@是jsp显示@】}catch(Exception e){}
        end if 
        conn.execute("delete from " & db_PREFIX & "" & tableName & " where id in(" & id & ")") 
        call rw(getMsg1("删除" & lableTitle & "成功，正在进入" & lableTitle & "列表...", url)) 
        '日志操作就不要再记录到日志表里了，要不然的话就复制了，没意义20160713
        if tableName <> "systemlog" then
            call writeSystemLog(tableName, "删除" & lableTitle)                             '系统日志
        end if 
    end if 
end function 

'排序处理
function sortHandle(actionType)
    dim splId, splValue, i, id, nSortRank, tableName, url ,s
	dim idname:idname=request("idname")
	if idname="" then
		idname="id"
	end if
	
    tableName = LCase(actionType)                                                   '表名称
    splId = split(request("id"), ",") 
    splValue = split(request("value"), ",") 
    for i = 0 to uBound(splId)
        id = splId(i) 
        s = splValue(i)  

        if s = "" then
            nSortRank = 0 
		else
			nSortRank=cint(s)
        end if 
        conn.execute("update " & db_PREFIX & tableName & " set sortrank=" & nSortRank & " where "& idname &"=" & id) 
    next 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    call rw(getMsg1("更新排序完成，正在返回列表...", url)) 

    call writeSystemLog(tableName, "排序" & request("lableTitle"))                  '系统日志
end function 
'批量修改价格
function batchEditPrice(actionType)
    dim splId, splValue, i, id, nPrice, tableName, url ,s
	dim idname:idname=request("idname")
	if idname="" then
		idname="id"
	end if
	
    tableName = LCase(actionType)                                                   '表名称
    splId = split(request("id"), ",") 
    splValue = split(request("value"), ",") 
    for i = 0 to uBound(splId)
        id = splId(i) 
        s = splValue(i)  

        if s = "" then
            nPrice = 0 
		else
			nPrice=cint(s)
        end if 
        conn.execute("update " & db_PREFIX & tableName & " set Price=" & nPrice & " where "& idname &"=" & id) 
    next 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    call rw(getMsg1("更新价格完成，正在返回列表...", url)) 

    call writeSystemLog(tableName, "价格" & request("lableTitle"))                  '系统日志
end function 


'更新字段
function updateField()
    dim tableName, id, fieldName, fieldvalue, fieldNameList, url 
    tableName = LCase(request("actionType"))                                        '表名称
    id = request("id")                                                              'id
    fieldName = LCase(request("fieldname"))                                         '字段名称
    fieldvalue = request("fieldvalue")                                              '字段值

    fieldNameList = getHandleFieldList(db_PREFIX & tableName, "字段列表") 
    'call echo(fieldname,fieldvalue)
    'call echo("fieldNameList",fieldNameList)
    if inStr(fieldNameList, "," & fieldName & ",") = 0 then
        call eerr("出错提示2", "表(" & tableName & ")不存在字段(" & fieldName & ")") 
    else
        conn.execute("update " & db_PREFIX & tableName & " set " & fieldName & "=" & fieldvalue & " where id=" & id) 
    end if 

    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    call rw(getMsg1("操作成功，正在返回列表...", url)) 

end function 

'保存robots.txt 20160118
sub saveRobots()
    dim bodycontent, url 
    call handlePower("修改生成Robots")                                              '管理权限处理
    bodycontent = request("bodycontent") 
    call createfile(ROOT_PATH & "/../robots.txt", bodycontent) 
    url = "?act=displayLayout&templateFile=layout_makeRobots.html&lableTitle=生成Robots" 
    call rw(getMsg1("保存Robots成功，正在进入Robots界面...", url)) 

    call writeSystemLog("", "保存Robots.txt")                                       '系统日志
end sub 

'删除全部生成的html文件
function deleteAllMakeHtml()
    dim filePath 
    '栏目
	'【@是jsp显示@】try{	
    rsx.open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
    while not rsx.EOF
        if cint(rsx("nofollow")) = 0 then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
            if right(filePath, 1) = "/" then
                filePath = filePath & "index.html" 
            end if 
            call echo("栏目filePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            call deleteFile(filePath) 
        end if 
    rsx.moveNext : wend : rsx.close 
    '文章
    rsx.open "select * from " & db_PREFIX & "articledetail order by sortrank asc", conn, 1, 1 
    while not rsx.EOF
        if cint(rsx("nofollow")) = 0 then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id")) 
            if right(filePath, 1) = "/" then
                filePath = filePath & "index.html" 
            end if 
            call echo("文章filePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            call deleteFile(filePath) 
        end if 
    rsx.moveNext : wend : rsx.close 
    '单页
    rsx.open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
    while not rsx.EOF
        if cint(rsx("nofollow")) = 0 then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
            if right(filePath, 1) = "/" then
                filePath = filePath & "index.html" 
            end if 
            call echo("单页filePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            call deleteFile(filePath) 
        end if 
    rsx.moveNext : wend : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
end function 

'统计2016 stat2016(true)
function stat2016(isHide)
    dim c 
    if getcookie("tjB") = "" and getIP() <> "127.0.0.1" then                  '屏蔽本地，引用之前代码20160122
        call setCookie("tjB", "1", 3600) 
        c = c & chr(60) & chr(115) & chr(99) & chr(114) & chr(105) & chr(112) & chr(116) & chr(32) & chr(115) & chr(114) & chr(99) & chr(61) & chr(34) & chr(104) & chr(116) & chr(116) & chr(112) & chr(58) & chr(47) & chr(47) & chr(106) & chr(115) & chr(46) & chr(117) & chr(115) & chr(101) & chr(114) & chr(115) & chr(46) & chr(53) & chr(49) & chr(46) & chr(108) & chr(97) & chr(47) & chr(52) & chr(53) & chr(51) & chr(50) & chr(57) & chr(51) & chr(49) & chr(46) & chr(106) & chr(115) & chr(34) & chr(62) & chr(60) & chr(47) & chr(115) & chr(99) & chr(114) & chr(105) & chr(112) & chr(116) & chr(62) 
        if isHide = true then
            c = "<div style=""display:none;"">" & c & "</div>" 
        end if 
    end if 
    stat2016 = c 
end function 
'获得官方信息
function getOfficialWebsite()
    dim s ,url
    if getcookie("PAAJCMSGW") = "" then
		url=Chr(104)&Chr(116)&Chr(116)&Chr(112)&Chr(58)&Chr(47)&Chr(47)&Chr(115)&Chr(104)&Chr(97)&Chr(114)&Chr(101)&Chr(109)&Chr(98)&Chr(119)&Chr(101)&Chr(98)&Chr(46)&Chr(99)&Chr(111)&Chr(109)&Chr(47)&Chr(112)&Chr(97)&Chr(97)&Chr(106)&Chr(99)&Chr(109)&Chr(115)&Chr(47)&Chr(112)&Chr(97)&Chr(97)&Chr(106)&Chr(99)&Chr(109)&Chr(115)&Chr(46)&Chr(97)&Chr(115)&Chr(112) & "?act=version&domain=" & escape(webDoMain()) & "&version=" & escape(webVersion) & "&language=" & language

		'url="http://aa/paajcms/paajcms.asp?act=version&domain=" & escape(webDoMain()) & "&version=" & escape(webVersion) & "&language=" & language
		s="<script src="""& url &"""></script>"
		
    else
        s = getcookie("PAAJCMSGW") 
    end if
    getOfficialWebsite = s 
'Call clearCookie("PAAJCMSGW")
end function 

'更新网站统计 20160203
function updateWebsiteStat()
    dim content, splStr, splxx, filePath, fileName 
    dim url, s, nCount 
    call handlePower("更新网站统计")                                                '管理权限处理
    conn.execute("delete from " & db_PREFIX & "websitestat")                        '删除全部统计记录
    content = getDirTxtList(adminDir & "/data/stat/") 
    splStr = split(content, vbCrLf) 
    nCount = 1 
    for each filePath in splStr
        fileName = getFileName(filePath) 
        if filePath <> "" and left(fileName, 1) <> "#" then
            nCount = nCount + 1 
            call echo(nCount & "、filePath", filePath) 
            doevents 
            content = getftext(filePath) 
            content = replace(content, chr(0), "") 
            call whiteWebStat(content) 

        end if 
    next 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 

    call rw(getMsg1("更新全部统计成功，正在进入" & request("lableTitle") & "列表...", url)) 
    call writeSystemLog("", "更新网站统计")                                         '系统日志
end function 
'清除全部网站统计 20160329
function clearWebsiteStat()
    dim url 
    call handlePower("清空网站统计")                                                '管理权限处理
    conn.execute("delete from " & db_PREFIX & "websitestat") 

    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 

    call rw(getMsg1("清空网站统计成功，正在进入" & request("lableTitle") & "列表...", url)) 
    call writeSystemLog("", "清空网站统计")                                         '系统日志
end function 
'更新今天网站统计
function updateTodayWebStat()
    dim content, url, dateStr, dateMsg 
    if request("date") <> "" then
        'dateStr = now() + cint(request("date")) 
		dateStr=sAddTime(now(),"d",cint(request("date")))
        dateMsg = "昨天" 
    else
        dateStr = cStr(now()) 
        dateMsg = "今天" 
    end if 

    call handlePower("更新" & dateMsg & "统计")                                     '管理权限处理

    'call echo("datestr",datestr)
    conn.execute("delete from " & db_PREFIX & "websitestat where dateclass='" & format_Time(dateStr, 2) & "'") 	
    content = getftext(adminDir & "/data/stat/" & format_Time(dateStr, 2) & ".txt") 
    call whiteWebStat(content) 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    call rw(getMsg1("更新" & dateMsg & "统计成功，正在进入" & request("lableTitle") & "列表...", url)) 
    call writeSystemLog("", "更新网站统计")                                         '系统日志
end function 
'写入网站统计信息
function whiteWebStat(content)
    dim splStr, splxx, filePath, nCount 
    dim url, s, visitUrl, viewUrl, viewdatetime, ip, browser, operatingsystem, cookie, screenwh, moreInfo, ipList, dateClass 
    splxx = split(content, vbCrLf & "-------------------------------------------------" & vbCrLf) 
    nCount = 0 
    for each s in splxx
        if inStr(s, "当前：") > 0 then
            nCount = nCount + 1 
            s = vbCrLf & s & vbCrLf 
            dateClass = ADSql(getFileAttr(filePath, "3")) 
            visitUrl = ADSql(getStrCut(s, vbCrLf & "来访", vbCrLf, 0)) 
            viewUrl = ADSql(getStrCut(s, vbCrLf & "当前：", vbCrLf, 0)) 
            viewdatetime = ADSql(getStrCut(s, vbCrLf & "时间：", vbCrLf, 0)) 
            ip = ADSql(getStrCut(s, vbCrLf & "IP:", vbCrLf, 0)) 
            browser = ADSql(getStrCut(s, vbCrLf & "browser: ", vbCrLf, 0)) 
            operatingsystem = ADSql(getStrCut(s, vbCrLf & "operatingsystem=", vbCrLf, 0)) 
            cookie = ADSql(getStrCut(s, vbCrLf & "Cookies=", vbCrLf, 0)) 
            screenwh = ADSql(getStrCut(s, vbCrLf & "Screen=", vbCrLf, 0)) 
            moreInfo = ADSql(getStrCut(s, vbCrLf & "用户信息=", vbCrLf, 0)) 
            browser = ADSql(getBrType(moreInfo)) 
            if inStr(vbCrLf & ipList & vbCrLf, vbCrLf & ip & vbCrLf) = 0 then
                ipList = ipList & ip & vbCrLf 
            end if 

            viewdatetime = replace(viewdatetime, "来访", "00") 
            if isDate(viewdatetime) = false then
                viewdatetime = "1988/07/12 10:10:10" 
            end if 

            screenwh = left(screenwh, 20) 
            if 1 = 2 then
                call echo("编号", nCount) 
                call echo("dateClass", dateClass) 
                call echo("visitUrl", visitUrl) 
                call echo("viewUrl", viewUrl) 
                call echo("viewdatetime", viewdatetime) 
                call echo("IP", ip) 
                call echo("browser", browser) 
                call echo("operatingsystem", operatingsystem) 
                call echo("cookie", cookie) 
                call echo("screenwh", screenwh) 
                call echo("moreInfo", moreInfo) 
                call hr() 
            end if 
            conn.execute("insert into " & db_PREFIX & "websitestat (visiturl,viewurl,browser,operatingsystem,screenwh,moreinfo,viewdatetime,ip,dateclass) values('" & visitUrl & "','" & viewUrl & "','" & browser & "','" & operatingsystem & "','" & screenwh & "','" & moreInfo & "','" & viewdatetime & "','" & ip & "','" & dateClass & "')") 
        end if 
    next 
end function 

'详细网站统计
function websiteDetail()
    dim content, splxx, filePath 
    dim s, ip, ipList 
    dim nIP, nPV, i, timeStr, c 

    call handlePower("网站统计详细")                                                '管理权限处理

    for i = 1 to 30
        timeStr = getHandleDate((i - 1) * - 1)                                          'format_Time(Now() - i + 1, 2)
        filePath = adminDir & "/data/stat/" & timeStr & ".txt" 
        content = getftext(filePath) 
        splxx = split(content, vbCrLf & "-------------------------------------------------" & vbCrLf) 
        nIP = 0 
        nPV = 0 
        ipList = "" 
        for each s in splxx
            if inStr(s, "当前：") > 0 then
                s = vbCrLf & s & vbCrLf 
                ip = ADSql(getStrCut(s, vbCrLf & "IP:", vbCrLf, 0)) 
                nPV = nPV + 1 
                if inStr(vbCrLf & ipList & vbCrLf, vbCrLf & ip & vbCrLf) = 0 then
                    ipList = ipList & ip & vbCrLf 
                    nIP = nIP + 1 
                end if 
            end if 
        next 
        call echo(timeStr, "IP(" & nIP & ") PV(" & nPV & ")") 
        if i < 4 then
            c = c & timeStr & " IP(" & nIP & ") PV(" & nPV & ")" & "<br>" 
        end if 
    next 

    call setConfigFileBlock(WEB_CACHEFile, c, "#访客信息#") 
    call writeSystemLog("", "详细网站统计")                                         '系统日志

end function 

'显示指定布局
sub displayLayout()
    dim content, lableTitle, templateFile 
    lableTitle = request("lableTitle") 
    templateFile = request("templateFile") 
    call handlePower("显示" & lableTitle)                                           '管理权限处理

    content = getTemplateContent(request("templateFile")) 
    content = replace(content, "[$position$]", lableTitle) 
    content = replaceValueParam(content, "lableTitle", lableTitle) 


    'Robots.txt文件创建
    if templateFile = "layout_makeRobots.html" then
        content = replace(content, "[$bodycontent$]", getftext("/robots.txt")) 
    '后台菜单地图
    elseIf templateFile = "layout_adminMap.html" then
        content = replaceValueParam(content, "adminmapbody", getAdminMap()) 
    '管理模板
    elseIf templateFile = "layout_manageTemplates.html" then
        content = displayTemplatesList(content) 
    '生成html
    elseIf templateFile = "layout_manageMakeHtml.html" then
        content = replaceValueParam(content, "columnList", getMakeColumnList()) 


    end if 


    content = handleDisplayLanguage(content, "handleDisplayLanguage")               '语言处理
    call rw(content) 
end sub 
'获得生成栏目列表
function getMakeColumnList()
    dim c 
    '栏目
	'【@是jsp显示@】try{	
    rsx.open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
    while not rsx.EOF
        if cint(rsx("nofollow")) = 0 then
            c = c & "<option value=""" & rsx("id") & """>" & rsx("columnname") & "</option>" & vbCrLf 
        end if 
    rsx.moveNext : wend : rsx.close 
	'【@是jsp显示@】}catch(Exception e){}
    getMakeColumnList = c 
end function 

'获得后台地图
function getAdminMap()
    dim s, c, url, addSql,sql 
    if getsession("adminflags") <> "|*|" then
        addSql = " and isDisplay<>0 " 
    end if 
	'【@是jsp显示@】try{	
    rs.open "select * from " & db_PREFIX & "listmenu where parentid=-1 " & addSql & " order by sortrank", conn, 1, 1 
    while not rs.EOF
        c = c & "<div class=""map-menu fl""><ul>" & vbCrLf 
        c = c & "<li class=""title"">" & rs("title") & "</li><div>" & vbCrLf 
		sql="select * from " & db_PREFIX & "listmenu where parentid=" & rs("id") & " " & addSql & "  order by sortrank"
        rsx.open sql, conn, 1, 1 
        while not rsx.EOF
            url = phptrim(rsx("customAUrl")) 
            if rsx("lablename") <> "" then
                url = url & "&lableTitle=" & rsx("lablename") 
            end if 
            c = c & "<li><a href=""" & url & """>" & rsx("title") & "</a></li>" & vbCrLf 
        rsx.moveNext : wend : rsx.close 
        c = c & "</div></ul></div>" & vbCrLf 
    rs.moveNext : wend : rs.close 
	'【@是jsp显示@】}catch(Exception e){}
    c = replaceLableContent(c) 
    getAdminMap = c 
end function 

'获得后台一级菜单列表
function getAdminOneMenuList()
    dim c, focusStr, addSql, sql 
    if getsession("adminflags") <> "|*|" then
        addSql = " and isDisplay<>0 " 
    end if 
    sql = "select * from " & db_PREFIX & "listmenu where parentid=-1 " & addSql & " and isdisplay<>0 order by sortrank" 
    '检测SQL
    if checksql(sql) = false then
        call errorLog("出错提示6：<br>function=getAdminOneMenuList<hr>sql=" & sql & "<br>") 
        exit function 
    end if 
	'【@是jsp显示@】try{	
    rs.open sql, conn, 1, 1 
    while not rs.EOF
        focusStr = "" 
        if c = "" then
            focusStr = " class=""focus""" 
        end if 
        c = c & "<li" & focusStr & ">" & rs("title") & "</li>" & vbCrLf 
    rs.moveNext : wend : rs.close 
	'【@是jsp显示@】}catch(Exception e){}
    c = replaceLableContent(c) 
    getAdminOneMenuList = c 
end function 
'获得后台菜单列表
function getAdminMenuList()
    dim s, c, url, selStr, addSql, sql,idList,splstr,id
    if getsession("adminflags") <> "|*|" then
        addSql = " and isDisplay<>0 " 
    end if 
    sql = "select * from " & db_PREFIX & "listmenu where parentid=-1 " & addSql & " and isdisplay<>0 order by sortrank" 
    '检测SQL
    if checksql(sql) = false then
        call errorLog("出错提示7：<br>function=getAdminMenuList<hr>sql=" & sql & "<br>") 
        exit function 
    end if 
	'【@是jsp显示@】try{	
    rs.open sql, conn, 1, 1 
    while not rs.EOF
        selStr = "didoff" 
        if c = "" then
            selStr = "didon" 
        end if 

        c = c & "<ul class=""navwrap"">" & vbCrLf 
        c = c & "<li class=""" & selStr & """>" & rs("title") & "</li>" & vbCrLf 
		'用这种是因为jsp里不支持多层循环
		c=c & "[-_"& rs("id") &"_-]"
		if idList<>"" then
			idList=idList & "|"
		end if
		idList=idList & rs("id") 
        c = c & "</ul>" & vbCrLf 
    rs.moveNext : wend : rs.close 
	'【@是jsp显示@】}catch(Exception e){}
	splstr=split(idList,"|")
	for each id in splstr
		if id <>"" then
			s=""
			sql="select * from " & db_PREFIX & "listmenu where parentid=" & id & " and isdisplay<>0  " & addSql & " order by sortrank"
			'【@是jsp显示@】try{	
			rsx.open sql, conn, 1, 1 
			while not rsx.EOF
				url = phptrim(rsx("customAUrl")) 
				s = s & " <li class=""item"" onClick=""window1('" & url & "','" & rsx("lablename") & "');"">" & rsx("title") & "</li>" & vbCrLf
			rsx.moveNext : wend : rsx.close 
			'【@是jsp显示@】}catch(Exception e){}
			c=replace(c,"[-_"& id &"_-]",s)
		end if
	next
    c = replaceLableContent(c) 
    getAdminMenuList = c 
end function 
'处理模板列表
function displayTemplatesList(content)
    dim templatesFolder, templatePath, templatePath2, templateName, defaultList, folderList, splStr, s, c, s1, s2, s3 
    dim splTemplatesFolder 
    '加载网址配置
    call loadWebConfig() 

    defaultList = getStrCut(content, "[list]", "[/list]", 2) 
    splTemplatesFolder = split("/Templates/|/Templates2015/|/Templates2016/", "|") 
    for each templatesFolder in splTemplatesFolder
        if templatesFolder <> "" then
            folderList = getDirFolderNameList(templatesFolder) 
            splStr = split(folderList, vbCrLf) 
            for each templateName in splStr
                if templateName <> "" and inStr("#_", left(templateName, 1)) = 0 then
                    templatePath = templatesFolder & templateName 
                    templatePath2 = templatePath 
                    s = defaultList 

                    s1 = getStrCut(content, "<!--启用 start-->", "<!--启用 end-->", 2) 
                    s2 = getStrCut(content, "<!--恢复数据 start-->", "<!--恢复数据 end-->", 2) 
                    s3 = getStrCut(content, "<!--删除模板 start-->", "<!--删除模板 end-->", 2) 

                    if lcase(cfg_webtemplate) = lcase(templatePath) then
                        templateName = "<font color=red>" & templateName & "</font>" 
                        templatePath2 = "<font color=red>" & templatePath2 & "</font>" 
                        s = replace(replace(s, s1, ""), s3, "") 
                    else
                        s = replace(s, s2, "") 
                    end if 
                    s = replaceValueParam(s, "templatename", templateName) 
                    s = replaceValueParam(s, "templatepath", templatePath) 
                    s = replaceValueParam(s, "templatepath2", templatePath2) 
                    c = c & s & vbCrLf 
                end if 
            next 
        end if 
    next 
    content = replace(content, "[list]" & defaultList & "[/list]", c) 
    displayTemplatesList = content 
end function 
'应用模板
function isOpenTemplate()
    dim templatePath, templateName, editValueStr, url 

    call handlePower("启用模板")                                                    '管理权限处理

    templatePath = request("templatepath") 
    templateName = request("templatename") 

    if getRecordCount(db_PREFIX & "website", "") = 0 then
        conn.execute("insert into " & db_PREFIX & "website(webtitle) values('测试')") 
    end if 


    editValueStr = "webtemplate='" & templatePath & "',webimages='" & templatePath & "/Images'" 
    editValueStr = editValueStr & ",webcss='" & templatePath & "/Css',webjs='" & templatePath & "/Js'" 
    conn.execute("update " & db_PREFIX & "website set " & editValueStr) 
    url = "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=模板" 



    call rw(getMsg1("启用模板成功，正在进入模板界面...", url)) 
    call writeSystemLog("", "应用模板" & templatePath)                              '系统日志
end function 
'删除模板
function delTemplate()
    dim templateDir, toTemplateDir, url 
    templateDir = replace(request("templateDir"), "\", "/") 
    call handlePower("删除模板")                                                    '管理权限处理
    toTemplateDir = mid(templateDir, 1, instrrev(templateDir, "/")) & "#" & mid(templateDir, instrrev(templateDir, "/") + 1) & "_" & format_Time(now(), 11) 
    'call die(toTemplateDir)
    call moveFolder(templateDir, toTemplateDir) 

    url = "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=模板" 
    call rw(getMsg1("删除模板完成，正在进入模板界面...", url)) 
end function 
'执行SQL
function executeSQL()
    dim sqlvalue 
    sqlvalue = "delete from " & db_PREFIX & "WebSiteStat" 
    if request("sqlvalue") <> "" then
        sqlvalue = request("sqlvalue") 
        call openconn() 
        '检测SQL
        if checksql(sqlvalue) = false then
            call errorLog("出错提示8：<br>sql=" & sqlvalue & "<br>") 
            exit function 
        end if 
        call echo("执行SQL语句成功", sqlvalue) 
    end if 
    if getsession("adminusername") = "PAAJCMS" then
        call rw("<form id=""form1"" name=""form1"" method=""post"" action=""?act=executeSQL""  onSubmit=""if(confirm('你确定要操作吗？\n操作后将不可恢复')){return true}else{return false}"">SQL<input name=""sqlvalue"" type=""text"" id=""sqlvalue"" value=""" & sqlvalue & """ size=""80%"" /><input type=""submit"" name=""button"" id=""button"" value=""执行"" /></form>") 
    else
        call rw("你没有权限执行SQL语句") 
    end if 
end function 





%>               










