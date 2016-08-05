<%
'************************************************************
'���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2016-08-05
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'��̨�������ĳ��� ��� ɾ�� �޸� �б�

'����function�ļ�����
function callFunction()
    select case request("stype")
        case "updateWebsiteStat" : updateWebsiteStat()                                  '������վͳ��
        case "clearWebsiteStat" : clearWebsiteStat()                                    '�����վͳ��
        case "updateTodayWebStat" : updateTodayWebStat()                                '������վ����ͳ��
        case "websiteDetail" : websiteDetail()                                          '��ϸ��վͳ��
        case "displayAccessDomain" : displayAccessDomain()                              '��ʾ��������
        case "delTemplate" : delTemplate()                                              'ɾ��ģ��
 
        case else : call eerr("function1ҳ��û�ж���", request("stype"))
    end select
end function 

'��ʾ��������
function displayAccessDomain()
    dim visitWebSite, visitWebSiteList, urlList, nOK 
    call handlePower("��ʾ��������") 
    call openconn() 
    nOK = 0 
    rs.open "select * from " & db_PREFIX & "websitestat", conn, 1, 1 
    while not rs.EOF
        visitWebSite = lcase(getWebSite(rs("visiturl"))) 
        'call echo("visitWebSite",visitWebSite)
        if instr(vbCrLf & visitWebSiteList & vbCrLf, vbCrLf & visitWebSite & vbCrLf) = 0 then
            if visitWebSite <> lcase(getWebSite(webDoMain())) then
                visitWebSiteList = visitWebSiteList & visitWebSite & vbCrLf 
                nOK = nOK + 1 
                urlList = urlList & nOK & "��<a href='" & rs("visiturl") & "' target='_blank'>" & rs("visiturl") & "</a><br>" 
            end if 
        end if 
    rs.movenext : wend : rs.close 
    call echo("��ʾ��������", "������� <a href='javascript:history.go(-1)'>�������</a>") 
    call rwend(visitWebSiteList & "<br><hr><br>" & urlList) 
end function 
'��ô������б� 20160313
function getHandleTableList()
    dim s, lableStr 
    lableStr = "���б�[" & request("mdbpath") & "]" 
    if WEB_CACHEContent = "" then
        WEB_CACHEContent = getftext(WEB_CACHEFile) 
    end if 
    s = getConfigContentBlock(WEB_CACHEContent, "#" & lableStr & "#") 
    if s = "" then
        s = LCase(getTableList()) 
        s = "|" & replace(s, vbCrLf, "|") & "|" 
        WEB_CACHEContent = setConfigFileBlock(WEB_CACHEFile, s, "#" & lableStr & "#") 
        if isCacheTip = true then
            call echo("����", lableStr) 
        end if 
    end if 
    getHandleTableList = s 
end function 

'��ô�����ֶ��б�   getHandleFieldList("ArticleDetail","�ֶ��б�")
function getHandleFieldList(tableName, sType)
    dim s 
    if WEB_CACHEContent = "" then
        WEB_CACHEContent = getftext(WEB_CACHEFile) 
    end if 
    s = getConfigContentBlock(WEB_CACHEContent, "#" & tableName & sType & "#") 

    if s = "" then
        if sType = "�ֶ������б�" then
            s = LCase(getFieldConfigList(tableName)) 
        else
            s = LCase(getFieldList(tableName)) 
        end if 
        WEB_CACHEContent = setConfigFileBlock(WEB_CACHEFile, s, "#" & tableName & sType & "#") 
        if isCacheTip = true then
            call echo("����", tableName & sType) 
        end if 
    end if 
    getHandleFieldList = s 
end function 
'��ģ������ 20160310
function getTemplateContent(templateFileName)
    call loadWebConfig() 
    '��ģ��
    dim templateFile, customTemplateFile, c 
    customTemplateFile = ROOT_PATH & "template/" & db_PREFIX & "/" & templateFileName 
    'Ϊ�ֻ���
    if checkMobile() = true or request("m") = "mobile" then
        templateFile = ROOT_PATH & "/Template/mobile/" & templateFileName 
    end if 
    '�ж��ֻ����ļ��Ƿ����20160330
    if checkFile(templateFile) = false then
        if checkFile(customTemplateFile) = true then
            templateFile = customTemplateFile 
        else
            templateFile = ROOT_PATH & templateFileName 
        end if 
    end if 
    c = getFText(templateFile) 
    c = replaceLableContent(c) 
    getTemplateContent = c 
end function 
'�滻��ǩ����
function replaceLableContent(content)
    dim s, c, splStr, list 
    content = replace(content, "{$webVersion$}", webVersion)                        '��վ�汾
    content = replace(content, "{$Web_Title$}", cfg_webTitle)                       '��վ����
    content = replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'ASP��PHP
    content = replace(content, "{$adminDir$}", adminDir)                            '��̨Ŀ¼

    content = replace(content, "[$adminId$]", session("adminId"))              '����ԱID
    content = replace(content, "{$adminusername$}", session("adminusername"))       '�����˺�����
    content = replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        '��������
    content = replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      'ǰ̨
    content = replace(content, "{$webVersion$}", webVersion)                        '�汾
    content = replace(content, "{$WebsiteStat$}", getConfigFileBlock(WEB_CACHEFile, "#�ÿ���Ϣ#")) '����ÿ���Ϣ


    content = replace(content, "{$databaseType$}", databaseType)                          '����Ϊ����
    content = replace(content, "{$DB_PREFIX$}", db_PREFIX)                          '��ǰ׺
    content = replace(content, "{$adminflags$}", IIF(session("adminflags") = "|*|", "��������Ա", "��ͨ����Ա")) '����Ա����
    content = replace(content, "{$SERVER_SOFTWARE$}", request.serverVariables("SERVER_SOFTWARE")) '�������汾
    content = replace(content, "{$SERVER_NAME$}", request.serverVariables("SERVER_NAME")) '��������ַ
    content = replace(content, "{$LOCAL_ADDR$}", request.serverVariables("LOCAL_ADDR")) '������IP
    content = replace(content, "{$SERVER_PORT$}", request.serverVariables("SERVER_PORT")) '�������˿�
    content = replaceValueParam(content, "mdbpath", request("mdbpath")) 
    content = replaceValueParam(content, "webDir", webDir) 
    content = replaceValueParam(content, "EDITORTYPE", EDITORTYPE)                        'ASP��PHP
 
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
        content = replace(content, "{$EDITORTYPE_PHP$}", "php")                         '��phpinc/��
    end if 
    content = replace(content, "{$EDITORTYPE_PHP$}", "")                            '��phpinc/��

    replaceLableContent = content 
end function 

'�����б���
function displayFlags(flags)
    dim c 
    'ͷ��[h]
    if inStr("|" & flags & "|", "|h|") > 0 then
        c = c & "ͷ " 
    end if 
    '�Ƽ�[c]
    if inStr("|" & flags & "|", "|c|") > 0 then
        c = c & "�� " 
    end if 
    '�õ�[f]
    if inStr("|" & flags & "|", "|f|") > 0 then
        c = c & "�� " 
    end if 
    '�ؼ�[a]
    if inStr("|" & flags & "|", "|a|") > 0 then
        c = c & "�� " 
    end if 
    '����[s]
    if inStr("|" & flags & "|", "|s|") > 0 then
        c = c & "�� " 
    end if 
    '�Ӵ�[b]
    if inStr("|" & flags & "|", "|b|") > 0 then
        c = c & "�� " 
    end if 
    if c <> "" then c = "[<font color=""red"">" & c & "</font>]" 

    displayFlags = c 
end function 


'��Ŀ���ѭ������        showColumnList(parentid, "webcolumn", ,"",0, defaultStr,3,"")   nCountΪ���ֵ   thisPIdΪ�����id
function showColumnList(byVal parentid, byVal tableName, showFieldName, byVal thisPId, nCount, byVal action)
    dim i, s, c, selectcolumnname, selStr, url, isFocus, sql, addSql, listLableStr, topnav 
    dim thisColumnName, sNavheaderStr, sNavfooterStr 
	
	
	
    parentid = trim(parentid) 
    listLableStr = "list" 

    topnav = getStrCut(action, "[topnav]", "[/topnav]", 2) 
    thisColumnName = getColumnName(parentid) 
    'call echo(parentid,topnav)

    if parentid <> topnav then
        if instr(action, "[small-list") > 0 then
            listLableStr = "small-list" 
        end if 
    end if 
    'call echo("listLableStr",listLableStr)
    dim rs : set rs = createObject("Adodb.RecordSet")
	'��@��.netc��ʾ@��OleDbDataReader rs=null;				//Ҫ��������
        dim fieldNameList, splFieldName, nK, fieldName, replaceStr, startStr, endStr, nTop, nModI, title 
        dim subHeaderStr, subFooterStr, subHeaderStartStr, subHeaderEndStr, subFooterStartStr, subFooterEndStr 


        fieldNameList = getHandleFieldList(db_PREFIX & tableName, "�ֶ��б�") 
        splFieldName = split(fieldNameList, ",") 
        sql = "select * from " & db_PREFIX & tableName & " where parentid=" & parentid 
        'call echo("sql1111111111111",tableName)
        '����׷��SQL
        startStr = "[sql-" & nCount & "]" : endStr = "[/sql-" & nCount & "]" 
        if inStr(action, startStr) = false and inStr(action, endStr) = 0 then
            startStr = "[sql]" : endStr = "[/sql]" 
        end if 
        addSql = getStrCut(action, startStr, endStr, 2) 
        if addSql <> "" then
            sql = getWhereAnd(sql, addSql) 
        end if 
        sql = sql & " order by sortrank asc" 
        rs.open sql, conn, 1, 1 
        'call echo(sql,rs.recordcount)
        for i = 1 to rs.recordCount
            if not rs.EOF then
                startStr = "" : endStr = "" 
                selStr = "" 
                isFocus = false 
                if CStr(rs("id")) = CStr(thisPId) then
                    selStr = " selected " 
                    isFocus = true 
                end if 
                '��ַ�ж�
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

                '�����ʱ����ǰ����20160202
                if i = nTop and isFocus = false then
                    startStr = "[" & listLableStr & "-end]" : endStr = "[/" & listLableStr & "-end]" 
                end if 
                '��[list-mod2]  [/list-mod2]    20150112
                for nModI = 6 to 2 step - 1
                    if inStr(action, startStr) = false and i mod nModI = 0 then
                        startStr = "[" & listLableStr & "-mod" & nModI & "]" : endStr = "[/" & listLableStr & "-mod" & nModI & "]" 
                        if inStr(action, startStr) > 0 then
                            exit for 
                        end if 
                    end if 
                next 

                'û������Ĭ��
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
                        selectcolumnname = copystr("&nbsp;&nbsp;", nCount) & "����" & selectcolumnname 
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

                    'url = WEB_VIEWURL & "?act=nav&columnName=" & rs(showFieldName)             '����Ŀ������ʾ�б�
                    url = WEB_VIEWURL & "?act=nav&id=" & rs("id")                               '����ĿID��ʾ�б�



                    '�Զ�����ַ
                    if trim(rs("customaurl")) <> "" then
                        url = trim(rs("customaurl")) 
                    end if 
                    s = replace(s, "[$viewWeb$]", url) 
                    s = replaceValueParam(s, "url", url) 

                    '��վ��Ŀû��pageλ�ô��� ׷����20160716 home
                    url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" & rs("id") & "&n=" & getRnd(11) 
                    s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span")                    '�����Ƿ���������޸Ĺ�����


                    if EDITORTYPE = "php" then
                        s = replace(s, "[$phpArray$]", "[]") 
                    else
                        s = replace(s, "[$phpArray$]", "") 
                    end if 

                    's=copystr("",nCount) & rs("columnname") & "<hr>"
                    if rs("parentid") = "-1" and instr(action, "[navheader]") > 0 then
                        sNavheaderStr = getStrCut(action, "[navheader]", "[/navheader]", 2) 
                        sNavfooterStr = getStrCut(action, "[navfooter]", "[/navfooter]", 2) 
                    'call die(sNavfooterStr)
                    end if 
                    c = c & sNavheaderStr & s & vbCrLf 
                    s = showColumnList(rs("id"), tableName, showFieldName, thisPId, nCount + 1, action) & sNavfooterStr 


                    subHeaderStartStr = "[subheader-" & rs("columnname") & "]" : subHeaderEndStr = "[/subheader-" & rs("columnname") & "]" 
                    if instr(action, subHeaderStartStr) = 0 and instr(action, subHeaderEndStr) = 0 then
                        subHeaderStartStr = "[subheader]" : subHeaderEndStr = "[/subheader]" 

                    end if 
                    subFooterStartStr = "[subfooter-" & rs("columnname") & "]" : subFooterEndStr = "[/subfooter-" & rs("columnname") & "]" 
                    if instr(action, subFooterStartStr) = 0 and instr(action, subFooterStartStr) = 0 then
                        subFooterStartStr = "[subfooter]" : subFooterEndStr = "[/subfooter]" 
                    end if 
                    subHeaderStr = getStrCut(action, subHeaderStartStr, subHeaderEndStr, 2) 
                    subFooterStr = getStrCut(action, subFooterStartStr, subFooterEndStr, 2) 
                    'call echo(rs("columnname"),"����")

                    if s <> "" then s = vbCrLf & subHeaderStr & s & subFooterStr 
                    c = c & s 
                end if 
            end if 
        rs.moveNext : next : rs.close 
        showColumnList = c 
end function
'msg1  ����
function getMsg1(msgStr, url)
    dim content 
    content = getFText(ROOT_PATH & "msg.html") 
    msgStr = msgStr & "<br>" & jsTiming(url, 5) 
    content = replace(content, "[$msgStr$]", msgStr) 
    content = replace(content, "[$url$]", url) 


    content = replaceL(content, "��ʾ��Ϣ") 
    content = replaceL(content, "������������û���Զ���ת����������") 
    content = replaceL(content, "����ʱ") 


    getMsg1 = content 
end function 

'���Ȩ��
function checkPower(powerName)
	dim sql
	checkPower=false
    if session("adminId") <> "" then
        call openconn()                                                                 '�����ݿ� Ҫ��Ȼ��php������
        '����������������ʱ��
		sql="select * from " & db_PREFIX & "admin where id=" & session("adminId")
        rss.open sql, conn, 1, 1 
        if not rss.eof then
            session("adminflags") = rss("flags") 
        end if : rss.close 
        if inStr("|" & session("adminflags") & "|", "|" & powerName & "|") > 0 or inStr("|" & session("adminflags") & "|", "|*|") > 0 then
            checkPower = true 
        else
            checkPower = false 
        end if 
    else
        checkPower = true 
    end if 
end function 
'�����̨����Ȩ��
function handlePower(powerName)
    if checkPower(powerName) = false then
        call eerr("��ʾ", "��û�С�" & powerName & "��Ȩ�ޣ�<a href='javascript:history.go(-1);'>�������</a>") 
    end if 
end function 
'��ʾ�����б�
function dispalyManage(actionName, lableTitle, byVal nPageSize, addSql)
    call handlePower("��ʾ" & lableTitle)                                           '����Ȩ�޴���
    call loadWebConfig() 
    dim content, i, s, c, fieldNameList, sql, action 
    dim nX, url, nCount, nPage 
    dim idInputName 

    dim tableName, j, splxx 
    dim fieldName                                                                   '�ֶ�����
    dim splFieldName                                                                '�ָ��ֶ�
    dim searchfield, keyWord                                                        '�����ֶΣ������ؼ���
    dim parentid                                                                    '��Ŀid

    dim replaceStr                                                                  '�滻�ַ�
    tableName = LCase(actionName)                                                   '������

    searchfield = request("searchfield")                                            '��������ֶ�ֵ
    keyWord = request("keyword")                                                    '��������ؼ���ֵ
    if request.form("parentid") <> "" then
        parentid = request.form("parentid") 
    else
        parentid = request.queryString("parentid") 
    end if 

    dim id 
    dim focusid                                                                     '���жϴ�������id�Ƿ��ڵ�ǰ�б����ǽ���20160715 home
    id = rq("id") 
    focusid = rq("focusid") 

    fieldNameList = getHandleFieldList(db_PREFIX & tableName, "�ֶ��б�") 

    fieldNameList = specialStrReplace(fieldNameList)                                '�����ַ�����
    splFieldName = split(fieldNameList, ",")                                        '�ֶηָ������

    '��ģ��
    content = getTemplateContent("manage_" & tableName & ".html") 

    action = getStrCut(content, "[list]", "[/list]", 2) 
    '��վ��Ŀ��������      ��Ŀ��һ��20160301
    if actionName = "WebColumn" then
        action = getStrCut(content, "[action]", "[/action]", 1) 
        content = replace(content, action, showColumnList( "-1", "WebColumn", "columnname", "", 0, action)) 
    elseIf actionName = "ListMenu" then
        action = getStrCut(content, "[action]", "[/action]", 1) 
        content = replace(content, action, showColumnList( "-1", "listmenu", "title", "", 0, action)) 
    else
        if keyWord <> "" and searchfield <> "" then
            addSql = getWhereAnd(" where " & searchfield & " like '%" & keyWord & "%' ", addSql) 
        end if 
        if parentid <> "" then
            addSql = getWhereAnd(" where parentid=" & parentid & " ", addSql) 
        end if 
        'call echo(tableName,addsql)
        sql = "select * from " & db_PREFIX & tableName & " " & addSql 
        '���SQL
        if checksql(sql) = false then
            call errorLog("������ʾ��<br>action=" & action & "<hr>sql=" & sql & "<br>") 
            exit function 
        end if 
        rs.open sql, conn, 1, 1 
        '��PHP��ɾ��rs
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
            nX = getRsPageNumber(rs, nCount, nPageSize, nPage)     		'��@����asp����@��
			
		elseif EDITORTYPE = "aspx" then 	
				
			'��@��.netc��ʾ@��int  nCountPage = getCountPage(nCount, nPageSize);
			'��@��.netc��ʾ@��if(nPage<=1){
			'��@��.netc��ʾ@��	nX=nPageSize;
			'��@��.netc��ʾ@��	if(nX>nCount){
			'��@��.netc��ʾ@��		nX=nCount;
			'��@��.netc��ʾ@��	} 
			'��@��.netc��ʾ@��}else{
			'��@��.netc��ʾ@��	for(int nI2=0;nI2<nPageSize*(nPage-1);nI2++){
            '��@��.netc��ʾ@��		rs.Read();
			'��@��.netc��ʾ@��	}
			'��@��.netc��ʾ@��	if(nPage<nCountPage){
			'��@��.netc��ʾ@��		nX=nPageSize;
			'��@��.netc��ʾ@��	}else{
			'��@��.netc��ʾ@��		nX=nCount-nPageSize*(nPage-1);
			'��@��.netc��ʾ@��	}
			'��@��.netc��ʾ@��} 
			
        else									 
            if nPage<>0 then'��@��.netc����@��
                nPage = nPage - 1 '��@��.netc����@��
            end if '��@��.netc����@��
            sql = "select * from " & db_PREFIX & "" & tableName & " " & addSql & " limit " & nPageSize * nPage & "," & nPageSize '��@��.netc����@��
            rs.open sql, conn, 1, 1 '��@��.netc����@��
            '��PHP��ɾ��rs
            nX = rs.recordCount '��@��.netc����@��
        end if 
        for i = 1 to nX
            '��PHP��$rs=mysql_fetch_array($rsObj);                                            //��PHP�ã���Ϊ�� asptophpת��������  ����
            '��@��.netc��ʾ@��rs.Read();
			s = replace(action, "[$id$]", rs("id")) 
            for j = 0 to uBound(splFieldName)
                if splFieldName(j) <> "" then
                    splxx = split(splFieldName(j) & "|||", "|") 
                    fieldName = splxx(0) 
                    replaceStr = rs(fieldName) & "" 
                    '�������촦��
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
            url = "��NO��" 
            if actionName = "ArticleDetail" then
                url = WEB_VIEWURL & "?act=detail&id=" & rs("id") 
            elseIf actionName = "OnePage" then
                url = WEB_VIEWURL & "?act=onepage&id=" & rs("id") 
            '�����ۼ�Ԥ��=����  20160129
            elseIf actionName = "TableComment" then
                url = WEB_VIEWURL & "?act=detail&id=" & rs("itemid") 
            end if 
            '�������Զ����ֶ�
            if inStr(fieldNameList, "customaurl") > 0 then
                '�Զ�����ַ
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
        content = replace(content, "[list]" & action & "[/list]", c) 
        '���ύ����parentid(��ĿID) searchfield(�����ֶ�) keyword(�ؼ���) addsql(����)
        url = "?page=[id]&addsql=" & request("addsql") & "&keyword=" & request("keyword") & "&searchfield=" & request("searchfield") & "&parentid=" & request("parentid") 
        url = getUrlAddToParam(getUrl(), url, "replace") 
        'call echo("url",url)
        content = replace(content, "[list]" & action & "[/list]", c) 

    end if 

    if inStr(content, "[$input_parentid$]") > 0 then
        action = "[list]<option value=""[$id$]""[$selected$]>[$selectcolumnname$]</option>[/list]" 
        c = "<select name=""parentid"" id=""parentid""><option value="""">�� ѡ����Ŀ ��</option>" & showColumnList( "-1", "webcolumn", "columnname", parentid, 0, action) & vbCrLf & "</select>" 
        content = replace(content, "[$input_parentid$]", c)                        '�ϼ���Ŀ
    end if 

    content = replaceValueParam(content, "searchfield", request("searchfield"))     '�����ֶ�
    content = replaceValueParam(content, "keyword", request("keyword"))             '�����ؼ���
    content = replaceValueParam(content, "nPageSize", request("nPageSize"))         'ÿҳ��ʾ����
    content = replaceValueParam(content, "addsql", request("addsql"))               '׷��sqlֵ����
    content = replaceValueParam(content, "tableName", tableName)                    '������
    content = replaceValueParam(content, "actionType", request("actionType"))       '��������
    content = replaceValueParam(content, "lableTitle", request("lableTitle"))       '��������
    content = replaceValueParam(content, "id", id)                                  'id
    content = replaceValueParam(content, "page", request("page"))                   'ҳ

    content = replaceValueParam(content, "parentid", request("parentid"))           '��Ŀid
    content = replaceValueParam(content, "focusid", focusid) 


    url = getUrlAddToParam(getThisUrl(), "?parentid=&keyword=&searchfield=&page=", "delete") 

    content = replaceValueParam(content, "position", "ϵͳ���� > <a href='" & url & "'>" & lableTitle & "�б�</a>") 'positionλ��


    content = replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'asp��phh
    content = replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      'ǰ�������ַ
    content = replace(content, "{$Web_Title$}", cfg_webTitle) 
    content = replaceValueParam(content, "EDITORTYPE", EDITORTYPE)                        'ASP��PHP

    content = content & stat2016(true) 

    content = handleDisplayLanguage(content, "handleDisplayLanguage")               '���Դ���

    call rw(content) 
end function 

'����޸Ľ���
function addEditDisplay(actionName, lableTitle, byVal fieldNameList)
    dim content, addOrEdit, splxx, i, j, s, c, tableName, url, aStr 
    dim fieldName                                                                   '�ֶ�����
    dim splFieldName                                                                '�ָ��ֶ�
    dim fieldSetType                                                                '�ֶ���������
    dim fieldValue                                                                  '�ֶ�ֵ
    dim sql                                                                         'sql���
    dim defaultList                                                                 'Ĭ���б�
    dim flagsInputName                                                              '��input���Ƹ�ArticleDetail��
    dim titlecolor                                                                  '������ɫ
    dim flags                                                                       '��
    dim splStr, fieldConfig, defaultFieldValue, postUrl 
    dim subTableName, subFileName                                                   '���б�ı����ƣ����б��ֶ�����
    dim templateListStr, listStr, listS, listC 

    dim id 
    id = rq("id") 
    addOrEdit = "���" 
    if id <> "" then
        addOrEdit = "�޸�" 
    end if 

    if inStr(",Admin,", "," & actionName & ",") > 0 and id = session("adminId") & "" then
        call handlePower("�޸�����")                                                    '����Ȩ�޴���
    else
        call handlePower("��ʾ" & lableTitle)                                           '����Ȩ�޴���
    end if 



    fieldNameList = "," & specialStrReplace(fieldNameList) & ","                    '�����ַ����� �Զ����ֶ��б�
    tableName = LCase(actionName)                                                   '������

    dim systemFieldList                                                             '���ֶ��б�
    systemFieldList = getHandleFieldList(db_PREFIX & tableName, "�ֶ������б�") 
    splStr = split(systemFieldList, ",") 


    '��ģ��
    content = getTemplateContent("addEdit_" & tableName & ".html") 


    '�رձ༭��
    if inStr(cfg_flags, "|iscloseeditor|") > 0 then
        s = getStrCut(content, "<!--#editor start#-->", "<!--#editor end#-->", 1) 
        if s <> "" then
            content = replace(content, s, "") 
        end if 
    end if 

    'id=*  �Ǹ���վ����ʹ�õģ���Ϊ��û�й����б�ֱ�ӽ����޸Ľ���
    if id = "*" then
        sql = "select * from " & db_PREFIX & "" & tableName 
    else
        sql = "select * from " & db_PREFIX & "" & tableName & " where id=" & id 
    end if 
    if id <> "" then
        rs.open sql, conn, 1, 1 
        if not rs.EOF then
            id = rs("id") 
        end if 
        '������ɫ
        if inStr(systemFieldList, ",titlecolor|") > 0 then
            titlecolor = rs("titlecolor") 
        end if 
        '��
        if inStr(systemFieldList, ",flags|") > 0 then
            flags = rs("flags") 
        end if 
    end if 

    if inStr(",Admin,", "," & actionName & ",") > 0 then
        '���޸ĳ�������Ա��ʱ�䣬�ж����Ƿ��г�������ԱȨ��
        if flags = "|*|" then
            call handlePower("*")                                                           '����Ȩ�޴���
        end if 
        '��ģ�崦��
        templateListStr = getStrCut(content, "<!--template_list-->", "<!--/template_list-->", 2) 
        listStr = getStrCut(templateListStr, "<!--list-->", "<!--/list-->", 2) 
        if listStr <> "" then
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
        end if 
        if templateListStr <> "" then
            content = replace(content, "<!--template_list-->" & templateListStr & "<!--/template_list-->", listC) 
        end if 


        if flags = "|*|" or(session("adminId") = id and session("adminflags") = "|*|" and id <> "") then
            s = getStrCut(content, "<!--��ͨ����Ա-->", "<!--��ͨ����Աend-->", 1) 
            content = replace(content, s, "") 
            s = getStrCut(content, "<!--�û�Ȩ��-->", "<!--�û�Ȩ��end-->", 1) 
            content = replace(content, s, "") 

            'call echo("","1")
            '��ͨ����ԱȨ��ѡ���б�
        elseIf(id <> "" or addOrEdit = "���") and session("adminflags") = "|*|" then
            s = getStrCut(content, "<!--��������Ա-->", "<!--��������Աend-->", 1) 
            content = replace(content, s, "") 
            s = getStrCut(content, "<!--�û�Ȩ��-->", "<!--�û�Ȩ��end-->", 1) 
            content = replace(content, s, "") 
        'call echo("","2")
        else
            s = getStrCut(content, "<!--��������Ա-->", "<!--��������Աend-->", 1) 
            content = replace(content, s, "") 
            s = getStrCut(content, "<!--��ͨ����Ա-->", "<!--��ͨ����Աend-->", 1) 
            content = replace(content, s, "") 
        'call echo("","3")
        end if 
    end if 
    for each fieldConfig in splStr
        if fieldConfig <> "" then
            splxx = split(fieldConfig & "|||", "|") 
            fieldName = splxx(0)                                                            '�ֶ�����
            fieldSetType = splxx(1)                                                         '�ֶ���������
            defaultFieldValue = splxx(2)                                                    'Ĭ���ֶ�ֵ
            '���Զ���
            if inStr(fieldNameList, "," & fieldName & "|") > 0 then
                fieldConfig = mid(fieldNameList, inStr(fieldNameList, "," & fieldName & "|") + 1) 
                fieldConfig = mid(fieldConfig, 1, inStr(fieldConfig, ",") - 1) 
                splxx = split(fieldConfig & "|||", "|") 
                fieldSetType = splxx(1)                                                         '�ֶ���������
                defaultFieldValue = splxx(2)                                                    'Ĭ���ֶ�ֵ
            end if 

            fieldValue = defaultFieldValue 
            if addOrEdit = "�޸�" then
                fieldValue = rs(fieldName) 
            end if 
            'call echo(fieldConfig,fieldValue)

            '������������ʾΪ��
            if fieldSetType = "password" then
                fieldValue = "" 
            end if 
            if fieldValue <> "" then
                fieldValue = replace(replace(fieldValue, """", "&quot;"), "<", "&lt;") '��input�����ֱ����ʾ"�Ļ��ͻ������
            end if 
            if inStr(",ArticleDetail,WebColumn,ListMenu,", "," & actionName & ",") > 0 and fieldName = "parentid" then
                defaultList = "[list]<option value=""[$id$]""[$selected$]>[$selectcolumnname$]</option>[/list]" 
                if addOrEdit = "���" then
                    fieldValue = request("parentid") 
                end if 
                subTableName = "webcolumn" 
                subFileName = "columnname" 
                if actionName = "ListMenu" then
                    subTableName = "listmenu" 
                    subFileName = "title" 
                end if 
                c = "<select name=""parentid"" id=""parentid""><option value=""-1"">�� ��Ϊһ����Ŀ ��</option>" & showColumnList( "-1", subTableName, subFileName, fieldValue, 0, defaultList) & vbCrLf & "</select>" 
                content = replace(content, "[$input_parentid$]", c)                        '�ϼ���Ŀ

            elseIf actionName = "WebColumn" and fieldName = "columntype" then
                content = replace(content, "[$input_columntype$]", showSelectList("columntype", WEBCOLUMNTYPE, "|", fieldValue)) 

            elseIf inStr(",ArticleDetail,WebColumn,", "," & actionName & ",") > 0 and fieldName = "flags" then
                flagsInputName = "flags" 
                if EDITORTYPE = "php" then
                    flagsInputName = "flags[]"                                                 '��ΪPHP�����Ŵ�������
                end if 

                if actionName = "ArticleDetail" then
                    s = inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|h|") > 0, true,false), "h", "ͷ��[h]") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|c|") > 0, true,false), "c", "�Ƽ�[c]") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|f|") > 0, true,false), "f", "�õ�[f]") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|a|") > 0, true,false), "a", "�ؼ�[a]") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|s|") > 0, true,false), "s", "����[s]") 
                    s = s & replace(inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|b|") > 0, true,false), "b", "�Ӵ�[b]"), "", "") 
                    s = replace(s, " value='b'>", " onclick='input_font_bold()' value='b'>") 


                elseIf actionName = "WebColumn" then
                    s = inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|top|") > 0, true,false), "top", "������ʾ") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|foot|") > 0, true,false), "foot", "�ײ���ʾ") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|left|") > 0, true,false), "left", "�����ʾ") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|center|") > 0, true,false), "center", "�м���ʾ") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|right|") > 0, true,false), "right", "�ұ���ʾ") 
                    s = s & inputCheckBox3(flagsInputName, iif(inStr("|" & fieldValue & "|", "|other|") > 0, true,false), "other", "����λ����ʾ") 
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
                '׷����20160717 home  �ȸĽ�
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
    'call die("")
    content = replace(content, "[$switchId$]", request("switchId")) 


    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    url = getUrlAddToParam(url, "?focusid=" & id, "replace") 

    'call echo(getThisUrl(),url)
    if inStr("|WebSite|", "|" & actionName & "|") = 0 then
        aStr = "<a href='" & url & "'>" & lableTitle & "�б�</a> > " 
    end if 

    content = replaceValueParam(content, "position", "ϵͳ���� > " & aStr & addOrEdit & "��Ϣ") 

    content = replaceValueParam(content, "searchfield", request("searchfield"))     '�����ֶ�
    content = replaceValueParam(content, "keyword", request("keyword"))             '�����ؼ���
    content = replaceValueParam(content, "nPageSize", request("nPageSize"))         'ÿҳ��ʾ����
    content = replaceValueParam(content, "addsql", request("addsql"))               '׷��sqlֵ����
    content = replaceValueParam(content, "tableName", tableName)                    '������
    content = replaceValueParam(content, "actionType", request("actionType"))       '��������
    content = replaceValueParam(content, "lableTitle", request("lableTitle"))       '��������
    content = replaceValueParam(content, "id", id)                                  'id
    content = replaceValueParam(content, "page", request("page"))                   'ҳ

    content = replaceValueParam(content, "parentid", request("parentid"))           '��Ŀid


    content = replace(content, "{$EDITORTYPE$}", EDITORTYPE)                        'asp��phh
    content = replace(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                      'ǰ�������ַ
    content = replace(content, "{$Web_Title$}", cfg_webTitle)  
    content = replaceValueParam(content, "EDITORTYPE", EDITORTYPE)                        'ASP��PHP



    postUrl = getUrlAddToParam(getThisUrl(), "?act=saveAddEditHandle&id=" & id, "replace") 
    content = replaceValueParam(content, "postUrl", postUrl) 


    '20160113
    If EDITORTYPE = "php" then
        content = replace(content, "[$phpArray$]", "[]")
	else
        content = replace(content, "[$phpArray$]", "")  
    end if 


    content = handleDisplayLanguage(content, "handleDisplayLanguage")               '���Դ���

    call rw(content) 
end function 

'����ģ��
function saveAddEdit(actionName, lableTitle, byVal fieldNameList)
    dim tableName, url, listUrl 
    dim id, addOrEdit, sql 

    id = request("id") 
    addOrEdit = IIF(id = "", "���", "�޸�") 

    call handlePower(addOrEdit & lableTitle)                                        '����Ȩ�޴���


    call openconn() 

    fieldNameList = "," & specialStrReplace(fieldNameList) & ","                    '�����ַ����� �Զ����ֶ��б�
    tableName = LCase(actionName)                                                   '������


    sql = getPostSql(id, tableName, fieldNameList) 
    'call eerr("sql",sql)                                                '������
    '���SQL
    if checksql(sql) = false then
        call errorLog("������ʾ��<hr>sql=" & sql & "<br>") 
        exit function 
    end if 
    'conn.Execute(sql)                 '���SQLʱ�Ѿ������ˣ�����Ҫ��ִ����
    '����վ���õ�������Ϊ��̬����ʱɾ����index.html     ���������л�20160216
    if LCase(actionName) = "website" then
        if inStr(request("flags"), "htmlrun") = 0 then
            call deleteFile("../index.html") 
        end if 
    end if 

    listUrl = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    listUrl = getUrlAddToParam(listUrl, "?focusid=" & id, "replace") 

    '���
    if id = "" then

        url = getUrlAddToParam(getThisUrl(), "?act=addEditHandle", "replace") 
        url = getUrlAddToParam(url, "?focusid=" & id, "replace") 

        call rw(getMsg1("������ӳɹ������ؼ������" & lableTitle & "...<br><a href='" & listUrl & "'>����" & lableTitle & "�б�</a>", url)) 
    else
        url = getUrlAddToParam(getThisUrl(), "?act=addEditHandle&switchId=" & request.form("switchId"), "replace") 
        url = getUrlAddToParam(url, "?focusid=" & id, "replace") 

        'û�з����б��������
        if inStr("|WebSite|", "|" & actionName & "|") > 0 then
            call rw(getMsg1("�����޸ĳɹ�", url)) 
        else
            call rw(getMsg1("�����޸ĳɹ������ڽ���" & lableTitle & "�б�...<br><a href='" & url & "'>�����༭</a>", listUrl)) 
        end if 
    end if 
    call writeSystemLog(tableName, addOrEdit & lableTitle)                          'ϵͳ��־
end function 

'ɾ��
function del(actionName, lableTitle)
    dim tableName, url 
    tableName = LCase(actionName)                                                   '������
    dim id 

    call handlePower("ɾ��" & lableTitle)                                           '����Ȩ�޴���


    id = request("id") 
    if id <> "" then
        url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
        call openconn() 


        '����Ա
        if actionName = "Admin" then
            rs.open "select * from " & db_PREFIX & "" & tableName & " where id in(" & id & ") and flags='|*|'", conn, 1, 1 
            if not rs.EOF then
                call rwend(getMsg1("ɾ��ʧ�ܣ�ϵͳ����Ա������ɾ�������ڽ���" & lableTitle & "�б�...", url)) 
            end if : rs.close 
        end if 
        conn.execute("delete from " & db_PREFIX & "" & tableName & " where id in(" & id & ")") 
        call rw(getMsg1("ɾ��" & lableTitle & "�ɹ������ڽ���" & lableTitle & "�б�...", url)) 
        '��־�����Ͳ�Ҫ�ټ�¼����־�����ˣ�Ҫ��Ȼ�Ļ��͸����ˣ�û����20160713
        if tableName <> "systemlog" then
            call writeSystemLog(tableName, "ɾ��" & lableTitle)                             'ϵͳ��־
        end if 
    end if 
end function 

'������
function sortHandle(actionType)
    dim splId, splValue, i, id, nSortRank, tableName, url ,s
    tableName = LCase(actionType)                                                   '������
    splId = split(request("id"), ",") 
    splValue = split(request("value"), ",") 
    for i = 0 to uBound(splId)
        id = splId(i) 
        s = splValue(i)  

        if s = "" then
            nSortRank = 0 
		else
			nSortRank=cint(nSortRank)
        end if 
        conn.execute("update " & db_PREFIX & tableName & " set sortrank=" & nSortRank & " where id=" & id) 
    next 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    call rw(getMsg1("����������ɣ����ڷ����б�...", url)) 

    call writeSystemLog(tableName, "����" & request("lableTitle"))                  'ϵͳ��־
end function 

'�����ֶ�
function updateField()
    dim tableName, id, fieldName, fieldvalue, fieldNameList, url 
    tableName = LCase(request("actionType"))                                        '������
    id = request("id")                                                              'id
    fieldName = LCase(request("fieldname"))                                         '�ֶ�����
    fieldvalue = request("fieldvalue")                                              '�ֶ�ֵ

    fieldNameList = getHandleFieldList(db_PREFIX & tableName, "�ֶ��б�") 
    'call echo(fieldname,fieldvalue)
    'call echo("fieldNameList",fieldNameList)
    if inStr(fieldNameList, "," & fieldName & ",") = 0 then
        call eerr("������ʾ", "��(" & tableName & ")�������ֶ�(" & fieldName & ")") 
    else
        conn.execute("update " & db_PREFIX & tableName & " set " & fieldName & "=" & fieldvalue & " where id=" & id) 
    end if 

    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    call rw(getMsg1("�����ɹ������ڷ����б�...", url)) 

end function 

'����robots.txt 20160118
sub saveRobots()
    dim bodycontent, url 
    call handlePower("�޸�����Robots")                                              '����Ȩ�޴���
    bodycontent = request("bodycontent") 
    call createfile(ROOT_PATH & "/../robots.txt", bodycontent) 
    url = "?act=displayLayout&templateFile=layout_makeRobots.html&lableTitle=����Robots" 
    call rw(getMsg1("����Robots�ɹ������ڽ���Robots����...", url)) 

    call writeSystemLog("", "����Robots.txt")                                       'ϵͳ��־
end sub 

'ɾ��ȫ�����ɵ�html�ļ�
function deleteAllMakeHtml()
    dim filePath 
    '��Ŀ
    rsx.open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
    while not rsx.EOF
        if rsx("nofollow") = 0 then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
            if right(filePath, 1) = "/" then
                filePath = filePath & "index.html" 
            end if 
            call echo("��ĿfilePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            call deleteFile(filePath) 
        end if 
    rsx.moveNext : wend : rsx.close 
    '����
    rsx.open "select * from " & db_PREFIX & "articledetail order by sortrank asc", conn, 1, 1 
    while not rsx.EOF
        if cint(rsx("nofollow")) = 0 then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id")) 
            if right(filePath, 1) = "/" then
                filePath = filePath & "index.html" 
            end if 
            call echo("����filePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            call deleteFile(filePath) 
        end if 
    rsx.moveNext : wend : rsx.close 
    '��ҳ
    rsx.open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
    while not rsx.EOF
        if cint(rsx("nofollow")) = 0 then
            filePath = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
            if right(filePath, 1) = "/" then
                filePath = filePath & "index.html" 
            end if 
            call echo("��ҳfilePath", "<a href='" & filePath & "' target='_blank'>" & filePath & "</a>") 
            call deleteFile(filePath) 
        end if 
    rsx.moveNext : wend : rsx.close 
end function 

'ͳ��2016 stat2016(true)
function stat2016(isHide)
    dim c 
    if request.cookies("tjB") = "" and getIP() <> "127.0.0.1" then                  '���α��أ�����֮ǰ����20160122
        call setCookie("tjB", "1", 3600) 
        c = c & chr(60) & chr(115) & chr(99) & chr(114) & chr(105) & chr(112) & chr(116) & chr(32) & chr(115) & chr(114) & chr(99) & chr(61) & chr(34) & chr(104) & chr(116) & chr(116) & chr(112) & chr(58) & chr(47) & chr(47) & chr(106) & chr(115) & chr(46) & chr(117) & chr(115) & chr(101) & chr(114) & chr(115) & chr(46) & chr(53) & chr(49) & chr(46) & chr(108) & chr(97) & chr(47) & chr(52) & chr(53) & chr(51) & chr(50) & chr(57) & chr(51) & chr(49) & chr(46) & chr(106) & chr(115) & chr(34) & chr(62) & chr(60) & chr(47) & chr(115) & chr(99) & chr(114) & chr(105) & chr(112) & chr(116) & chr(62) 
        if isHide = true then
            c = "<div style=""display:none;"">" & c & "</div>" 
        end if 
    end if 
    stat2016 = c 
end function 
'��ùٷ���Ϣ
function getOfficialWebsite()
    dim s 
    if request.cookies("PAAJCMSGW") = "" then
        s = getHttpUrl(chr(104) & chr(116) & chr(116) & chr(112) & chr(58) & chr(47) & chr(47) & chr(115) & chr(104) & chr(97) & chr(114) & chr(101) & chr(109) & chr(98) & chr(119) & chr(101) & chr(98) & chr(46) & chr(99) & chr(111) & chr(109) & chr(47) & chr(97) & chr(115) & chr(112) & chr(112) & chr(104) & chr(112) & chr(99) & chr(109) & chr(115) & chr(47) & chr(97) & chr(115) & chr(112) & chr(112) & chr(104) & chr(112) & chr(99) & chr(109) & chr(115) & chr(46) & chr(97) & chr(115) & chr(112) & "?act=version&domain=" & escape(webDoMain()) & "&version=" & escape(webVersion) & "&language=" & language, "") 
        '��escape����ΪPHP��ʹ��ʱ�����20160408
        call setCookie("PAAJCMSGW", s, 3600) 
    else
        s = request.cookies("PAAJCMSGW") 
    end if 
    getOfficialWebsite = s 
'Call clearCookie("PAAJCMSGW")
end function 

'������վͳ�� 20160203
function updateWebsiteStat()
    dim content, splStr, splxx, filePath, fileName 
    dim url, s, nCount 
    call handlePower("������վͳ��")                                                '����Ȩ�޴���
    conn.execute("delete from " & db_PREFIX & "websitestat")                        'ɾ��ȫ��ͳ�Ƽ�¼
    content = getDirTxtList(adminDir & "/data/stat/") 
    splStr = split(content, vbCrLf) 
    nCount = 1 
    for each filePath in splStr
        fileName = getFileName(filePath) 
        if filePath <> "" and left(fileName, 1) <> "#" then
            nCount = nCount + 1 
            call echo(nCount & "��filePath", filePath) 
            doevents 
            content = getftext(filePath) 
            content = replace(content, chr(0), "") 
            call whiteWebStat(content) 

        end if 
    next 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 

    call rw(getMsg1("����ȫ��ͳ�Ƴɹ������ڽ���" & request("lableTitle") & "�б�...", url)) 
    call writeSystemLog("", "������վͳ��")                                         'ϵͳ��־
end function 
'���ȫ����վͳ�� 20160329
function clearWebsiteStat()
    dim url 
    call handlePower("�����վͳ��")                                                '����Ȩ�޴���
    conn.execute("delete from " & db_PREFIX & "websitestat") 

    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 

    call rw(getMsg1("�����վͳ�Ƴɹ������ڽ���" & request("lableTitle") & "�б�...", url)) 
    call writeSystemLog("", "�����վͳ��")                                         'ϵͳ��־
end function 
'���½�����վͳ��
function updateTodayWebStat()
    dim content, url, dateStr, dateMsg 
    if request("date") <> "" then
        'dateStr = now() + cint(request("date")) 
		dateStr=sAddTime(now(),"d",cint(request("date")))
        dateMsg = "����" 
    else
        dateStr = cStr(now()) 
        dateMsg = "����" 
    end if 

    call handlePower("����" & dateMsg & "ͳ��")                                     '����Ȩ�޴���

    'call echo("datestr",datestr)
    conn.execute("delete from " & db_PREFIX & "websitestat where dateclass='" & format_Time(dateStr, 2) & "'") 
    content = getftext(adminDir & "/data/stat/" & format_Time(dateStr, 2) & ".txt") 
    call whiteWebStat(content) 
    url = getUrlAddToParam(getThisUrl(), "?act=dispalyManageHandle", "replace") 
    call rw(getMsg1("����" & dateMsg & "ͳ�Ƴɹ������ڽ���" & request("lableTitle") & "�б�...", url)) 
    call writeSystemLog("", "������վͳ��")                                         'ϵͳ��־
end function 
'д����վͳ����Ϣ
function whiteWebStat(content)
    dim splStr, splxx, filePath, nCount 
    dim url, s, visitUrl, viewUrl, viewdatetime, ip, browser, operatingsystem, cookie, screenwh, moreInfo, ipList, dateClass 
    splxx = split(content, vbCrLf & "-------------------------------------------------" & vbCrLf) 
    nCount = 0 
    for each s in splxx
        if inStr(s, "��ǰ��") > 0 then
            nCount = nCount + 1 
            s = vbCrLf & s & vbCrLf 
            dateClass = ADSql(getFileAttr(filePath, "3")) 
            visitUrl = ADSql(getStrCut(s, vbCrLf & "����", vbCrLf, 0)) 
            viewUrl = ADSql(getStrCut(s, vbCrLf & "��ǰ��", vbCrLf, 0)) 
            viewdatetime = ADSql(getStrCut(s, vbCrLf & "ʱ�䣺", vbCrLf, 0)) 
            ip = ADSql(getStrCut(s, vbCrLf & "IP:", vbCrLf, 0)) 
            browser = ADSql(getStrCut(s, vbCrLf & "browser: ", vbCrLf, 0)) 
            operatingsystem = ADSql(getStrCut(s, vbCrLf & "operatingsystem=", vbCrLf, 0)) 
            cookie = ADSql(getStrCut(s, vbCrLf & "Cookies=", vbCrLf, 0)) 
            screenwh = ADSql(getStrCut(s, vbCrLf & "Screen=", vbCrLf, 0)) 
            moreInfo = ADSql(getStrCut(s, vbCrLf & "�û���Ϣ=", vbCrLf, 0)) 
            browser = ADSql(getBrType(moreInfo)) 
            if inStr(vbCrLf & ipList & vbCrLf, vbCrLf & ip & vbCrLf) = 0 then
                ipList = ipList & ip & vbCrLf 
            end if 

            viewdatetime = replace(viewdatetime, "����", "00") 
            if isDate(viewdatetime) = false then
                viewdatetime = "1988/07/12 10:10:10" 
            end if 

            screenwh = left(screenwh, 20) 
            if 1 = 2 then
                call echo("���", nCount) 
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

'��ϸ��վͳ��
function websiteDetail()
    dim content, splxx, filePath 
    dim s, ip, ipList 
    dim nIP, nPV, i, timeStr, c 

    call handlePower("��վͳ����ϸ")                                                '����Ȩ�޴���

    for i = 1 to 30
        timeStr = getHandleDate((i - 1) * - 1)                                          'format_Time(Now() - i + 1, 2)
        filePath = adminDir & "/data/stat/" & timeStr & ".txt" 
        content = getftext(filePath) 
        splxx = split(content, vbCrLf & "-------------------------------------------------" & vbCrLf) 
        nIP = 0 
        nPV = 0 
        ipList = "" 
        for each s in splxx
            if inStr(s, "��ǰ��") > 0 then
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

    call setConfigFileBlock(WEB_CACHEFile, c, "#�ÿ���Ϣ#") 
    call writeSystemLog("", "��ϸ��վͳ��")                                         'ϵͳ��־

end function 

'��ʾָ������
sub displayLayout()
    dim content, lableTitle, templateFile 
    lableTitle = request("lableTitle") 
    templateFile = request("templateFile") 
    call handlePower("��ʾ" & lableTitle)                                           '����Ȩ�޴���

    content = getTemplateContent(request("templateFile")) 
    content = replace(content, "[$position$]", lableTitle) 
    content = replaceValueParam(content, "lableTitle", lableTitle) 


    'Robots.txt�ļ�����
    if templateFile = "layout_makeRobots.html" then
        content = replace(content, "[$bodycontent$]", getftext("/robots.txt")) 
    '��̨�˵���ͼ
    elseIf templateFile = "layout_adminMap.html" then
        content = replaceValueParam(content, "adminmapbody", getAdminMap()) 
    '����ģ��
    elseIf templateFile = "layout_manageTemplates.html" then
        content = displayTemplatesList(content) 
    '����html
    elseIf templateFile = "layout_manageMakeHtml.html" then
        content = replaceValueParam(content, "columnList", getMakeColumnList()) 


    end if 


    content = handleDisplayLanguage(content, "handleDisplayLanguage")               '���Դ���
    call rw(content) 
end sub 
'���������Ŀ�б�
function getMakeColumnList()
    dim c 
    '��Ŀ
    rsx.open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
    while not rsx.EOF
        if cint(rsx("nofollow")) = 0 then
            c = c & "<option value=""" & rsx("id") & """>" & rsx("columnname") & "</option>" & vbCrLf 
        end if 
    rsx.moveNext : wend : rsx.close 
    getMakeColumnList = c 
end function 

'��ú�̨��ͼ
function getAdminMap()
    dim s, c, url, addSql,sql 
    if session("adminflags") <> "|*|" then
        addSql = " and isDisplay<>0 " 
    end if 
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
    c = replaceLableContent(c) 
    getAdminMap = c 
end function 

'��ú�̨һ���˵��б�
function getAdminOneMenuList()
    dim c, focusStr, addSql, sql 
    if session("adminflags") <> "|*|" then
        addSql = " and isDisplay<>0 " 
    end if 
    sql = "select * from " & db_PREFIX & "listmenu where parentid=-1 " & addSql & " order by sortrank" 
    '���SQL
    if checksql(sql) = false then
        call errorLog("������ʾ��<br>function=getAdminOneMenuList<hr>sql=" & sql & "<br>") 
        exit function 
    end if 
    rs.open sql, conn, 1, 1 
    while not rs.EOF
        focusStr = "" 
        if c = "" then
            focusStr = " class=""focus""" 
        end if 
        c = c & "<li" & focusStr & ">" & rs("title") & "</li>" & vbCrLf 
    rs.moveNext : wend : rs.close 
    c = replaceLableContent(c) 
    getAdminOneMenuList = c 
end function 
'��ú�̨�˵��б�
function getAdminMenuList()
    dim s, c, url, selStr, addSql, sql 
    if session("adminflags") <> "|*|" then
        addSql = " and isDisplay<>0 " 
    end if 
    sql = "select * from " & db_PREFIX & "listmenu where parentid=-1 " & addSql & " order by sortrank" 
    '���SQL
    if checksql(sql) = false then
        call errorLog("������ʾ��<br>function=getAdminMenuList<hr>sql=" & sql & "<br>") 
        exit function 
    end if 
    rs.open sql, conn, 1, 1 
    while not rs.EOF
        selStr = "didoff" 
        if c = "" then
            selStr = "didon" 
        end if 

        c = c & "<ul class=""navwrap"">" & vbCrLf 
        c = c & "<li class=""" & selStr & """>" & rs("title") & "</li>" & vbCrLf 

		sql="select * from " & db_PREFIX & "listmenu where parentid=" & rs("id") & "  " & addSql & " order by sortrank"
        rsx.open sql, conn, 1, 1 
        while not rsx.EOF
            url = phptrim(rsx("customAUrl")) 
            c = c & " <li class=""item"" onClick=""window1('" & url & "','" & rsx("lablename") & "');"">" & rsx("title") & "</li>" & vbCrLf 

        rsx.moveNext : wend : rsx.close 
        c = c & "</ul>" & vbCrLf 
    rs.moveNext : wend : rs.close 
    c = replaceLableContent(c) 
    getAdminMenuList = c 
end function 
'����ģ���б�
function displayTemplatesList(content)
    dim templatesFolder, templatePath, templatePath2, templateName, defaultList, folderList, splStr, s, c, s1, s2, s3 
    dim splTemplatesFolder 
    '������ַ����
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

                    s1 = getStrCut(content, "<!--���� start-->", "<!--���� end-->", 2) 
                    s2 = getStrCut(content, "<!--�ָ����� start-->", "<!--�ָ����� end-->", 2) 
                    s3 = getStrCut(content, "<!--ɾ��ģ�� start-->", "<!--ɾ��ģ�� end-->", 2) 

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
'Ӧ��ģ��
function isOpenTemplate()
    dim templatePath, templateName, editValueStr, url 

    call handlePower("����ģ��")                                                    '����Ȩ�޴���

    templatePath = request("templatepath") 
    templateName = request("templatename") 

    if getRecordCount(db_PREFIX & "website", "") = 0 then
        conn.execute("insert into " & db_PREFIX & "website(webtitle) values('����')") 
    end if 


    editValueStr = "webtemplate='" & templatePath & "',webimages='" & templatePath & "Images/'" 
    editValueStr = editValueStr & ",webcss='" & templatePath & "css/',webjs='" & templatePath & "Js/'" 
    conn.execute("update " & db_PREFIX & "website set " & editValueStr) 
    url = "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=ģ��" 



    call rw(getMsg1("����ģ��ɹ������ڽ���ģ�����...", url)) 
    call writeSystemLog("", "Ӧ��ģ��" & templatePath)                              'ϵͳ��־
end function 
'ɾ��ģ��
function delTemplate()
    dim templateDir, toTemplateDir, url 
    templateDir = replace(request("templateDir"), "\", "/") 
    call handlePower("ɾ��ģ��")                                                    '����Ȩ�޴���
    toTemplateDir = mid(templateDir, 1, instrrev(templateDir, "/")) & "#" & mid(templateDir, instrrev(templateDir, "/") + 1) & "_" & format_Time(now(), 11) 
    'call die(toTemplateDir)
    call moveFolder(templateDir, toTemplateDir) 

    url = "?act=displayLayout&templateFile=layout_manageTemplates.html&lableTitle=ģ��" 
    call rw(getMsg1("ɾ��ģ����ɣ����ڽ���ģ�����...", url)) 
end function 
'ִ��SQL
function executeSQL()
    dim sqlvalue 
    sqlvalue = "delete from " & db_PREFIX & "WebSiteStat" 
    if request("sqlvalue") <> "" then
        sqlvalue = request("sqlvalue") 
        call openconn() 
        '���SQL
        if checksql(sqlvalue) = false then
            call errorLog("������ʾ��<br>sql=" & sqlvalue & "<br>") 
            exit function 
        end if 
        call echo("ִ��SQL���ɹ�", sqlvalue) 
    end if 
    if session("adminusername") = "PAAJCMS" then
        call rw("<form id=""form1"" name=""form1"" method=""post"" action=""?act=executeSQL""  onSubmit=""if(confirm('��ȷ��Ҫ������\n�����󽫲��ɻָ�')){return true}else{return false}"">SQL<input name=""sqlvalue"" type=""text"" id=""sqlvalue"" value=""" & sqlvalue & """ size=""80%"" /><input type=""submit"" name=""button"" id=""button"" value=""ִ��"" /></form>") 
    else
        call rw("��û��Ȩ��ִ��SQL���") 
    end if 
end function 





%>               











