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
<!--#Include File = "Inc/Config.Asp"-->
<!--#Include File = "inc/admin_function.asp"--> 
<% 
'asp������

call openconn() 
'=========


'������   ReplaceValueParamΪ�����ַ���ʾ��ʽ
function handleAction(content)
    dim startStr, endStr, actionList, splStr, action, s, isHnadle 
    startStr = "{$" : endStr = "$}" 
    actionList = getArrayNew(content, startStr, endStr, true, true) 
    'Call echo("ActionList ", ActionList)
    splStr = split(actionList, "$Array$") 
    for each s in splStr
        action = trim(s) 
        action = handleInModule(action, "start")                                        '����\'�滻��
        if action <> "" then
            action = trim(mid(action, 3, len(action) - 4)) & " " 
            'call echo("s",s)
            isHnadle = true                                                                 '����Ϊ��
            '{VB #} �����Ƿ���ͼƬ·���Ŀ����Ϊ����VB�ﲻ�������·��
            if checkFunValue(action, "# ") = true then
                action = "" 
            '����
            elseIf checkFunValue(action, "GetLableValue ") = true then
                action = XY_getLableValue(action) 
            '�����������������б�
            elseIf checkFunValue(action, "TitleInSearchEngineList ") = true then
                action = XY_TitleInSearchEngineList(action) 

            '�����ļ�
            elseIf checkFunValue(action, "Include ") = true then
                action = XY_Include(action) 
            '��Ŀ�б�
            elseIf checkFunValue(action, "ColumnList ") = true then
                action = XY_AP_ColumnList(action) 
            '�����б�
            elseIf checkFunValue(action, "ArticleList ") = true or checkFunValue(action, "CustomInfoList ") = true then
                action = XY_AP_ArticleList(action) 
            '�����б�
            elseIf checkFunValue(action, "CommentList ") = true then
                action = XY_AP_CommentList(action) 
            '����ͳ���б�
            elseIf checkFunValue(action, "SearchStatList ") = true then
                action = XY_AP_SearchStatList(action) 
            '���������б�
            elseIf checkFunValue(action, "Links ") = true then
                action = XY_AP_Links(action) 

            '��ʾ��ҳ����
            elseIf checkFunValue(action, "GetOnePageBody ") = true or checkFunValue(action, "MainInfo ") = true then
                action = XY_AP_GetOnePageBody(action) 
            '��ʾ��������
            elseIf checkFunValue(action, "GetArticleBody ") = true then
                action = XY_AP_GetArticleBody(action) 
            '��ʾ��Ŀ����
            elseIf checkFunValue(action, "GetColumnBody ") = true then
                action = XY_AP_GetColumnBody(action) 

            '�����ĿURL
            elseIf checkFunValue(action, "GetColumnUrl ") = true then
                action = XY_GetColumnUrl(action) 
            '�����ĿID
            elseIf checkFunValue(action, "GetColumnId ") = true then
                action = XY_GetColumnId(action) 
            '�������URL
            elseIf checkFunValue(action, "GetArticleUrl ") = true then
                action = XY_GetArticleUrl(action) 
            '��õ�ҳURL
            elseIf checkFunValue(action, "GetOnePageUrl ") = true then
                action = XY_GetOnePageUrl(action) 
            '��õ���URL
            elseIf checkFunValue(action, "GetNavUrl ") = true then
                action = XY_GetNavUrl(action) 



                '------------------- ģ��ģ���� -----------------------
            '��ʾ������ ���ò���
            elseIf checkFunValue(action, "DisplayWrap ") = true then
                action = XY_DisplayWrap(action) 
            '��ʾ����
            elseIf checkFunValue(action, "Layout ") = true then
                action = XY_Layout(action) 
            '��ʾģ��
            elseIf checkFunValue(action, "Module ") = true then
                action = XY_Module(action) 
            '��ģ������
            elseIf checkFunValue(action, "ReadTemplateModule ") = true then
                action = XY_ReadTemplateModule(action) 
            '�������ģ�� 20150108
            elseIf checkFunValue(action, "GetContentModule ") = true then
                action = XY_ReadTemplateModule(action) 
            '��ģ����ʽ�����ñ���������   ������и���ĿStyle��������
            elseIf checkFunValue(action, "ReadColumeSetTitle ") = true then
                action = XY_ReadColumeSetTitle(action) 


                '------------------- ������ -----------------------
            '��ʾJS��ȾASP/PHP/VB�ȳ���ı༭��
            elseIf checkFunValue(action, "displayEditor ") = true then
                action = displayEditor(action) 
            'Js����վͳ��
            elseIf checkFunValue(action, "JsWebStat ") = true then
                action = XY_JsWebStat(action) 

                '------------------- ������ -----------------------
            '��ͨ����A
            elseIf checkFunValue(action, "HrefA ") = true then
                action = XY_HrefA(action) 
            '��Ŀ�˵�(���ú�̨��Ŀ����)
            elseIf checkFunValue(action, "ColumnMenu ") = true then
                action = XY_AP_ColumnMenu(action) 

                '------------------- ѭ������ -----------------------
            'Forѭ������
            elseIf checkFunValue(action, "ForArray ") = true then
                action = XY_ForArray(action) 

                '------------------- ������ -----------------------
            '��վ�ײ�
            elseIf checkFunValue(action, "WebSiteBottom ") = true or checkFunValue(action, "WebBottom ") = true then
                action = XY_AP_WebSiteBottom(action) 
            '��ʾ��վ��Ŀ 20160331
            elseIf checkFunValue(action, "DisplayWebColumn ") = true then
                action = XY_DisplayWebColumn(action) 
            'URL����
            elseIf checkFunValue(action, "escape ") = true then
                action = XY_escape(action) 
            'URL����
            elseIf checkFunValue(action, "unescape ") = true then
                action = XY_unescape(action) 
            'asp��php�汾
            elseIf checkFunValue(action, "EDITORTYPE ") = true then
                action = XY_EDITORTYPE(action) 

            '�����ַ
            elseIf checkFunValue(action, "getUrl ") = true then
                action = XY_getUrl(action) 

            '����λ����ʾ��Ϣ{}Ϊ�ж�����
            elseIf checkFunValue(action, "detailPosition ") = true then
                action = XY_detailPosition(action) 




            '��ʱ������
            elseIf checkFunValue(action, "copyTemplateMaterial ") = true then
                action = "" 
            elseIf checkFunValue(action, "clearCache ") = true then
                action = "" 
            else
                isHnadle = false                                                                '����Ϊ��
            end if 
            'ע���������е�����ʾ �� And IsNul(action)=False
            if isNul(action) = true then action = "" 
            if isHnadle = true then
                content = replace(content, s, action) 
            end if 
        end if 
    next 
    handleAction = content 
end function 

'��ʾ��վ��Ŀ �°� ��֮ǰ��վ ��������Ľ�������
function XY_DisplayWebColumn(action)
    dim i, c, s, url, sql, dropDownMenu, focusType, addSql 
    dim isConcise                                                                   '�����ʾ20150212
    dim styleId, styleValue                                                         '��ʽID����ʽ����
    dim cssNameAddId 
    dim shopnavidwrap                                                               '�Ƿ���ʾ��ĿID��

    styleId = PHPTrim(rParam(action, "styleID")) 
    styleValue = PHPTrim(rParam(action, "styleValue")) 
    addSql = PHPTrim(rParam(action, "addSql")) 
    shopnavidwrap = PHPTrim(rParam(action, "shopnavidwrap")) 
    'If styleId <> "" Then
    'Call ReadNavCSS(styleId, styleValue)
    'End If

    'Ϊ�������� ���Զ���ȡ��ʽ����  20150615
    if checkStrIsNumberType(styleValue) then
        cssNameAddId = "_" & styleValue                                                 'Css����׷��Id���
    end if 
    sql = "select * from " & db_PREFIX & "webcolumn" 
    '׷��sql
    if addSql <> "" then
        sql = getWhereAnd(sql, addSql) 
    end if 
    if checkSql(sql) = false then call eerr("Sql", sql) 
    rs.open sql, conn, 1, 1 
    dropDownMenu = LCase(rParam(action, "DropDownMenu")) 
    focusType = LCase(rParam(action, "FocusType")) 
    isConcise = IIF(LCase(rParam(action, "isConcise")) = "true", false, true) 

    if isConcise = true then c = c & copyStr(" ", 4) & "<li class=left></li>" & vbCrLf 
    for i = 1 to rs.recordCount

        '��PHP��$rs=mysql_fetch_array($rsObj);                                            //��PHP�ã���Ϊ�� asptophpת��������
        url = getColumnUrl(rs("columnname"), "name") 
        if rs("columnName") = glb_columnName then
            if focusType = "a" then
                s = copyStr(" ", 8) & "<li class=focus><a href=""" & url & """>" & rs("columname") & "</a>" 
            else
                s = copyStr(" ", 8) & "<li class=focus>" & rs("columnname") 
            end if 
        else
            s = copyStr(" ", 8) & "<li><a href=""" & url & """>" & rs("columnname") & "</a>" 
        end if 
        '��վ��Ŀû��pageλ�ô���
        url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" & rs("id") & "&n=" & getRnd(11) 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span")                    '�����Ƿ���������޸Ĺ�����

        c = c & s 

        'С��

        c = c & copyStr(" ", 8) & "</li>" & vbCrLf 

        if isConcise = true then c = c & copyStr(" ", 8) & "<li class=line></li>" & vbCrLf 
    rs.moveNext : next : rs.close 
    if isConcise = true then c = c & copyStr(" ", 8) & "<li class=right></li>" & vbCrLf 

    if styleId <> "" then
        c = "<ul class='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</ul>" & vbCrLf 
    end if 
    if shopnavidwrap = "1" or shopnavidwrap = "true" then
        c = "<div id='nav" & styleId & cssNameAddId & "'>" & vbCrLf & c & vbCrLf & "</div>" & vbCrLf 
    end if 

    XY_DisplayWebColumn = c 
end function 

'�滻ȫ�ֱ��� {$cfg_websiteurl$}
function replaceGlobleVariable(byVal content)
    content = handleRGV(content, "{$cfg_webSiteUrl$}", cfg_webSiteUrl)              '��ַ
    content = handleRGV(content, "{$cfg_webTemplate$}", cfg_webTemplate)            'ģ��
    content = handleRGV(content, "{$cfg_webImages$}", cfg_webImages)                'ͼƬ·��
    content = handleRGV(content, "{$cfg_webCss$}", cfg_webCss)                      'css·��
    content = handleRGV(content, "{$cfg_webJs$}", cfg_webJs)                        'js·��
    content = handleRGV(content, "{$cfg_webTitle$}", cfg_webTitle)                  '��վ����
    content = handleRGV(content, "{$cfg_webKeywords$}", cfg_webKeywords)            '��վ�ؼ���
    content = handleRGV(content, "{$cfg_webDescription$}", cfg_webDescription)      '��վ����

    content = handleRGV(content, "{$cfg_webSiteBottom$}", cfg_webSiteBottom)        '��վ�ײ�����

    content = handleRGV(content, "{$glb_columnId$}", glb_columnId)                  '��ĿId
    content = handleRGV(content, "{$glb_columnName$}", glb_columnName)              '��Ŀ����
    content = handleRGV(content, "{$glb_columnType$}", glb_columnType)              '��Ŀ����
    content = handleRGV(content, "{$glb_columnENType$}", glb_columnENType)          '��ĿӢ������

    content = handleRGV(content, "{$glb_Table$}", glb_table)                        '��
    content = handleRGV(content, "{$glb_Id$}", glb_id)                              'id

    content = handleRGV(content, "[$ģ��Ŀ¼$]", "Module/")                    'Module


    '���ݾɰ汾 ��������ȥ��
    content = handleRGV(content, "{$WebImages$}", cfg_webImages)                    'ͼƬ·��
    content = handleRGV(content, "{$WebCss$}", cfg_webCss)                          'css·��
    content = handleRGV(content, "{$WebJs$}", cfg_webJs)                            'js·��
    content = handleRGV(content, "{$Web_Title$}", cfg_webTitle) 
    content = handleRGV(content, "{$Web_KeyWords$}", cfg_webKeywords) 
    content = handleRGV(content, "{$Web_Description$}", cfg_webDescription) 


    content = handleRGV(content, "{$EDITORTYPE$}", EDITORTYPE)                      '��׺
    content = handleRGV(content, "{$WEB_VIEWURL$}", WEB_VIEWURL)                    '��ҳ��ʾ��ַ
    '�����õ�
    content = handleRGV(content, "{$glb_articleAuthor$}", glb_articleAuthor)        '��������
    content = handleRGV(content, "{$glb_articleAdddatetime$}", glb_articleAdddatetime) '�������ʱ��
    content = handleRGV(content, "{$glb_articlehits$}", glb_articlehits)            '�������ʱ��

    content = handleRGV(content, "{$glb_upArticle$}", glb_upArticle)                '��һƪ����
    content = handleRGV(content, "{$glb_downArticle$}", glb_downArticle)            '��һƪ����
    content = handleRGV(content, "{$glb_aritcleRelatedTags$}", glb_aritcleRelatedTags) '���±�ǩ��
    content = handleRGV(content, "{$glb_aritcleBigImage$}", glb_aritcleBigImage)    '���´�ͼ
    content = handleRGV(content, "{$glb_aritcleSmallImage$}", glb_aritcleSmallImage) '����Сͼ
    content = handleRGV(content, "{$glb_searchKeyWord$}", glb_searchKeyWord)        '��ҳ��ʾ��ַ


    replaceGlobleVariable = content 
end function 

'�����滻
function handleRGV(byVal content, findStr, replaceStr)
    dim lableName 
    '��[$$]����
    lableName = mid(findStr, 3, len(findStr) - 4) & " " 
    lableName = mid(lableName, 1, inStr(lableName, " ") - 1) 
    content = replaceValueParam(content, lableName, replaceStr) 
    content = replaceValueParam(content, LCase(lableName), replaceStr) 
    'ֱ���滻{$$}���ַ�ʽ������֮ǰ��վ
    content = replace(content, findStr, replaceStr) 
    content = replace(content, LCase(findStr), replaceStr) 
    handleRGV = content 
end function 

'������վ������Ϣ
sub loadWebConfig()
    dim templatedir 
    call openconn() 
    rs.open "select * from " & db_PREFIX & "website", conn, 1, 1 
    if not rs.EOF then
        cfg_webSiteUrl = phptrim(rs("webSiteUrl"))                                      '��ַ
        cfg_webTemplate = webDir & phptrim(rs("webTemplate"))                           'ģ��·��
        cfg_webImages = webDir & phptrim(rs("webImages"))                               'ͼƬ·��
        cfg_webCss = webDir & phptrim(rs("webCss"))                                     'css·��
        cfg_webJs = webDir & phptrim(rs("webJs"))                                       'js·��
        cfg_webTitle = rs("webTitle")                                                   '��ַ����
        cfg_webKeywords = rs("webKeywords")                                             '��վ�ؼ���
        cfg_webDescription = rs("webDescription")                                       '��վ����
        cfg_webSiteBottom = rs("webSiteBottom")                                         '��վ�ص�
        cfg_flags = rs("flags")                                                         '��

        '�Ļ�ģ��20160202
        if request("templatedir") <> "" then
            'ɾ������Ŀ¼ǰ���Ŀ¼������Ҫ�Ǹ�����20160414
            templatedir = replace(handlePath(request("templatedir")), handlePath("/"), "/") 
            'call eerr("templatedir",templatedir)

            if(inStr(templatedir, ":") > 0 or inStr(templatedir, "..") > 0) and getIP() <> "127.0.0.1" then
                call eerr("��ʾ", "ģ��Ŀ¼�зǷ��ַ�") 
            end if 
            templatedir = handlehttpurl(replace(templatedir, handlePath("/"), "/")) 

            cfg_webImages = replace(cfg_webImages, cfg_webTemplate, templatedir) 
            cfg_webCss = replace(cfg_webCss, cfg_webTemplate, templatedir) 
            cfg_webJs = replace(cfg_webJs, cfg_webTemplate, templatedir) 
            cfg_webTemplate = templatedir 
        end if  
    end if : rs.close 
end sub 

'��վλ�� ������
function thisPosition(content)
    dim c 
    c = "<a href=""" & getColumnUrl("��ҳ", "type") & """>��ҳ</a>" 
    if glb_columnName <> "" then
        c = c & " >> <a href=""" & getColumnUrl(glb_columnName, "name") & """>" & glb_columnName & "</a>" 
    end if 
    '20160330
    if glb_locationType = "detail" then
        c = c & " >> �鿴����" 
    end if 
    'β��׷������
    c = c & positionEndStr 

    'call echo("glb_locationType",glb_locationType)

    content = replace(content, "[$detailPosition$]", c) 
    content = replace(content, "[$detailTitle$]", glb_detailTitle) 
    content = replace(content, "[$detailContent$]", glb_bodyContent) 

    thisPosition = content 
end function 

'��ʾ�����б�
function getDetailList(action, content, actionName, lableTitle, byVal fieldNameList, nPageSize, sPage, addSql)
    call openconn() 
    dim defaultStr, i, s, c, tableName, j, splxx, sql, nPage 
    dim nX, url, nCount 
    dim pageInfo, nModI, startStr, endStr 

    dim fieldName                                                                   '�ֶ�����
    dim splFieldName                                                                '�ָ��ֶ�

    dim replaceStr                                                                  '�滻�ַ�
    tableName = LCase(actionName)                                                   '������
    dim listFileName                                                                '�б��ļ�����
    listFileName = rParam(action, "listFileName") 
    dim abcolorStr                                                                  'A�Ӵֺ���ɫ
    dim atargetStr                                                                  'A���Ӵ򿪷�ʽ
    dim atitleStr                                                                   'A���ӵ�title20160407
    dim anofollowStr                                                                'A���ӵ�nofollow

    dim id, idPage 
    id = rq("id") 
    call checkIDSQL(request("id")) 

    if fieldNameList = "*" then
        fieldNameList = getHandleFieldList(db_PREFIX & tableName, "�ֶ��б�") 
    end if 

    fieldNameList = specialStrReplace(fieldNameList)                                '�����ַ�����
    splFieldName = split(fieldNameList, ",")                                        '�ֶηָ������


    defaultStr = getStrCut(content, "<!--#body start#-->", "<!--#body end#-->", 2) 



    pageInfo = getStrCut(content, "[page]", "[/page]", 1) 
    if pageInfo <> "" then
        content = replace(content, pageInfo, "") 
    end if 
    'call eerr("pageInfo",pageInfo)

    sql = "select * from " & db_PREFIX & tableName & " " & addSql 
    '���SQL
    if checksql(sql) = false then
        call errorLog("������ʾ��<br>sql=" & sql & "<br>") 
        exit function 
    end if 
    rs.open sql, conn, 1, 1 
    '��PHP��ɾ��rs
    nCount = rs.recordCount 

    'Ϊ��̬��ҳ��ַ
    if isMakeHtml = true then
        url = "" 
        if len(listFileName) > 5 then
            url = mid(listFileName, 1, len(listFileName) - 5) & "[id].html" 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        end if 
    else
        url = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    end if 
    content = replace(content, "[$pageInfo$]", webPageControl(nCount, nPageSize, cstr(sPage), url, pageInfo)) 

    if EDITORTYPE = "asp" then
        nX = getRsPageNumber(rs, nCount, nPageSize, sPage)                              '��@����asp����@��   ���Rsҳ��      ��¼����
   

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
        if sPage <> "" then'��@��.netc����@��
            nPage = cint(sPage) - 1 '��@��.netc����@��
        end if '��@��.netc����@��
        sql = "select * from " & db_PREFIX & "" & tableName & " " & addSql & " limit " & nPageSize * nPage & "," & nPageSize '��@��.netc����@��
        rs.open sql, conn, 1, 1 '��@��.netc����@��
        '��PHP��ɾ��rs
        nX = rs.recordCount '��@��.netc����@��
    end if 
    'call echo("sql",sql)
    for i = 1 to nX
        '��PHP��$rs=mysql_fetch_array($rsObj);                                            //��PHP�ã���Ϊ�� asptophpת��������
		'��@��.netc��ʾ@��rs.Read();
        startStr = "[list-" & i & "]" : endStr = "[/list-" & i & "]" 

        '�����ʱ����ǰ����20160202
        if i = nX then
            startStr = "[list-end]" : endStr = "[/list-end]" 
        end if 

        '��[list-mod2]  [/list-mod2]    20150112
        for nModI = 6 to 2 step - 1
            if inStr(defaultStr, startStr) = false and i mod nModI = 0 then
                startStr = "[list-mod" & nModI & "]" : endStr = "[/list-mod" & nModI & "]" 
                if inStr(defaultStr, startStr) > 0 then
                    exit for 
                end if 
            end if 
        next 

        'û������Ĭ��
        if inStr(defaultStr, startStr) = false or startStr = "" then
            startStr = "[list]" : endStr = "[/list]" 
        end if 

        if inStr(defaultStr, startStr) > 0 and inStr(defaultStr, endStr) > 0 then
            s = strCut(defaultStr, startStr, endStr, 2) 

            's = defaultStr
            s = replace(s, "[$id$]", rs("id")) 
            for j = 0 to uBound(splFieldName)
                if splFieldName(j) <> "" then
                    splxx = split(splFieldName(j) & "|||", "|") 
                    fieldName = splxx(0) 
                    replaceStr = rs(fieldName) & "" 
                    s = replaceValueParam(s, fieldName, replaceStr) 
                end if 

                if isMakeHtml = true then
                    url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/detail/detail" & rs("id")) 
                else
                    url = handleWebUrl("?act=detail&id=" & rs("id")) 
                    if rs("customAUrl") <> "" then
                        url = rs("customAUrl") 
                    end if 
                end if 

                'A���������ɫ
                abcolorStr = "" 
                if inStr(fieldNameList, ",titlecolor,") > 0 then
                    'A������ɫ
                    if rs("titlecolor") <> "" then
                        abcolorStr = "color:" & rs("titlecolor") & ";" 
                    end if 
                end if 
                if inStr(fieldNameList, ",flags,") > 0 then
                    'A���ӼӴ�
                    if inStr(rs("flags"), "|b|") > 0 then
                        abcolorStr = abcolorStr & "font-weight:bold;" 
                    end if 
                end if 
                if abcolorStr <> "" then
                    abcolorStr = " style=""" & abcolorStr & """" 
                end if 

                '�򿪷�ʽ2016
                if inStr(fieldNameList, ",target,") > 0 then
                    atargetStr = IIF(rs("target") <> "", " target=""" & rs("target") & """", "") 
                end if 

                'A��title
                if inStr(fieldNameList, ",title,") > 0 then
                    atitleStr = IIF(rs("title") <> "", " title=""" & rs("title") & """", "") 
                end if 

                'A��nofollow
                if inStr(fieldNameList, ",nofollow,") > 0 then
                    anofollowStr = IIF(rs("nofollow") <> 0, " rel=""nofollow""", "") 
                end if 



                s = replaceValueParam(s, "url", url) 
                s = replaceValueParam(s, "abcolor", abcolorStr)                                 'A���Ӽ���ɫ��Ӵ�
                s = replaceValueParam(s, "atitle", atitleStr)                                   'A����title
                s = replaceValueParam(s, "anofollow", anofollowStr)                             'A����nofollow
                s = replaceValueParam(s, "atarget", atargetStr)                                 'A���Ӵ򿪷�ʽ


            next 
        end if 
        'call echo("tableName",tableName)
        idPage = getThisIdPage(db_PREFIX & tableName, rs("id"), 10) 
        '�����ԡ�
        if tableName = "guestbook" then
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=GuestBook&lableTitle=����&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page=" & idPage & "&id=" & rs("id") & "&n=" & getRnd(11) 

        '��Ĭ����ʾ���¡�
        else
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=" & idPage & "&parentid=" & rs("parentid") & "&id=" & rs("id") & "&n=" & getRnd(11) 
        end if 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span") 

        c = c & s 
    rs.moveNext : next : rs.close 
    content = replace(content, "<!--#body start#-->" & defaultStr & "<!--#body end#-->", c) 

    if isMakeHtml = true then
        url = "" 
        if len(listFileName) > 5 then
            url = mid(listFileName, 1, len(listFileName) - 5) & "[id].html" 
            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
        end if 
    else
        url = getUrlAddToParam(getUrl(), "?page=[id]", "replace") 
    end if 

    getDetailList = content 
end function 


'****************************************************
'Ĭ���б�ģ��
function defaultListTemplate(sType, sName)
    dim c, templateHtml, listTemplate, startStr, endStr, lableName 

    templateHtml = getFText(cfg_webTemplate & "/" & templateName) 
    '����Ŀ��������������Ŀ���ͣ���Ĭ��20160630
    lableName = sName & "list" 
    startStr = "<!--#" & lableName & " start#-->" 
    endStr = "<!--#" & lableName & " end#-->" 
    if inStr(templateHtml, startStr) = false or inStr(templateHtml, endStr) = false then
        lableName = sType & "list" 
        startStr = "<!--#" & lableName & " start#-->" 
        endStr = "<!--#" & lableName & " end#-->" 
    end if 
    if inStr(templateHtml, startStr) = false or inStr(templateHtml, endStr) = false then
        lableName = "list" 
        startStr = "<!--#" & lableName & " start#-->" 
        endStr = "<!--#" & lableName & " end#-->" 
    end if 

    'call rwend(templateHtml)
    if inStr(templateHtml, startStr) > 0 and inStr(templateHtml, endStr) > 0 then
        listTemplate = strCut(templateHtml, startStr, endStr, 2) 
    else
        startStr = "<!--#" & lableName 
        endStr = "#-->" 
        if inStr(templateHtml, startStr) > 0 and inStr(templateHtml, endStr) > 0 then
            listTemplate = strCut(templateHtml, startStr, endStr, 2) 
        end if 
    end if 
    if listTemplate = "" then
        c = "<ul class=""list""><!--#body start#-->" & vbCrLf 
        c = c & "[list]    <li><a href=""[$url$]""[$atitle$][$atarget$][$abcolor$][$anofollow$]>[$title$]</a><span class=""time"">[$adddatetime format_time='1'$]</span></li>" & vbCrLf 
        c = c & "[/list]11111111111<!--#body end#--> " & vbCrLf 
        c = c & "</ul>" & vbCrLf 
        c = c & "<div class=""clear10""></div>" & vbCrLf 
        c = c & "<div>[$pageInfo$]</div>" & vbCrLf 
        listTemplate = c 
    end if 
    'call rwend(listTemplate)

    defaultListTemplate = listTemplate 
end function 

call loadRun()                                                                  '��@��.netc����@��
'���ؾ�����
sub loadRun()
    '����Ϊ�˸�.netʹ�õģ���Ϊ��.net����ȫ�ֱ��������б���
	WEB_CACHEFile=replace(replace(WEB_CACHEFile,"[adminDir]",adminDir),"[EDITORTYPE]",EDITORTYPE)	
	
    '���崦��20160622
    cacheHtmlFilePath = "/cache/html/" & setFileName(getThisUrlFileParam()) & ".html" 
    '���û���
    if request("cache") <> "false" and isOnCacheHtml = true then
        if checkFile(cacheHtmlFilePath) = true then
            'call echo("��ȡ�����ļ�","OK")
            call rwend(getftext(cacheHtmlFilePath)) 
        end if 
    end if 

    '��¼��ǰ׺
    if request("db_PREFIX") <> "" then
        db_PREFIX = request("db_PREFIX") 
    elseIf session("db_PREFIX") <> "" then
        db_PREFIX = session("db_PREFIX") 
    end if 
    '������ַ����
    call loadWebConfig() 
    isMakeHtml = false                                                              'Ĭ������HTMLΪ�ر�
    if request("isMakeHtml") = "1" or request("isMakeHtml") = "true" then
        isMakeHtml = true 
    end if 
    templateName = request("templateName")                                          'ģ������

    '�������ݴ���ҳ
    select case request("act")
        case "savedata" : saveData(request("stype")) : response.end()                   '��������
        ''վ��ͳ�� | ����IP[653] | ����PV[9865] | ��ǰ����[65]')
        case "webstat" : webStat(adminDir & "/Data/Stat/") : response.end()             '��վͳ��

        case "saveSiteMap" : isMakeHtml = true : saveSiteMap() : response.end()         '����sitemap.xml

        case "handleAction":
            if request("ishtml") = "1" then
                isMakeHtml = true 
            end if
            rwend(handleAction(request("content")))                                         '������
    end select


    '����html
    if request("act") = "makehtml" then
        call echo("makehtml", "makehtml") 
        isMakeHtml = true 
        call makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' id='" & request("id") & "' ") 
        call createFileGBK("index.html", code) 

    '����Html����վ
    elseIf request("act") = "copyHtmlToWeb" then
        call copyHtmlToWeb() 
    'ȫ������
    elseIf request("act") = "makeallhtml" then
        call makeAllHtml("", "", request("id")) 

    '���ɵ�ǰҳ��
    elseIf request("isMakeHtml") <> "" and request("isSave") <> "" then

        call handlePower("���ɵ�ǰHTMLҳ��")                                            '����Ȩ�޴���
        call writeSystemLog("", "���ɵ�ǰHTMLҳ��")                                     'ϵͳ��־

        isMakeHtml = true 


        call checkIDSQL(request("id")) 
        call rw(makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' columnType='" & request("columnType") & "' id='" & request("id") & "' npage='" & request("page") & "' ")) 
        glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
        if right(glb_filePath, 1) = "/" then
            glb_filePath = glb_filePath & "index.html" 
        elseIf glb_filePath = "" and glb_columnType = "��ҳ" then
            glb_filePath = "index.html" 
        end if 
        '�ļ���Ϊ��  ���ҿ�������html
        if glb_filePath <> "" and glb_isonhtml = true then
            call createDirFolder(getFileAttr(glb_filePath, "1")) 
            call createFileGBK(glb_filePath, code) 
            if request("act") = "detail" then
                conn.execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & request("id")) 
            elseIf request("act") = "nav" then
                if request("id") <> "" then
                    conn.execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & request("id")) 
                else
                    conn.execute("update " & db_PREFIX & "WebColumn set ishtml=true where columnname='" & request("columnName") & "'") 
                end if 
            end if 
            call echo("�����ļ�·��", "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>") 

            '�������������� 20160216
            if glb_columnType = "����" then
                call makeAllHtml("", "", glb_columnId) 
            end if 

        end if 

    'ȫ������
    elseIf request("act") = "Search" then
        call rw(makeWebHtml("actionType='Search' npage='"& IIF(request("page")="","1",request("page")) &"' ")) 
    else
        if LCase(request("issave")) = "1" then
            call makeAllHtml(request("columnType"), request("columnName"), request("columnId")) 
        else
            call checkIDSQL(request("id")) 
            call rw(makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' columnType='" & request("columnType") & "' id='" & request("id") & "' npage='" & request("page") & "' ")) 
        end if 
    end if 
    '��������html
    if isOnCacheHtml = true then
        call createFile(cacheHtmlFilePath, code)                                        '���浽�����ļ���20160622
    end if 
end sub 
'���ID�Ƿ�SQL��ȫ
function checkIDSQL(id)
    if checkNumber(id) = false and id <> "" then
        call eerr("��ʾ", "id���зǷ��ַ�") 
    end if 
end function 
'http://127.0.0.1/aspweb.asp?act=nav&columnName=ASP
'http://127.0.0.1/aspweb.asp?act=detail&id=75
'����html��̬ҳ
function makeWebHtml(action)
    dim actionType, npagesize, npage, s, url, addSql, sortSql, sortFieldName, ascOrDesc 
    dim serchKeyWordName, parentid                                                  '׷����20160716 home
    actionType = rParam(action, "actionType") 
    s = rParam(action, "npage") 
    s = getnumber(s) 
    if s = "" then
        npage = 1 
    else
        npage = CInt(s) 
    end if 
    '����
    if actionType = "nav" then
        glb_columnType = rParam(action, "columnType") 
        glb_columnName = rParam(action, "columnName") 
        glb_columnId = rParam(action, "columnId") 
        if glb_columnId = "" then
            glb_columnId = rParam(action, "id") 
        end if 
        if glb_columnType <> "" then
            addSql = "where columnType='" & glb_columnType & "'" 
        end if 
        if glb_columnName <> "" then
            addSql = getWhereAnd(addSql, "where columnName='" & glb_columnName & "'") 
        end if 
        if glb_columnId <> "" then
            addSql = getWhereAnd(addSql, "where id=" & glb_columnId & "") 
        end if 
        'call echo("addsql",addsql)
        rs.open "Select * from " & db_PREFIX & "webcolumn " & addSql, conn, 1, 1 
        if not rs.EOF then
            glb_columnId = rs("id") 
            glb_columnName = rs("columnname") 
            glb_columnType = rs("columntype") 
            glb_bodyContent = rs("bodycontent") 
            glb_detailTitle = glb_columnName 
            glb_flags = rs("flags") 
            npagesize = rs("npagesize")                                                     'ÿҳ��ʾ����
            glb_isonhtml = IIF(rs("isonhtml") = true, true, false)                          '�Ƿ����ɾ�̬��ҳ
            sortSql = " " & rs("sortsql")                                                   '����SQL

            if rs("webTitle") <> "" then
                cfg_webTitle = rs("webTitle")                                                   '��ַ����
            end if 
            if rs("webKeywords") <> "" then
                cfg_webKeywords = rs("webKeywords")                                             '��վ�ؼ���
            end if 
            if rs("webDescription") <> "" then
                cfg_webDescription = rs("webDescription")                                       '��վ����
            end if 
            if templateName = "" then
                if trim(rs("templatePath")) <> "" then
                    templateName = rs("templatePath") 
                elseIf rs("columntype") <> "��ҳ" then
                    templateName = getDateilTemplate(rs("id"), "List") 
                end if 
            end if 
        end if : rs.close 
        glb_columnENType = handleColumnType(glb_columnType) 
        glb_url = getColumnUrl(glb_columnName, "name") 

        '�������б�
        if inStr("|��Ʒ|����|��Ƶ|����|����|", "|" & glb_columnType & "|") > 0 then
            glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "��Ŀ�б�", "*", npagesize, cstr(npage), "where parentid=" & glb_columnId & sortSql) 
        '�������б�
        elseIf inStr("|����|", "|" & glb_columnType & "|") > 0 then
            glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "GuestBook", "�����б�", "*", npagesize, cstr(npage), " where isthrough<>0 " & sortSql) 
        elseIf glb_columnType = "�ı�" then
            '������Ŀ�ӹ���
            if request("gl") = "edit" then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            end if 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=WebColumn&lableTitle=��վ��Ŀ&nPageSize=10&page=&id=" & glb_columnId & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 

        end if 
    'ϸ��
    elseIf actionType = "detail" then
        glb_locationType = "detail" 
        rs.open "Select * from " & db_PREFIX & "articledetail where id=" & rParam(action, "id"), conn, 1, 1 
        if not rs.EOF then
            glb_columnName = getColumnName(rs("parentid")) 
            glb_detailTitle = rs("title") 
            glb_flags = rs("flags") 
            glb_isonhtml = rs("isonhtml")                                                   '�Ƿ����ɾ�̬��ҳ
            glb_id = rs("id")                                                               '����ID
            if isMakeHtml = true then
                glb_url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/detail/detail" & rs("id")) 
            else
                glb_url = handleWebUrl("?act=detail&id=" & rs("id")) 
            end if 

            if rs("webTitle") <> "" then
                cfg_webTitle = rs("webTitle")                                                   '��ַ����
            end if 
            if rs("webKeywords") <> "" then
                cfg_webKeywords = rs("webKeywords")                                             '��վ�ؼ���
            end if 
            if rs("webDescription") <> "" then
                cfg_webDescription = rs("webDescription")                                       '��վ����
            end if 

            '�Ľ�20160628
            sortFieldName = "id" 
            ascOrDesc = "asc" 
            addSql = trim(getWebColumnSortSql(rs("parentid"))) 
            if addSql <> "" then
                sortFieldName = trim(replace(replace(replace(addSql, "order by", ""), " desc", ""), " asc", "")) 
                if instr(addSql, " desc") > 0 then
                    ascOrDesc = "desc" 
                end if 
            end if 
            glb_articleAuthor = rs("author") 
            glb_articleAdddatetime = rs("adddatetime") 
            glb_upArticle = upArticle(rs("parentid"), sortFieldName, rs(sortFieldName), ascOrDesc) 
            glb_downArticle = downArticle(rs("parentid"), sortFieldName, rs(sortFieldName), ascOrDesc) 
            glb_aritcleRelatedTags = aritcleRelatedTags(rs("relatedtags")) 
            glb_aritcleSmallImage = rs("smallimage") 
            glb_aritcleBigImage = rs("bigimage") 
            glb_articlehits = rs("hits") 
            conn.execute("update " & db_PREFIX & "articledetail set hits=hits+1 where id=" & rs("id")) '���µ����
            '��������
            'glb_bodyContent = "<div class=""articleinfowrap"">[$articleinfowrap$]</div>" & rs("bodycontent") & "[$relatedtags$]<ul class=""updownarticlewrap"">[$updownArticle$]</ul>"
            '��һƪ���£���һƪ����
            'glb_bodyContent = Replace(glb_bodyContent, "[$updownArticle$]", upArticle(rs("parentid"), "sortrank", rs("sortrank")) & downArticle(rs("parentid"), "sortrank", rs("sortrank")))
            'glb_bodyContent = Replace(glb_bodyContent, "[$articleinfowrap$]", "��Դ��" & rs("author") & " &nbsp; ����ʱ�䣺" & format_Time(rs("adddatetime"), 1))
            'glb_bodyContent = Replace(glb_bodyContent, "[$relatedtags$]", aritcleRelatedTags(rs("relatedtags")))

            glb_bodyContent = rs("bodycontent") 

            '������ϸ�ӿ���
            if request("gl") = "edit" then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            end if 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=" & rParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 

            if templateName = "" then
                if trim(rs("templatePath")) <> "" then
                    templateName = rs("templatePath") 
                else
                    templateName = getDateilTemplate(rs("parentid"), "Detail") 
                end if 
            end if 

        end if : rs.close 

    '��ҳ
    elseIf actionType = "onepage" then
        rs.open "Select * from " & db_PREFIX & "onepage where id=" & rParam(action, "id"), conn, 1, 1 
        if not rs.EOF then
            glb_detailTitle = rs("title") 
            glb_isonhtml = rs("isonhtml")                                                   '�Ƿ����ɾ�̬��ҳ
            if isMakeHtml = true then
                glb_url = getHandleRsUrl(rs("fileName"), rs("customAUrl"), "/page/page" & rs("id")) 
            else
                glb_url = handleWebUrl("?act=detail&id=" & rs("id")) 
            end if 

            if rs("webTitle") <> "" then
                cfg_webTitle = rs("webTitle")                                                   '��ַ����
            end if 
            if rs("webKeywords") <> "" then
                cfg_webKeywords = rs("webKeywords")                                             '��վ�ؼ���
            end if 
            if rs("webDescription") <> "" then
                cfg_webDescription = rs("webDescription")                                       '��վ����
            end if 
            '����
            glb_bodyContent = rs("bodycontent") 


            '������ϸ�ӿ���
            if request("gl") = "edit" then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            end if 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=" & rParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 


            if templateName = "" then
                if trim(rs("templatePath")) <> "" then
                    templateName = rs("templatePath") 
                else
                    templateName = "Main_Model.html" 
                'call echo(templateName,"templateName")
                end if 
            end if 

        end if : rs.close 

    '����
    elseIf actionType = "Search" then
        templateName = "Main_Model.html" 
        serchKeyWordName = request("keywordname") 
        parentid = request("parentid") 
        if serchKeyWordName = "" then
            serchKeyWordName = "wd" 
        end if 
        glb_searchKeyWord = replace(request(serchKeyWordName), "<", "&lt;") 
        addSql = "" 
        if parentid <> "" then
            addSql = " where parentid=" & parentid 
        end if 
        addSql = getWhereAnd(addSql, " where title like '%" & glb_searchKeyWord & "%'") 
        npagesize = 20  
        glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "��վ��Ŀ", "*", npagesize, cstr(npage), addSql) 
        positionEndStr = " >> �������ݡ�" & glb_searchKeyWord & "��" 
    '���صȴ�
    elseIf actionType = "loading" then
        call rwend("ҳ�����ڼ����С�����") 
    end if 
    'ģ��Ϊ�գ�����Ĭ����ҳģ��
    if templateName = "" then
        templateName = "Index_Model.html"                                               'Ĭ��ģ��
    end if 
    '��⵱ǰ·���Ƿ���ģ��
    if inStr(templateName, "/") = false then
        templateName = cfg_webTemplate & "/" & templateName 
    end if 
    'call echo("templateName",templateName)
    if checkFile(templateName) = false then
        call eerr("δ�ҵ�ģ���ļ�", templateName) 
    end if 
    code = getftext(templateName) 

    code = handleAction(code)                                                       '������
    code = thisPosition(code)                                                       'λ��
    code = replaceGlobleVariable(code)                                              '�滻ȫ�ֱ�ǩ
    code = handleAction(code)                                                       '������    '����һ�Σ��������������ﶯ��

    'call die(code)
    code = handleAction(code)                                                       '������
    code = handleAction(code)                                                       '������
    code = thisPosition(code)                                                       'λ��
    code = replaceGlobleVariable(code)                                              '�滻ȫ�ֱ�ǩ
    code = delTemplateMyNote(code)                                                  'ɾ����������
    code = handleAction(code)                                                       '������

    '��ʽ��HTML
    if inStr(cfg_flags, "|formattinghtml|") > 0 then
        'code = HtmlFormatting(code)        '��
        code = handleHtmlFormatting(code, false, 0, "ɾ������")                         '�Զ���
    '��ʽ��HTML�ڶ���
    elseIf inStr(cfg_flags, "|formattinghtmltow|") > 0 then
        code = htmlFormatting(code)                                                     '��
        code = handleHtmlFormatting(code, false, 0, "ɾ������")                         '�Զ���
    'ѹ��HTML
    elseIf inStr(cfg_flags, "|ziphtml|") > 0 then
        code = ziphtml(code) 

    end if 
    '�պϱ�ǩ
    if inStr(cfg_flags, "|labelclose|") > 0 then
        code = handleCloseHtml(code, true, "")                                          'ͼƬ�Զ���alt  "|*|",
    end if 

    '���߱༭20160127
    if rq("gl") = "edit" then
        if inStr(code, "</head>") > 0 then
            if inStr(lcase(code), "jquery.min.js") = false then
                code = replace(code, "</head>", "<script src=""/Jquery/jquery.Min.js""></script></head>") 
            end if 
            code = replace(code, "</head>", "<script src=""/Jquery/Callcontext_menu.js""></script></head>") 
        end if 
        if inStr(code, "<body>") > 0 then
        'Code = Replace(Code,"<body>", "<body onLoad=""ContextMenu.intializeContextMenu()"">")
        end if 
    end if 
    'call echo(templateName,templateName)
    makeWebHtml = code 
end function 

'���Ĭ��ϸ��ģ��ҳ
function getDateilTemplate(parentid, templateType)
    dim templateName 
    templateName = "Main_Model.html" 
    rsx.open "select * from " & db_PREFIX & "webcolumn where id=" & parentid, conn, 1, 1 
    if not rsx.EOF then
        'call echo("columntype",rsx("columntype"))
        if rsx("columntype") = "����" then
            '����ϸ��ҳ
            if checkFile(cfg_webTemplate & "/News_" & templateType & ".html") = true then
                templateName = "News_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "��Ʒ" then
            '��Ʒϸ��ҳ
            if checkFile(cfg_webTemplate & "/Product_" & templateType & ".html") = true then
                templateName = "Product_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "����" then
            '����ϸ��ҳ
            if checkFile(cfg_webTemplate & "/Down_" & templateType & ".html") = true then
                templateName = "Down_" & templateType & ".html" 
            end if 

        elseIf rsx("columntype") = "��Ƶ" then
            '��Ƶϸ��ҳ
            if checkFile(cfg_webTemplate & "/Video_" & templateType & ".html") = true then
                templateName = "Video_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "����" then
            '��Ƶϸ��ҳ
            if checkFile(cfg_webTemplate & "/GuestBook_" & templateType & ".html") = true then
                templateName = "Video_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "�ı�" then
            '��Ƶϸ��ҳ
            if checkFile(cfg_webTemplate & "/Page_" & templateType & ".html") = true then
                templateName = "Page_" & templateType & ".html" 
            end if 
        end if 
    end if : rsx.close 
    'call echo(templateType,templateName)
    getDateilTemplate = templateName 

end function 


'����ȫ��htmlҳ��
sub makeAllHtml(columnType, columnName, columnId)
    dim action, s, i, nPageSize, nCountSize, nPage, addSql, url, articleSql 
    call handlePower("����ȫ��HTMLҳ��")                                            '����Ȩ�޴���
    call writeSystemLog("", "����ȫ��HTMLҳ��")                                     'ϵͳ��־

    isMakeHtml = true 
    '��Ŀ
    call echo("��Ŀ", "") 
    if columnType <> "" then
        addSql = "where columnType='" & columnType & "'" 
    end if 
    if columnName <> "" then
        addSql = getWhereAnd(addSql, "where columnName='" & columnName & "'") 
    end if 
    if columnId <> "" then
        addSql = getWhereAnd(addSql, "where id in(" & columnId & ")") 
    end if 
    rss.open "select * from " & db_PREFIX & "webcolumn " & addSql & " order by sortrank asc", conn, 1, 1 
    while not rss.EOF
        glb_columnName = "" 
        '��������html
        if rss("isonhtml") = true then
            if inStr("|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����|", "|" & rss("columntype") & "|") > 0 then
                if rss("columntype") = "����" then
                    nCountSize = getRecordCount(db_PREFIX & "guestbook", "")                        '��¼��
                else
                    nCountSize = getRecordCount(db_PREFIX & "articledetail", " where parentid=" & rss("id")) '��¼��
                end if 
                nPageSize = rss("npagesize") 
                nPage = getPageNumb(CInt(nCountSize), CInt(nPageSize)) 
                if nPage <= 0 then
                    nPage = 1 
                end if 
                for i = 1 to nPage
                    url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/nav" & rss("id")) 
                    glb_filePath = replace(url, cfg_webSiteUrl, "") 
                    if right(glb_filePath, 1) = "/" or glb_filePath = "" then
                        glb_filePath = glb_filePath & "index.html" 
                    end if 
                    'call echo("glb_filePath",glb_filePath)
                    action = " action actionType='nav' columnName='" & rss("columnname") & "' npage='" & i & "' listfilename='" & glb_filePath & "' " 
                    'call echo("action",action)
                    call makeWebHtml(action) 
                    if i > 1 then
                        glb_filePath = mid(glb_filePath, 1, len(glb_filePath) - 5) & i & ".html" 
                    end if 
                    s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
                    call echo(action, s) 
                    if glb_filePath <> "" then
                        call createDirFolder(getFileAttr(glb_filePath, "1")) 
                        call createFileGBK(glb_filePath, code) 
                    end if 
                    doevents() 
                    templateName = ""                                                               '���ģ���ļ�����
                next 
            else
                action = " action actionType='nav' columnName='" & rss("columnname") & "'" 
                call makeWebHtml(action) 
                glb_filePath = replace(getColumnUrl(rss("columnname"), "name"), cfg_webSiteUrl, "") 
                if right(glb_filePath, 1) = "/" or glb_filePath = "" then
                    glb_filePath = glb_filePath & "index.html" 
                end if 
                s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
                call echo(action, s) 
                if glb_filePath <> "" then
                    call createDirFolder(getFileAttr(glb_filePath, "1")) 
                    call createFileGBK(glb_filePath, code) 
                end if 
                doevents() 
                templateName = "" 
            end if 
            conn.execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & rss("id")) '���µ���Ϊ����״̬
        end if 
    rss.moveNext : wend : rss.close 

    '��������ָ����Ŀ��Ӧ����
    if columnId <> "" then
        articleSql = "select * from " & db_PREFIX & "articledetail where parentid=" & columnId & " order by sortrank asc" 
    '������������
    elseIf addSql = "" then
        articleSql = "select * from " & db_PREFIX & "articledetail order by sortrank asc" 
    end if 
    if articleSql <> "" then
        '����
        call echo("����", "") 
        rss.open articleSql, conn, 1, 1 
        while not rss.EOF
            glb_columnName = "" 
            action = " action actionType='detail' columnName='" & rss("parentid") & "' id='" & rss("id") & "'" 
            'call echo("action",action)
            call makeWebHtml(action) 
            glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
            if right(glb_filePath, 1) = "/" then
                glb_filePath = glb_filePath & "index.html" 
            end if 
            s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
            call echo(action, s) 
            '�ļ���Ϊ��  ���ҿ�������html
            if glb_filePath <> "" and rss("isonhtml") = true then
                call createDirFolder(getFileAttr(glb_filePath, "1")) 
                call createFileGBK(glb_filePath, code) 
                conn.execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & rss("id")) '��������Ϊ����״̬
            end if 
            templateName = ""                                                               '���ģ���ļ�����
        rss.moveNext : wend : rss.close 
    end if 

    if addSql = "" then
        '��ҳ
        call echo("��ҳ", "") 
        rss.open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
        while not rss.EOF
            glb_columnName = "" 
            action = " action actionType='onepage' id='" & rss("id") & "'" 
            'call echo("action",action)
            call makeWebHtml(action) 
            glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
            if right(glb_filePath, 1) = "/" then
                glb_filePath = glb_filePath & "index.html" 
            end if 
            s = "<a href=""" & glb_filePath & """ target='_blank'>" & glb_filePath & "</a>(" & rss("isonhtml") & ")" 
            call echo(action, s) 
            '�ļ���Ϊ��  ���ҿ�������html
            if glb_filePath <> "" and rss("isonhtml") = true then
                call createDirFolder(getFileAttr(glb_filePath, "1")) 
                call createFileGBK(glb_filePath, code) 
                conn.execute("update " & db_PREFIX & "onepage set ishtml=true where id=" & rss("id")) '���µ�ҳΪ����״̬
            end if 
            templateName = ""                                                               '���ģ���ļ�����
        rss.moveNext : wend : rss.close 

    end if 
end sub 

'����html����վ
sub copyHtmlToWeb()
    dim webDir, toWebDir, toFilePath, filePath, fileName, fileList, splStr, content, s, s1, c, webImages, webCss, webJs, splJs 
    dim webFolderName, jsFileList, setFileCode, nErrLevel, jsFilePath, url 

    setFileCode = request("setcode")                                                '�����ļ��������

    call handlePower("��������HTMLҳ��")                                            '����Ȩ�޴���
    call writeSystemLog("", "��������HTMLҳ��")                                     'ϵͳ��־

    webFolderName = cfg_webTemplate 
    if left(webFolderName, 1) = "/" then
        webFolderName = mid(webFolderName, 2) 
    end if 
    if right(webFolderName, 1) = "/" then
        webFolderName = mid(webFolderName, 1, len(webFolderName) - 1) 
    end if 
    if inStr(webFolderName, "/") > 0 then
        webFolderName = mid(webFolderName, inStr(webFolderName, "/") + 1) 
    end if 
    webDir = "/htmladmin/" & webFolderName & "/" 
    toWebDir = "/htmlw" & "eb/viewweb/" 
    call createDirFolder(toWebDir) 
    toWebDir = toWebDir & pinYin2(webFolderName) & "/" 

    call deleteFolder(toWebDir)                                                     'ɾ��
    call createFolder("/htmlweb/web")                                               '�����ļ��� ��ֹweb�ļ��в�����20160504
    call deleteFolder(webDir) 
    call createDirFolder(webDir) 
    webImages = webDir & "Images/" 
    webCss = webDir & "Css/" 
    webJs = webDir & "Js/" 
    call copyFolder(cfg_webImages, webImages) 
    call copyFolder(cfg_webCss, webCss) 
    call createFolder(webJs)                                                        '����Js�ļ���


    '����Js�ļ���
    splJs = split(getDirJsList(webJs), vbCrLf) 
    for each filePath in splJs
        if filePath <> "" then
            toFilePath = webJs & getFileName(filePath) 
            call echo("js", filePath) 
            call moveFile(filePath, toFilePath) 
        end if 
    next 
    '����Css�ļ���
    splStr = split(getDirCssList(webCss), vbCrLf) 
    for each filePath in splStr
        if filePath <> "" then
            content = getftext(filePath) 
            content = replace(content, cfg_webImages, "../images/") 

            content = deleteCssNote(content) 
            content = phptrim(content) 
            '����Ϊutf-8���� 20160527
            if lcase(setFileCode) = "utf-8" then
                content = replace(content, "gb2312", "utf-8") 
            end if 
            call writeToFile(filePath, content, setFileCode) 
            call echo("css", cfg_webImages) 
        end if 
    next 
    '������ĿHTML
    isMakeHtml = true 
    rss.open "select * from " & db_PREFIX & "webcolumn where isonhtml=true", conn, 1, 1 
    while not rss.EOF
        glb_filePath = replace(getColumnUrl(rss("columnname"), "name"), cfg_webSiteUrl, "") 

        if right(glb_filePath, 1) = "/" or right(glb_filePath, 1) = "" then
            glb_filePath = glb_filePath & "index.html" 
        end if 
        if right(glb_filePath, 5) = ".html" then
            if right(glb_filePath, 11) = "/index.html" then
                fileList = fileList & glb_filePath & vbCrLf 
            else
                fileList = glb_filePath & vbCrLf & fileList 
            end if 
            fileName = replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            call copyfile(glb_filePath, toFilePath) 
            call echo("����", glb_filePath) 
        end if 
    rss.moveNext : wend : rss.close 
    '��������HTML
    rss.open "select * from " & db_PREFIX & "articledetail where isonhtml=true", conn, 1, 1 
    while not rss.EOF
        glb_url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/detail/detail" & rss("id")) 
        glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
        if right(glb_filePath, 1) = "/" or right(glb_filePath, 1) = "" then
            glb_filePath = glb_filePath & "index.html" 
        end if 
        if right(glb_filePath, 5) = ".html" then
            if right(glb_filePath, 11) = "/index.html" then
                fileList = fileList & glb_filePath & vbCrLf 
            else
                fileList = glb_filePath & vbCrLf & fileList 
            end if 
            fileName = replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            call copyfile(glb_filePath, toFilePath) 
            call echo("����" & rss("title"), glb_filePath) 
        end if 
    rss.moveNext : wend : rss.close 
    '���Ƶ���HTML
    rss.open "select * from " & db_PREFIX & "onepage where isonhtml=true", conn, 1, 1 
    while not rss.EOF
        glb_url = getHandleRsUrl(rss("fileName"), rss("customAUrl"), "/page/page" & rss("id")) 
        glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
        if right(glb_filePath, 1) = "/" or right(glb_filePath, 1) = "" then
            glb_filePath = glb_filePath & "index.html" 
        end if 
        if right(glb_filePath, 5) = ".html" then
            if right(glb_filePath, 11) = "/index.html" then
                fileList = fileList & glb_filePath & vbCrLf 
            else
                fileList = glb_filePath & vbCrLf & fileList 
            end if 
            fileName = replace(glb_filePath, "/", "_") 
            toFilePath = webDir & fileName 
            call copyfile(glb_filePath, toFilePath) 
            call echo("��ҳ" & rss("title"), glb_filePath) 
        end if 
    rss.moveNext : wend : rss.close 
    '��������html�ļ��б�
    'call echo(cfg_webSiteUrl,cfg_webTemplate)
    'call rwend(fileList)
    dim sourceUrl, replaceUrl 
    splStr = split(fileList, vbCrLf) 
    for each filePath in splStr
        if filePath <> "" then
            filePath = webDir & replace(filePath, "/", "_") 
            call echo("filePath", filePath) 
            content = getftext(filePath) 
            for each s in splStr
                s1 = s 
                if right(s1, 11) = "/index.html" then
                    s1 = left(s1, len(s1) - 11) & "/" 
                end if 
                sourceUrl = cfg_webSiteUrl & s1 
                replaceUrl = cfg_webSiteUrl & replace(s, "/", "_") 
                'Call echo(sourceUrl, replaceUrl)                             '����  ���������ʾ20160613
                content = replace(content, sourceUrl, replaceUrl) 
            next 
            content = replace(content, cfg_webSiteUrl, "")                                  'ɾ����ַ
            content = replace(content, cfg_webTemplate & "/", "")                           'ɾ��ģ��·�� ��

            'content=nullLinkAddDefaultName(content)
            for each s in splJs
                if s <> "" then
                    fileName = getFileName(s) 
                    content = replace(content, "Images/" & fileName, "js/" & fileName) 
                end if 
            next 
            if inStr(content, "/Jquery/Jquery.Min.js") > 0 then
                content = replace(content, "/Jquery/Jquery.Min.js", "js/Jquery.Min.js") 
                call copyfile("/Jquery/Jquery.Min.js", webJs & "/Jquery.Min.js") 
            end if 
            content = replace(content, "<a href="""" ", "<a href=""index.html"" ")    '����ҳ��index.html

            call createFileGBK(filePath, content) 
        end if 
    next 

    '�Ѹ�����վ���µ�images/�ļ����µ�js�Ƶ�js/�ļ�����  20160315
    dim htmlFileList, splHtmlFile, splJsFile, htmlFilePath, jsFileName 
    jsFileList = getDirJsNameList(webImages) 
    htmlFileList = getDirHtmlList(webDir) 
    splHtmlFile = split(htmlFileList, vbCrLf) 
    splJsFile = split(jsFileList, vbCrLf) 
    for each htmlFilePath in splHtmlFile
        content = getftext(htmlFilePath) 
        for each jsFileName in splJsFile
            content = regExp_Replace(content, "Images/" & jsFileName, "js/" & jsFileName) 
        next 

        nErrLevel = 0 
        content = handleHtmlFormatting(content, false, nErrLevel, "|ɾ������|")         '|ɾ������|
        content = handleCloseHtml(content, true, "")                                    '�պϱ�ǩ
        'nErrLevel = checkHtmlFormatting(content) 
        if checkHtmlFormatting(content) = false then
            call echored(htmlFilePath & "(��ʽ������)", nErrLevel)                          'ע��
        end if 
        '����Ϊutf-8����
        if lcase(setFileCode) = "utf-8" then
            content = replace(content, "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />", "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />") 
        end if 
        content = phptrim(content) 
        call writeToFile(htmlFilePath, content, setFileCode) 
    next 
    'images��js�ƶ���js��
    for each jsFileName in splJsFile
        jsFilePath = webImages & jsFileName 
        content = getftext(jsFilePath) 
        content = phptrim(content) 
        call writeToFile(webJs & jsFileName, content, setFileCode) 
        call deleteFile(jsFilePath) 
    next 

    call copyFolder(webDir, toWebDir) 
    'ʹhtmlWeb�ļ�����phpѹ��
    if request("isMakeZip") = "1" then
        call makeHtmlWebToZip(webDir) 
    end if 
    'ʹ��վ��xml���20160612
    if request("isMakeXml") = "1" then
        call makeHtmlWebToXmlZip("/htmladmin/", webFolderName) 
    end if 
    '�����ַ
    url = "http://10.10.10.57/" & toWebDir 
    call echo("���", "<a href='" & url & "' target='_blank'>" & url & "</a>") 
end sub 
'ʹhtmlWeb�ļ�����phpѹ��
function makeHtmlWebToZip(webDir)
    dim content, splStr, filePath, c, arrayFile, fileName, fileType, isTrue 
    dim webFolderName 
    dim cleanFileList 
    splStr = split(webDir, "/") 
    webFolderName = splStr(2) 
    'content = getFileFolderList(webDir, true, "ȫ��", "", "ȫ���ļ���", "", "") 		'��������
	content=getDirAllFileList(webDir,"")
    splStr = split(content, vbCrLf) 
    for each filePath in splStr
        if checkfolder(filePath) = false then
            arrayFile = handleFilePathArray(filePath) 
            fileName = LCase(arrayFile(2)) 
            fileType = LCase(arrayFile(4)) 
            fileName = remoteNumber(fileName) 
            isTrue = true 

            if inStr("|" & cleanFileList & "|", "|" & fileName & "|") > 0 and fileType = "html" then
                isTrue = false 
            end if 
            if isTrue = true then
                'call echo(fileType,fileName)
                if c <> "" then c = c & "|" 
                c = c & replace(filePath, handlePath("/"), "") 
                cleanFileList = cleanFileList & fileName & "|" 
            end if 
        end if 
    next 
    call rw(c) 
    c = c & "|||||" 
    call createFileGBK("htmlweb/1.txt", c) 
    call echo("<hr>cccccccccccc", c) 
    '���ж�����ļ�����20160309
    if checkFile("/myZIP.php") = true then
        call echo("", XMLPost(getHost() & "/myZIP.php?webFolderName=" & webFolderName, "content=" & escape(c))) 
    end if 

end function 
'ʹ��վ��xml���20160612
function makeHtmlWebToXmlZip(sNewWebDir, rootDir)
    dim xmlFileName, xmlSize 
    xmlFileName = setFileName(getIP()) & "_update.xml"                              '���ip�п���Ϊ��:: ����ʱ��������

    'sNewWebDir="\Templates2015\"
    'rootDir="\sharembweb\"

    dim objXmlZIP : set objXmlZIP = new xmlZIP
        call objXmlZIP.callRun(handlePath(sNewWebDir), handlePath(sNewWebDir & rootDir), false, xmlFileName) 
        call echo(handlePath(sNewWebDir), handlePath(sNewWebDir & rootDir)) 
    set objXmlZIP = nothing 
    doevents 
    xmlSize = getFSize(xmlFileName) 
    xmlSize = printSpaceValue(xmlSize) 
    call echo("����xml����ļ�", "<a href=/tools/downfile.asp?act=download&downfile=" & xorEnc("/" & xmlFileName, 31380) & " title='�������'>�������" & xmlFileName & "(" & xmlSize & ")</a>") 
end function 


'���ɸ���sitemap.xml 20160118
sub saveSiteMap()
    dim isWebRunHtml                                                                '�Ƿ�Ϊhtml��ʽ��ʾ��վ
    dim changefreg                                                                  '����Ƶ��
    dim priority                                                                    '���ȼ�
    dim s, c, url,sql 
    call handlePower("�޸�����SiteMap")                                             '����Ȩ�޴���

    changefreg = request("changefreg") 
    priority = request("priority") 
    call loadWebConfig()                                                            '��������
    'call eerr("cfg_flags",cfg_flags)
    if inStr(cfg_flags, "|htmlrun|") > 0 then
        isWebRunHtml = true 
    else
        isWebRunHtml = false 
    end if 

    c = c & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf 
    c = c & vbTab & "<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & vbCrLf 
    dim rsx : set rsx = createObject("Adodb.RecordSet")
        '��Ŀ
        rsx.open "select * from " & db_PREFIX & "webcolumn where isonhtml<>0 order by sortrank asc", conn, 1, 1 
        while not rsx.EOF
            if rsx("nofollow") = 0 then
                c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                if isWebRunHtml = true then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
                    url = handleAction(url) 
                else
                    url = escape("?act=nav&columnName=" & rsx("columnname")) 
                end if 
                url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                'call echo(cfg_webSiteUrl,url)

                c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
                c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
                call echo("��Ŀ", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
            end if 
        rsx.moveNext : wend : rsx.close 

        '����
		sql="select * from " & db_PREFIX & "articledetail  where isonhtml<>0 order by sortrank asc"
        rsx.open sql, conn, 1, 1 
        while not rsx.EOF
            if rsx("nofollow") = 0 then
                c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                if isWebRunHtml = true then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/detail/detail" & rsx("id")) 
                    url = handleAction(url) 
                else
                    url = "?act=detail&id=" & rsx("id") 
                end if 
                url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                'call echo(cfg_webSiteUrl,url)

                c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
                c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
                call echo("����", "<a href=""" & url & """>" & url & "</a>") 
            end if 
        rsx.moveNext : wend : rsx.close 

        '��ҳ
        rsx.open "select * from " & db_PREFIX & "onepage where isonhtml<>0 order by sortrank asc", conn, 1, 1 
        while not rsx.EOF
            if rsx("nofollow") = 0 then
                c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                if isWebRunHtml = true then
                    url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
                    url = handleAction(url) 
                else
                    url = "?act=onepage&id=" & rsx("id") 
                end if 
                url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                'call echo(cfg_webSiteUrl,url)

                c = c & copystr(vbTab, 3) & "<loc>" & url & "</loc>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<lastmod>" & format_Time(rsx("updatetime"), 2) & "</lastmod>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<changefreq>" & changefreg & "</changefreq>" & vbCrLf 
                c = c & copystr(vbTab, 3) & "<priority>" & priority & "</priority>" & vbCrLf 
                c = c & copystr(vbTab, 2) & "</url>" & vbCrLf 
                call echo("��ҳ", "<a href=""" & url & """>" & url & "</a>") 
            end if 
        rsx.moveNext : wend : rsx.close 
        c = c & vbTab & "</urlset>" & vbCrLf 
        call loadWebConfig() 
        call createFile("sitemap.xml", c) 
        call echo("����sitemap.xml�ļ��ɹ�", "<a href='/sitemap.xml' target='_blank'>���Ԥ��sitemap.xml</a>") 


        '�ж��Ƿ�����sitemap.html
        if request("issitemaphtml") = "1" then
            c = "" 
            '�ڶ���
            '��Ŀ
            rsx.open "select * from " & db_PREFIX & "webcolumn order by sortrank asc", conn, 1, 1 
            while not rsx.EOF
                if rsx("nofollow") = 0 then
                    if isWebRunHtml = true then
                        url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/nav" & rsx("id")) 
                        url = handleAction(url) 
                    else
                        url = escape("?act=nav&columnName=" & rsx("columnname")) 
                    end if 
                    url = urlAddHttpUrl(cfg_webSiteUrl, url) 

                    '�ж��Ƿ�����html
                    if rsx("isonhtml") = true then
                        s = "<a href=""" & url & """>" & rsx("columnname") & "</a>" 
                    else
                        s = "<span>" & rsx("columnname") & "</span>" 
                    end if 
                    c = c & "<li style=""width:20%;"">" & s & vbCrLf & "<ul>" & vbCrLf 

                    '����
					sql="select * from " & db_PREFIX & "articledetail where parentId=" & rsx("id") & " order by sortrank asc"
                    rss.open sql, conn, 1, 1 
                    while not rss.EOF
                        if rss("nofollow") = 0 then
                            if isWebRunHtml = true then
                                url = getRsUrl(rss("fileName"), rss("customAUrl"), "/detail/detail" & rss("id")) 
                                url = handleAction(url) 
                            else
                                url = "?act=detail&id=" & rss("id") 
                            end if 
                            url = urlAddHttpUrl(cfg_webSiteUrl, url) 
                            '�ж��Ƿ�����html
                            if rss("isonhtml") = true then
                                s = "<a href=""" & url & """>" & rss("title") & "</a>" 
                            else
                                s = "<span>" & rss("title") & "</span>" 
                            end if 
                            c = c & "<li style=""width:20%;"">" & s & "</li>" & vbCrLf 
                        end if 
                    rss.moveNext : wend : rss.close 
                    c = c & "</ul>" & vbCrLf & "</li>" & vbCrLf 


                end if 
            rsx.moveNext : wend : rsx.close 

            '����
            c = c & "<li style=""width:20%;""><a href=""javascript:;"">�����б�</a>" & vbCrLf & "<ul>" & vbCrLf 
            rsx.open "select * from " & db_PREFIX & "onepage order by sortrank asc", conn, 1, 1 
            while not rsx.EOF
                if rsx("nofollow") = 0 then
                    c = c & copystr(vbTab, 2) & "<url>" & vbCrLf 
                    if isWebRunHtml = true then
                        url = getRsUrl(rsx("fileName"), rsx("customAUrl"), "/page/detail" & rsx("id")) 
                        url = handleAction(url) 
                    else
                        url = "?act=onepage&id=" & rsx("id") 
                    end if 
                    '�ж��Ƿ�����html
                    if rsx("isonhtml") = true then
                        s = "<a href=""" & url & """>" & rsx("title") & "</a>" 
                    else
                        s = "<span>" & rsx("title") & "</span>" 
                    end if 

                    c = c & "<li style=""width:20%;"">" & s & "</li>" & vbCrLf                'target=""_blank""  ȥ��
                end if 
            rsx.moveNext : wend : rsx.close 
            c = c & "</ul>" & vbCrLf & "</li>" & vbCrLf 

            dim templateContent 
            templateContent = getftext(adminDir & "/template_SiteMap.html") 


            templateContent = replace(templateContent, "{$content$}", c) 
            templateContent = replace(templateContent, "{$Web_Title$}", cfg_webTitle) 


            call createFile("sitemap.html", templateContent) 
            call echo("����sitemap.html�ļ��ɹ�", "<a href='/sitemap.html' target='_blank'>���Ԥ��sitemap.html</a>") 
        end if 
        call writeSystemLog("", "����sitemap.xml")                                      'ϵͳ��־
end sub
%>   
