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
<!--#Include File = "Inc/Config.Asp"-->
<!--#Include File = "inc/admin_function.asp"--> 
<% 
'asp������   �ķ�����111
call openconn()    
'========= 
dim makeHtmlFileSetCode:makeHtmlFileSetCode="gb2312"			'����html�ļ�����  

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
            '��ҳ�б� 20170518
            elseIf checkFunValue(action, "GetOnePageList ") = true then
                action = XY_AP_GetOnePageList(action) 
            '��Ա�б� 20170518
            elseIf checkFunValue(action, "GetMemberList ") = true then
                action = XY_AP_GetMemberList(action) 
            '�Զ�����б� 20170603
            elseIf checkFunValue(action, "CustomTableList ") = true then
                action = XY_AP_CustomTableList(action) 
				
				
				
				
            '��÷ָ��б� 20170518
            elseIf checkFunValue(action, "GetSplitList ") = true then
                action = XY_AP_GetSplitList(action) 

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
            '��ģ����ʽ�����ñ��������� ����
            elseIf checkFunValue(action, "ReadTop ") = true then
                action = XY_ReadTop(action) 
            '��ģ����ʽ�����ñ��������� ����
            elseIf checkFunValue(action, "ReadNav ") = true then
                action = XY_ReadNav(action) 
            '��ģ����ʽ�����ñ��������� �ײ�
            elseIf checkFunValue(action, "ReadFoot ") = true then
                action = XY_ReadFoot(action) 


                '------------------- ������ -----------------------
            '��ʾJS��ȾASP/PHP/VB�ȳ���ı༭��
            elseIf checkFunValue(action, "displayEditor ") = true then
                action = displayEditor(action) 
            'Js����վͳ��
            elseIf checkFunValue(action, "JsWebStat ") = true then
                action = XY_JsWebStat(action) 
            '��ñ�ָ��ֵ
            elseIf checkFunValue(action, "NewGetFieldValue ") = true then
                action = XY_AP_NewGetFieldValue(action) 
            '��������ָ��ֵ
            elseIf checkFunValue(action, "SubGetFieldValue ") = true then
                action = XY_AP_SubGetFieldValue(action) 


 
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

			

                '------------------- ��Ա -----------------------
			'��ע 
            elseIf checkFunValue(action, "thisArticleFollow ") = true then
                action = XY_thisArticleFollow(action) 
			
                '------------------- ���� -----------------------
			'��ע 
            elseIf checkFunValue(action, "Request ") = true then
                action = XY_Request(action)  

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
	
	isConcise=false		'�����ʾΪ��
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
    if checkSql(sql) = false then 
		call eerr("Sql", sql) 
	end if
	'��@��jsp��ʾ@��try{
    rs.open sql, conn, 1, 1 
    dropDownMenu = LCase(rParam(action, "DropDownMenu")) 
    focusType = LCase(rParam(action, "FocusType")) 
    isConcise = IIF(LCase(rParam(action, "isConcise")) = "true", false, true) 

    if isConcise = true then 
		c = c & copyStr(" ", 4) & "<li class=left></li>" & vbCrLf 
	end if
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

        if isConcise = true  then
			c = c & copyStr(" ", 8) & "<li class=line></li>" & vbCrLf 
    	end if
	rs.moveNext : next : rs.close 
	'��@��jsp��ʾ@��}catch(Exception e){} 
    if isConcise = true then 
		c = c & copyStr(" ", 8) & "<li class=right></li>" & vbCrLf 
	end if
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
    content = handleRGV(content, "{$glb_columnRootId$}", glb_columnRootId)                  '��Ŀ��Id
    content = handleRGV(content, "{$glb_columnParentId$}", glb_columnParentId)                  '��һ����ĿId
    content = handleRGV(content, "{$glb_columnRootName$}", glb_columnRootName)                  '��Ŀ������
    content = handleRGV(content, "{$glb_columnRootEnName$}", glb_columnRootEnName)                  '��Ŀ��Ӣ������

	
	
    content = handleRGV(content, "{$glb_columnName$}", glb_columnName)              '��Ŀ����
    content = handleRGV(content, "{$glb_columnEnName$}", glb_columnEnName)              '��ĿӢ������
    content = handleRGV(content, "{$glb_columnType$}", glb_columnType)              '��Ŀ����
    content = handleRGV(content, "{$glb_columnENType$}", glb_columnENType)          '��ĿӢ������

    content = handleRGV(content, "{$glb_Table$}", glb_table)                        '��
    content = handleRGV(content, "{$glb_Id$}", glb_id)                              'id
	
	'��Ա
    content = handleRGV(content, "{$member_id$}", getsession("member_id"))          '��Աid
    content = handleRGV(content, "{$member_user$}", getsession("member_user"))          '��Ա�˺� 
	

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
    content = handleRGV(content, "{$glb_articlehits$}", glb_articlehits)            '���µ������

    content = handleRGV(content, "{$glb_upArticle$}", glb_upArticle)                '��һƪ����
    content = handleRGV(content, "{$glb_downArticle$}", glb_downArticle)            '��һƪ����
    content = handleRGV(content, "{$glb_upArticleUrl$}", glb_upArticleUrl)                '��һƪ��������
    content = handleRGV(content, "{$glb_downArticleUrl$}", glb_downArticleUrl)            '��һƪ��������
	
    content = handleRGV(content, "{$glb_aritcleRelatedTags$}", glb_aritcleRelatedTags) '���±�ǩ��
    content = handleRGV(content, "{$glb_aritcleBigImage$}", glb_aritcleBigImage)    '���´�ͼ
    content = handleRGV(content, "{$glb_aritcleSmallImage$}", glb_aritcleSmallImage) '����Сͼ
    content = handleRGV(content, "{$glb_searchKeyWord$}", glb_searchKeyWord)        '��ҳ��ʾ��ַ
	
	
	
    content = handleRGV(content, "{$member_expiredatetime$}", getSession("member_expiredatetime"))        '��Ա��ʱ
    content = handleRGV(content, "{$member_user$}", getSession("member_user"))        '��Ա�˺�
	
	
	
    content = handleRGV(content, "{$pageInfo$}", gbl_PageInfo)        '��ʾ��ҳ��Ϣ
    content = handleRGV(content, "{$detailTitle$}", glb_detailTitle)        '��ҳ��ϸҳ����
    content = handleRGV(content, "{$detailContent$}", glb_bodyContent)        '��ʾ��ϸҳ����
	
	
    content = handleRGV(content, "{$glb_detailTitle$}", glb_detailTitle)        '����ϸ�ڱ���
    content = handleRGV(content, "{$glb_flags$}", glb_flags)        '��
    content = handleRGV(content, "{$glb_smallimage$}", glb_smallimage)        'Сͼ
    content = handleRGV(content, "{$glb_bigimage$}", glb_bigimage)        '��ͼ
    content = handleRGV(content, "{$glb_bannerimage$}", glb_bannerimage)        '��ͼ
	
	
    content = handleRGV(content, "{$glb_author$}", glb_author)        '����
    content = handleRGV(content, "{$glb_sortrank$}", glb_sortrank)        '����
    content = handleRGV(content, "{$glb_aboutcontent$}", glb_aboutcontent)        '����
    content = handleRGV(content, "{$glb_bodyContent$}", glb_bodyContent)        '����
 
 
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
	'��@��jsp��ʾ@��try{
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
		templatesetcode=rs("templatesetcode")											'����ģ����� 
		isMemberVerification=IIF(rs("isMemberVerification")=1,true,false)				'��Ա����ж�
		

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
	'��@��jsp��ʾ@��}catch(Exception e){} 
end sub 



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
	'��@��jsp��ʾ@��try{
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
	gbl_PageInfo=webPageControl(nCount, nPageSize, cstr(sPage), url, pageInfo)
    content = replace(content, "[$pageInfo$]", gbl_PageInfo) 

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
   
	elseif EDITORTYPE = "jsp" then
	
		'��@��jsp��ʾ@��int  nCountPage = getCountPage(nCount, nPageSize);
		'��@��jsp��ʾ@��rs = Conn.executeQuery(sql);			
		'��@��jsp��ʾ@��if(nPage<=1){
		'��@��jsp��ʾ@��	nX=nPageSize;
		'��@��jsp��ʾ@��	if(nX>nCount){
		'��@��jsp��ʾ@��		nX=nCount;
		'��@��jsp��ʾ@��	} 
		'��@��jsp��ʾ@��}else{
		'��@��jsp��ʾ@��	for(int nI2=0;nI2<nPageSize*(nPage-1);nI2++){
		'��@��jsp��ʾ@��		rs.next();
		'��@��jsp��ʾ@��	}
		'��@��jsp��ʾ@��	if(nPage<nCountPage){
		'��@��jsp��ʾ@��		nX=nPageSize;
		'��@��jsp��ʾ@��	}else{
		'��@��jsp��ʾ@��		nX=nCount-nPageSize*(nPage-1);
		'��@��jsp��ʾ@��	}
		'��@��jsp��ʾ@��}  

	else									 
		if nPage<>0 then'��@��.netc����@��'��@��jsp����@��
			nPage = nPage - 1 '��@��.netc����@��'��@��jsp����@��
		end if '��@��.netc����@��'��@��jsp����@��
		sql = "select * from " & db_PREFIX & "" & tableName & " " & addSql & " limit " & nPageSize * nPage & "," & nPageSize '��@��.netc����@��'��@��jsp����@��
		rs.open sql, conn, 1, 1 '��@��.netc����@��'��@��jsp����@��
		'��PHP��ɾ��rs
		nX = rs.recordCount '��@��.netc����@��'��@��jsp����@��
    end if 
    'call echo("sql",sql)
    for i = 1 to nX
        '��PHP��$rs=mysql_fetch_array($rsObj);                                            //��PHP�ã���Ϊ�� asptophpת��������
		'��@��.netc��ʾ@��rs.Read();
		'��@��jsp��ʾ@��rs.next();
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
                elseif inStr(fieldNameList, ",fontcolor,") > 0 then
                    'A������ɫ
                    if rs("fontcolor") <> "" then
                        abcolorStr = "color:" & rs("fontcolor") & ";" 
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



				s = replaceValueParam(s, "i", i)                                                'ѭ�����
				s = replaceValueParam(s, "���", i)                                             'ѭ�����
			
                s = replaceValueParam(s, "url", url) 
                s = replaceValueParam(s, "abcolor", abcolorStr)                                 'A���Ӽ���ɫ��Ӵ�
                s = replaceValueParam(s, "atitle", atitleStr)                                   'A����title
                s = replaceValueParam(s, "anofollow", anofollowStr)                             'A����nofollow
                s = replaceValueParam(s, "atarget", atargetStr)                                 'A���Ӵ򿪷�ʽ


            next 
        end if 
		if EDITORTYPE <> "jsp" then
        idPage = getThisIdPage(db_PREFIX & tableName, rs("id"), 10) 
        '�����ԡ�
        if tableName = "guestbook" then
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=GuestBook&lableTitle=����&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page=" & idPage & "&id=" & rs("id") & "&n=" & getRnd(11) 
        '����Ƹ��
        elseif tableName = "job" then
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=Job&lableTitle=��Ƹ&nPageSize=10&parentid=&searchfield=bodycontent&keyword=&addsql=&page=" & idPage & "&id=" & rs("id") & "&n=" & getRnd(11) 
        '��Ĭ����ʾ���¡�
        else
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=" & idPage & "&parentid=" & rs("parentid") & "&id=" & rs("id") & "&n=" & getRnd(11) 
        end if 
        s = handleDisplayOnlineEditDialog(url, s, "", "div|li|span") 
		end if
        c = c & s 
    rs.moveNext : next : rs.close 
	'��@��jsp��ʾ@��}catch(Exception e){} 
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

    templateHtml=readFile(cfg_webTemplate & "/" & templateName,templatesetcode)
    'templateHtml = getFText(cfg_webTemplate & "/" & templateName) 
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

call loadRun()                                                                  '��@��.netc����@��'��@��jsp����@��
'���ؾ�����
sub loadRun()
	dim c
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
    elseIf getsession("db_PREFIX") <> "" then
        db_PREFIX = getsession("db_PREFIX") 
    end if 
    '������ַ����
    call loadWebConfig() 
    isMakeHtml = false                                                              'Ĭ������HTMLΪ�ر�
    if request("isMakeHtml") = "1" or request("isMakeHtml") = "true" then
        isMakeHtml = true 
    end if 
    templateName = request("templateName")                                          'ģ������


    '�������ݴ���ҳ
	if request("act")= "savedata" then
		call saveData(request("stype")) : response.end()                   '��������
	''վ��ͳ�� | ����IP[653] | ����PV[9865] | ��ǰ����[65]')
	elseif request("act")= "webstat" then 
		call webStat(adminDir & "/Data/Stat/") : response.end()             '��վͳ��
	
	elseif request("act")= "saveSiteMap" then
		isMakeHtml = true : call saveSiteMap() : response.end()         '����sitemap.xml
	
	elseif request("act")= "member" then
		call callMember()	:call eerr("","sdf")											'����Member�ļ�����
	
	elseif request("act")= "handleAction" then
		if request("ishtml") = "1" then
			isMakeHtml = true 
		end if
		call rwend(handleAction(request("content")))                                         '������ 
    '����html
    elseif request("act") = "makehtml" then
        call echo("makehtml", "makehtml") 
        isMakeHtml = true 
        call makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' id='" & request("id") & "' template='"& request("template") &"' ") 
        call writeToFile("index.html", code,makeHtmlFileSetCode) 

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
        call rw(makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' columnType='" & request("columnType") & "' id='" & request("id") & "' npage='" & request("page") & "' template='"& request("template") &"' ")) 
        glb_filePath = replace(glb_url, cfg_webSiteUrl, "") 
        if right(glb_filePath, 1) = "/" then
            glb_filePath = glb_filePath & "index.html" 
        elseIf glb_filePath = "" and glb_columnType = "��ҳ" then
            glb_filePath = "index.html" 
        end if 
        '�ļ���Ϊ��  ���ҿ�������html
        if glb_filePath <> "" and glb_isonhtml = true then
            call createDirFolder(getFileAttr(glb_filePath, "1")) 
            call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
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
        call rw(makeWebHtml("actionType='Search' npage='"& IIF(request("page")="","1",request("page")) &"'  template='"& request("template") &"'")) 
    else
        if LCase(request("issave")) = "1" then
            call makeAllHtml(request("columnType"), request("columnName"), request("columnId")) 
        else
            call checkIDSQL(request("id")) 
			c=makeWebHtml(" action actionType='" & request("act") & "' columnName='" & request("columnName") & "' columnType='" & request("columnType") & "' id='" & request("id") & "' npage='" & request("page") & "'  template='"& request("template") &"'")
            
			'������
			if host()="http://aa/" then
				'call echo("",host())
				c=replace(c,"http://sharembweb.com/asptoLanguage/","")
				c=replace(c,"http://sharembweb.com/AspToLanguage/","")
				c=replace(c,"http://sharembweb.com/toolslist/","/toolslist/")
				c=replace(c,"http://sharembweb.com/Tools/FormattingTools/","/����/")
			end if
			call rw(c) 
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
    dim actionType, npagesize, npage, s, url, addSql, sortSql, sortFieldName, ascOrDesc,thisTable,thisSql,customTemplate
    dim parentid                                                  '׷����20160716 home
	dim isOK,relatedtags,sortFieldValue,sql
	isOK=false			'���
    actionType = rParam(action, "actionType") 
    s = rParam(action, "npage") 
    s = getnumber(s) 
    if s = "" then
        npage = 1 
    else
        npage = CInt(s) 
    end if
	
   	customTemplate = rParam(action, "template") 					'�Զ���ģ��
	'call echo("action",action)
	
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
		'��@��jsp��ʾ@��try{
		thisTable="webcolumn"
		thisSql="Select * from " & db_PREFIX & thisTable & " " & addSql
        rs.open thisSql, conn, 1, 1
        if not rs.EOF then
            glb_columnId = rs("id")
            glb_columnName = rs("columnname")
			glb_columnEnName=rs("columnenname")
            glb_columnType = rs("columntype") 
            glb_bodyContent = rs("bodycontent") 
            glb_detailTitle = glb_columnName 
			glb_memberusercheck=rs("memberusercheck")									'��Ա���
            glb_flags = rs("flags") 
            npagesize = rs("npagesize")                                                     'ÿҳ��ʾ����
            glb_isonhtml = IIF(rs("isonhtml") = 0, false, true)                          '�Ƿ����ɾ�̬��ҳ
            sortSql = " " & rs("sortsql")                                                   '����SQL
			glb_bigimage=rs("bigimage")														'��ͼ
			glb_smallimage=rs("smallimage")													'Сͼ
			glb_bannerimage=rs("bannerimage")												'bannerͼ
			
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
		call handleColumnRoot(glb_columnId) 	'�����Ŀ��ID��������
		
		'��@��jsp��ʾ@��}catch(Exception e){} 
        glb_columnENType = handleColumnType(glb_columnType) 
        glb_url = getColumnUrl(glb_columnName, "name") 

        '�������б�
        if inStr("|��Ʒ|����|��Ƶ|����|����|", "|" & glb_columnType & "|") > 0 then
		
            glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "��Ŀ�б�", "*", npagesize, cstr(npage), "where parentid in("& getColumnIdList(glb_columnId,"addthis") &") " & sortSql)  
			
        '�������б�
        elseIf inStr("|����|", "|" & glb_columnType & "|") > 0 then
            glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "GuestBook", "�����б�", "*", npagesize, cstr(npage), " where isthrough<>0 " & sortSql) 
        '��Ƹ���б�
        elseIf inStr("|��Ƹ|", "|" & glb_columnType & "|") > 0 then 
            glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "Job", "��Ƹ�б�", "*", npagesize, cstr(npage), " where isthrough<>0 " & sortSql) 
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
		thisTable="articledetail"		
		thisSql="Select * from " & db_PREFIX & thisTable & " where id=" & rParam(action, "id")
		'��@��jsp��ʾ@��try{
        rs.open thisSql, conn, 1, 1 
        if not rs.EOF then
			isOK=true
			
            glb_columnId = rs("parentid")  
			
			if glb_columnRootId="" then
				glb_columnRootId=getColumnRootId(rs("id"))					'�������ĿID Ϊ��ʱ������
			end if
			glb_columnParentId=getParentColumnId(rs("parentid"))		'�����һ����ĿID
			glb_columnParentName=getParentColumnName(rs("parentid"))		'�����һ����Ŀ����
			
			
			
            glb_detailTitle = rs("title") 
            glb_flags = rs("flags") 
            glb_smallimage = rs("smallimage") 
            glb_bigimage = rs("bigimage") 
            glb_author = rs("author") 
            glb_sortrank = rs("sortrank")
            glb_aboutcontent = rs("aboutcontent")
            glb_bodyContent = rs("bodycontent")
			 
            glb_isonhtml =IIF(rs("isonhtml") = 0, false, true)   
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


            glb_articleAuthor = rs("author") 
            glb_articleAdddatetime = rs("adddatetime")
			parentid=rs("parentid")
			relatedtags=rs("relatedtags")
			
			glb_relatedTags=relatedtags
            glb_aritcleSmallImage = rs("smallimage") 
            glb_aritcleBigImage = rs("bigimage") 
            glb_articlehits = rs("hits") 
 
            glb_bodyContent = rs("bodycontent") 
			 
            if templateName = "" then
                if trim(rs("templatePath")) <> "" then
                    templateName = rs("templatePath") 
                else
                    templateName = getDateilTemplate(rs("parentid"), "Detail") 
                end if 
            end if  
        end if : rs.close 
		if isOK=true then
			call handleColumnRoot(glb_columnId) 	'�����Ŀ��ID��������
		
            '�Ľ�20160628
            sortFieldName = "id" 
            ascOrDesc = "asc" 
            addSql = trim(getWebColumnSortSql(parentid)) 
            if addSql <> "" then
                sortFieldName = trim(replace(replace(replace(addSql, "order by", ""), " desc", ""), " asc", "")) 
                if instr(addSql, " desc") > 0 then
                    ascOrDesc = "desc" 
                end if 
            end if
			'��������Ӧ�ַ���ֵ
			rs.open thisSql,conn,1,1
			if not rs.eof then
				sortFieldValue=rs(sortFieldName) 
			end if:rs.close
			
            glb_columnName = getColumnName(parentid) 
            glb_upArticle = upArticle(parentid, sortFieldName, sortFieldValue, ascOrDesc,glb_upArticleUrl) 
            glb_downArticle = downArticle(parentid, sortFieldName, sortFieldValue, ascOrDesc,glb_downArticleUrl) 
            glb_aritcleRelatedTags = aritcleRelatedTags(relatedtags)
            conn.execute("update " & db_PREFIX & "articledetail set hits=hits+1 where id=" & glb_id) '���µ����

            '������ϸ�ӿ���
            if request("gl") = "edit" then
                glb_bodyContent = "<span>" & glb_bodyContent & "</span>" 
            end if 
            url = WEB_ADMINURL & "?act=addEditHandle&actionType=ArticleDetail&lableTitle=������Ϣ&nPageSize=10&page=&parentid=&id=" & rParam(action, "id") & "&n=" & getRnd(11) 
            glb_bodyContent = handleDisplayOnlineEditDialog(url, glb_bodyContent, "", "span") 
			
		end if
		'��@��jsp��ʾ@��}catch(Exception e){} 
	
    '��ҳ
    elseIf actionType = "onepage" then
		'��@��jsp��ʾ@��try{
		
		thisTable="onepage"		
		thisSql="Select * from " & db_PREFIX & thisTable & " where id=" & rParam(action, "id") 
        rs.open thisSql, conn, 1, 1 
        if not rs.EOF then
            glb_detailTitle = rs("title") 
            glb_isonhtml = IIF(rs("isonhtml") = 0, false, true)                   '�Ƿ����ɾ�̬��ҳ
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
                end if 
            end if 

        end if : rs.close 
		'��@��jsp��ʾ@��}catch(Exception e){}

    '����
    elseIf actionType = "Search" then
		templateName="Main_Model.html"
		if request("template") <>"" then
			templateName=request("template") & ".html"
		end if 
        searchKeyWordName = request("keywordname") 
		glb_searchFieldName= request("searchfieldname")		'�����ֶ�����
        parentid = request("parentid") 
        if searchKeyWordName = "" then
            searchKeyWordName = "wd" 
        end if
		if glb_searchFieldName="" then
			glb_searchFieldName="title"
		end if
		'call echo("searchKeyWordName",searchKeyWordName)
        glb_searchKeyWord = replace(request(searchKeyWordName), "<", "&lt;") 
        addSql = "" 
        if parentid <> "" then
            addSql = " where parentid=" & parentid
		elseif request("subparentid")<>"" then
			glb_columnRootId=trim(request("subparentid"))		'��¼�£��������ã���
			addSql = " where parentid in(" & getColumnIdList(request("subparentid"),"addthis") & ")"
			'call eerr("addSql",addSql)
        end if 
        addSql = getWhereAnd(addSql, " where "& glb_searchFieldName &" like '%" & glb_searchKeyWord & "%'")
		'call echo("addsql",addsql)
        npagesize = 20  
        glb_bodyContent = getDetailList(action, defaultListTemplate(glb_columnType, glb_columnName), "ArticleDetail", "��վ��Ŀ", "*", npagesize, cstr(npage), addSql) 
        positionEndStr = " >> �������ݡ�" & glb_searchKeyWord & "��" 
    '���صȴ�
    elseIf actionType = "loading" then
        call rwend("ҳ�����ڼ����С�����") 
	'�� index.asp?act=table_HuoYuan&id=1
    elseIf left(actionType,6) = "table_" then
		thisTable=mid(actionType,7)
		thisSql="Select * from " & db_PREFIX & thisTable & " where id=" & rParam(action, "id") 
 
    end if 
    '20170603׷��
	if customTemplate<>"" then
		templateName=customTemplate & ".html"
	'ģ��Ϊ�գ�����Ĭ����ҳģ��
    elseif templateName = "" then
        templateName = "Index_Model.html"                                               'Ĭ��ģ��
    end if 
	'�����Ա��֤20170522
	templateName=getMemberVerificationTemplateName(templateName) 
	
    '��⵱ǰ·���Ƿ���ģ��
    if inStr(templateName, "/") = false then
        templateName = cfg_webTemplate & "/" & templateName 
    end if 
    'call echo("templateName",templateName)
    if checkFile(templateName) = false then
        call eerr("δ�ҵ�ģ���ļ�", templateName) 
    end if
	
	code=readFile(templateName,templatesetcode)			'ģ����Զ������ 
    'code = getftext(templateName) 

    code = handleAction(code)                                                       '������
    'code = thisPosition(code)                                                       'λ��
    code = replaceGlobleVariable(code)                                              '�滻ȫ�ֱ�ǩ
    code = handleAction(code)                                                       '������    '����һ�Σ��������������ﶯ��

    'call die(code)
    code = handleAction(code)                                                       '������
    code = handleAction(code)                                                       '������
    'code = thisPosition(code)                                                       'λ��
    code = replaceGlobleVariable(code)                                              '�滻ȫ�ֱ�ǩ
    code = delTemplateMyNote(code)                                                  'ɾ����������
    code = handleAction(code)                                                       '������
	
	if thisTable <>"" then
		'call eerr(thisTable,thisSql)
		code=handleReplaceTableFieldList(code,thisTable,thisSql,"this_glb_","")
	end if
 
'call echo("templateName",templateName)
'call eerr("customTemplate",customTemplate)

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
        if inStr(code, "<head>") > 0 then
            if inStr(lcase(code), "jquery.min.js") = false then
                code = replace(code, "<head>", "<head><script src=""/Jquery/jquery.Min.js""></script>") 
            end if 
		end if
		if inStr(code, "</head>") > 0 then
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
    dim templateName,tempS
    templateName = "Main_Model.html" 
	'��@��jsp��ʾ@��try{
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
        elseIf rsx("columntype") = "��Ƹ" then
            '����ϸ��ҳ
            if checkFile(cfg_webTemplate & "/Job_" & templateType & ".html") = true then
                templateName = "Job_" & templateType & ".html" 
            end if 

        elseIf rsx("columntype") = "��Ƶ" then
            '��Ƶϸ��ҳ
            if checkFile(cfg_webTemplate & "/Video_" & templateType & ".html") = true then
                templateName = "Video_" & templateType & ".html" 
            end if 
        elseIf rsx("columntype") = "����" then
            '��Ƶϸ��ҳ
            if checkFile(cfg_webTemplate & "/GuestBook_" & templateType & ".html") = true then
                templateName = "GuestBook_" & templateType & ".html" 
            end if 
        
		elseIf rsx("columntype") = "�ı�" then
            '��Ƶϸ��ҳ
            if checkFile(cfg_webTemplate & "/Page_" & templateType & ".html") = true then
                templateName = "Page_" & templateType & ".html" 
            end if 
        end if  
		if rsx("columntype") <> "�ı�" and (templateType="List" Or templateType="Detail") and templateName="Main_Model.html"  then			 
			tempS="Default_" & templateType & ".html" 
			if checkFile(cfg_webTemplate & "/" & tempS) = true then
				templateName =tempS
			end if
			'call echo("templateName",templateName) 
		end if
		'call echo("templateType",templateType)
    end if : rsx.close 
	'��@��jsp��ʾ@��}catch(Exception e){}
    getDateilTemplate = templateName 

end function 


'����ȫ��htmlҳ��
sub makeAllHtml(columnType, columnName, columnId)
    dim action, s, i, nPageSize, nCountSize, nPage, addSql, url, articleSql 
    call handlePower("����ȫ��HTMLҳ��")                                            '����Ȩ�޴���
    call writeSystemLog("", "����ȫ��HTMLҳ��")                                     'ϵͳ��־

    isMakeHtml = true 
    '��Ŀ
    call echo("��Ŀ", columnName) 
    if columnType <> "" then
        addSql = "where columnType='" & columnType & "'" 
    end if 
    if columnName <> "" then
        addSql = getWhereAnd(addSql, "where columnName='" & columnName & "'") 
    end if 
    if columnId <> "" then
        addSql = getWhereAnd(addSql, "where id in(" & columnId & ")") 
    end if 
	'��@��jsp��ʾ@��try{
    rss.open "select * from " & db_PREFIX & "webcolumn " & addSql & " order by sortrank asc", conn, 1, 1 
    while not rss.EOF
        glb_columnName = "" 
        '��������html
        if rss("isonhtml") <>0 then
            if inStr("|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����|", "|" & rss("columntype") & "|") > 0 then
                if rss("columntype") = "����" then
                    nCountSize = getRecordCount(db_PREFIX & "guestbook", "")                        '��¼��
                elseif rss("columntype") = "��Ƹ" then
                    nCountSize = getRecordCount(db_PREFIX & "job", "")                        '��¼��
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
                        call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
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
                    call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
                end if 
                doevents() 
                templateName = "" 
            end if 
            conn.execute("update " & db_PREFIX & "WebColumn set ishtml=true where id=" & rss("id")) '���µ���Ϊ����״̬
        end if 
    rss.moveNext : wend : rss.close 
	'��@��jsp��ʾ@��}catch(Exception e){} 

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
		'��@��jsp��ʾ@��try{
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
            if glb_filePath <> "" and rss("isonhtml") <>0 then
                call createDirFolder(getFileAttr(glb_filePath, "1")) 
                call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
                conn.execute("update " & db_PREFIX & "ArticleDetail set ishtml=true where id=" & rss("id")) '��������Ϊ����״̬
            end if 
            templateName = ""                                                               '���ģ���ļ�����
        rss.moveNext : wend : rss.close 
		'��@��jsp��ʾ@��}catch(Exception e){} 
    end if 

    if addSql = "" then
        '��ҳ
        call echo("��ҳ", "") 
		'��@��jsp��ʾ@��try{
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
            if glb_filePath <> "" and rss("isonhtml") <>0 then
                call createDirFolder(getFileAttr(glb_filePath, "1")) 
                call writeToFile(glb_filePath, code,makeHtmlFileSetCode)  
                conn.execute("update " & db_PREFIX & "onepage set ishtml=true where id=" & rss("id")) '���µ�ҳΪ����״̬
            end if 
            templateName = ""                                                               '���ģ���ļ�����
        rss.moveNext : wend : rss.close 
		'��@��jsp��ʾ@��}catch(Exception e){} 

    end if 
end sub 

'����html����վ
sub copyHtmlToWeb()
    dim webDir, toWebDir, toFilePath, filePath, fileName, fileList, splStr, content, s, s1, c, webImages, webCss, webJs, splJs 
    dim webFolderName, jsFileList, setFileCode, nErrLevel, jsFilePath, url,isSetFileNameType,isPinYin

    setFileCode = request("setcode")                                                '�����ļ��������
	if setFileCode="" then
		setFileCode="gb2312"	'Ĭ��
	end if
	
	isSetFileNameType=request("isSetFileNameType")									'���ɺ���ļ����� �Ƿ��-ת_
	isPinYin=request("isPinYin")													'�ļ���תƴ��

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
            content=readFile(filePath,templatesetcode)
			'content = getftext(filePath)
			 
            content = replace(content, cfg_webImages, "../Images") 

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
	'��@��jsp��ʾ@��try{
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
			if isSetFileNameType="1" then
            	fileName = replace(fileName, "-", "_") 
			end if
			if isPinYin="1" then
				fileName=pinYin2(fileName)
			end if
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
			if isSetFileNameType="1" then
            	fileName = replace(fileName, "-", "_") 
			end if
			if isPinYin="1" then
				fileName=pinYin2(fileName)
			end if
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
			if isSetFileNameType="1" then
            	fileName = replace(fileName, "-", "_") 
			end if
			if isPinYin="1" then
				fileName=pinYin2(fileName)
			end if
            toFilePath = webDir & fileName 
            call copyfile(glb_filePath, toFilePath) 
            call echo("��ҳ" & rss("title"), glb_filePath) 
        end if 
    rss.moveNext : wend : rss.close 
	'��@��jsp��ʾ@��}catch(Exception e){}
	
	 
    '��������html�ļ��б�
    'call echo(cfg_webSiteUrl,cfg_webTemplate)
    'call rwend(fileList)
    dim sourceUrl, replaceUrl 
    splStr = split(fileList, vbCrLf) 
    for each filePath in splStr
        if filePath <> "" then
            filePath =  replace(filePath, "/", "_") 
			if isSetFileNameType="1" then
				filePath = replace(filePath, "-", "_") 
			end if
			if isPinYin="1" then
				filePath=pinYin2(filePath)			'ע��
			end if
            filePath = webDir & filePath
            call echo("filePath", filePath) 
            content = getftext(filePath) 
            for each s in splStr
                s1 = s 
                if right(s1, 11) = "/index.html" then
                    s1 = left(s1, len(s1) - 11) & "/" 
                end if 
                sourceUrl = cfg_webSiteUrl & s1 
				s=replace(s, "/", "_") 
				if isSetFileNameType="1" then
					s = replace(s, "-", "_") 
				end if
				if isPinYin="1" then
					s=pinYin2(s)
				end if
                replaceUrl = cfg_webSiteUrl & s
				
                'Call echo(sourceUrl, replaceUrl)                             '����  ���������ʾ20160613
                content = replace(content, sourceUrl, replaceUrl) 
            next 
            content = replace(content, cfg_webSiteUrl, "")                                  'ɾ����ַ
            content = replace(content, cfg_webTemplate & "/", "")                           'ɾ��ģ��·�� ��
			
            content = replace(replace(content, "\Inc/YZM_7.asp", "Images/YZM_7.jpg"), "/Inc/YZM_7.asp", "Images/YZM_7.jpg")  '�滻��֤��20160916
			content = replace(content, "/inc/yzm_7.asp", "Images/YZM_7.jpg")
			
			
            'content=nullLinkAddDefaultName(content)
            for each s in splJs
                if s <> "" then
                    fileName = getFileName(s) 
                    content = replace(content, "Images/" & fileName, "Js/" & fileName) 
                end if 
            next 
            if inStr(content, "/Jquery/Jquery.Min.js") > 0 then
                content = replace(content, "/Jquery/Jquery.Min.js", "Js/Jquery.Min.js") 
                call copyfile("/Jquery/Jquery.Min.js", webJs & "/Jquery.Min.js") 
            end if 
            content = replace(content, "<a href="""" ", "<a href=""index.html"" ")    '����ҳ��index.html

            call writeToFile(filePath, content,makeHtmlFileSetCode)  
			
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
            content = regExp_Replace(content, "Images/" & jsFileName, "Js/" & jsFileName) 
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
    call echo("���", "<a href='" & toWebDir & "' target='_blank'>" & toWebDir & "</a>") 
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
    call writeToFile("htmlweb/1.txt", c,makeHtmlFileSetCode) 
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
		'��@��jsp��ʾ@��try{
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
                call echo("��Ŀ2", "<a href=""" & url & """ target='_blank'>" & url & "</a>") 
            end if 
        rsx.moveNext : wend : rsx.close 
		'��@��jsp��ʾ@��}catch(Exception e){} 

        '����
		sql="select * from " & db_PREFIX & "articledetail  where isonhtml<>0 order by sortrank asc"
		'��@��jsp��ʾ@��try{
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
		'��@��jsp��ʾ@��}catch(Exception e){} 

        '��ҳ
		'��@��jsp��ʾ@��try{
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
		'��@��jsp��ʾ@��}catch(Exception e){} 
        c = c & vbTab & "</urlset>" & vbCrLf 
        call loadWebConfig() 
        call createFile("sitemap.xml", c) 
        call echo("����sitemap.xml�ļ��ɹ�", "<a href='/sitemap.xml' target='_blank'>���Ԥ��sitemap.xml</a>") 


        '�ж��Ƿ�����sitemap.html
        if request("issitemaphtml") = "1" then
            c = "" 
            '�ڶ���
            '��Ŀ
			'��@��jsp��ʾ@��try{
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
                    if rsx("isonhtml") <>0 then
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
                            if rss("isonhtml") <>0 then
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
			'��@��jsp��ʾ@��}catch(Exception e){} 

            '����
            c = c & "<li style=""width:20%;""><a href=""javascript:;"">�����б�</a>" & vbCrLf & "<ul>" & vbCrLf 
			'��@��jsp��ʾ@��try{
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
                    if rsx("isonhtml") <>0 then
                        s = "<a href=""" & url & """>" & rsx("title") & "</a>" 
                    else
                        s = "<span>" & rsx("title") & "</span>" 
                    end if 

                    c = c & "<li style=""width:20%;"">" & s & "</li>" & vbCrLf                'target=""_blank""  ȥ��
                end if 
            rsx.moveNext : wend : rsx.close 
			'��@��jsp��ʾ@��}catch(Exception e){} 

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
 
'��û�Ա��֤��ģ�� 20170422
function getMemberVerificationTemplateName(templateName) 
	templateName=getCheckMemberLoginTemplate(templateName)
	getMemberVerificationTemplateName=templateName
end function
%>   