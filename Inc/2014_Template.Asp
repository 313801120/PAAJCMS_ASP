<%
'************************************************************
'���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2016-08-04
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'ɾ�����ֶ����ǩ <R#��������BlockName��վ���� start#>  <R#��������BlockName��վ���� end#>
'��ģ�崦��



'��ģ������
function XY_ReadTemplateModule(action)
    dim moduleId, filePath, c, i 
    dim sourceList                                                                  'Դ�����б� 20150109
    dim replaceList                                                                 '�滻�����б�
    dim splSource, splReplace, sourceStr, replaceStr 
    filePath = rParam(action, "File") 
    moduleId = rParam(action, "ModuleId") 
    sourceList = rParam(action, "SourceList") 
    replaceList = rParam(action, "ReplaceList")  
    'Call Echo(SourceList,ReplaceList)

    if moduleId = "" then moduleId = rParam(action, "ModuleName")                   '�ÿ�����
    filePath = filePath & ".html" 
    'Call Echo("FilePath",FilePath)
    'Call Echo("ModuleId",ModuleId)
    c = readTemplateModuleStr(filePath, "", moduleId) 
    '���滻��20160331
    if sourceList <> "" and replaceList <> "" then
        splSource = split(sourceList, "[Array]") 
        splReplace = split(replaceList, "[Array]") 
        for i = 0 to uBound(splSource)
            sourceStr = splSource(i) 
            replaceStr = splReplace(i) 
            c = replace(c, sourceStr, replaceStr) 
        next 
    end if 
    XY_ReadTemplateModule = c 
end function 

'��ģ������
function readTemplateModuleStr(filePath, defaultContent, moduleId)
    dim startStr, endStr, content 
	readTemplateModuleStr="" 
    startStr = "<!--#Module " & moduleId & " start#-->" 
    endStr = "<!--#Module " & moduleId & " end#-->" 
    'FilePath = ReplaceGlobleLable(FilePath)                '�滻ȫ����ǩ        '������2014 12 11

    '�ļ������ڣ���׷��ģ��·�� 20150616 ��VB��������
    if checkFile(filePath) = false then
        filePath = cfg_webTemplate & "/" & filePath 
    end if 

    'û��handleRGV���������
    'filePath = handleRGV(filePath, "[$ģ��Ŀ¼$]", "Module/")                              'Module

    if defaultContent <> "" then
        content = defaultContent 
    elseIf checkFile(filePath) = true then
        content = getFText(filePath) 
    else
        content = code                                                                  'Ĭ��������ָ������
    end if 

    if inStr(content, startStr) = false then
        startStr = "<!--#Module " & moduleId & " Start#-->" 
    end if 
    if inStr(content, endStr) = false then
        endStr = "<!--#Module " & moduleId & " End#-->" 
    end if 
    if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
        readTemplateModuleStr = strCut(content, startStr, endStr, 2) 
    else
        readTemplateModuleStr = "ģ��[" & moduleId & "]������,·��=" & filePath 
        call die(readTemplateModuleStr & content) 
    end if 
end function 
'��ģ���Ӧ����
function findModuleStr(content, valueStr)
    dim startStr, endStr, yuanStr, replaceStr, i, block, blockFile, action 
    for i = 1 to 9
        startStr = "[$�������� " : endStr = "$]" 
        if inStr(valueStr, startStr) > 0 and inStr(valueStr, endStr) > 0 then
            action = strCut(valueStr, startStr, endStr, 2) 
            block = rParam(action, "Block") 
            blockFile = rParam(action, "File") 
            if inStr(vbCrLf & readBlockList & vbCrLf, vbCrLf & block & vbCrLf) = false then
                readBlockList = readBlockList & block & vbCrLf 
            end if 
            '���ļ����� ���������
            if blockFile <> "" then
                content = getFText(blockFile) 
            end if 
            yuanStr = startStr & action & endStr 
            replaceStr = "" 

            startStr = "<R#��������" & block & " start#>" : endStr = "<R#��������" & block & " end#>" 
            if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
                replaceStr = strCut(content, startStr, endStr, 2) 
            else
                startStr = "<!--#��������" & block : endStr = "#-->" 
                if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
                    replaceStr = strCut(content, startStr, endStr, 2) 
                end if 
            end if 
            'Call Eerr(YuanStr,ReplaceStr)
            valueStr = replace(valueStr, yuanStr, replaceStr) 
        'Call Echo("ValueStr",ValueStr)
        else
            'û��ģ��Ҫ������ ���˳�
            exit for 
        end if 
    next 
    findModuleStr = valueStr 
end function 

'����Leftģ����ʽ        �������ַ� ' ���ظ��ƻ���������������� \|*|\ ���洦��
function XY_ReadColumeSetTitle(action)
    dim startStr, endStr, style, title, valueStr, moreClass, moreUrl, moreStr, aStr, c 
    action = handleInModule(action, "start") 
    style = rParam(action, "style") 
    title = rParam(action, "Title") 
    'Call Echo("ContentHeight",ContentHeight)
    'ValueStr = RParam(Action,"value")
    '����ģ��������
    valueStr = moduleFindContent(action, "value") 
    'Call Eerr("ValueStr",ValueStr)
    valueStr = findModuleStr(code, valueStr)                                        '��ģ���Ӧ����

    moreClass = rParam(action, "MoreClass") 
    moreUrl = phpTrim(rParam(action, "MoreUrl")) 
    moreStr = rParam(action, "MoreStr") 
    valueStr = handleInModule(valueStr, "end") 
    c = readColumeSetTitle(action, style, title, valueStr) 

    if moreClass = "" then moreClass = "more"                                       'More����Ϊ�� ����Ĭ�ϴ���
    'If MoreUrl="" Then MoreUrl="#"                    'More������ַΪ�� ����Ĭ��#����
    'More������ʽ����Ϊ�գ���Ϊû����ʽ���Ͳ�����More�������
    if moreUrl <> "" and moreStr <> "" then
        'AStr = "<a href='"& MoreUrl &"' class='"& MoreClass &"'>"& MoreStr &"</a>"
        aStr = "<a " & AHref(moreUrl, title, "") & " class='" & moreClass & "'>" & moreStr & "</a>" 
        c = replace(c, "<!--#AMore#-->", aStr) 
    end if 
    XY_ReadColumeSetTitle = c 
end function 

'����Ŀ��������������ֵ
function readColumeSetTitle(action, styleID, columeTitle, columeContent)
    dim titleWidth                                                                  '�������
    dim titleHeight                                                                 '����߶�
    dim contentHeight                                                               '���ݸ߶�
    dim contentWidth                                                                '���ݿ���
    dim contentCss 

    titleWidth = rParam(action, "TitleWidth")                                       '��ñ���߶�    ��Ӧ��20150715
    titleHeight = rParam(action, "TitleHeight")                                     '��ñ������
    contentWidth = rParam(action, "ContentWidth")                                   '������ݿ���
    contentHeight = rParam(action, "ContentHeight")                                 '������ݸ߶�

    '�����
    titleWidth = trim(titleWidth) 
    '�Զ���px��λ�����ӻ���Ч�� 20150115
    if right(titleHeight, 1) <> "%" and right(titleHeight, 2) <> "px" and titleHeight <> "" and titleHeight <> "auto" then
        titleHeight = titleHeight & "px" 
    end if 
    if right(titleWidth, 1) <> "%" and right(titleWidth, 2) <> "px" and titleWidth <> "" and titleWidth <> "auto" then
        titleWidth = titleWidth & "px" 
    end if 
    '���ݸ�
    contentHeight = trim(contentHeight) 
    '�Զ���px��λ�����ӻ���Ч�� 20150115
    if right(contentHeight, 1) <> "%" and right(contentHeight, 2) <> "px" and contentHeight <> "" and contentHeight <> "auto" then
        contentHeight = contentHeight & "px" 
    end if 
    '���ݿ�
    contentWidth = trim(contentWidth) 
    '�Զ���px��λ�����ӻ���Ч�� 20150115
    if right(contentWidth, 1) <> "%" and right(contentWidth, 2) <> "px" and contentWidth <> "" and contentWidth <> "auto" then
        contentWidth = contentWidth & "px" 
    end if 

    if contentHeight <> "" then
        contentCss = "height:" & contentHeight & ";" 
    end if 
    if contentWidth <> "" then
        contentCss = contentCss & "width:" & contentWidth & ";" 
    end if 

    dim content 
    content = readColumn(styleID) 
    '�����
    if titleWidth <> "" then
        content = replace(content, "<div class=""tvalue"">", "<div class=""tvalue"" style='width:" & titleWidth & ";'>") 
    end if 
    '���ݸ�
    if contentCss <> "" then
        content = replace(content, "<div class=""ccontent"">", "<div class=""ccontent"" style='" & contentCss & "'>") 
    end if 
    'call echo(ContentWidth,ContentCss)

    content = replace(content, "��Ŀ����", columeTitle) 
    content = replace(content, "��Ŀ����", columeContent) 
    readColumeSetTitle = content 
end function 

'����Ŀģ��
function readColumn(sID)
    dim templateFilePath, startStr, endStr, s  
    templateFilePath = cfg_webTemplate & "\Template_Left.html" 
    startStr = "/*columnlist" & sID & "Start*/" 
    endStr = "/*columnlist" & sID & "End*/" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    if s = "[$NO$]" then
        s = "Left��ʽID[" & sID & "]������" 
    end if 
    readColumn = s 
end function 


'��ģ���ز�
function readTemplateSource(sID)
    dim templateFilePath, startStr, endStr, s 
    templateFilePath = cfg_webTemplate & "\TemplateSource.html" 
    startStr = "<!--#sourceHtml" & sID & "Start#-->" 
    endStr = "<!--#sourceHtml" & sID & "End#-->" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    if s = "[$NO$]" then
        s = "ģ����ԴID[" & sID & "]������" 
    end if 
    readTemplateSource = s 
end function 



'��ģ���ļ���ĳģ��
function readTemplateFileModular(templateFilePath, startStr, endStr)
    dim content 
    readTemplateFileModular = "" 
    content = getFText(templateFilePath) 
    if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
        readTemplateFileModular = strCut(content, startStr, endStr, 2) 
    else
        readTemplateFileModular = "[$NO$]" 
    end if 
end function 

'���ļ�ģ���ز�
function readTemplateFileSource(templateFilePath, sID)
    dim startStr, endStr, s, c 
	readTemplateFileSource=""
    startStr = "<!--#sourceHtml" & replace(sID, ".html", "") & "Start#-->" 
    endStr = "<!--#sourceHtml" & replace(sID, ".html", "") & "End#-->" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    if s = "[$NO$]" then
        '��һ����ȡ���ı��������б�(20150815)
        c = getStrCut(pubCode, startStr, endStr, 2) 
        if c <> "" then
            readTemplateFileSource = c 
            'call rwend(c)
            exit function 
        end if 
        c = getftext(templateFilePath) 
        '���� <!--#TemplateSplitStart#-->  �ͷ��ص�ǰȫ������
        if inStr(c, "<!--#DialogStart#-->") > 0 then
            readTemplateFileSource = c 
            exit function 
        end if 

        s = "ģ����ԴID[" & sID & "]������,·��TemplateFilePath=" & handlePath(templateFilePath) 
    end if 
    readTemplateFileSource = s 
end function 
'�����ļ�չʾ�б���Դ
function readArticleListStyleSource(sID)
    dim filePath 
    filePath = cfg_webImages & "\����չʾ��ʽ\" & sID 
    if checkFile(filePath) = false then
        filePath = cfg_webTemplate & "\Resources\" & sID 
    end if 
    'call echo(checkfile(filePath),filePath)
    readArticleListStyleSource = readTemplateFileSource(filePath, sID) 

end function 
'�����ļ���Ϣ�б���Դ
function readArticleInfoStyleSource(sID)
    dim filePath 
    filePath = cfg_webImages & "\������Ϣչʾ��ʽ\" & sID 
    if checkFile(filePath) = false then
        filePath = cfg_webTemplate & "\Resources\" & sID 
    end if 
    readArticleInfoStyleSource = readTemplateFileSource(filePath, sID) 
end function 


%>     

