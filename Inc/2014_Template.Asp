<%
'************************************************************
'作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-08-05
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'删除这种多余标签 <R#读出内容BlockName网站公告 start#>  <R#读出内容BlockName网站公告 end#>
'对模板处理



'读模块内容
function XY_ReadTemplateModule(action)
    dim moduleId, filePath, c, i 
    dim sourceList                                                                  '源内容列表 20150109
    dim replaceList                                                                 '替换内容列表
    dim splSource, splReplace, sourceStr, replaceStr 
    filePath = rParam(action, "File") 
    moduleId = rParam(action, "ModuleId") 
    sourceList = rParam(action, "SourceList") 
    replaceList = rParam(action, "ReplaceList")  
    'Call Echo(SourceList,ReplaceList)

    if moduleId = "" then moduleId = rParam(action, "ModuleName")                   '用块名称
    filePath = filePath & ".html" 
    'Call Echo("FilePath",FilePath)
    'Call Echo("ModuleId",ModuleId)
    c = readTemplateModuleStr(filePath, "", moduleId) 
    '加替换于20160331
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

'读模块内容
function readTemplateModuleStr(filePath, defaultContent, moduleId)
    dim startStr, endStr, content 
	readTemplateModuleStr="" 
    startStr = "<!--#Module " & moduleId & " start#-->" 
    endStr = "<!--#Module " & moduleId & " end#-->" 
    'FilePath = ReplaceGlobleLable(FilePath)                '替换全部标签        '添加于2014 12 11

    '文件不存在，则追加模板路径 20150616 给VB软件里用
    if checkFile(filePath) = false then
        filePath = cfg_webTemplate & "/" & filePath 
    end if 

    '没有handleRGV这个函数了
    'filePath = handleRGV(filePath, "[$模块目录$]", "Module/")                              'Module

    if defaultContent <> "" then
        content = defaultContent 
    elseIf checkFile(filePath) = true then
        content = getFText(filePath) 
    else
        content = code                                                                  '默认用内容指定内容
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
        readTemplateModuleStr = "模块[" & moduleId & "]不存在,路径=" & filePath 
        call die(readTemplateModuleStr & content) 
    end if 
end function 
'找模块对应内容
function findModuleStr(content, valueStr)
    dim startStr, endStr, yuanStr, replaceStr, i, block, blockFile, action 
    for i = 1 to 9
        startStr = "[$读出内容 " : endStr = "$]" 
        if inStr(valueStr, startStr) > 0 and inStr(valueStr, endStr) > 0 then
            action = strCut(valueStr, startStr, endStr, 2) 
            block = rParam(action, "Block") 
            blockFile = rParam(action, "File") 
            if inStr(vbCrLf & readBlockList & vbCrLf, vbCrLf & block & vbCrLf) = false then
                readBlockList = readBlockList & block & vbCrLf 
            end if 
            '块文件存在 则读出内容
            if blockFile <> "" then
                content = getFText(blockFile) 
            end if 
            yuanStr = startStr & action & endStr 
            replaceStr = "" 

            startStr = "<R#读出内容" & block & " start#>" : endStr = "<R#读出内容" & block & " end#>" 
            if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
                replaceStr = strCut(content, startStr, endStr, 2) 
            else
                startStr = "<!--#读出内容" & block : endStr = "#-->" 
                if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
                    replaceStr = strCut(content, startStr, endStr, 2) 
                end if 
            end if 
            'Call Eerr(YuanStr,ReplaceStr)
            valueStr = replace(valueStr, yuanStr, replaceStr) 
        'Call Echo("ValueStr",ValueStr)
        else
            '没有模块要处理了 则退出
            exit for 
        end if 
    next 
    findModuleStr = valueStr 
end function 

'读出Left模板样式        这里面字符 ' 来回复制会出错，所以我们用 \|*|\ 代替处理
function XY_ReadColumeSetTitle(action)
    dim startStr, endStr, style, title, valueStr, moreClass, moreUrl, moreStr, aStr, c 
    action = handleInModule(action, "start") 
    style = rParam(action, "style") 
    title = rParam(action, "Title") 
    'Call Echo("ContentHeight",ContentHeight)
    'ValueStr = RParam(Action,"value")
    '根据模块找内容
    valueStr = moduleFindContent(action, "value") 
    'Call Eerr("ValueStr",ValueStr)
    valueStr = findModuleStr(code, valueStr)                                        '找模块对应内容

    moreClass = rParam(action, "MoreClass") 
    moreUrl = phpTrim(rParam(action, "MoreUrl")) 
    moreStr = rParam(action, "MoreStr") 
    valueStr = handleInModule(valueStr, "end") 
    c = readColumeSetTitle(action, style, title, valueStr) 

    if moreClass = "" then moreClass = "more"                                       'More链接为空 则用默认代替
    'If MoreUrl="" Then MoreUrl="#"                    'More链接网址为空 则用默认#代替
    'More链接样式不能为空，因为没有样式它就不能让More在最近边
    if moreUrl <> "" and moreStr <> "" then
        'AStr = "<a href='"& MoreUrl &"' class='"& MoreClass &"'>"& MoreStr &"</a>"
        aStr = "<a " & AHref(moreUrl, title, "") & " class='" & moreClass & "'>" & moreStr & "</a>" 
        c = replace(c, "<!--#AMore#-->", aStr) 
    end if 
    XY_ReadColumeSetTitle = c 
end function 

'读栏目并赋标题与内容值
function readColumeSetTitle(action, styleID, columeTitle, columeContent)
    dim titleWidth                                                                  '标题宽度
    dim titleHeight                                                                 '标题高度
    dim contentHeight                                                               '内容高度
    dim contentWidth                                                                '内容宽度
    dim contentCss 

    titleWidth = rParam(action, "TitleWidth")                                       '获得标题高度    待应用20150715
    titleHeight = rParam(action, "TitleHeight")                                     '获得标题宽度
    contentWidth = rParam(action, "ContentWidth")                                   '获得内容宽度
    contentHeight = rParam(action, "ContentHeight")                                 '获得内容高度

    '标题宽
    titleWidth = trim(titleWidth) 
    '自动加px单位，不加会无效果 20150115
    if right(titleHeight, 1) <> "%" and right(titleHeight, 2) <> "px" and titleHeight <> "" and titleHeight <> "auto" then
        titleHeight = titleHeight & "px" 
    end if 
    if right(titleWidth, 1) <> "%" and right(titleWidth, 2) <> "px" and titleWidth <> "" and titleWidth <> "auto" then
        titleWidth = titleWidth & "px" 
    end if 
    '内容高
    contentHeight = trim(contentHeight) 
    '自动加px单位，不加会无效果 20150115
    if right(contentHeight, 1) <> "%" and right(contentHeight, 2) <> "px" and contentHeight <> "" and contentHeight <> "auto" then
        contentHeight = contentHeight & "px" 
    end if 
    '内容宽
    contentWidth = trim(contentWidth) 
    '自动加px单位，不加会无效果 20150115
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
    '标题宽
    if titleWidth <> "" then
        content = replace(content, "<div class=""tvalue"">", "<div class=""tvalue"" style='width:" & titleWidth & ";'>") 
    end if 
    '内容高
    if contentCss <> "" then
        content = replace(content, "<div class=""ccontent"">", "<div class=""ccontent"" style='" & contentCss & "'>") 
    end if 
    'call echo(ContentWidth,ContentCss)

    content = replace(content, "栏目标题", columeTitle) 
    content = replace(content, "栏目内容", columeContent) 
    readColumeSetTitle = content 
end function 

'读栏目模块
function readColumn(sID)
    dim templateFilePath, startStr, endStr, s  
    templateFilePath = cfg_webTemplate & "\Template_Left.html" 
    startStr = "/*columnlist" & sID & "Start*/" 
    endStr = "/*columnlist" & sID & "End*/" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    if s = "[$NO$]" then
        s = "Left样式ID[" & sID & "]不存在" 
    end if 
    readColumn = s 
end function 


'读模板素材
function readTemplateSource(sID)
    dim templateFilePath, startStr, endStr, s 
    templateFilePath = cfg_webTemplate & "\TemplateSource.html" 
    startStr = "<!--#sourceHtml" & sID & "Start#-->" 
    endStr = "<!--#sourceHtml" & sID & "End#-->" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    if s = "[$NO$]" then
        s = "模板资源ID[" & sID & "]不存在" 
    end if 
    readTemplateSource = s 
end function 



'读模板文件中某模块
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

'读文件模板素材
function readTemplateFileSource(templateFilePath, sID)
    dim startStr, endStr, s, c 
	readTemplateFileSource=""
    startStr = "<!--#sourceHtml" & replace(sID, ".html", "") & "Start#-->" 
    endStr = "<!--#sourceHtml" & replace(sID, ".html", "") & "End#-->" 
    s = readTemplateFileModular(templateFilePath, startStr, endStr) 
    if s = "[$NO$]" then
        '加一个读取本文本里配置列表(20150815)
        c = getStrCut(pubCode, startStr, endStr, 2) 
        if c <> "" then
            readTemplateFileSource = c 
            'call rwend(c)
            exit function 
        end if 
        c = getftext(templateFilePath) 
        '存在 <!--#TemplateSplitStart#-->  就返回当前全部内容
        if inStr(c, "<!--#DialogStart#-->") > 0 then
            readTemplateFileSource = c 
            exit function 
        end if 

        s = "模板资源ID[" & sID & "]不存在,路径TemplateFilePath=" & handlePath(templateFilePath) 
    end if 
    readTemplateFileSource = s 
end function 
'读出文件展示列表资源
function readArticleListStyleSource(sID)
    dim filePath 
    filePath = cfg_webImages & "\文章展示样式\" & sID 
    if checkFile(filePath) = false then
        filePath = cfg_webTemplate & "\Resources\" & sID 
    end if 
    'call echo(checkfile(filePath),filePath)
    readArticleListStyleSource = readTemplateFileSource(filePath, sID) 

end function 
'读出文件信息列表资源
function readArticleInfoStyleSource(sID)
    dim filePath 
    filePath = cfg_webImages & "\文章信息展示样式\" & sID 
    if checkFile(filePath) = false then
        filePath = cfg_webTemplate & "\Resources\" & sID 
    end if 
    readArticleInfoStyleSource = readTemplateFileSource(filePath, sID) 
end function 


%>     


