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
'Css

'【删除.netc start】

'Css压缩 cssCompression(Content,0)
function cssCompression(content, level)
    level = cStr(level)                                                             '转成字符好判断
    'Css高级压缩
    if level = "1" then
        content = regExp_Replace(content, "\/\*(.|" & vbCrLf & ")*?\*\/", "") 
        content = regExp_Replace(content, "\s*([\{\}\:\;\,])\s*", "$1") 
        content = regExp_Replace(content, "\,[\s\.\#\d]*\{", "{") 
        content = regExp_Replace(content, ";\s*;", ";") 
    'Css简单压缩
    else
        if level >= 2 then
            content = regExp_Replace(content, "\/\*(.|" & vbCrLf & ")*?\*\/", "")           '//删除注释
        end if 
        content = regExp_Replace(content, "\s*([\{\}\:\;\,])\s*", "$1") 
        content = regExp_Replace(content, "\,[\s\.\#\d]*\{", "{")                  '//容错处理
        content = regExp_Replace(content, ";\s*;", ";")                                 '//清除连续分号
        content = regExp_Replace(content, ";\s*}", "}")                                 '//清除末尾分号和大括号
        content = regExp_Replace(content, "([^\s])\{([^\s])", "$1{$2") 
        content = regExp_Replace(content, "([^\s])\}([^" & vbCrLf & "]s*)", "$1}" & vbCrLf & "$2") 

    end if 
    content = trimVBcrlf(content) 
    cssCompression = content 
end function 

'删除Css里注释
function deleteCssNote(content)
    deleteCssNote = regExp_Replace(content, "\/\*(.|" & vbCrLf & ")*?\*\/", "")     '//删除注释
end function 

'Css格式化  展开CSS
function unCssCompression(content)
    content = regExp_Replace(content, "\s*([\{\}\:\;\,])\s*", "$1") 
    content = regExp_Replace(content, ";\s*;", ";")                                 '//清除连续分号
    content = regExp_Replace(content, "\,[\s\.\#\d]*{", "{") 
    content = regExp_Replace(content, "([^\s])\{([^\s])", "$1 {" & vbCrLf & "" & vbTab & "$2") 
    content = regExp_Replace(content, "([^\s])\}([^" & vbCrLf & "]*)", "$1" & vbCrLf & "}" & vbCrLf & "$2") 
    content = regExp_Replace(content, "([^\s]);([^\s\}])", "$1;" & vbCrLf & "" & vbTab & "$2") 
    unCssCompression = content 
end function 

'【删除.netc end】

'去掉字符串头尾的连续的回车和空格
function trimVbCrlf(str)
    trimVbCrlf = phpRTrim(phpLTrim(str)) 
end function 

'php里Trim方法
function phpTrim(str)
    phpTrim = trimVbCrlf(str) 
end function 

'去掉字符串开头的连续的回车和空格
function phpLTrim(str)
    dim nPos, isBlankChar 
    nPos = 1 
    isBlankChar = true 
    while isBlankChar
        if mid(str, nPos, 1) = " " or mid(str, nPos, 1) = vbTab then                      '把vbTab也去掉
            nPos = nPos + 1 
        elseIf mid(str, nPos, 2) = vbCrLf then
            nPos = nPos + 2 
        else
            isBlankChar = false 
        end if 
    wend 
    phpLTrim = right(str, len(str) - nPos + 1) 
end function 

'去掉字符串末尾的连续的回车和空格
function phpRTrim(str)
    dim nPos, isBlankChar 
    nPos = len(str) 
    isBlankChar = true 
    while isBlankChar and nPos >= 2
        if mid(str, nPos, 1) = " " or mid(str, nPos, 1) = vbTab then                      '把vbTab也去掉
            nPos = nPos - 1 
        elseIf mid(str, nPos - 1, 2) = vbCrLf then
            nPos = nPos - 2 
        else
            isBlankChar = false 
        end if 
    wend 
    phpRTrim = rTrim(left(str, nPos)) 
end function 


'--------------- 有用 暂时用这文件里 ------------------
'去掉字符串头尾的连续的Tab退格和空格
function trimVbTab(str)
    trimVbTab = rTrimVBTab(lTrimVbTab(str)) 
end function 


'去掉字符串开头的连续的Tab退格和空格
function lTrimVbTab(str)
    dim nPos, isBlankChar 
    nPos = 1 
    isBlankChar = true 
    while isBlankChar
        if mid(str, nPos, 1) = " " then
            nPos = nPos + 1 
        elseIf mid(str, nPos, 1) = vbTab then
            nPos = nPos + 1 
        else
            isBlankChar = false 
        end if 
    wend 
    lTrimVbTab = right(str, len(str) - nPos + 1) 
end function 

'去掉字符串末尾的连续的Tab退格和空格
function rTrimVBTab(str)
    dim nPos, isBlankChar 
    nPos = len(str) 
    isBlankChar = true 
    while isBlankChar and nPos >= 2
        if mid(str, nPos, 1) = " " then
            nPos = nPos - 1 
        elseIf mid(str, nPos - 1, 1) = vbTab then
            nPos = nPos - 1 
        else
            isBlankChar = false 
        end if 
    wend 
    rTrimVBTab = rTrim(left(str, nPos)) 
end function 


'找Html页中Css Style <style></style> 里css内容
function getHtmlCssStyle(byVal content)
    getHtmlCssStyle = handleHtmlStyleCss(content, "", "") 
end function 
'处理html里的css 替换路径
function handleHtmlStyleCss(content, sType, imgPath)
    dim splStr, s, isStyle, styleStartStr, styleEndStr, styleStr, cssStr, sNewCssStr 
    splStr = split(content, vbCrLf)                                                 '分割行
    isStyle = false                                                                'Css样式默认为假
	handleHtmlStyleCss=""
    '循环分行
    for each s in splStr
        if isStyle = false then
            if inStr(lCase(s), "<style") > 0 then
                styleStartStr = mid(s, inStr(lCase(s), "<style")) 
                styleStartStr = mid(styleStartStr, 1, inStr(styleStartStr, ">")) 
                styleEndStr = mid(s, inStr(lCase(s), styleStartStr) + len(styleStartStr)) 
                'HTML中定义的Css在一行
                if inStr(styleEndStr, "</style>") > 0 then
                    styleStr = mid(styleEndStr, 1, inStr(styleEndStr, "</style>") - 1) 
                    cssStr = cssStr & styleStr & vbCrLf 
                else
                    cssStr = cssStr & styleEndStr & vbCrLf 
                    isStyle = true                                                       '收集CssStyle样式开始
                end if 
                'Call Echo("StyleStartStr",ShowHtml(StyleStartStr))
                'Call Echo("StyleEndStr",ShowHtml(StyleEndStr))
                'Call Echo("StyleStr",ShowHtml(StyleStr))
                'Call Echo("CssStr",ShowHtml(CssStr))
            'Call RwEnd("")
            end if 
        elseIf isStyle = true then
            if inStr(lCase(s), "</style>") > 0 then
                styleStr = mid(s, 1, inStr(lCase(s), "</style>") - 1) 
                cssStr = cssStr & styleStr & vbCrLf 

                if sType = "替换路径" then
                    sNewCssStr = replaceCssImgPath(cssStr, imgPath) 
                    content = replace(content, phptrim(cssStr), sNewCssStr) 
                end if 

                isStyle = false                                                                '收集CssStyle样式结束
            else
                cssStr = cssStr & s & vbCrLf 
            end if 
        end if 
    next 
    if sType = "替换路径" then
        handleHtmlStyleCss = content 
    else
        handleHtmlStyleCss = cssStr 
    end if 
end function 

'替换css内容里图片路径 20160718home
function replaceCssImgPath(byval cssStr, byval imgPath)
    dim content, splStr, s, c, fileName, toImgPath, lCaseS 
    content = getArray(cssStr, "\(", "\)", false, false) 
    splStr = split(content, "$Array$") 
    for each s in splStr
        if s <> "" then
			'对等号前面不处理，如   filter:progid:DXImageTransform.Microsoft.AlphaImageLoader (enabled=true, sizingMethod=scale, src='../images/ad-gallery/opa75.png');}
			if instr(s,"=")>0 then
				s=mid(s,instrrev(s,"=")+1) 
			end if
            lCaseS = lcase(s) 
            fileName = trim(getFileAttr(s, "2")) 
			'call echo("css s",s)
			'call echo("css filename",fileName)
            '引用图片 基础判断方法，有待改进，改进也是成长的一部分20160719
            if instr(lCaseS, "data:image") = 0 and(instr(lCaseS, ".jpg") > 0 or instr(lCaseS, ".gif") > 0 or instr(lCaseS, ".png") > 0 or instr(lCaseS, ".bmp") > 0) then
                '修复于20160719
                if instr(""",", left(fileName, 1)) > 0 then
                    fileName = mid(fileName, 2) 
                end if 
                if instr(""",", right(fileName, 1)) > 0 then
                    fileName = mid(fileName, 1, len(fileName) - 1) 
                end if 
                'call echo(s,imgPath  & " and " & fileName)
                cssStr = replace(cssStr, "(" & s & ")", "(" & imgPath & fileName & ")") 
            'fileName="data:image/" & fileName
            end if 

        end if 
    next 
    replaceCssImgPath = cssStr 
end function 

'处理成干净的Css内容  CSS格式化
function handleCleanCss(byVal content)
    dim splStr, s, c, isAddStr, customS 
    content = replace(content, "{", vbCrLf & "{" & vbCrLf) 
    content = replace(content, "}", vbCrLf & "}" & vbCrLf) 
    content = replace(content, ";", ";" & vbCrLf) 

    splStr = split(content, vbCrLf) 
    isAddStr = false                                                               '追加字符默认为假
    for each s in splStr
        s = trimVbCrlf(s) 
        customS = ""                                                                    '自定义S值
        if s <> "" then
            if inStr(s, "{") > 0 and inStr(s, "}") = 0 then
                isAddStr = true 
                customS = s 
            elseIf inStr(s, "}") > 0 then
                isAddStr = false 
            end if 
            if left(s, 1) <> "{" then c = c & vbCrLf 
            if isAddStr = true then s = "    " & s 
            if customS <> "" then s = customS                                               '自定义值不为空则用自定义内容
            c = c & s 

        end if 
    next 
    c = trimVbCrlf(c) 
    c = replace(c, "    ;" & vbCrLf, "")                                            '清除多于的分号
    c = replace(c, ";" & vbCrLf & "}", vbCrLf & "}")                                '最后一个参数不要分号
    handleCleanCss = c 
end function 



'移除内容中多除行
function removeExcessRow(content)
    dim splStr, s, c, tempS 
    splStr = split(content, vbCrLf)                                                 '分割行
    for each s in splStr
        tempS = replace(replace(s, " ", ""), vbTab, "") 
        if tempS <> "" then
            c = c & s & vbCrLf 
        end if 
    next 
    if c <> "" then c = left(c, len(c) - 2) 
    removeExcessRow = c 
end function 


'2014 11 30
'向Css里追加样式  a=CssAddToStyle(GetFText("1.html")," .test {color:#FF0f000; font-size:10px; float:left}")
function cssAddToStyle(content, addToStyle)
    dim styleName, yunStyleStr, replaceStyleStr, c 
    if inStr(addToStyle, "{") > 0 then
        styleName = trim(mid(addToStyle, 1, inStr(addToStyle, "{") - 1)) 
    end if 
    yunStyleStr = findCssStyle(content, styleName) 
    replaceStyleStr = cssStyleAddToParam(yunStyleStr, addToStyle)                   'Css样式累加参数
    content = replace(content, yunStyleStr, replaceStyleStr) 
    'C = C & "<hr>Content=" & Content
    cssAddToStyle = content 
    'CssAddToStyle = YunStyleStr
'CssAddToStyle = "StyleName=" & StyleName & "<hr>YunStyleStr=" & YunStyleStr & "<hr>ReplaceStyleStr=" & ReplaceStyleStr
end function 

'检测Css内容中是否有指定样式
function checkCssStyle(content, styleStr)
    dim styleName 
    checkCssStyle = true 
    if inStr(styleStr, "{") > 0 then
        styleName = trim(mid(styleStr, 1, inStr(styleStr, "{") - 1)) 
    end if 
    if styleName = "" then
        checkCssStyle = false 
    elseIf findCssStyle(content, styleName) = "" then
        checkCssStyle = false 
    end if 
end function 


'Css样式累加参数
function cssStyleAddToParam(byVal cssStyleStr, byVal cssStyleStrTwo)
    dim splStr, cssStr, s, paramList, paramName, cssStyleName 
    cssStyleName = mid(cssStyleStr, 1, inStr(cssStyleStr, "{")) 
    if inStr(cssStyleStr, "{") > 0 then
        cssStyleStr = mid(cssStyleStr, inStr(cssStyleStr, "{") + 1) 
    end if 
    if inStr(cssStyleStr, "}") > 0 then
        cssStyleStr = mid(cssStyleStr, 1, inStr(cssStyleStr, "}") - 1) 
    end if 
    if inStr(cssStyleStrTwo, "{") > 0 then
        cssStyleStrTwo = mid(cssStyleStrTwo, inStr(cssStyleStrTwo, "{") + 1) 
    end if 
    if inStr(cssStyleStrTwo, "}") > 0 then
        cssStyleStrTwo = mid(cssStyleStrTwo, 1, inStr(cssStyleStrTwo, "}") - 1) 
    end if 
    splStr = split(replace(cssStyleStr & ";" & cssStyleStrTwo, vbCrLf, ""), ";") 
    for each s in splStr
        s = trim(s) 
        if inStr(s, ":") > 0 and s <> "" then
            paramName = trim(mid(s, 1, inStr(s, ":") - 1)) 
            if inStr("|" & paramList & "|", "|" & paramName & "|") = 0 then
                paramList = paramList & paramName & "|" 
                'Call Echo("ParamName",ParamName)
                cssStr = cssStr & "    " & s & ";" & vbCrLf 
            end if 
        end if 
    next 
    if cssStyleName <> "" then
        cssStr = cssStyleName & vbCrLf & cssStr & "}" 
    end if 
    cssStyleAddToParam = cssStr 
'Call Echo(CssStyleStr,CssStyleStrTwo)
end function 

'根据Css名称找到对应Css块
function findCssStyle(byVal content, byVal styleName)
    dim splStr, s, tempS, findStyleName 
    'CAll Echo("StyleName",StyleName)
    'CAll Echo("Content",Content)
	findCssStyle=""
    styleName = trim(styleName) 
    splStr = split(content, vbCrLf) 
    for each s in splStr
        if inStr(s, styleName) > 0 then
            findStyleName = trim(s) 
            if inStr(findStyleName, "{") > 0 then
                findStyleName = trim(mid(findStyleName, 1, inStr(findStyleName, "{") - 1)) 
            end if 
            if findStyleName = styleName then
                'Call Eerr( FindStyleName , StyleName)
                if inStr(s, "}") > 0 then
                    findCssStyle = mid(s, 1, inStr(s, "}") + 1) 
                    'Call EErr(s,FindCssStyle)
                    exit function 
                else
                    tempS = mid(content, inStr(content, s & vbCrLf) + 1) 
                    tempS = mid(tempS, 1, inStr(tempS, "}") + 1) 
                    findCssStyle = tempS 
                    exit function 
                end if 
            'Call Eerr("temps",Temps)
            end if 
        'Call Echo(FindStyleName,StyleName)
        end if 
    next 
end function 

'================================================
'导入网站前端要用到
'================================================
'处理截取到的Css
function handleCutCssCode(dirPath, cssStr)
    dim content, startStr, endStr, splStr, imageFile, fileName, listStr 
    startStr = "url\(" : endStr = "\)" 
    content = getArray(cssStr, startStr, endStr, false, false) 
    splStr = split(content, "$Array$") 
    for each imageFile in splStr
        if imageFile <> "" and inStr(imageFile, ".") > 0 and inStr(vbCrLf & listStr & vbCrLf, vbCrLf & imageFile & vbCrLf) = 0 then'对重复使用的图片处理
            listStr = listStr & imageFile & vbCrLf 
            fileName = replace(replace(replace(imageFile, """", ""), "'", ""), "\", "/") 
            if inStr(fileName, "/") > 0 then
                fileName = mid(fileName, inStrRev(fileName, "/") + 1) 
            end if 
            cssStr = replace(cssStr, imageFile, dirPath & fileName) 
        end if 
    next 
    handleCutCssCode = cssStr 
end function 

'处理截取到的HtmlDiv
function handleCutDivCode(dirPath, divStr)
    dim content, startStr, endStr, splStr, imageFile, toImageFile, fileName, isHandle 
    startStr = "url\(" : endStr = "\)" 
    content = getArray(divStr, startStr, endStr, false, false) 
    splStr = split(content, "$Array$") 
    for each imageFile in splStr

        if imageFile <> "" and inStr(imageFile, ".") > 0 and inStr(imageFile, "{$#") = 0 then
            '判断是否有域名 20150202
            if getWebSite(imageFile) = "" then
                fileName = replace(replace(replace(imageFile, """", ""), "'", ""), "\", "/") 
                if inStr(fileName, "/") > 0 then
                    fileName = mid(fileName, inStrRev(fileName, "/") + 1) 
                end if 
                divStr = replace(divStr, imageFile, dirPath & imageFile) 
            end if 
        end if 
    next 
    '图片处理
    'Content = GetIMG(DivStr) & vbCrlf & GetHtmlBackGroundImgList(DivStr)        '再加个Html背景图片
    content = getImgJsUrl(divStr, "不重复") & vbCrLf & getHtmlBackGroundImgList(divStr) '再加个Html背景图片  加强版20150126
    splStr = split(content, vbCrLf) 
    for each imageFile in splStr
        if imageFile <> "" then                                                         '当链接地址当前为HTTP:时则不处理20150313
            isHandle = false 

            if left(imageFile, 1) = "\" then
            '等处理20150817
            elseIf inStr(imageFile, ".") > 0 and left(imageFile, 5) <> "HTTP:" and inStr(imageFile, "{$#") = 0 then
                isHandle = true 
            end if 
            if isHandle = true then
                toImageFile = dirPath & removeFileDir(imageFile)                                '移除文件路径目录
                'html中图片路径替换
                divStr = replace(divStr, """" & imageFile & """", """" & toImageFile & """") 
                divStr = replace(divStr, "'" & imageFile & "'", """" & toImageFile & """") 
                divStr = replace(divStr, "=" & imageFile & " ", """" & toImageFile & """") 
                divStr = replace(divStr, "=" & imageFile & ">", """" & toImageFile & """") 
            end if 
        end if 
    next 
    handleCutDivCode = divStr 
end function 

'获得HTMl里背景图片 20150116  如： <td width="980" height="169" background="kslx3bg.jpg">
function getHtmlBackGroundImgList(byVal content)
    content = getArray(content, " background=""", """", false, false) 
    content = replace(content, "$Array$", vbCrLf) 
    getHtmlBackGroundImgList = content 
end function 


'完善html里link css链接 Content = getHandleWebHtmlLink("/aa/bb/",Content)  外部调用
function getHandleWebHtmlLink(rootPath, content)
    dim startStr, endStr, imgList, splStr, c, cssUrl, sNewCssUrl, cssStr 
    startStr = "<link " 
    cssStr = "" 
    endStr = ">" 
    imgList = getArray(content, startStr, endStr, false, false) 
    'Call RwEnd(ImgList)
    splStr = split(imgList, "$Array$") 
    for each cssUrl in splStr
        if cssUrl <> "" and inStr(lCase(cssUrl), "stylesheet") > 0 then
            '获得Css加强版，改于20141125
            cssUrl = lCase(replace(replace(replace(cssUrl, """", ""), "'", ""), ">", " ")) & " " 
            startStr = "href=" : endStr = " " 
            if inStr(cssUrl, startStr) > 0 and inStr(cssUrl, endStr) > 0 then
                cssUrl = strCut(cssUrl, startStr, endStr, 2) 
            end if 
            sNewCssUrl = handleHttpUrl(cssUrl) 
            if inStr(sNewCssUrl, "/") > 0 then
                sNewCssUrl = mid(sNewCssUrl, inStrRev(sNewCssUrl, "/") + 1) 
            end if 
            if lCase(sNewCssUrl) <> "common.css" and lCase(sNewCssUrl) <> "public.css" then
                sNewCssUrl = rootPath & sNewCssUrl 
                cssStr = cssStr & "<link href=""" & sNewCssUrl & """ rel=""stylesheet"" type=""text/css"" />" & vbCrLf 
            end if 
        end if 
    next 
    if cssStr <> "" then cssStr = left(cssStr, len(cssStr) - 2) 
    getHandleWebHtmlLink = cssStr 
end function 


'获得css链接地址列表(20150824)  找<link>
function getCssListUrlList(content)
    dim startStr, endStr, imgList, splStr, c, cssUrl, cssStr, urlList 
    startStr = "<link " 
    cssStr = "" 
    endStr = ">" 
    imgList = getArray(content, startStr, endStr, false, false) 
    'Call RwEnd(ImgList)
    splStr = split(imgList, "$Array$") 
    for each cssUrl in splStr
        if cssUrl <> "" and inStr(lCase(cssUrl), "stylesheet") > 0 then
            '获得Css加强版，改于20141125
            cssUrl = lCase(replace(replace(replace(cssUrl, """", ""), "'", ""), ">", " ")) & " " 
            startStr = "href=" : endStr = " " 
            if inStr(cssUrl, startStr) > 0 and inStr(cssUrl, endStr) > 0 then
                cssUrl = strCut(cssUrl, startStr, endStr, 2) 
            end if 
            if inStr(vbCrLf & urlList & vbCrLf, vbCrLf & cssUrl & vbCrLf) = 0 then
                if urlList <> "" then urlList = urlList & vbCrLf 
                urlList = urlList & cssUrl & vbCrLf 
            end if 
        end if 
    next 
    getCssListUrlList = urlList 
end function 
'call rw(handleHtmlStyle(getftext("1.html")))
'压缩html里的style样式 (20151008)
function handleHtmlStyle(content)
    dim serchS, replaceS, nLength 
    serchS = content 
    nLength = inStr(lCase(serchS), "</style>") + 7 
    serchS = mid(serchS, 1, nLength) 

    nLength = inStrRev(lCase(serchS), "<style") 
    if nLength > 0 then
        serchS = mid(serchS, nLength) 
    end if 
    replaceS = serchS 
    replaceS = cssCompression(replaceS, 0) & vbCrLf                                 '格式化CSS
    replaceS = removeBlankLines(replaceS) 

    content = replace(content, serchS, replaceS) 
    handleHtmlStyle = content 
end function 



'读Css文件内容并处理(20150824) 加特殊记录标签   不常用
'如call rwend(handleReadCssContent("E:\E盘\WEB网站\至前网站\DataDir\VB模块\服务器\Template\模块功能列表\B站页面设计\home\home.css","aa",true))
function handleReadCssContent(cssFilePath, labelName, isHandleCss)
    dim c, startStr, endStr 
    c = getFText(cssFilePath) 
    '截取CSS
    startStr = "/*CssCodeStart*/" 
    endStr = "/*CssCodeEnd*/" 
    if inStr(c, startStr) > 0 and inStr(c, endStr) > 0 then
        c = strCut(c, startStr, endStr, 2) 
    end if 
    '处理CSS
    if isHandleCss = true then
        c = cssCompression(c, 0) 
    end if 
    if labelName <> "" then
        c = "/*" & labelName & " start*/" & c & "/*" & labelName & " end*/" 
    end if 
    handleReadCssContent = c 
end function 

'处理Css样式里PX或%
function handleCssPX(byVal sValue)
    sValue = lCase(trim(sValue)) 
    if right(sValue, 1) <> "%" and right(sValue, 2) <> "px" then
        sValue = sValue & "px" 
    end if 
    handleCssPX = sValue 
end function 

%>   
