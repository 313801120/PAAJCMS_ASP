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
'动态生静态(2013,12,17)

'================ 快速获得网站操作 ==================
'在线修改 修改单文本
'MainStr = DisplayOnlineED2(WEB_ADMINURL &"MainInfo.Asp?act=ShowEdit&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainStr, "<li|<a ")
'在线修改 产品大类
'DidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditBigClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), DidStr, "<li|<a ")
'在线修改 产品小类
'SidStr = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditSmallClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), SidStr, "<li|<a ")
'在线修改 产品子类
'S = DisplayOnlineED2(WEB_ADMINURL &"ProductClassManage.Asp?act=ShowEditThreeClass&Id=" & TempRs("Id") & "&n=" & GetRnd(11), S, "<li|<a ")
'在线修改  文章
'ProStr = DisplayOnlineED2(WEB_ADMINURL &"Product.Asp?act=ShowEditProduct&Id=" & TempRs("Id") & "&n=" & GetRnd(11), ProStr, "<li|<a ")
'在线修改 导航大类
'NavDidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavDidStr, "<li|<a ")
'在线修改 导航小类
'NavSidStr = DisplayOnlineED2(WEB_ADMINURL &"NavManage.Asp?act=EditNavSmall&Id=" & TempRs("Id") & "&n=" & GetRnd(11), NavSidStr, "<li|<a ")

'-------------------------------- 下面为网站后台常用快捷标签代码区 -------------------------------------------

'符加文字颜色
function infoColor(str, color)
    if color <> "" then str = "<font color=" & color & ">" & str & "</font>" 
    infoColor = str 
end function 
'图片加载失败显示默认图片
function imgError()
    imgError = " onerror=""this.src='/UploadFiles/NoImg.jpg'""" 
end function 
'获得target样式
function handleTargetStr(byVal sType)
	handleTargetStr=""
    if sType <> "" then
        handleTargetStr = " target='" & sType & "'" 
    end if 
end function 
'打开方式  (辅助)
function aTarget(sType)
    aTarget = handleTargetStr(sType) 
end function 
'获得链接Title样式
function aTitle(byVal title)
	aTitle=""
    if title <> "" then
        aTitle = " Title='" & title & "'" 
    end if 
end function 
'获得链接Title
function imgAlt(byVal alt)
	imgAlt=""
    if alt <> "" then
        imgAlt = " alt='" & alt & "'" 
    end if 
end function 
'图片标题与Alt
function imgTitleAlt(byVal str)
	imgTitleAlt=""
    if str <> "" then
        imgTitleAlt = " alt='" & str & "' title='" & str & "'" 
    end if 
end function 
'获得A Rel值
function aRel(byVal isType)
	aRel=""
    if isType = true then
        aRel = " rel='nofollow'" 
    end if 
end function 
'获得target样式
function styleClass(byVal className)
	styleClass=""
    if className <> "" then
        styleClass = " class='" & className & "'" 
    end if 
end function 
'文本加粗
function textFontB(byVal text, isFontB)
    if isFontB = true then
        text = "<strong>" & text & "</strong>" 
    end if 
    textFontB = text 
end function 
'文本加颜色
function textFontColor(byVal text, color)
    if color <> "" then
        text = "<font color='" & color & "'>" & text & "</font>" 
    end if 
    textFontColor = text 
end function 
'处理文本颜色与加粗
function fontColorFontB(title, isFontB, fontColor)
    fontColorFontB = textFontColor(textFontB(title, isFontB), fontColor) 
end function 
'获得默认文章信息文件名称
function getDefaultFileName()
    getDefaultFileName = format_Time(now(), 6) 
end function 
'获得链接  例：'"<a " & AHref(Url, TempRs("BigClassName"), TempRs("Target")) & ">" & TempRs("BigClassName") & "</a>"
function aHref(url, title, target)
    url = handleHttpUrl(url)                                                        '处理一下URL 让之完整
    aHref = "href='" & url & "'" & aTitle(title) & aTarget(target) 
end function 
'获得图片路径
function imgSrc(url, title, target)
    url = handleHttpUrl(url)                                                        '处理一下URL 让之完整
    imgSrc = "src='" & url & "'" & aTitle(title) & imgAlt(title) & aTarget(target) 
end function 

'============== 网站后台使用 ==================

'选择Target打开方式
function selectTarget(target)
    dim c, sel 
    c = c & "<select name=""Target"" id=""Target"">" & vbCrLf 
    c = c & "  <option value=''>链接打开方式</option>" & vbCrLf 
    if target = "" then sel = " selected" else sel = "" 
    c = c & "  <option" & sel & " value=''>本页打开</option>" & vbCrLf 
    if target = "_blank" then sel = " selected" else sel = "" 
    c = c & "  <option value=""_blank""" & sel & ">新页打开</option>" & vbCrLf 
    if target = "Index" then sel = " selected" else sel = "" 
    c = c & "  <option value=""Index""" & sel & ">Index页打开</option>" & vbCrLf 
    if target = "Main" then sel = " selected" else sel = "" 
    c = c & "  <option value=""Main""" & sel & ">Main页打开</option>" & vbCrLf 
    c = c & "</select>" & vbCrLf 
    selectTarget = c  
end function 
'选择文本颜色
function selectFontColor(fontColor)
    dim c, sel 
    c = c & "  <select name=""FontColor"" id=""FontColor"">" & vbCrLf 
    c = c & "    <option value=''>文本颜色</option>" & vbCrLf 
    if fontColor = "Red" then sel = " selected" else sel = "" 
    c = c & "    <option value=""Red"" class=""FontColor_Red""" & sel & ">红色</option>" & vbCrLf 
    if fontColor = "Blue" then sel = " selected" else sel = "" 
    c = c & "    <option value=""Blue"" class=""FontColor_Blue""" & sel & ">蓝色</option>" & vbCrLf 
    if fontColor = "Green" then sel = " selected" else sel = "" 
    c = c & "    <option value=""Green"" class=""FontColor_Green""" & sel & ">绿色</option>" & vbCrLf 
    if fontColor = "Black" then sel = " selected" else sel = "" 
    c = c & "    <option value=""Black"" class=""FontColor_Black""" & sel & ">黑色</option>" & vbCrLf 
    if fontColor = "White" then sel = " selected" else sel = "" 
    c = c & "    <option value=""White"" class=""FontColor_White""" & sel & ">白色</option>" & vbCrLf 
    c = c & "  </select>" & vbCrLf 
    selectFontColor = c 
end function 
'选择男女
function selectSex(sex)
    dim c, sel 
    c = c & "  <select name=""FontColor"" id=""FontColor"">" & vbCrLf 
    c = c & "    <option value=""男"">男</option>" & vbCrLf 
    sel = IIf(sex = "女", " selected", "") 
    c = c & "    <option value=""女""" & sel & ">女</option>" & vbCrLf 
    c = c & "  </select>" & vbCrLf 
    selectSex = c 
end function 
'选择Session或Cookies验证
function selectSessionCookies(verificationMode)
    dim c, sel 
    c = c & "  <select name=""VerificationMode"" id=""VerificationMode"">" & vbCrLf 
    c = c & "    <option value=""1"">Session验证</option>" & vbCrLf 
    sel = IIf(verificationMode = "0", " selected", "") 
    c = c & "    <option value=""0""" & sel & ">Cookies验证</option>" & vbCrLf 
    c = c & "  </select>" & vbCrLf 
    selectSessionCookies = c 
end function 
'显示选择分割内容  showSelectList("aa","aa|bb|cc","|","bb")
function showSelectList(IDName, content, sSplType, thisValue)
    dim c, sel, splStr, s 
    IDName = trim(IDName) 
    if sSplType = "" then sSplType = "|_-|" 
    if IDName <> "" then c = c & "  <select name=""" & IDName & """ id=""" & IDName & """>" & vbCrLf 

    splStr = split(content, sSplType) 
    for each s in splStr
        sel = "" 
        if s = thisValue then sel = " selected" 
        c = c & "    <option value=""" & s & """" & sel & ">" & s & "</option>" & vbCrLf 
    next 
    if IDName <> "" then c = c & "  </select>" & vbCrLf 
    showSelectList = c 
end function 

'显示文章展示列表样式 20150114   例 Call Rw(ShowArticleListStyle("下载列表二.html"))
function showArticleListStyle(byVal thisValue)
    showArticleListStyle = handleArticleListStyleOrInfoStyle("文章展示样式", "ArticleListStyle", thisValue) 
end function 
'显示文章信息展示样式 20150114   例 Call Rw(ShowArticleInfoStyle("下载列表二.html"))
function showArticleInfoStyle(byVal thisValue)
    showArticleInfoStyle = handleArticleListStyleOrInfoStyle("文章信息展示样式", "ArticleInfoStyle", thisValue) 
end function 
'处理文章展示列表样式和文章信息样式
function handleArticleListStyleOrInfoStyle(folderName, inputName, byVal thisValue)
    dim resourceDir, content, c, splStr, fileName, sel  

    resourceDir = cfg_webImages & "\" & folderName & "\" 

    content = getDirHtmlNameList(resourceDir) 

    thisValue = LCase(thisValue)                                                    '转成小写 好对比

    c = c & "  <select name=""" & inputName & """ id=""" & inputName & """>" & vbCrLf 
    c = c & "    <option value=""""></option>" & vbCrLf 
    splStr = split(content, vbCrLf) 
    for each fileName in splStr
        if fileName <> "" then
            sel = IIF(LCase(fileName) = thisValue, " selected", "") 
            c = c & "    <option value=""" & fileName & """" & sel & ">" & fileName & "</option>" & vbCrLf 
        end if 
    next 
    c = c & "  </select>" & vbCrLf 

    handleArticleListStyleOrInfoStyle = c 
end function 

'获得模块皮肤 ShowWebModuleSkins("ModuleSkins", ModuleSkins)
function showWebModuleSkins(inputName, byVal thisValue)
    dim resourceDir, content, c, splStr, fileName, sel 
    resourceDir = cfg_webImages & "\Index\column" 
    'Call Echo("ResourceDir",ResourceDir)
    content = getDirFolderNameList(resourceDir) 
    'Call Echo("Content",Content)

    thisValue = LCase(thisValue)                                                    '转成小写 好对比

    c = c & "  <select name=""" & inputName & """ id=""" & inputName & """>" & vbCrLf 
    c = c & "    <option value=""""></option>" & vbCrLf 
    splStr = split(content, vbCrLf) 
    for each fileName in splStr
        if fileName <> "" then
            sel = IIF(LCase(fileName) = thisValue, " selected", "") 
            c = c & "    <option value=""" & fileName & """" & sel & ">" & fileName & "</option>" & vbCrLf 
        end if 
    next 
    c = c & "  </select>" & vbCrLf 

    showWebModuleSkins = c 
end function 

'显示单选项列表
function showRadioList(IDName, content, sSplType, thisValue)
    dim c, sel, splStr, s, i 
    IDName = trim(IDName) 
    if sSplType = "" then sSplType = "|_-|" 
    i = 0 
    splStr = split(content, sSplType) 
    for each s in splStr
        sel = "" : i = i + 1 
        if s = thisValue then sel = " checked" 
        c = c & "<input type=""radio"" name=""" & IDName & """ id=""" & IDName & i & """ value=""radio"" " & sel & "><label for=""" & IDName & i & """>" & s & "</label>" & vbCrLf 
    next 

    showRadioList = c 
end function 
'显示Input复选 InputCheckBox("Id",ID,"")
function inputCheckBox(textName, isChecked, helpStr)
    'Dim sel
    'If CStr(valueStr) = "True" Or CStr(isChecked) = "1" Then sel = " isChecked" Else sel = ""
    'inputCheckBox = "<input type='checkbox' name='" & textName & "' id='" & textName & "'" & sel & " value='1'>"
    'If helpStr <> "" Then inputCheckBox = "<label for='" & textName & "'>" & inputCheckBox & helpStr & "</label> "
    inputCheckBox = handleInputCheckBox(textName, isChecked, "1", helpStr, "") 
end function 
'显示Input复选 InputCheckBox("Id",ID,"")
function inputCheckBox3(textName, isChecked, valueStr, helpStr)
    inputCheckBox3 = handleInputCheckBox(textName, isChecked, valueStr, helpStr, "newidname") 
end function 
function handleInputCheckBox(textName, isChecked, valueStr, helpStr, sType)
    dim s, sel, idName 
    if CStr(valueStr) = "True" or isChecked = true then sel = " checked" else sel = "" 
    idName = textName                                                               'id名等于文件名称
    sType = "|" & sType & "|" 
    if inStr(sType, "|newidname|") > 0 then
        idName = textName & phprand(1, 9999) 
    end if 
    s = "<input type='checkbox' name='" & textName & "' id='" & idName & "'" & sel & " value='" & valueStr & "'>" 
    if helpStr <> "" then s = "<label for='" & idName & "'>" & s & helpStr & "</label> " 
    handleInputCheckBox = s 
end function 

'显示Input文本  InputText("FolderName", FolderName, "40px", "帮助文字")
function inputText(textName, valueStr, width, helpStr)
    dim css 

    width = trim(LCase(width)) 
    if width <> "" then
        if right(width, 1) <> "%" and right(width, 2) <> "px" then
            width = width & "px" 
        end if 
        css = " style='width:" & width & ";'" 
    end if 
    inputText = "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & css & " />" & helpStr 
end function 
'显示Input文本  InputText("FolderName", FolderName, "40px", "帮助文字")
function inputText2(textName, valueStr, width, className, helpStr)
    dim css 
    if className <> "" then
        className = " class=""" & className & """" 
    end if 
    width = trim(LCase(width)) 
    if width <> "" then
        if right(width, 1) <> "%" and right(width, 2) <> "px" then
            width = width & "px" 
        end if 
        css = " style='width:" & width & ";'" 
    end if 
    inputText2 = "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & css & className & " />" & helpStr 
end function 
'显示Input文本在左边  InputLeftText(TextName, ValueStr, "98%", "")
function inputLeftText(textName, valueStr, width, helpStr)
    dim css 
    width = trim(LCase(width)) 
    if width <> "" then
        if right(width, 1) <> "%" and right(width, 2) <> "px" then
            width = width & "px" 
        end if 
        css = " style='width:" & width & ";'" 
    end if 
    inputLeftText = helpStr & "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & css & " />" & vbCrLf 
end function 
'显示Input文本在左边 帮助文字在右边
function inputLeftTextHelpTextRight(textName, valueStr, width, helpStr)
    dim css 
    width = trim(LCase(width)) 
    if width <> "" then
        if right(width, 1) <> "%" and right(width, 2) <> "px" then
            width = width & "px" 
        end if 
        css = " style='width:" & width & ";'" 
    end if 
    inputLeftTextHelpTextRight = "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & css & " />" & helpStr 
end function 
'显示Input文本在中边 提示文本在左边
function inputLeftTextContent(textName, valueStr, width, helpStr)
    inputLeftTextContent = handleInputLeftRightTextContent("左边", textName, valueStr, width, helpStr) 
end function 
'显示Input文本在中边 提示文本在右边
function inputRightTextContent(textName, valueStr, width, helpStr)
    inputRightTextContent = handleInputLeftRightTextContent("右边", textName, valueStr, width, helpStr) 
end function 
'显示Input文本在中边 提示文本在左边 或 提示文本在右边 20150114
function handleInputLeftRightTextContent(sType, textName, valueStr, width, helpStr)
    dim css 
    width = trim(LCase(width)) 
    if width <> "" then
        if right(width, 1) <> "%" and right(width, 2) <> "px" then
            width = width & "px" 
        end if 
        css = " style='width:" & width & ";'" 
    end if 
    if css = "" then
        css = " style='text-align:center;'" 
    else
        css = replace(css, ";'", ";text-align:center;'") 
    end if 
    handleInputLeftRightTextContent = "<input name=""" & textName & """ type=""text"" id=""" & textName & """ value=""" & valueStr & """" & css & " />" 

    if sType = "左边" then
        handleInputLeftRightTextContent = helpStr & handleInputLeftRightTextContent & vbCrLf 
    else
        handleInputLeftRightTextContent = handleInputLeftRightTextContent & helpStr 
    end if 

end function 

'显示Input文本在左边密码
function inputLeftPassText(textName, valueStr, width, helpStr)
    dim css 
    width = trim(LCase(width)) 
    if width <> "" then
        if right(width, 1) <> "%" and right(width, 2) <> "px" then
            width = width & "px" 
        end if 
        css = " style='width:" & width & ";'" 
    end if 
    inputLeftPassText = helpStr & "<input name=""" & textName & """ type=""password"" id=""" & textName & """ value=""" & valueStr & """" & css & " />" & vbCrLf 
end function 
'显示Input文本在左边密码类型
function inputLeftPassTextContent(textName, valueStr, width, helpStr)
    dim css 
    width = trim(LCase(width)) 
    if width <> "" then
        if right(width, 1) <> "%" and right(width, 2) <> "px" then
            width = width & "px" 
        end if 
        css = " style='width:" & width & ";'" 
    end if 
    if css = "" then
        css = " style='text-align:center;'" 
    else
        css = replace(css, ";'", ";text-align:center;'") 
    end if 
    inputLeftPassTextContent = helpStr & "<input name=""" & textName & """ type=""password"" id=""" & textName & """ value=""" & valueStr & """" & css & " />" & vbCrLf 
end function 
'显示Input隐藏文本
function inputHiddenText(textName, valueStr)
    inputHiddenText = "<input name=""" & textName & """ type=""hidden"" id=""" & textName & """ value=""" & valueStr & """ />" & vbCrLf 
end function 
'显示Input文本域 InputTextArea("FindTpl", FindTpl, "60%" , "120px", "")
function inputTextArea(textName, valueStr, width, height, helpStr)
    dim css, heightStr 
    width = trim(LCase(width)) 
    if width <> "" then
        if right(width, 1) <> "%" and right(width, 2) <> "px" then
            width = width & "px" 
        end if 
        css = " style='width:" & width & ";'" 
    end if 
    if height <> "" then
        if checkNumber(height) then                                                     '自动加个px像素
            height = height & "px" 
        end if 
        heightStr = "height:" & height & ";" 
        if css <> "" then
            css = replace(css, ";'", ";" & heightStr & ";'") 
        else
            css = " style='height:" & height & ";'" 
        end if 
    end if 
    css = replace(css, ";;", ";")                                                   '去掉多余的值
    inputTextArea = "<textarea name=""" & textName & """ type=""text"" id=""" & textName & """" & css & ">" & valueStr & "</textarea>" & helpStr 
end function 
'显示隐藏Input文本域 InputTextArea("WebDescription", WebDescription, "99%", "100px", "")
function inputHiddenTextArea(textName, valueStr, width, height, helpStr)
    inputHiddenTextArea = handleInputHiddenTextArea(textName, valueStr, width, height, "", helpStr) 
end function 
'显示隐藏Input文本域 InputTextArea("WebDescription", WebDescription, "99%", "100px", "")
function handleInputHiddenTextArea(textName, valueStr, width, height, className, helpStr)
    dim css, heightStr 
    if className <> "" then
        className = " class=""" & className & """" 
    end if 
    if width <> "" then css = " style='width:" & width & ";'" 
    if height <> "" then
        heightStr = "height:" & height & ";" 
        if css <> "" then
            css = replace(css, ";'", ";" & heightStr & ";'") 
        else
            css = " style='height:" & height & ";display:none;'" 
        end if 
    end if 
    handleInputHiddenTextArea = "<textarea name=""" & textName & """ type=""text"" id=""" & textName & """" & css & className & ">" & valueStr & "</textarea>" & helpStr 
end function 
'显示目录列表 以Select方式显示
function showSelectDirList(folderPath, valueStr)
    dim splStr, c, fileName, sel 
    splStr = split(getDirFileSort(folderPath), vbCrLf) 
    for each fileName in splStr
        if fileName <> "" then
            sel = IIF(valueStr = fileName, " selected", "") 
            c = c & "<option value=""" & folderPath & fileName & """ " & sel & ">" & fileName & "</option>" & vbCrLf 
        end if 
    next 
    showSelectDirList = c 
end function 
'给Input加个Disabled不可操作
function inputDisabled(byVal content)
    inputDisabled = replace(content, "<input ", "<input disabled=""disabled"" ") 
end function 

'给Input加个rel关系内容
function inputAddAlt(byVal content, altStr)
    dim searchStr, replaceStr 
    searchStr = "<input " 
    replaceStr = searchStr & "alt=""" & altStr & """ " 
    if inStr(content, searchStr) > 0 then
        content = replace(content, searchStr, replaceStr) 
    else
        searchStr = "<textarea " 
        replaceStr = searchStr & "alt=""" & altStr & """ " 
        if inStr(content, searchStr) > 0 then
            content = replace(content, searchStr, replaceStr) 
        end if 
    end if 
    inputAddAlt = content 
end function 



'快速调用设置====================================================

'网站描述
function webTitle_InputTextArea(webTitle)
    webTitle_InputTextArea = inputText("WebTitle", webTitle, "70%", "  多个关键词用-隔开") '不填为网站默认标题
end function 
'网站关键词
function webKeywords_InputText(webKeywords)
    webKeywords_InputText = inputText("WebKeywords", webKeywords, "70%", " 请以，隔开(中文逗号)") 
end function 
'网站描述
function webDescription_InputTextArea(webDescription)
    webDescription_InputTextArea = inputTextArea("WebDescription", webDescription, "99%", "100px", "") 
end function 
'静态文件夹名
function folderName_InputText(folderName)
    folderName_InputText = inputText("FolderName", folderName, "40%", "") 
end function 
'静态文件名
function fileName_InputText(fileName)
    fileName_InputText = inputText("FileName", fileName, "40%", ".html 也可以是网络上的链接地址") 
end function 
'模板文件名

function templatePath_InputText(templatePath)
    templatePath_InputText = inputText("TemplatePath", templatePath, "40%", " 不填为默认") 
end function 
'获得拼音按钮内容
function clickPinYinHTMLStr(did)
    clickPinYinHTMLStr = "<a href=""javascript:GetPinYin('FolderName','" & did & "','AjAx.Asp?act=GetPinYin')"" >获得拼音</a>" 
end function 
'选择文本颜色与文本加粗
function showFontColorFontB(fontColor, isFontB)
    showFontColorFontB = selectFontColor(fontColor) & inputCheckBox("FontB", isFontB, "加粗") 
end function 
'显示文本TEXT排序
function showSort(sort)
    showSort = inputText("Sort", sort, "30px", "") 
    showSort = replace(showSort, ";'", ";text-align:center;'") 
end function 
'网站导航类型顶部底部等
function showWebNavType(isNavTop, isNavButtom, isNavLeft, isNavContent, isNavRight, isNavOthre)
    dim c 
    c = c & inputCheckBox("NavTop", isNavTop, "顶部导航") 
    c = c & inputCheckBox("NavButtom", isNavButtom, "底部导航") 
    c = c & inputCheckBox("NavLeft", isNavLeft, "左边导航") 
    c = c & inputCheckBox("NavContent", isNavContent, "中间导航") 
    c = c & inputCheckBox("NavRight", isNavRight, "右边导航") 
    c = c & inputCheckBox("NavOthre", isNavOthre, "其它导航") 
    showWebNavType = c 
end function 
function showOnHtml(isOnHtml)
    showOnHtml = inputCheckBox("OnHtml", isOnHtml, "生成HTML") 
end function 
function showThrough(isThrough)
    showThrough = inputCheckBox("Through", isThrough, "审核") 
end function 
function showRecommend(isRecommend)
    showRecommend = inputCheckBox("Recommend", isRecommend, "推荐") 
end function 
'显示开户与关闭图片
function showOnOffImg(id, table, fieldName, isRecommend, url)
    dim temp, img, aUrl 
    if rq("page") <> "" then temp = "&page=" & rq("page") else temp = "" 
    if isRecommend = true then
        img = "<img src=""" & adminDir & "Images/yes.gif"">" 
    else
        img = "<img src=""" & adminDir & "Images/webno.gif"">" 
    end if 
    'Call Echo(GetUrl(),""& adminDir &"HandleDatabase.Asp?act=SetTrueFalse&Table=" & Table & "&FieldName=" & FieldName & "&Url=" & Url & "&Id=" & Id)
    aUrl = getUrlAddToParam(getUrl(), "" & adminDir & "HandleDatabase.Asp?act=SetTrueFalse&Table=" & table & "&FieldName=" & fieldName & "&Url=" & url & "&Id=" & id, "replace") 
    showOnOffImg = "<a href=""" & aUrl & """>" & img & "</a>" 
    '旧版
'ShowOnOffImg = "<a href="& adminDir &"HandleDatabase.Asp?act=SetTrueFalse&Table=" & Table & "&FieldName=" & FieldName & "&Url=" & Url & "&Id=" & Id & Temp & ">" & Img & "</a>"
end function 
'显示开户与关闭图片
function newShowOnOffImg(id, table, fieldName, isRecommend, url)
    dim temp, img 
    if rq("page") <> "" then temp = "&page=" & rq("page") else temp = "" 
    if isRecommend = true then
        img = "<img src=""/Images/yes.gif"">" 
    else
        img = "<img src=""/Images/webno.gif"">" 
    end if 
    newShowOnOffImg = "<a href=/WebAdmin/ZAction.Asp?act=Through&Table=" & table & "&FieldName=" & fieldName & "&Url=" & url & "&Id=" & id & temp & ">" & img & "</a>" 
end function 


'获得控制Css样式 20150128  暂时不用
function controlDialogCss()
    dim c 
    c = "<style>" & vbCrLf 
    c = c & "/*控制Css20150128*/" & vbCrLf 
    c = c & ".controlDialog{" & vbCrLf 
    c = c & "    position:relative;" & vbCrLf 
    c = c & "    height:50px;" & vbCrLf 
    c = c & "    width:auto;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & ".controlDialog .menu{" & vbCrLf 
    c = c & "    position:absolute;" & vbCrLf 
    c = c & "    right:0px;" & vbCrLf 
    c = c & "    top:0px;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & ".controlDialog .menu a{" & vbCrLf 
    c = c & "    color:#FF0000;" & vbCrLf 
    c = c & "    font-size:14px;" & vbCrLf 
    c = c & "    text-decoration:none;" & vbCrLf 
    c = c & "    background-color:#FFFFFF;" & vbCrLf 
    c = c & "    border:1px solid #003300;" & vbCrLf 
    c = c & "    padding:4px;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & ".controlDialog .menu a:hover{" & vbCrLf 
    c = c & "    color:#C60000;" & vbCrLf 
    c = c & "    text-decoration:underline;" & vbCrLf 
    c = c & "}" & vbCrLf 
    c = c & "</style>" & vbCrLf 
    controlDialogCss = c 
end function 


'删除里暂存代码
function batchDeleteTempStr(content, startStr, endStr)
    dim i, s 
    for i = 1 to 9
        if inStr(content, startStr) = 0 then
            exit for 
        end if 
        s = getStrCut(content, startStr, endStr, 1) 
        content = replace(content, s, "") 
    next 
    batchDeleteTempStr = content 
end function 
%>     
