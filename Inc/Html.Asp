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
'Html 处理HTML代码 (2014,1,3)
'显示HTML结构        call rw(displayHTmL("<br>aasdfds<br>"))
'关闭显示HTML结构   call rwend(unDisplayHtml("&lt;br&gt;aasdfds&lt;br&gt;"))

'显示HTML结构
function displayHtml(str)
    str = replace(str, "<", "&lt;") 
    str = replace(str, ">", "&gt;") 
    displayHtml = str 
end function 
'关闭显示HTML结构
function unDisplayHtml(str)
    str = replace(str, "&lt;", "<") 
    str = replace(str, "&gt;", ">") 
    unDisplayHtml = str 
end function 

'处理闭合HTML标签(20150902)  比上面的更好用 第二种
function handleCloseHtml(content, isImgAddAlt, action)
    dim i, endStr, s, s2, c, labelName, startLabel, endLabel 
    action = "|" & action & "|" 
    startLabel = "<" 
    endLabel = ">" 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        endStr = mid(content, i) 
        if s = "<" then
            if inStr(endStr, ">") > 0 then
                s = mid(endStr, 1, inStr(endStr, ">")) 
                i = i + len(s) - 1 
                s = mid(s, 2, len(s) - 2) 
                s = phptrim(s) 
                if right(s, 1) = "/" then
                    s = phptrim(left(s, len(s) - 1)) 
                end if 
                endStr = right(endStr, len(endStr) - len(s) - 2)                                '最后字符减去当前标签  -2是因为它有<>二个字符
                '注意之前放在labelName下面
                labelName = mid(s, 1, inStr(s & " ", " ") - 1) 
                labelName = LCase(labelName) 
                'call eerr("s",s)

                if inStr(action, "|处理A链接|") > 0 then
                    s = handleHtmlAHref(s, labelName, "http://127.0.0.1/TestWeb/Web/", "处理A链接") '处理干净html标签
                elseIf inStr(action, "|处理A链接第二种|") > 0 then
                    s = handleHtmlAHref(s, labelName, "http://127.0.0.1/debugRemoteWeb.asp?url=", "处理A链接") '处理干净html标签
                end if 
                'call echo(s,labelName)   param与embed是Flash用到，不过embed有结束标签的
                if inStr("|meta|link|embed|param|input|img|br|hr|rect|line|area|script|div|span|a|", "|" & labelName & "|") > 0 then
                    s = replace(replace(replace(replace(s, " class=""""", ""), " alt=""""", ""), " title=""""", ""), " name=""""", "") '临时这么做一下，以后要完整系统的做
                    s = replace(replace(replace(replace(s, " class=''", ""), " alt=''", ""), " title=''", ""), " name=''", "") 

                    '给vb.net软件用的 要不然它会报错，晕
                    if labelName = "img" and isImgAddAlt = true then
                        if inStr(s, " alt") = 0 then
                            s = s & " alt=""""" 
                        end if 
                        s = trim(s) 
                        s = s & " /" 
                    '补齐<script>20160106  暂时不能用这个，等改进
                    elseIf labelName = "script" then
                        if inStr(s, " type") = 0 then
                            s = s & " type=""text/javascript""" 
                        end if 
                    elseIf right(trim(s), 1) <> "/" and inStr("|meta|link|embed|param|input|img|br|hr|rect|line|area|", "|" & labelName & "|") > 0 then
                        s = trim(s) 
                        s = s & " /" 
                    end if 
                end if 
                s = startLabel & s & endLabel 
                '处理javascript script部分
                if labelName = "script" then
                    s2 = mid(endStr, 1, inStr(endStr, "</"&"script>") + 8) 

                    'call eerr("",s2)
                    i = i + len(s2) 
                    s = s & s2 
                end if 
            'call echo("s",replace(s,"<","&lt;"))
            end if 
        end if 
        c = c & s 
    next 
    handleCloseHtml = c 
end function 
'处理htmlA标签的Href链接  配合上面函数
function handleHtmlAHref(byVal content, labelName, addToHttpUrl, action)
    dim i, s, c, temp 
    dim isValue                                                               '是否为内容值
    dim valueStr                                                              '存储内容值
    dim yinghaoLabel                                                          '引号类型如'"
    dim parentName                                                            '参数名称
    dim behindStr                                                             '后面全部字符
    dim sNotDanYinShuangYinStr                                                  '不是单引号和双引号字符
    action = "|" & action & "|" 
    content = replace(content & " ", vbTab, " ")                              '退格替换成空格，最后加一个空格，方便计算
    content = replace(replace(content, " =", "="), " =", "=") 
    isValue = false                                                           '默认内容为假，因为先是获得标签名称
    for i = 1 to len(content)
        s = mid(content, i, 1)                                                       '获得当前一个字符
        behindStr = mid(content, i)                                                  '后面字符
        if s = "=" and isValue = false then                                          '不是内容值，并为=号
            isValue = true 
            valueStr = "" 
            yinghaoLabel = "" 
            if c <> "" and right(c, 1) <> " " then c = c & " " 
            parentName = LCase(temp)                           '参数名称转小写
            c = c & parentName & s 
            temp = "" 
        '获得值第一个字符，因为它是引号类型
        elseIf isValue = true and yinghaoLabel = "" then
            if s <> " " then
                if s <> "'" and s <> """" then
                    sNotDanYinShuangYinStr = s                                                        '不是单引号和双引号字符
                    s = " " 
                end if 
                yinghaoLabel = s 
            'call echo("yinghaoLabel",yinghaoLabel)
            end if 
        elseIf isValue = true and yinghaoLabel <> "" then
            '为引号结束
            if yinghaoLabel = s then
                isValue = false 
                if labelName = "a" and parentName = "href" and inStr(action, "|处理A链接|") > 0 then
                    '处理
                    if inStr(valueStr, "?") > 0 then
                        valueStr = replace(valueStr, "?", "WenHao") & ".html" 
                    end if 
                    if inStr("|asp|php|aspx|jsp|", "|" & LCase(mid(valueStr, inStrRev(valueStr, ".") + 1)) & "|") > 0 then
                        valueStr = valueStr & ".html" 
                    end if 
                    valueStr = addToOrAddHttpUrl(addToHttpUrl, valueStr, "替换") 

                end if 
                'call echo("labelName",labelName)
                if yinghaoLabel = " " then
                    c = c & """" & sNotDanYinShuangYinStr & valueStr & """ "                          '追加 不是单引号和双引号字符            补全
                else
                    c = c & yinghaoLabel & valueStr & yinghaoLabel                                  '追加 不是单引号和双引号字符
                end if 
                yinghaoLabel = "" 
                sNotDanYinShuangYinStr = ""                                                       '不是单引号和双引号字符 清空
            else
                valueStr = valueStr & s 
            end if 
        '为 分割
        elseIf s = " " then
            '暂存内容不为空
            if temp <> "" then
                if left(trim(behindStr) & " ", 1) = "=" then
                '后面一个字符等于=不处理
                else
                    '为标签
                    if isValue = false then
                        temp = LCase(temp) & " "                                  '标签类型名称转小写
                    end if 
                    c = c & temp 
                    temp = "" 
                end if 
            end if 
        else
            temp = temp & s 
        end if 

    next 
    c = trim(c) 
    handleHtmlAHref = c 
end function 
'追加或替换网址(20150922) 配合上面   addToOrAddHttpUrl("http://127.0.0.1/aa/","http://127.0.0.1/4.asp","替换") = http://127.0.0.1/aa/4.asp
function addToOrAddHttpUrl(httpurl, byVal url, action)
    dim s 
    action = "|" & action & "|" 
    if inStr(action, "|替换|") > 0 then
        s = getwebsite(url) 
        if s <> "" then
            url = replace(url, s, "") 
        end if 
    end if 
    if inStr(url, httpurl) = 0 then
        if right(httpurl, 1) = "/" and(left(url, 1) = "/" or left(url, 1) = "\") then
            url = mid(url, 2) 
        end if 
        url = httpurl & url 
    end if 

    addToOrAddHttpUrl = url 
end function 

'获得HTML标签名 call rwend(getHtmlLableName("<img src><a href=>",0))    输出  img
function getHtmlLableName(content, nThisLabel)
    dim i, endStr, s, c, labelName, nLabelCount 
    nLabelCount = 0 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        endStr = mid(content, i) 
        if s = "<" then
            if inStr(endStr, ">") > 0 then
                s = mid(endStr, 1, inStr(endStr, ">")) 
                i = i + len(s) - 1 
                s = mid(s, 2, len(s) - 2) 
                s = phptrim(s) 
                if right(s, 1) = "/" then
                    s = phptrim(left(s, len(s) - 1)) 
                end if 
                endStr = right(endStr, len(endStr) - len(s) - 2)                                '最后字符减去当前标签  -2是因为它有<>二个字符
                '注意之前放在labelName下面
                labelName = mid(s, 1, inStr(s & " ", " ") - 1) 
                labelName = LCase(labelName) 
                if nThisLabel = nLabelCount then
                    exit for 
                end if 
                nLabelCount = nLabelCount + 1 
            end if 
        end if 
        c = c & s 
    next 
    getHtmlLableName = labelName 
end function 

'删除html里空行 最笨的方法 删除空行
function removeBlankLines(content)
    dim s, c, splStr 
    splStr = split(content, vbCrLf) 
    for each s in splStr
        if replace(replace(s, vbTab, ""), " ", "") <> "" then
            if c <> "" then c = c & vbCrLf 
            c = c & s 
        end if 
    next 
    removeBlankLines = c 
end function 


'call echo("webtitle",getHtmlValue(content,"webtitle"))
'call echo("webdescription",getHtmlValue(content,"webdescription"))
'call echo("webkeywords",getHtmlValue(content,"webkeywords"))
'获得html里面指定值20160520 call echo("webtitle",getHtmlValue(content,"webtitle"))
function getHtmlValue(content, sType)
    dim i, endStr, s, labelName, startLabel, endLabel, LCaseEndStr, paramName 
    startLabel = "<" 
    endLabel = ">" 
	getHtmlValue=""
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        endStr = mid(content, i) 
        if s = "<" then
            if inStr(endStr, ">") > 0 then
                s = mid(endStr, 1, inStr(endStr, ">")) 
                i = i + len(s) - 1 
                s = mid(s, 2, len(s) - 2) 
                s = phptrim(s) 
                if right(s, 1) = "/" then
                    s = phptrim(left(s, len(s) - 1)) 
                end if 
                endStr = right(endStr, len(endStr) - len(s) - 2)                                '最后字符减去当前标签  -2是因为它有<>二个字符
                '注意之前放在labelName下面
                labelName = mid(s, 1, inStr(s & " ", " ") - 1) 
                labelName = LCase(labelName) 

                if labelName = "title" and sType = "webtitle" then
                    LCaseEndStr = LCase(endStr) 
                    if inStr(LCaseEndStr, "</title>") > 0 then
                        s = mid(endStr, 1, inStr(LCaseEndStr, "</title>") - 1) 
                    else
                        s = "" 
                    end if 
                    getHtmlValue = s 
                    exit function 
                elseIf labelName = "meta" and(sType = "webkeywords" or sType = "webdescription") then
                    LCaseEndStr = LCase(endStr) 
                    paramName = phptrim(LCase(getParamValue(s, "name"))) 
                    if sType="web" & paramName then
                        getHtmlValue = getParamValue(s, "content") 
                        exit function 
                    end if 


                end if 

            end if 
        end if 
    next 
    getHtmlValue = "" 
end function 

'获得参数值20160520  call rwend(getParamValue("meta name=""keywords"" content=""{$web_keywords$}""","name"))
function getParamValue(content, paramName)
    dim LCaseContent, s, sArrayStart, sArrayEnd, i, startStr, endStr 
    LCaseContent = LCase(content) 
	getParamValue=""
	
    sArrayStart = array("=""", "='", "=") 
    sArrayEnd = array("""", "'", ">") 
    for i = 0 to uBound(sArrayStart)
        startStr = paramName & sArrayStart(i) 
        endStr = sArrayEnd(i) 
        if inStr(LCaseContent, startStr) > 0 and inStr(LCaseContent, endStr) > 0 then
            s = strCut(content, startStr, endStr, 2) 
            if s <> "" then
                getParamValue = s 
                exit function 
            end if 
        end if 
    next 
end function 
 




'检测Html内容是否设置了编码(当前页没有调用过20150119)  暂时留着
function checkContentCharset(byVal content, byVal char_Set)
    content = LCase(content) 
    char_Set = LCase(char_Set) 
    checkContentCharset = true 
    if inStr(content, "<meta ") > 0 then
        dim i, s, tempS, lalType, lalStr 
        for i = 1 to len(content)
            s = mid(content, i, 1) 
            if s = "<" then
                tempS = LCase(mid(content, i)) 
                lalType = LCase(mid(tempS, 1, inStr(tempS, " "))) 
                if lalType = "<meta " then
                    lalStr = mid(tempS, 1, inStr(tempS, ">")) 
                    if inStr(lalStr, "charset") > 0 then
                        if inStr(lalStr, char_Set) = 0 then checkContentCharset = false 
                    end if 
                end if 
            end if 
        next 
    end if 
end function 

%>     


