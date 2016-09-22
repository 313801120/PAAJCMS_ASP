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
dim wordReadContent 
'获得关键词
function getKeyWord(byVal content)
    getKeyWord = getHandleKeyWord(content, "") 
end function 
'获得处理后关键词  20150224
function getHandleKeyWord(byVal content, sType)
    dim nOK, splStr, i, j, tempContent, c, s 
    if wordReadContent = "" then
        wordReadContent = vbCrLf & getFText("/DataDir/keyWord.txt") & vbCrLf 
    end if 
    'Call Die(WordReadContent)
    nOK = 0 
    if CStr(sType) = "1" then
        tempContent = content & "    " 
        for i = 1 to len(content)
            for j = 2 to 4
                s = mid(tempContent, i, j) 
                if inStr(" " & c & " ", " " & s & " ") = false then
                    if inStr(wordReadContent, vbCrLf & s & vbCrLf) > 0 then
                        c = c & s & " " 
                        nOK = nOK + 1 
                    end if 
                end if 
            next 
        next 
    else                                       '这种方法比上面方法快N倍
        splStr = split(wordReadContent, vbCrLf) 
        for each s in splStr
            'Call Echo("S",S)
            if s <> "" then
                if inStr(" " & c & " ", " " & s & " ") = false then
                    if inStr(content, s) > 0 then
                        c = c & s & " " 
                        nOK = nOK + 1 
                    end if 
                end if 
            end if 
        next 
    end if 
    getHandleKeyWord = c 
end function 


'获得中文汉字内容
function getChinaStr(content)
    dim i, c, j, s 
    for i = 1 to len(content)
        j = asc(mid(content, i, 1)) 
        s = mid(content, i, 1) 
        '是汉字累加
        if j < 0 then
            if(j <= -22033 and j >= -24158) = false and inStr("�? ", s) = false then
                c = c & s 
            end if 
        '不是汉字 判断前一个字符是否为空格
        elseIf right(c, 1) <> " " then
            c = c & " " 
        end if 
    next 

    getChinaStr = c 
end function 
%>  
