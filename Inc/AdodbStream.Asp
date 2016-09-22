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
'AdodbStream Adodb.Stream组件操作大全 (2013,9,27)

'检测文件编码
function checkCode(byVal fileName)
    checkCode = getFileCharset(fileName) 
end function 
'获得文件编码
function getFileCharset(byVal fileName)
    on error resume next 
    dim objStream, bintou 
    call handlePath(fileName)                                                       '获得完整路径
    set objStream = createObject("ADODB.Stream")
        objStream.type = 1 
        objStream.mode = 3 
        objStream.open 
        objStream.position = 0 
        call objStream.loadFromFile(fileName) 
        bintou = objStream.read(2)                                                      '以二进制方式读出二个字符出来(不太太)
        if typeName(bintou) = "Null" then getFileCharset = "gb2312" : exit function 
        if ascB(midB(bintou, 1, 1)) = &HEF and ascB(midB(bintou, 2, 1)) = &HBB then
            getFileCharset = "utf-8" 
        elseIf ascB(midB(bintou, 1, 1)) = &HFF and ascB(midB(bintou, 2, 1)) = &HFE then
            getFileCharset = "unicode" 
        else
            getFileCharset = "gb2312" 
        end if 
        objStream.close 
    set objStream = nothing 
    if err then doError err.description, "getFileCharset，获得文件编码 函数出错，FileName=" & fileName 

end function 
'获得文件编码Senior
function seniorCheckCode(byVal fileName)
    '暂时不这样用，因为慢，但是还要测试20150119
    'CheckCode = checkFileCode(FileName)            '获得文件编码(这种慢 但精准)
    'Exit Function
    'On Error Resume Next
    dim objStream, bintou 
    call handlePath(fileName)                                                       '获得完整路径
    set objStream = createObject("ADODB.Stream")
        objStream.type = 1 
        objStream.mode = 3 
        objStream.open 
        objStream.position = 0 
        call objStream.loadFromFile(fileName) 
        bintou = objStream.read(2)                                                      '以二进制方式读出二个字符出来(不太太)
        if typeName(bintou) = "Null" then seniorCheckCode = checkFileCode(fileName) : exit function '改进2015019
        'N1 = AscB(MidB(bintou, 1, 1)) : N2 = AscB(MidB(bintou, 2, 1))
        'N = N1 & N2
        if ascB(midB(bintou, 1, 1)) = &HEF and ascB(midB(bintou, 2, 1)) = &HBB then
            seniorCheckCode = "UTF-8" 
        elseIf ascB(midB(bintou, 1, 1)) = &HFF and ascB(midB(bintou, 2, 1)) = &HFE then
            seniorCheckCode = "UNICODE" 
        else
            seniorCheckCode = checkFileCode(fileName)                                       '另一种超强获得文件编码方法
        end if 
        'CAll Rw(FileName & " >> Char_Set=" & CheckCode & " Code=" & N1  & "["& Chr(N1) &"]" &  "," & N2 & "["& Chr(N2) &"]" &"(utf-8(239,187),unicode(255,254)<br>")
        objStream.close 
    set objStream = nothing 
    if err then doError err.description, "CheckCode，获得文件编码 函数出错，FileName=" & fileName 
end function 
'数据流读出内容
function readFile(byVal fileName, byVal char_Set)
    on error resume next 
    if char_Set = "1" or uCase(char_Set) = "GB2312" then
        char_Set = "GB2312" 
    elseIf char_Set = "0" or uCase(char_Set) = "UTF-8" then
        char_Set = "UTF-8" 
    elseIf char_Set = "2" or uCase(char_Set) = "UNICODE" then
        char_Set = "UNICODE" 
    else
        char_Set = checkCode(fileName) 
        if char_Set = "" then exit function 
    end if 
    dim str, stm, f, fso 
    call handlePath(fileName)                                                       '获得完整路径
    if checkFile(fileName) = false then
        readFile = "" 
        exit function 
    end if 
    set stm = createObject("ADODB.Stream")
        stm.type = 2                                                                    '以本模式读取
        stm.mode = 3 
        stm.charset = char_Set 
        stm.open 
        set fso = createObject("Scripting.FileSystemObject")
            set f = fso.getFile(fileName)
                if f.size > 0 then
                    call stm.loadFromFile(fileName) 
                end if 
                str = stm.readText 
                stm.close 
            set stm = nothing 
            readFile = str 
            if err then doError err.description, "ReadFile，数据流读出内容 函数出错，FileName=" & fileName 
        set fso = nothing 
    set stm = nothing 
end function 
'数据流读出内容  (辅助)
function getStext(fileName)
    getStext = readFile(fileName, "") 
end function 
'以UTF-8方法读文件内容
function getFTextUTF(fileName)
    getFTextUTF = readFile(fileName, "utf-8") 
end function 
'写入内容
function writeToFile(byVal fileName, byVal content, byVal char_Set)
    on error resume next 
    if char_Set = "1" or uCase(char_Set) = "GB2312" then
        char_Set = "GB2312" 
    elseIf char_Set = "0" or uCase(char_Set) = "UTF-8" then
        char_Set = "UTF-8" 
    elseIf char_Set = "2" or uCase(char_Set) = "UNICODE" then
        char_Set = "UNICODE" 
    else
        char_Set = "GB2312" 
    end if 
    'Call Echo("Char_Set",Char_Set)
    dim stm 
    call handlePath(fileName)                                                       '获得完整路径
    set stm = createObject("ADODB.Stream")
        stm.type = 2                                                                    '以本模式读取
        stm.mode = 3 
        stm.charset = char_Set 
        stm.open 
        call stm.writeText(content) 
        call stm.saveToFile(fileName, 2) 
        stm.flush 
        stm.close 
        writeToFile = fileName & "写入成功" 
    set stm = nothing 
    if err then doError err.description, "WriteToFile，数据流写入内容 函数出错，FileName=" & fileName & "，Content字符" & len(content) 
end function 
'创建gb2312文件
function createFileGBK(byval fileName, byval content)
    call writeToFile(fileName, content, "gb2312") 
end function 
'创建UTF-8文件（20151201）
function createFileUTF(byval fileName, byval content)
    call writeToFile(fileName, content, "utf-8") 
end function 
'创建累加文件2
function createAddFile2(filePath, addToStr)
    dim content 
    content = readFile(filePath, "gb2312") 
    content = content & addToStr 
    call writeToFile(filePath, content, "gb2312") 
end function 
'是UTF或GB2312
'自动获得编码并且判断Html内容是否与读出文件编码一致，一致则退出内容
function getUGText(byVal path)
    dim char_Set, fanCharSet, tempCharSet 
    dim isContent 
    char_Set = checkCode(path) 
    isContent = false 
    '二次处理，防读取错误
    if char_Set = "utf-8" then
        fanCharSet = "gb2312" 
    elseIf char_Set = "gb2312" then
        fanCharSet = "utf-8" 
    end if 
    tempCharSet = fanCharSet 
    if fanCharSet <> "" then
        if checkStr(LCase(getUGText), "<meta [^>]*charset *= *" & fanCharSet) = true then
            getUGText = readFile(path, tempCharSet) 
            isContent = true 
        end if 
    end if 
    '没有读出内容则用文件自身编码读出内容
    if isContent = false then
        getUGText = readFile(path, char_Set) 
    end if 
end function 
'检测Html内容是否设置了编码(当前页没有调用过20150119)
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
'超级文件编码获取
function checkFileCode(filePath)
    dim splStr, c, content, toPath, yuanStr, s, nLen 
    yuanStr = binaryReadFile(filePath) 

    'If Len(YuanStr) > 500 Then
    'YuanStr = Right(YuanStr, Len(YuanStr) - 30)
    'End If
    '只提取2500个字符
    if len(yuanStr) > 2500 then
        yuanStr = right(yuanStr, 2500) 
    end if 
    nLen = len(yuanStr) 
    toPath = filePath & "_Temp" 
    splStr = split("GB2312|UTF-8|UNICODE", "|") 
    for each s in splStr
        content = readFile(filePath, s) 
        call writeToFile(toPath, content, s) 
        c = binaryReadFile(toPath) 
        c = right(c, nLen) 
        if yuanStr = c then
            checkFileCode = s 
            call deleteFile(toPath)                                                         '删除这个临时文件
            exit for 
        end if 
    next 
    call deleteFile(toPath)                                                         '删除这个临时文件
end function 
'获得文件二进制内容 20150312
function getFileBinaryContent(filePath)
    dim stream 
    call handlePath(filePath)                                                       '获得完整路径
    set stream = createObject("ADODB.Stream")
        stream.mode = 3 
        stream.type = 1 
        stream.open 
        call stream.loadFromFile(filePath) 
        getFileBinaryContent = stream.read 
        stream.close 
    set stream = nothing 
'Response.BinaryWrite Stream.Read            '显示二进制文件内容
end function 
'读取二进制文件处理成数字方式
function binaryReadFile(filePath)
    dim i, nR, c, stm 
    call handlePath(filePath)                                                       '获得完整路径
    set stm = createObject("ADODB.Stream")
        stm.mode = 3 
        stm.type = 1 
        stm.open 
        call stm.loadFromFile(filePath) 
        c = "" : i = 0 : nR = 1 
        while i < stm.size
            c = c & ascB(stm.read(nR)) 
            i = i + nR 
            doEvents 
        wend 
    set stm = nothing 
    binaryReadFile = c 
end function 
'读取二进制文件并加密
function readBinary(filePath, pass)
    dim i, nR, c, stm 
    if pass = "" then pass = 0                                                      '如果为空则为零20150119
    call handlePath(filePath)                                                       '获得完整路径
    set stm = createObject("ADODB.Stream")
        stm.mode = 3 
        stm.type = 1 
        stm.open 
        call stm.loadFromFile(filePath) 
        c = "" : i = 0 : nR = 1                                                         '每次读取字符数
        while i < stm.size
            c = c &(ascB(stm.read(nR)) + pass) & "|" 
            i = i + nR 
        wend 
        if c <> "" then c = left(c, len(c) - 1) 
    set stm = nothing 
    readBinary = c 
end function 
'解密保存二进制文件
function decSaveBinary(filePath, content, pass)
    dim splStr, buf(), i 
    if pass = "" then pass = 0                                                      '如果为空则为零20150119
    splStr = split(content, "|") 
    reDim buf(uBound(splStr)) 
    for i = 0 to uBound(splStr)
        buf(i) = CByte(CInt(splStr(i)) - pass) 
    next 
    call WriteBinary(filePath, buf) 
end function 
'保存二进制文件   call writeBinary(filePath,Request.BinaryRead (Request.TotalBytes))
sub writeBinary(filePath, buf)
    dim i, aBuf, size, bStream 
    call handlePath(filePath)                                                       '获得完整路径
    '不是二进制处理成二进制(20151211)
    if typeName(buf) <> "Byte()" then
        size = uBound(buf) : reDim aBuf(size \ 2) 
        for i = 0 to size - 1 step 2
            aBuf(i \ 2) = chrW(buf(i + 1) * 256 + buf(i)) 
        next 
        if i = size then aBuf(i \ 2) = chrW(buf(i)) 
        aBuf = join(aBuf, "") 
    else
        aBuf = buf 
    end if 

    set bStream = createObject("ADODB.Stream")
        bStream.type = 1 
        bStream.open 
        with createObject("ADODB.Stream")
            .type = 2 
            .open 
            call.writeText(aBuf) 
            .position = 2 
            call.copyTo(bStream) 
            .close 
        end with 
        call bStream.saveToFile(filePath, 2) 
        bStream.close 
    set bStream = nothing 
end sub 



'---------------------------- 待定 ----------------------------
'获得文件编码 超级快速版 (从ASPBOX引用过来)  不要用
public function getFileCharset2(byVal filePath)
    dim strFileHead 
    'On Error Resume Next
    if err then
        err.clear 
    end if 
    call handlePath(filePath)                                                       '获得完整路径
    dim objStream, loadBytes 
    set objStream = createObject("ADODB.Stream")
        with objStream
            .type = 1 
            .mode = 3 
            .open 
            .loadFromFile filePath 
            loadBytes =.read(3) 
            .close 
        end with 
    set objStream = nothing 
    if err.number <> 0 then
        err.clear : getFileCharset2 = "" : exit function 
    end if 

    strFileHead = binToHex(loadBytes) 
    if strFileHead = "EFBBBF" then
        getFileCharset2 = "UTF-8" 
    else
        strFileHead = left(strFileHead, 4) 
        if strFileHead = "FEFF" then
            getFileCharset2 = "UNICODE BIG" 
        elseIf strFileHead = "FFFE" then
            getFileCharset2 = "UNICODE" 
        elseIf strFileHead = "3C25" or strFileHead = "3C21" then
            getFileCharset2 = "GB2312" 
        else
            getFileCharset2 = "GB2312" 
        end if 
    end if 
    if err.number <> 0 then
        err.clear : getFileCharset2 = "" : exit function 
    end if 
end function 
public function binToHex(byVal vStream)
    dim reVal, i 
    reVal = 0 
    for i = 1 to lenB(vStream)
        reVal = reVal * 256 + ascB(midB(vStream, i, 1)) 
    next 
    binToHex = hex(reVal) 
end function 



'可以把字符或二进制内容转成字符，牛 从别人哪里复制过来20160112
function base64Encode(strData)
    dim objAds, objXd 
    set objAds = createobject("adodb.stream")
        objAds.type = 2 
        objAds.charset = "unicode" 
        objAds.mode = 3 
        call objAds.open() 
        call objAds.writeText(strData) 
        objAds.position = 0 
        objAds.type = 1 
        'objAds.Position=2
        set objXd = createobject("msxml.domdocument")
            call objXd.loadXml("<root/>") 
            objXd.documentElement.dataType = "bin.base64" 
            objXd.documentElement.nodeTypedValue = objAds.read() 
            base64Encode = objXd.documentElement.text 
end function
'转成二进制了Byte()
function base64Decode(strData)
    dim objXd 
    set objXd = createobject("msxml.domdocument")
        call objXd.loadXml("<root/>") 
        objXd.documentElement.dataType = "bin.base64" 
        objXd.documentElement.text = strData 
        base64Decode = objXd.documentElement.nodeTypedValue 
end function

%>  

