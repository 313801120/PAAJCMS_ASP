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
'AdodbStream Adodb.Stream���������ȫ (2013,9,27)

'����ļ�����
function checkCode(byVal fileName)
    checkCode = getFileCharset(fileName) 
end function 
'����ļ�����
function getFileCharset(byVal fileName)
    on error resume next 
    dim objStream, bintou 
    call handlePath(fileName)                                                       '�������·��
    set objStream = createObject("ADODB.Stream")
        objStream.type = 1 
        objStream.mode = 3 
        objStream.open 
        objStream.position = 0 
        call objStream.loadFromFile(fileName) 
        bintou = objStream.read(2)                                                      '�Զ����Ʒ�ʽ���������ַ�����(��̫̫)
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
    if err then doError err.description, "getFileCharset������ļ����� ����������FileName=" & fileName 

end function 
'����ļ�����Senior
function seniorCheckCode(byVal fileName)
    '��ʱ�������ã���Ϊ�������ǻ�Ҫ����20150119
    'CheckCode = checkFileCode(FileName)            '����ļ�����(������ ����׼)
    'Exit Function
    'On Error Resume Next
    dim objStream, bintou 
    call handlePath(fileName)                                                       '�������·��
    set objStream = createObject("ADODB.Stream")
        objStream.type = 1 
        objStream.mode = 3 
        objStream.open 
        objStream.position = 0 
        call objStream.loadFromFile(fileName) 
        bintou = objStream.read(2)                                                      '�Զ����Ʒ�ʽ���������ַ�����(��̫̫)
        if typeName(bintou) = "Null" then seniorCheckCode = checkFileCode(fileName) : exit function '�Ľ�2015019
        'N1 = AscB(MidB(bintou, 1, 1)) : N2 = AscB(MidB(bintou, 2, 1))
        'N = N1 & N2
        if ascB(midB(bintou, 1, 1)) = &HEF and ascB(midB(bintou, 2, 1)) = &HBB then
            seniorCheckCode = "UTF-8" 
        elseIf ascB(midB(bintou, 1, 1)) = &HFF and ascB(midB(bintou, 2, 1)) = &HFE then
            seniorCheckCode = "UNICODE" 
        else
            seniorCheckCode = checkFileCode(fileName)                                       '��һ�ֳ�ǿ����ļ����뷽��
        end if 
        'CAll Rw(FileName & " >> Char_Set=" & CheckCode & " Code=" & N1  & "["& Chr(N1) &"]" &  "," & N2 & "["& Chr(N2) &"]" &"(utf-8(239,187),unicode(255,254)<br>")
        objStream.close 
    set objStream = nothing 
    if err then doError err.description, "CheckCode������ļ����� ����������FileName=" & fileName 
end function 
'��������������
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
    call handlePath(fileName)                                                       '�������·��
    if checkFile(fileName) = false then
        readFile = "" 
        exit function 
    end if 
    set stm = createObject("ADODB.Stream")
        stm.type = 2                                                                    '�Ա�ģʽ��ȡ
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
            if err then doError err.description, "ReadFile���������������� ����������FileName=" & fileName 
        set fso = nothing 
    set stm = nothing 
end function 
'��������������  (����)
function getStext(fileName)
    getStext = readFile(fileName, "") 
end function 
'��UTF-8�������ļ�����
function getFTextUTF(fileName)
    getFTextUTF = readFile(fileName, "utf-8") 
end function 
'д������
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
    call handlePath(fileName)                                                       '�������·��
    set stm = createObject("ADODB.Stream")
        stm.type = 2                                                                    '�Ա�ģʽ��ȡ
        stm.mode = 3 
        stm.charset = char_Set 
        stm.open 
        call stm.writeText(content) 
        call stm.saveToFile(fileName, 2) 
        stm.flush 
        stm.close 
        writeToFile = fileName & "д��ɹ�" 
    set stm = nothing 
    if err then doError err.description, "WriteToFile��������д������ ����������FileName=" & fileName & "��Content�ַ�" & len(content) 
end function 
'����gb2312�ļ�
function createFileGBK(byval fileName, byval content)
    call writeToFile(fileName, content, "gb2312") 
end function 
'����UTF-8�ļ���20151201��
function createFileUTF(byval fileName, byval content)
    call writeToFile(fileName, content, "utf-8") 
end function 
'�����ۼ��ļ�2
function createAddFile2(filePath, addToStr)
    dim content 
    content = readFile(filePath, "gb2312") 
    content = content & addToStr 
    call writeToFile(filePath, content, "gb2312") 
end function 
'��UTF��GB2312
'�Զ���ñ��벢���ж�Html�����Ƿ�������ļ�����һ�£�һ�����˳�����
function getUGText(byVal path)
    dim char_Set, fanCharSet, tempCharSet 
    dim isContent 
    char_Set = checkCode(path) 
    isContent = false 
    '���δ���������ȡ����
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
    'û�ж������������ļ����������������
    if isContent = false then
        getUGText = readFile(path, char_Set) 
    end if 
end function 
'���Html�����Ƿ������˱���(��ǰҳû�е��ù�20150119)
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
'�����ļ������ȡ
function checkFileCode(filePath)
    dim splStr, c, content, toPath, yuanStr, s, nLen 
    yuanStr = binaryReadFile(filePath) 

    'If Len(YuanStr) > 500 Then
    'YuanStr = Right(YuanStr, Len(YuanStr) - 30)
    'End If
    'ֻ��ȡ2500���ַ�
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
            call deleteFile(toPath)                                                         'ɾ�������ʱ�ļ�
            exit for 
        end if 
    next 
    call deleteFile(toPath)                                                         'ɾ�������ʱ�ļ�
end function 
'����ļ����������� 20150312
function getFileBinaryContent(filePath)
    dim stream 
    call handlePath(filePath)                                                       '�������·��
    set stream = createObject("ADODB.Stream")
        stream.mode = 3 
        stream.type = 1 
        stream.open 
        call stream.loadFromFile(filePath) 
        getFileBinaryContent = stream.read 
        stream.close 
    set stream = nothing 
'Response.BinaryWrite Stream.Read            '��ʾ�������ļ�����
end function 
'��ȡ�������ļ����������ַ�ʽ
function binaryReadFile(filePath)
    dim i, nR, c, stm 
    call handlePath(filePath)                                                       '�������·��
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
'��ȡ�������ļ�������
function readBinary(filePath, pass)
    dim i, nR, c, stm 
    if pass = "" then pass = 0                                                      '���Ϊ����Ϊ��20150119
    call handlePath(filePath)                                                       '�������·��
    set stm = createObject("ADODB.Stream")
        stm.mode = 3 
        stm.type = 1 
        stm.open 
        call stm.loadFromFile(filePath) 
        c = "" : i = 0 : nR = 1                                                         'ÿ�ζ�ȡ�ַ���
        while i < stm.size
            c = c &(ascB(stm.read(nR)) + pass) & "|" 
            i = i + nR 
        wend 
        if c <> "" then c = left(c, len(c) - 1) 
    set stm = nothing 
    readBinary = c 
end function 
'���ܱ���������ļ�
function decSaveBinary(filePath, content, pass)
    dim splStr, buf(), i 
    if pass = "" then pass = 0                                                      '���Ϊ����Ϊ��20150119
    splStr = split(content, "|") 
    reDim buf(uBound(splStr)) 
    for i = 0 to uBound(splStr)
        buf(i) = CByte(CInt(splStr(i)) - pass) 
    next 
    call WriteBinary(filePath, buf) 
end function 
'����������ļ�   call writeBinary(filePath,Request.BinaryRead (Request.TotalBytes))
sub writeBinary(filePath, buf)
    dim i, aBuf, size, bStream 
    call handlePath(filePath)                                                       '�������·��
    '���Ƕ����ƴ����ɶ�����(20151211)
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



'---------------------------- ���� ----------------------------
'����ļ����� �������ٰ� (��ASPBOX���ù���)  ��Ҫ��
public function getFileCharset2(byVal filePath)
    dim strFileHead 
    'On Error Resume Next
    if err then
        err.clear 
    end if 
    call handlePath(filePath)                                                       '�������·��
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



'���԰��ַ������������ת���ַ���ţ �ӱ������︴�ƹ���20160112
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
'ת�ɶ�������Byte()
function base64Decode(strData)
    dim objXd 
    set objXd = createobject("msxml.domdocument")
        call objXd.loadXml("<root/>") 
        objXd.documentElement.dataType = "bin.base64" 
        objXd.documentElement.text = strData 
        base64Decode = objXd.documentElement.nodeTypedValue 
end function

%>  

