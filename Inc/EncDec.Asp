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
'加密解密(2014)


'特殊Html上传加密解密 20150121 specialHtmlUploadEncryptionDecrypt(Content,"Decrypt")
function specialHtmlUploadEncryptionDecrypt(content, sType)
    dim splStr, splxx, c, s 
    c = "·|[*-24156*]" & vbCrLf 
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

'加密ASP代码内容
function encAspContent(byVal content)
    dim splStr, c, s, THStr 
    c = "Str=Str&""|Str=Str & |If | Then|End If|&vbCrlf|Temp |Rs(|Rs.|.AddNew|(""Title"")|(""Content"")|=False|ElseIf|" 
    c = c & "Conn.Execute(""| Exit For|[Product]|.Open|.Close|Exit For|Exit Function|MoveNext:Next:|Str " 
    splStr = split(c, "|") 
    for each s in splStr
        if s <> "" then
            THStr = upperCaseORLowerCase(s) 
            content = replace(content, chr(9), "") 'Chr(9) = Tab
            content = replace(content, s, THStr) 
        end if 
    next 
    encAspContent = content 
end function 
'让大小写乱掉
function upperCaseORLowerCase(byVal content)
    dim i, s, c, nRnd 
    c = "" 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        randomize 
        nRnd = CInt(rnd() * 1) 
        if nRnd = 0 then
            c = c & LCase(s) 
        else
            c = c & uCase(s) 
        end if 
    next 
    upperCaseORLowerCase = c 
end function 
'加密  Encryption
function encCode(byVal content)
    dim i, c 
    c = "" 
    for i = 1 to len(content)
        c = c & "%" & hex(asc(mid(content, i, 1))) 
    next 
    encCode = c 
end function 
'解密  Decrypt
function decCode(byVal content)
    dim i, c, splStr,s
    c = "" 
    splStr = split(content, "%") 
    for i = 1 to uBound(splStr)
        if splStr(i) <> "" then
			s="&H" & splStr(i)
            c = c & chr(CInt(s)) 
        end if 
    next 
    decCode = c 
end function 
'将汉字等转换为&#开头的unicode字符串形式
public function toUnicode(str)
    dim i, j, c, p 
    toUnicode = "" 
    c = "" 
    p = "" 
    for i = 1 to len(str)
        c = mid(str, i, 1) 
        j = ascW(c) 
        if j < 0 then
            j = j + 65536 
        end if 
        if j >= 0 and j <= 128 then
            if p = "c" then
                toUnicode = " " & toUnicode 
                p = "e" 
            end if 
            toUnicode = toUnicode & c 
        else
            if p = "e" then
                toUnicode = toUnicode & " " 
                p = "c" 
            end if 
            toUnicode = toUnicode & "&#" & j & ";" 
        end if 
    next 
end function 
'日文26字母编码
function japan(byVal iStr, sType)
    japan = "" 
    if isNull(iStr) or isEmpty(iStr) then
        japan = "" 
        exit function 
    end if 
    dim arrF, i, arrE 
    if sType = "" then sType = "0"
    'F=array("ゴ","ガ","ギ","グ","ゲ","ザ","ジ","ズ","ヅ","デ",_
    '"ド","ポ","ベ","プ","ビ","パ","ヴ","ボ","ペ","ブ","ピ","バ",_
    '"ヂ","ダ","ゾ","ゼ")
    'E = Array("Jn0;", "Jn1;", "Jn2;", "Jn3;", "Jn4;", "Jn5;", "Jn6;", "Jn7;", "Jn8;", "Jn9;", "Jn10;", "Jn11;", "Jn12;", "Jn13;", "Jn14;", "Jn15;", "Jn16;", "Jn17;", "Jn18;", "Jn19;", "Jn20;", "Jn21;", "Jn22;", "Jn23;", "Jn24;", "Jn25;")
    arrE = split("Jn0;,Jn1;,Jn2;,Jn3;,Jn4;,Jn5;,Jn6;,Jn7;,Jn8;,Jn9;,Jn10;,Jn11;,Jn12;,Jn13;,Jn14;,Jn15;,Jn16;,Jn17;,Jn18;,Jn19;,Jn20;,Jn21;,Jn22;,Jn23;,Jn24;,Jn25;", ",") 

    'F = Array(Chr( -23116), Chr( -23124), Chr( -23122), Chr( -23120),    Chr(-23118), Chr( -23114), Chr( -23112), Chr( -23110),     Chr(-23099), Chr( -23097), Chr( -23095), Chr( -23075),   Chr(-23079), Chr( -23081), Chr( -23085), Chr( -23087),  Chr(-23052), Chr( -23076), Chr( -23078), Chr( -23082),  Chr(-23084), Chr( -23088), Chr( -23102), Chr( -23104), Chr(-23106), Chr( -23108))
    arrF = split(chr( -23116) & "," & chr( -23124) & "," & chr( -23122) & "," & chr( -23120) & "," & chr( -23118) & "," & chr( -23114) & "," & chr( -23112) & "," & chr( -23110) & "," & chr( -23099) & "," & chr( -23097) & "," & chr( -23095) & "," & chr( -23075) & "," & chr( -23079) & "," & chr( -23081) & "," & chr( -23085) & "," & chr( -23087) & "," & chr( -23052) & "," & chr( -23076) & "," & chr( -23078) & "," & chr( -23082) & "," & chr( -23084) & "," & chr( -23088) & "," & chr( -23102) & "," & chr( -23104) & "," & chr( -23106) & "," & chr( -23108), ",") 
    japan = iStr 
    for i = 0 to 25
        if sType = "0" then
            japan = replace(japan, arrF(i), arrE(i)) 
        else
            japan = replace(japan, arrE(i), arrF(i)) 
        end if 
    next  
end function 
'日文26字母 加密
function japan26(iStr)
    japan26 = japan(iStr, "0")   
end function   
'日文26字母 解密
function unJapan26(iStr)
    unJapan26 = japan(iStr, "1") 
end function 
'处理内容让它成为纯HTML代码
function handleHTML(byVal content)
    'Content = Replace(Content, "&", "&amp;")
    content = replace(content, "<", "&lt;") 
    handleHTML = content 
end function 
'解开 处理内容让它成为纯HTML代码
function unHandleHTML(byVal content)
    'Content = Replace(Content, "&amp;", "&")
    content = replace(content, "&lt;", "<") 
    unHandleHTML = content 
end function 
'小写加密   [可扩展为大写与数字]
function lcaseEnc(str)
    dim i, n, s, c 
    c = "" 
    for i = 1 to len(str)
        s = mid(str, i, 1) 
        n = ascW(s) 
        if n >= 97 and n <= 122 then
            c = c & chrW(n + 1) 
        else
            c = c & s 
        end if 
    next 
    c = replace(c, vbCrLf, "＠") 
    lcaseEnc = c 
end function 
'小写解密
function lcaseDec(str)
    dim i, n, s, c 
    c = "" 
    for i = 1 to len(str)
        s = mid(str, i, 1) 
        n = ascW(s) 
        if n >= 97 and n <= 123 then
            c = c & chrW(n - 1) 
        else
            c = c & s 
        end if 
    next 
    c = replace(c, "＠", vbCrLf) 
    lcaseDec = c 
end function 

'html转换成js
function htmlToJs(byVal c)
    c = replace("" & c, "\", "\\") 
    c = replace(c, "/", "\/") 
    c = replace(c, "'", "\'") 
    c = replace(c, """", "\""") 
    c = join(split(c, vbCrLf), """);" & vbCrLf & "document.write(""") 
    c = "document.write(""" & c & """);" 
    htmlToJs = c 
end function 
'js转换成html
function jsToHtml(byVal c)
    c = replace(c, "document.write(""", "") 
    c = replace(c, """);", "") 
    c = replace(c, "\""", """") 
    c = replace(c, "\'", "'") 
    c = replace(c, "\/", "/") 
    c = replace(c, "\\", "\") 
    jsToHtml = c 
end function 
'html转换成Asp
function htmlToAsp(byVal c)
    c = replace(c, """", """""") 
    c = join(split(c, vbCrLf), """)" & vbCrLf & "Response.Write(""") 
    c = "Response.Write(""" & c & """)" 
    htmlToAsp = c 
end function 
'Html转Asp变量存储
function htmlToAspDim(byVal c)
    c = replace(c, """", """""") 
    c = join(split(c, vbCrLf), """" & vbCrLf & "C=C & """) 
    c = "C=C & """ & c & """" 
    htmlToAspDim = c 
end function 
'Asp转换成html
function aspToHtml(byVal c)
    c = replace(c, "Response.Write(""", "") 
    c = replace(c, """""", """") 
    aspToHtml = c 
end function 
'检测创建文件名称是否合法
function checkCreateFileName(byval fileName)
	dim i,s,sArrayA
	sArrayA= array("\", "/", ":", "*", "?", """", "<", ">", "|", ",") 
	for i = 0 to uBound(sArrayA)
        s = sArrayA(i) 
		if instr(fileName,s)>0 then
			checkCreateFileName=false
			exit function
		end if
    next 
	checkCreateFileName=true
end function
'文件命名规则
function setFileName(byVal fileName)
    dim i, s, tStr, sArrayA, sArrayB 
    'sArrayA = array("\", "/", ":", "*", "?", """", "<", ">", "|", ".", ",")       '换这种方法是为了与PHP版通用 20160511
    'sArrayB = array("撇", "揦", "：", "星", "？", "“", "左", "右", "横", "。", "，") 
	sArrayA= array("\", "/", ":", "*", "?", """", "<", ">", "|", ".", ",") 
	sArrayB= array("撇", "揦", "：", "星", "？", "“", "左", "右", "横", "。", "，") 
    for i = 0 to uBound(sArrayA)
        s = sArrayA(i) 
        tStr = sArrayB(i) 
        fileName = replace(fileName, s, tStr) 
    next 
    fileName = replace(fileName, "&nbsp;", " ") 
    fileName = replace(fileName, "&quot;", "双") 
    fileName = replace(fileName, vbCrLf, "") 
    setFileName = fileName  
end function 
'文件命名规则解开
function unSetFileName(byVal fileName)
    dim i, s, tStr, sArrayA, sArrayB  
    sArrayA = array("\", "/", ":", "*", "?", """", "<", ">", "|", ".", ",") 
    sArrayB = array("撇", "揦", "：", "星", "？", "“", "左", "右", "横", "。", "，") 
    for i = 0 to uBound(sArrayA)
        s = sArrayA(i) 
        tStr = sArrayB(i) 
        fileName = replace(fileName, tStr, s) 
    next 

    unSetFileName = fileName 
end function 

'把Html转成ASP，并且字符转成Chr(*)样式
function HTMLToAscChr(title)
    dim i, s, c 
    c = "" 
    for i = 1 to len(title)
        s = mid(title, i, 1) 
        c = c & "Chr(" & asc(s) & ")&" 
    next 
    if c <> "" then c = left(c, len(c) - 1) 
    HTMLToAscChr = c 
'HTMLToAscChr = "<" & "%=" & C & "%" & ">"
end function 
'解密AscChr
function unHTMLToAscChr(content)
    dim i, s, c, splStr, temp 
    c = content : temp = "" 
    c = replace(c, "Chr(", "") 
    c = replace(c, ")&", " ") 
    c = replace(c, ")", " ") 
    splStr = split(c, " ") 
    for i = 0 to uBound(splStr) - 1
        s = splStr(i) 
        temp = temp & chr(s) 
    next 
    unHTMLToAscChr = temp 
end function 

'变量移位
function variableDisplacement(content, nPass)
    dim c, i, s, letterGroup, digitalGroup, nLetterGroup, nDigitalGroup, nLetterLen, nDigitalLen, nX 
    '字母组
    'LetterGroup="abcdefghijklmnopqrstuvwxyz"
    letterGroup = "yzoehijklmfgqrstuvpabnwxcd" 
    '字母长
    nLetterGroup = len(letterGroup) 
    '数字组
    'DigitalGroup="0123456789"
    digitalGroup = "4539671820" 
    '数字长
    nDigitalGroup = len(digitalGroup) 
    c = "" 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        nLetterLen = inStr(letterGroup, s) 
        nDigitalLen = inStr(digitalGroup, s) 
        '字母处理
        if nLetterLen > 0 then
            nX = nLetterLen + nPass 
            if nX > nLetterGroup then
                nX = nX - nLetterGroup 
            elseIf nX <= 0 then
                'Call Echo("nX",nX & "," & (nLetterGroup - nX) & "/" & nLetterGroup)
                nX = nX + nLetterGroup 
            end if 
            s = mid(letterGroup, nX, 1) 
        '数字处理
        elseIf nDigitalLen > 0 then
            nX = nDigitalLen + nPass 
            if nX > nDigitalGroup then
                nX = nX - nDigitalGroup 
            elseIf nX <= 0 then
                'Call Echo("nX",nX & "," & (nLetterGroup - nX) & "/" & nLetterGroup)
                nX = nX + nDigitalGroup 
            end if 
            s = mid(digitalGroup, nX, 1) 


        end if 
        c = c & s 
    next 
    variableDisplacement = c 
end function 

'【删除此行start】

'Xor加密
function xorEnc(code, n)
    dim c, s1, s2, s3, i 
    c = code 
    s1 = len(c) : s3 = "" 
    for i = 0 to s1 - 1
        s2 = ascW(right(c, s1 - i)) xor n 
        s3 = s3 & chrW(int(s2)) 
    next 
    'Chr(34) 就是等于(") 防止出错 因为"在ASP里出错
    s3 = replace(s3, chrW(34), "ㄨ") 
    xorEnc = s3 
end function 
'Xor解密
function xorDec(code, n)
    dim c, s1, s2, s3, i 
    c = code : s3 = "" 
    c = replace(c, "ㄨ", chrW(34)) 
    s1 = len(c) 
    for i = 0 to s1 - 1
        s2 = ascW(right(c, s1 - i)) xor n 
        s3 = s3 & chrW(int(s2)) 
    next 
    s3 = replace(s3, "＠", vbCrLf) 
    xorDec = s3 
end function 


'处理Eval加密
function handleEvalAddDec(byVal content, addStr)
    dim c 
    c = addStr & "Execute unDec_log(""" & encCode(content) & """)" & addStr 
    c = c & "Function unDec_log(ByVal Content)" & addStr 
    c = c & "    Dim I, C, SplStr " & addStr 
    c = c & "    SplStr = Split(Content, ""%"") " & addStr 
    c = c & "    For I = 1 To UBound(SplStr)" & addStr 
    c = c & "        C = C & Chr(CInt(""&H"" & SplStr(I))) " & addStr 
    c = c & "    Next " & addStr 
    c = c & "    unDec_log = C " & addStr 
    c = c & "End Function " & addStr 
    handleEvalAddDec = c 
end function 
'Eval加密
function evalAddDec(byVal content)
    evalAddDec = handleEvalAddDec(content, vbCrLf) 
end function 
'Eval加密
function vbEvalAddDec(byVal content)
    VBEvalAddDec = escape(handleEvalAddDec(content, ":")) 
end function 

'Xor解密    这个暂时留着
function XDec(code)
    dim c, s1, s2, s3, i 
    c = code 
    c = replace(c, "ㄨ", chr(34)) 
    s1 = len(c) : s2 = "" : s3 = "" 
    for i = 0 to s1 - 1
        s2 = asc(right(c, s1 - i)) xor 20 
        s3 = s3 & chr(int(s2)) 
    next 
    s3 = replace(s3, "＠", vbCrLf) 
    XDec = s3 
end function 


'RGB转十六进制
public function toHexColor(scrColor)
    dim hexColor 
    toHexColor = "" 
    hexColor = CStr(hex(scrColor)) 
    select case len(hexColor)
        case 1 : toHexColor = "&H00000" & hexColor
        case 2 : toHexColor = "&H0000" & hexColor
        case 3 : toHexColor = "&H000" & hexColor
        case 4 : toHexColor = "&H00" & hexColor
        case 5 : toHexColor = "&H0" & hexColor
        case 6 : toHexColor = "&H" & hexColor
    end select
end function 
'十六进制与RGB
public function colorRGB(color)
    dim byN, byN2 
    byN = "" : byN2 = "" 
    dim arrayStr(2) 
    arrayStr(0) =(color mod byN) 
    arrayStr(1) =((color mod byN2) \ byN) 
    arrayStr(2) =(color \ byN2) 
    colorRGB = arrayStr 
end function 

'简单加密
function simpleEnc(c)
    c = replace(c, "x", "輕") 
    c = replace(c, "f", "蚛") 
    c = replace(c, "s", "鋋") 
    c = replace(c, "h", "孿") 
    c = replace(c, "t", "謣") 
    c = replace(c, "o", "薺") 
    c = replace(c, "<", "讞") 
    c = replace(c, ">", "旻") 
    c = replace(c, "a", "艓") 
    c = replace(c, "c", "襋") 
    c = replace(c, "n", "贁") 
    c = replace(c, "m", "曁") 
    c = replace(c, vbCrLf, "樞") 
    simpleEnc = c 
end function 
'简单解密
function simpleDec(c)
    c = replace(c, "輕", "x") 
    c = replace(c, "蚛", "f") 
    c = replace(c, "鋋", "s") 
    c = replace(c, "孿", "h") 
    c = replace(c, "謣", "t") 
    c = replace(c, "薺", "o") 
    c = replace(c, "讞", "<") 
    c = replace(c, "旻", ">") 
    c = replace(c, "艓", "a") 
    c = replace(c, "襋", "c") 
    c = replace(c, "贁", "n") 
    c = replace(c, "曁", "m") 
    c = replace(c, "樞", vbCrLf) 
    simpleDec = c 
end function 
'加密ASP内容成Eval来运行
function encAspEvalRun(byVal content)
    dim splStr, i, c 
    '去除<% % >
    content = replace(content, "<" & "%", "") 
    content = replace(content, "%" & ">", "") 
    content = replace(content, """", """""") 
    content = replace(content, "x", "輕") 
    content = replace(content, "f", "蚛") 
    content = replace(content, "s", "鋋") 
    content = replace(content, "o", "薺") 
    content = replace(content, "e", "療") 
    splStr = split(content, vbCrLf) 
    c = "" 
    for i = 0 to uBound(splStr)
        if inStr(splStr(i), "'") = false and len(splStr(i)) > 1 then
            c = c & splStr(i) & "⊙" 
        end if 
    next 
    c = left(c, len(c) - 1) 
    c = "Execute MyEval(""" & c & """)" & vbCrLf 
    c = c & "Function MyEval(HtmlStr)" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(HtmlStr,""⊙"",vbCrlf)" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""輕"",""x"")" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""蚛"",""f"")" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""鋋"",""s"")" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""薺"",""o"")" & vbCrLf 
    c = c & vbTab & "MyEval=Replace(MyEval,""療"",""e"")" & vbCrLf 
    c = c & "End Function" & vbCrLf 
    encAspEvalRun = "<" & "%" & vbCrLf & c & "%" & ">" 
end function 

'【删除此行end】
%>     
