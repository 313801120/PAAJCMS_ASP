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


'�����滻�б� (ReplaceList�ĳ�handleReplaceList)Ϊ�˼���php����
function handleReplaceList(byVal content, yStr, tStr)
    dim splYuan, splTi, i, s 
    splYuan = split(yStr, "|") 
    splTi = split(tStr & "||||||||||||||||||||||||||||||||", "|") 
    for i = 0 to uBound(splYuan)
        s = splYuan(i) 
        if s <> "" then
            content = replace(content, s, splTi(i)) 
        end if 
    next 
    handleReplaceList = content 
end function 

'���ұ߿�ʼ��ȡ ���ã�ɾ����߿ո�
function leftSpace(byVal content, byVal nAdd)
    dim i, nCount 
    nCount = 0 
    for i = len(content) to 1 step - 1
        if mid(content, i, 1) <> " " then exit for 
        nCount = nCount + 1 
    next 
    if nCount > nAdd and replace(content, " ", "") <> "" then
        content = left(content, len(content) - nCount + nAdd) 
    elseIf nCount < nAdd and 1 = 1 then
        for i = nCount to nAdd - 1
            content = content & " " 
        next 
    end if 
    leftSpace = content 
end function 
'���ұ߿�ʼ��ȡ ���ã�ɾ���ұ߿ո�
function rightSpace(byVal content, byVal nAdd)
    dim i, nCount 
    nCount = 0 
    for i = 1 to len(content)
        if mid(content, i, 1) <> " " then exit for 
        nCount = nCount + 1 
    next 
    if nCount > nAdd then
        content = right(content, len(content) - nCount + nAdd) 
    elseIf nCount < nAdd then
        for i = nCount to nAdd - 1
            content = " " & content 
        next 
    end if 
    rightSpace = content 
end function 

'=============== ����Ϊ2013,11,1����
'���QQ
function getQQ(byVal content, nOK)
    dim splStr, i, j, s, c, qQ, nErr, isQQ 
    content = replace(replace(lCase(content), "&nbsp;", ""), " ", "") 
    nOK = 0 
    if inStr(content, "qq") > 0 then
        splStr = split(content, "qq") 
        for i = 1 to uBound(splStr)
            s = splStr(i) 
            qQ = "" : nErr = 14 : isQQ = false 
            for j = 1 to len(s)
                if inStr("0123456789", mid(s, j, 1)) = false then
                    if isQQ = true then exit for                                                   'qq��ʼ�ۼ�ʱ���˳�
                    if nErr = 0 then exit for                                                       '�������N�˳�
                    nErr = nErr - 1 
                    if mid(s, j, 1) = "Ⱥ" then exit for                                            'ΪQQȺ�˳�
                else
                    isQQ = true 
                    qQ = qQ & mid(s, j, 1) 
                end if 
                if j > 30 then exit for                                                         'j����20�˳�
            next 
            if len(qQ) >= 6 and len(qQ) <= 10 and inStr(vbCrLf & c, vbCrLf & qQ & vbCrLf) = false then
                if inStr(vbCrLf & c, vbCrLf & qQ & vbCrLf) = false then
                    c = c & qQ & vbCrLf 
                    nOK = nOK + 1 
                end if 
            end if 
        next 
    end if 
    getQQ = c 
end function 
'����ֻ�
function getTel(byVal content, nOK)
    dim splStr, i, j, s, c, tel, nErr, isTel 
    content = replace(replace(lCase(content), "&nbsp;", ""), "�ֻ�", "tel") 
    nOK = 0 
    if inStr(content, "tel") > 0 then
        splStr = split(content, "tel") 
        for i = 1 to uBound(splStr)
            s = splStr(i) 
            tel = "" : nErr = 14 : isTel = false 
            for j = 1 to len(s)
                if inStr("0123456789", mid(s, j, 1)) = false then
                    if isTel = true then exit for                                                  'Tel��ʼ�ۼ�ʱ���˳�
                    if nErr = 0 then exit for                                                       '�������N�˳�
                    nErr = nErr - 1 
                else
                    isTel = true 
                    tel = tel & mid(s, j, 1) 
                end if 
                if j > 30 then exit for                                                         'j����20�˳�
            next 
            if len(tel) = 11 then
                if inStr(vbCrLf & c, vbCrLf & tel & vbCrLf) = false then
                    c = c & tel & vbCrLf 
                    nOK = nOK + 1 
                end if 
            end if 
        next 
    end if 
    getTel = c 
end function 
'�������
function getMail(byVal content, nOK)
    dim splStr, i, j, s, c, mail, nErr, isEMail 
    content = replace(replace(lCase(content), "&nbsp;", ""), "����", "mail") 
    nOK = 0 
    if inStr(content, "mail") > 0 then
        splStr = split(content, "mail") 
        for i = 1 to uBound(splStr)
            s = splStr(i) 
            mail = "" : nErr = 14 : isEMail = false 
            for j = 1 to len(s)
                if inStr("0123456789abcdefghijklmnopqrstuvwxyz.@", mid(s, j, 1)) = false then
                    if isEMail = true then exit for                                                 'Mail��ʼ�ۼ�ʱ���˳�
                    if nErr = 0 then exit for                                                       '�������N�˳�
                    nErr = nErr - 1 
                else
                    isEMail = true 
                    mail = mail & mid(s, j, 1) 
                end if 
                if j > 30 then exit for                                                         'j����20�˳�
            next 
            if inStr(mail, ".") > 0 and inStr(mail, "@") > 0 then
                if inStr(vbCrLf & c, vbCrLf & mail & vbCrLf) = false then
                    c = c & mail & vbCrLf 
                    nOK = nOK + 1 
                end if 
            end if 
        next 
    end if 
    getMail = c 
end function 
'���ͼƬ�б�
function getImgStr(httpurl, content, nOK)
    dim splStr, i, c, url, urlList 
    content = getIMG(content) 
    splStr = split(content, vbCrLf) 
    nOK = 0 
    for each url in splStr
        if inStr("|" & urlList & "|", "|" & url & "|") = false then
            if left(url, 1) = "/" or(inStr(url, "http://") = false and inStr(url, "www.") = false) then
                url = fullHttpUrl(httpurl, url) 
            end if 
            nOK = nOK + 1 
            c = c & url & "<br>" 
        end if 
    next 
    getImgStr = c 
end function 
'���Css�б�
function getCssStr(httpurl, content, nOK)
    dim splStr, i, c, url, urlList 
    content = getCssHref(content) 
    splStr = split(content, vbCrLf) 
    nOK = 0 
    for each url in splStr
        if inStr("|" & urlList & "|", "|" & url & "|") = false then
            if left(url, 1) = "/" or(inStr(url, "http://") = false and inStr(url, "www.") = false) then
                url = fullHttpUrl(httpurl, url) 
            end if 
            nOK = nOK + 1 
            c = c & url & "<br>" 
        end if 
    next 
    getCssStr = c 
end function 
'���Js�б�
function getJsStr(httpurl, content, nOK)
    dim splStr, i, c, url, urlList 
    content = getJsSrc(content) 
    splStr = split(content, vbCrLf) 
    nOK = 0 
    for each url in splStr
        if inStr("|" & urlList & "|", "|" & url & "|") = false then
            if left(url, 1) = "/" or(inStr(url, "http://") = false and inStr(url, "www.") = false) then
                url = fullHttpUrl(httpurl, url) 
            end if 
            nOK = nOK + 1 
            c = c & url & "<br>" 
        end if 
    next 
    getJsStr = c 
end function 


'�����ַ�б� Rw(GetUrlStr("", "����", Content, "", 0))
function getUrlStr(httpurl, sType, content, searchValue, nOK)
    dim hrefList, splStr, title, url, c, s, lcaseUrl, urlList, titleList, sVal, sVal2 
    if sType = "����" then
        hrefList = getAURL(content) 
    elseIf sType = "���ӱ���" then
        hrefList = getATitle(content) 
    else
        hrefList = getAURLTitle(content) 
    end if 
    splStr = split(hrefList, vbCrLf) 
    nOK = 0 
    for each s in splStr
        if s <> "" then
            if sType = "����" then
                url = trim(s) 
                if inStr(lCase(url), "javascript:") = false and url <> "#" and url <> "" then
                    if inStr("|" & urlList & "|", "|" & url & "|") = false then
                        if left(url, 1) = "/" or(inStr(url, "http://") = false and inStr(url, "www.") = false) then
                            url = fullHttpUrl(httpurl, url) 
                        end if 
                        nOK = nOK + 1 
                        c = c & url & vbCrLf 
                    end if 
                end if 
            elseIf sType = "���ӱ���" then
                if inStr("|" & titleList & "|", "|" & s & "|") = false then
                    titleList = titleList & s & "|" 
                    nOK = nOK + 1 
                    c = c & s & vbCrLf 
                end if 
            else
                if inStr(s, "��_|-��") > 0 then
                    url = mid(s, 1, inStr(s, "��_|-��") - 1) : lcaseUrl = url 
                    title = mid(s, len(url) + 6) 
                    if inStr(lCase(lcaseUrl), "javascript:") = false and url <> "#" and url <> "" then
                        if inStr("|" & urlList & "|", "|" & url & "|") = false then
                            if left(url, 1) = "/" or(inStr(lcaseUrl, "http://") = false and inStr(lcaseUrl, "www.") = false) then
                                url = fullHttpUrl(httpurl, url) 
                            end if 
                            nOK = nOK + 1 
                            c = c & url & "   &nbsp;  " & title & vbCrLf 
                        end if 
                    end if 
                end if 
            end if 
        end if 
    next 
    if searchValue <> "" then
        searchValue = replace(searchValue, " and ", " And ") 
        splStr = split(searchValue, " And ") 
        sVal = splStr(0) 
        if uBound(splStr) > 0 then sVal2 = splStr(1) 
        splStr = split(c, vbCrLf) : nOK = 0 : c = "" 
        for each s in splStr
            if sVal2 <> "" then
                if inStr(s, sVal) > 0 and inStr(s, sVal2) > 0 then
                    if inStr(vbCrLf & c & vbCrLf, vbCrLf & s & vbCrLf) = false then
                        nOK = nOK + 1 
                        c = c & s & "<br>" 
                    end if 
                end if 
            else
                if inStr(s, sVal) > 0 then
                    if inStr(vbCrLf & c & vbCrLf, vbCrLf & s & vbCrLf) = false then
                        nOK = nOK + 1 
                        c = c & s & "<br>" 
                    end if 
                end if 
            end if 
        next 
    end if 
    getUrlStr = c 
end function 
'���Html�з�ҳ���� (�ɼ��õ�)
function getPageConfig(httpurl, content, sType)
    dim splStr, i, s, s2,  c, url, tempUrl, arrUrl(99), urlList, pageUrl, nLen, splxx,nIndex,n3
    content = getAURL(content) 
    splStr = split(content, vbCrLf) 
    for each url in splStr
        if url <> "" then
            tempUrl = url 
            if inStr(lCase(url), "javascript:") = false and url <> "#" then
                url = replace(replace(replace(replace(replace(url, "0", ""), "1", ""), "2", ""), "3", ""), "4", "") 
                url = replace(replace(replace(replace(replace(url, "5", ""), "6", ""), "7", ""), "8", ""), "9", "") 
                if sType = "ע��" and inStr(url, "?") > 0 then url = handlSqlInUrl(url) 
                c = c & url & vbCrLf 
                nLen = inStr(vbCrLf & urlList, vbCrLf & url & "����") 
                if nLen > 0 then
                    s = mid(urlList, nLen) 
                    s = mid(s, 1, inStr(s, vbCrLf) - 1) 
                    splxx = split(s, "����") 
                    s2 = splxx(0) 
                    n3 = cint(splxx(1)) + 1 
                    urlList = replace(urlList, s, s2 & "����" & n3 & "����" & splxx(2)) 
                    pageUrl = url 
                else
                    urlList = urlList & url & "����0����" & fullHttpUrl(httpurl, tempUrl) & vbCrLf 
                end if 
            end if 
        end if 
    next 
    splStr = split(urlList, vbCrLf) : c = "" 
    if sType = "ע��" then
        for each s in splStr
            if inStr(s, "����") > 0 then
                splxx = split(s, "����") 
                url = trim(splxx(2)) 
                if inStr(url, httpurl) > 0 then
                    if inStr(vbCrLf & c, vbCrLf & url & vbCrLf) = false then
                        c = c & url & vbCrLf 
                    end if 
                end if 
            end if 
        next 
        getPageConfig = c : exit function 
    end if 
    '������ҳ����
    for each s in splStr
        if inStr(s, "����") > 0 then
            splxx = split(s, "����")  
			nIndex=cint(splxx(1))
            if nIndex > 0 then arrUrl(nIndex) = splxx(0) & "  &nbsp; | &nbsp; " & splxx(2) 
        end if 
    next 
    for i = 99 to 0 step - 1
        if arrUrl(i) <> "" then
            if inStr(arrUrl(i), httpurl) > 0 then
                c = c & arrUrl(i) & "��   ����[" & i & "]��<br>" 
            end if 
        end if 
    next 
    getPageConfig = c 
end function 
'�����ҳ��ע����ַ
function getSqlInUrl(httpurl, content, sType)
    dim splStr, i, s, s2,  c, url, tempUrl, arrUrl(99), urlList, pageUrl, nLen, splxx,n3,nIndex
    content = getAURL(content) 
    splStr = split(content, vbCrLf) 
    for each url in splStr
        if url <> "" then
            tempUrl = url 
            if inStr(url, "?") > 0 then
                c = c & url & vbCrLf 
                url = handlSqlInUrl(url) 
                nLen = inStr(vbCrLf & urlList, vbCrLf & url & "����") 
                if nLen > 0 then
                    s = mid(urlList, nLen) 
                    s = mid(s, 1, inStr(s, vbCrLf) - 1) 
                    splxx = split(s, "����") 
                    s2 = splxx(0) 
                    n3 = cint(splxx(1)) + 1 
                    urlList = replace(urlList, s, s2 & "����" & n3 & "����" & splxx(2)) 
                    pageUrl = url 
                else
                    urlList = urlList & url & "����0����" & fullHttpUrl(httpurl, tempUrl) & vbCrLf 
                end if 
            end if 
        end if 
    next 
    splStr = split(urlList, vbCrLf) 
    for each s in splStr
        if inStr(s, "����") > 0 then
            splxx = split(s, "����") 
            if n3 > 0 then
				nIndex=cint(splxx(1))
                if sType = "ע��" then
                    arrUrl(nIndex) = splxx(2) 
                else
                    arrUrl(nIndex) = splxx(0) & "  &nbsp; | &nbsp; " & splxx(2) 
                end if 
            end if 
        end if 
    next 
    c = "" 
    for i = 99 to 0 step - 1
        if arrUrl(i) <> "" then
            if sType = "ע��" then
                c = c & arrUrl(i) & vbCrLf 
            else
                c = c & arrUrl(i) & "��   ����[" & i & "]��<br>" 
            end if 
        end if 
    next 
    getSqlInUrl = c 
end function 
'����ע����ַ�����û����վע����ַ
function handlSqlInUrl(httpurl)
    dim url, splStr, i, s 
    splStr = split(httpurl, "=") 
    for i = 0 to uBound(splStr)
        if i = uBound(splStr) then url = url & "=" : exit for 
        s = splStr(i) 
        if i mod 2 = 0 then
            url = url & splStr(i) 
        else
            if inStr(s, "&") > 0 then
                url = url & "=" & mid(s, inStr(s, "&")) 
            end if 
        end if 
    next 
    handlSqlInUrl = url 
end function 

'���ɨ����������б�
function getScanFunctionNameList(content)
    on error resume next 
    dim splStr, i, isASP, isWord, sx, s, wc, zc, s1, aspCode, nYHCount, upWord, funList, s2 
    dim upWordn, tempS, dimList 
    dim yesFunction                                                                 '�����Ƿ�Ϊ��
    isASP = false                                                                  '��ASP Ĭ��Ϊ��
    yesFunction = false                                                             '�Ǻ��� Ĭ��Ϊ��
    isWord = false                                                                 '�ǵ��� Ĭ��Ϊ��
    nYHCount = 0                                                                    '˫����Ĭ��Ϊ0
    splStr = split(content, vbCrLf)                                                 '�ָ���
    'ѭ������
    for each s in splStr
        'ѭ��ÿ���ַ�
        for i = 1 to len(s)
            sx = mid(s, i, 1) 
            'Asp��ʼ
            if sx = "<" and wc = "" then                                                 '����ı�����Ϊ�� WcΪ������� ��"<%" �ų� �޸���20140412
                if mid(s, i + 1, 1) = "%" then
                    isASP = true                                                                   'ASPΪ��
                    i = i + 1                                                                       '��1�����ܼ�2��Ҫ��Ȼ<%function Test() �ͽ�ȡ����
                end if 
            'ASP����
            elseIf sx = "%" and mid(s, i + 1, 1) = ">" and wc = "" then                     'WcΪ�������
                isASP = false                                                                  'ASPΪ��
                i = i + 1                                                                       '���ܼ�2��ֻ�ܼ�1����Ϊ���ﶨ��ASPΪ�٣���������һ����ʾ����� 'ASP����Ϊ��
            end if 
            if isASP = true then
                '�����ı�
                if(sx = """" or wc <> "") then
                    '˫�����ۼ�
                    if sx = """" then nYHCount = nYHCount + 1 
                    '�ж��Ƿ�"�����
                    if nYHCount mod 2 = 0 then
                        if mid(s, i + 1, 1) <> """" then
                            wc = wc & sx 
                            s1 = right(replace(mid(s, 1, i - len(wc)), " ", ""), 1)                         '����������Ҫ�������
                            if yesFunction = true then aspCode = aspCode & wc                               '���������ۼ�
                            nYHCount = 0 : wc = ""                                                          '���
                        else
                            wc = wc & sx 
                        end if 
                    else
                        wc = wc & sx 
                    end if 
                elseIf sx = "'" then                                                            'ע�����˳�
                    if yesFunction = true then aspCode = aspCode & mid(s, i) 
                    exit for 
                '��ĸ
                elseIf checkABC(sx) = true or(sx = "_" and zc <> "") or zc <> "" then
                    zc = zc & sx 
                    s1 = lCase(mid(s & " ", i + 1, 1)) 
                    if inStr("abcdefghijklmnopqrstuvwxyz0123456789", s1) = 0 and not(s1 = "_" and zc <> "") then'����ж�
                        tempS = mid(s, i + 1) 
                        if inStr("|function|sub|", "|" & lCase(zc) & "|")>0 then
                            '������ʼ
                            if yesFunction = false and lCase(upWord) <> "end" then
                                yesFunction = true 
                                s2 = mid(s, i + 2) 
                                s2 = mid(s2, 1, inStr(s2, "(") - 1) 
                                funList = funList & s2 & vbCrLf 
                            elseIf yesFunction = true and lCase(upWord) = "end" then                        '�����һ������
                                aspCode = aspCode & zc & vbCrLf 

                                yesFunction = false 
                            end if 

                        end if 
                        upWord = zc                                                                     '��ס��ǰ����
                        if yesFunction = true then aspCode = aspCode & zc 
                        zc = "" 
                    end if 

                end if 
            end if 
            doEvents 
        next 

    next 
    getScanFunctionNameList = funList 
end function 
'��ú�������20150402
function getFunName(byVal c)
    if inStr(c, "(") > 0 then
        c = mid(c, 1, inStr(c, "(") - 1) 
        c = pHPTrim(c) 
    end if 
    getFunName = c 
end function 
'��ú��������б���
function getFunDimName(byVal c)
    dim startStr, endStr, s 
    c = lCase(c) 
    startStr = "(" 
    endStr = ")" 
    if inStr(c, startStr) > 0 and inStr(c, endStr) > 0 then
        c = strCut(c, startStr, endStr, 2) 
    end if 
    if c <> "" then
        c = replace(replace(c, "byref ", ""), "byref,", "") 
        c = replace(replace(c, "byval ", ""), "byval,", "") 
        c = replace(c, " ", "") 
    end if 
    getFunDimName = c 
end function 
'��ñ������������
function getDimName(byVal c)
    dim startStr, endStr, s 
    c = lCase(c) 
    startStr = ":" 
    if inStr(c, startStr) > 0 then
        c = mid(c, 1, inStr(c, ":") - 1) 
    end if 
    if c <> "" then
        c = replace(replace(c, "byref ", ""), "byref,", "") 
        c = replace(replace(c, "byval ", ""), "byval,", "") 
        c = replace(c, " ", "") 
    end if 
    if inStr(c, "'") > 0 then
        c = pHPTrim(mid(c, 1, inStr(c, "'") - 1))                                       '�h�������'ע�Ͳ���    20150330
    end if 
    getDimName = c 
end function 

'���JS�������������
function getVarName(byVal content)
    dim splStr, i, s, c 
    content = lCase(content) 
    if inStr(content, ";") > 0 then
        content = mid(content, 1, inStr(content, ";") - 1) 
    end if 
    splStr = split(content, ",") 
    for each s in splStr
        if instr(s, "=") > 0 then
            s = mid(s, 1, instr(s, "=") - 1) 
        end if 
        s = trim(s) 
        if c <> "" then
            c = c & "," 
        end if 
        if s <> "" then
            c = c & s 
        end if 
    next 
    'call rwend(c)
    getVarName = c 
end function 


'���Set������������� ��ʱ�ò���
function getSetName(byVal c)
    c = pHPTrim(lCase(c)) 
    c = mid(c, 1, inStr(c, "=") - 1) 
    getSetName = trim(c) 
end function 
'�滻����
function replaceDim(byVal content)
    dim splStr, s, tempS, c, lCaseS, nDimInTHNumb 
    splStr = split(content, ",") 
    for each s in splStr
        s = trim(s) 
        lCaseS = lCase(s) 
        if s <> "" then
            '�Ա�����()������
            if inStr(s, "(") > 0 then
                s = mid(s, 1, inStr(s, "(") - 1) 
            end if 
            nDimInTHNumb = inStr(lCase(replaceDimList), "," & lCaseS & "=")                  '�滻����
            if nDimInTHNumb > 0 then                                                         '�滻����
                tempS = mid(replaceDimList, nDimInTHNumb + 1) 
                tempS = mid(tempS, 1, inStr(tempS, ",") - 1) 
                tempS = mid(tempS, inStr(tempS, "=") + 1) 
                if inStr(funDim & rootFunDim, "," & tempS & ",") = false then
                    s = tempS 
                end if 
            end if 
            c = c & s & "," 
        end if 
    next 
    if c <> "" then c = left(c, len(c) - 1) 
    replaceDim = c 
end function 
'�滻��������(���Ż�����) ��ǿ��20141017
function replaceDim2(byVal dimList, dimName)
    dim ZD, splStr, i, s, nMod, nInt, c 
    replaceDim2 = dimName 
    ZD = "abcdefghijklmnopqrstuvwxyz" 
    dimName = lCase(dimName) 
    splStr = split(dimList, ",") 
    for i = 0 to uBound(splStr)
        s = splStr(i) 
        if s = dimName then
            nMod =(i) mod len(ZD) + 1 
            nInt = fix((i) / len(ZD)) 

            if nMod <> 0 then
                c = c & mid(ZD, nMod, 1) 
            end if 
            if nInt <> 0 then
                c = c & copyStr(mid(ZD, nInt, 1), nInt) 
            end if 
            'Call Echo(I,Len(ZD))
            'Call Echo("nMod",nMod)
            'Call Echo("nInt",nInt)
            'Call Echo("C",C)
            replaceDim2 = c 
            exit function 
        end if 
    next 
end function 
'�ҵ�ǰ�ļ����ظ�����������
function findFolderRepeatFunction(folderPath)
    dim filePath, s, c, content, funs, funList, allFunList, nOK, nErr, splStr, splxx, nAllOK, nAllErr, nI 
    dim errFunList, allErrFunList 
    call handlePath(folderPath)                                                     '�������·��
    c = "�����ļ���" & folderPath & vbCrLf 
    content = getDirFileList(folderPath, "") 
    splStr = split(content, vbCrLf) : nI = 0 
    for each filePath in splStr
        nI = nI + 1 
        s = nI & "��" & filePath 
        content = getFText(filePath) 
        content = getScanFunctionNameList(content)                                      '���ASP���������б�
        nOK = 0 : nErr = 0 : nAllOK = 0 : nAllErr = 0 : funList = "" : errFunList = "" : allErrFunList = "" 
        splxx = split(content, vbCrLf) 
        for each funs in splxx
            if funs <> "" then
                if inStr("|" & funList & "|", "|" & funs & "|") = 0 then
                    funList = funList & funs & "|" 
                    nOK = nOK + 1 
                else
                    errFunList = errFunList & funs & "|" 
                    nErr = nErr + 1 
                end if 
                if inStr("|" & allFunList & "|", "|" & funs & "|") = 0 then
                    allFunList = allFunList & funs & "|"                                            'ȫ������
                    nAllOK = nAllOK + 1 
                else
                    allErrFunList = allErrFunList & funs & "|" 
                    nAllErr = nAllErr + 1 
                end if 
            end if 
            doEvents 
        next 
        'Call CreateFile("allfun.txt", AllFunList)
        c = c & s & "��������" & uBound(splStr) + 1 & "�����ظ�(" & nErr & "[" & errFunList & "])ȫ�������ظ�(" & nAllErr & "[" & allErrFunList & "])" & vbCrLf 
        doEvents 
    next 
    findFolderRepeatFunction = c 
end function 
'�ҵ�ǰ�ļ��ظ����� (����)
function findFileRepeatFunction(filePath)
    findFileRepeatFunction = handleContentRepeatFunction(getFText("FilePath"), "2") 
end function 
'�ҵ�ǰ�����ظ�����
function findContentRepeatFunction(content)
    findContentRepeatFunction = handleContentRepeatFunction(content, "2") 
end function 
'���������ظ������б� sTypeΪ0Ϊ����ʾ��1Ϊ��ʾ�����б���2Ϊ��ʾ�ظ������б���3Ϊ��ʾ�������ظ������б�
function handleContentRepeatFunction(content, sType)
    dim c, funs, funList, nOK, nErr, splxx, errFunList 
    sType = cStr(sType) '��@����asp����@�� 
    content = getScanFunctionNameList(content)                                      '���ASP���������б�
    nOK = 0 : nErr = 0 
    splxx = split(content, vbCrLf) 
    for each funs in splxx
        if funs <> "" then
            if inStr("|" & funList & "|", "|" & funs & "|") = 0 then
                funList = funList & funs & "|" 
                nOK = nOK + 1 
            else
                errFunList = errFunList & funs & vbCrLf 
                nErr = nErr + 1 
            end if 
        end if 
        doEvents 
    next 
    c = "�ҵ���������" & uBound(splxx) + 1 & "�����ظ�(" & nErr & ")" & vbCrLf 
    '�����б�
    if sType = "1" or sType = "3" then
        c = c & vbCrLf & "�����б�" & vbCrLf & funList 
    end if 
    '�ظ������б�
    if errFunList <> "" and(sType = "1" or sType = "3") then
        c = c & vbCrLf & "�ظ������б�" & vbCrLf & errFunList 
    end if 
    handleContentRepeatFunction = c 
end function 
'�滻�ַ�����2 �Լ�д��һ�� ���ִ�Сд
function replace2(content, searchStr, replaceStr)
    dim leftStr, rightStr 
    if inStr(content, searchStr) > 0 then
        leftStr = mid(content, 1, inStr(content, searchStr) - 1) 
        rightStr = mid(content, len(leftStr) + len(searchStr) + 1) 
        content = leftStr & replaceStr & rightStr 
    end if 
    replace2 = content 
end function 
'�滻ȫ���ַ�����2 �Լ�д��һ�� ���ִ�Сд
function allReplace(content, searchStr, replaceStr)
    dim leftStr, rightStr, i 
    for i = 1 to 99
        if inStr(content, searchStr) > 0 then
            leftStr = mid(content, 1, inStr(content, searchStr) - 1) 
            rightStr = mid(content, len(leftStr) + len(searchStr) + 1) 
            content = leftStr & replaceStr & rightStr 
        else
            exit for 
        end if 
    next 
    allReplace = content 
end function 
'�滻һ�Σ������ִ�Сд
function replaceOneNOLU(content, searchStr, replaceStr)
    dim leftStr, rightStr, lCaseContent 
    searchStr = lCase(searchStr) 
    lCaseContent = lCase(content) 
    if inStr(lCaseContent, searchStr) > 0 then
        leftStr = mid(content, 1, inStr(lCaseContent, searchStr) - 1) 
        rightStr = mid(content, len(leftStr) + len(searchStr) + 1) 
        content = leftStr & replaceStr & rightStr 
    end if 
    replaceOneNOLU = content 
end function 
'�Ż�ASP���� ɾ�����ҿո�
function optimizeAspCode(content)
    dim splStr, s, c, i 
    splStr = split(content, vbCrLf) 
    for each s in splStr
        s = trimVbTab(s) 
        if s <> "" then c = c & trimVbTab(s) & vbCrLf 
    next 
    optimizeAspCode = c 
end function 

'��ȡ�����������
function cutDaKuoHao(content) 
	cutDaKuoHao = cutFunctionvValue(content, "{", "}") 
	
end function 
'��ȡ������С����
function cutXianKuoHao(content)
    cutXianKuoHao = cutFunctionvValue(content, "(", ")") 
end function 

'��ȡ����ֵ20150716
function cutFunctionvValue(content, startStr, endStr)
    dim n1, n2, i, s, c 
    n1 = 1 : n2 = 0 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        c = c & s 
        if s = startStr then
            n1 = n1 + 1 
        elseIf s = endStr then 
            n2 = n2 + 1 
            if n1 = n2 then
                exit for 
            end if 
        end if 
    next 
    cutFunctionvValue = c 
end function 
'����ַ�����������ִ���20150721
function getStrIntContentNumb(content, findStr)
    dim splStr 
    getStrIntContentNumb = 0 
    if inStr(content, findStr) > 0 then
        splStr = split(content, findStr) 
        getStrIntContentNumb = uBound(splStr)  
    end if 
end function 
%>  
