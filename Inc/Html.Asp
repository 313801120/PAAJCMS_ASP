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
'Html ����HTML���� (2014,1,3)
'��ʾHTML�ṹ        call rw(displayHTmL("<br>aasdfds<br>"))
'�ر���ʾHTML�ṹ   call rwend(unDisplayHtml("&lt;br&gt;aasdfds&lt;br&gt;"))

'��ʾHTML�ṹ
function displayHtml(str)
    str = replace(str, "<", "&lt;") 
    str = replace(str, ">", "&gt;") 
    displayHtml = str 
end function 
'�ر���ʾHTML�ṹ
function unDisplayHtml(str)
    str = replace(str, "&lt;", "<") 
    str = replace(str, "&gt;", ">") 
    unDisplayHtml = str 
end function 

'�����պ�HTML��ǩ(20150902)  ������ĸ����� �ڶ���
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
                endStr = right(endStr, len(endStr) - len(s) - 2)                                '����ַ���ȥ��ǰ��ǩ  -2����Ϊ����<>�����ַ�
                'ע��֮ǰ����labelName����
                labelName = mid(s, 1, inStr(s & " ", " ") - 1) 
                labelName = LCase(labelName) 
                'call eerr("s",s)

                if inStr(action, "|����A����|") > 0 then
                    s = handleHtmlAHref(s, labelName, "http://127.0.0.1/TestWeb/Web/", "����A����") '�����ɾ�html��ǩ
                elseIf inStr(action, "|����A���ӵڶ���|") > 0 then
                    s = handleHtmlAHref(s, labelName, "http://127.0.0.1/debugRemoteWeb.asp?url=", "����A����") '�����ɾ�html��ǩ
                end if 
                'call echo(s,labelName)   param��embed��Flash�õ�������embed�н�����ǩ��
                if inStr("|meta|link|embed|param|input|img|br|hr|rect|line|area|script|div|span|a|", "|" & labelName & "|") > 0 then
                    s = replace(replace(replace(replace(s, " class=""""", ""), " alt=""""", ""), " title=""""", ""), " name=""""", "") '��ʱ��ô��һ�£��Ժ�Ҫ����ϵͳ����
                    s = replace(replace(replace(replace(s, " class=''", ""), " alt=''", ""), " title=''", ""), " name=''", "") 

                    '��vb.net�����õ� Ҫ��Ȼ���ᱨ������
                    if labelName = "img" and isImgAddAlt = true then
                        if inStr(s, " alt") = 0 then
                            s = s & " alt=""""" 
                        end if 
                        s = trim(s) 
                        s = s & " /" 
                    '����<script>20160106  ��ʱ������������ȸĽ�
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
                '����javascript script����
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
'����htmlA��ǩ��Href����  ������溯��
function handleHtmlAHref(byVal content, labelName, addToHttpUrl, action)
    dim i, s, c, temp 
    dim isValue                                                               '�Ƿ�Ϊ����ֵ
    dim valueStr                                                              '�洢����ֵ
    dim yinghaoLabel                                                          '����������'"
    dim parentName                                                            '��������
    dim behindStr                                                             '����ȫ���ַ�
    dim sNotDanYinShuangYinStr                                                  '���ǵ����ź�˫�����ַ�
    action = "|" & action & "|" 
    content = replace(content & " ", vbTab, " ")                              '�˸��滻�ɿո�����һ���ո񣬷������
    content = replace(replace(content, " =", "="), " =", "=") 
    isValue = false                                                           'Ĭ������Ϊ�٣���Ϊ���ǻ�ñ�ǩ����
    for i = 1 to len(content)
        s = mid(content, i, 1)                                                       '��õ�ǰһ���ַ�
        behindStr = mid(content, i)                                                  '�����ַ�
        if s = "=" and isValue = false then                                          '��������ֵ����Ϊ=��
            isValue = true 
            valueStr = "" 
            yinghaoLabel = "" 
            if c <> "" and right(c, 1) <> " " then c = c & " " 
            parentName = LCase(temp)                           '��������תСд
            c = c & parentName & s 
            temp = "" 
        '���ֵ��һ���ַ�����Ϊ������������
        elseIf isValue = true and yinghaoLabel = "" then
            if s <> " " then
                if s <> "'" and s <> """" then
                    sNotDanYinShuangYinStr = s                                                        '���ǵ����ź�˫�����ַ�
                    s = " " 
                end if 
                yinghaoLabel = s 
            'call echo("yinghaoLabel",yinghaoLabel)
            end if 
        elseIf isValue = true and yinghaoLabel <> "" then
            'Ϊ���Ž���
            if yinghaoLabel = s then
                isValue = false 
                if labelName = "a" and parentName = "href" and inStr(action, "|����A����|") > 0 then
                    '����
                    if inStr(valueStr, "?") > 0 then
                        valueStr = replace(valueStr, "?", "WenHao") & ".html" 
                    end if 
                    if inStr("|asp|php|aspx|jsp|", "|" & LCase(mid(valueStr, inStrRev(valueStr, ".") + 1)) & "|") > 0 then
                        valueStr = valueStr & ".html" 
                    end if 
                    valueStr = addToOrAddHttpUrl(addToHttpUrl, valueStr, "�滻") 

                end if 
                'call echo("labelName",labelName)
                if yinghaoLabel = " " then
                    c = c & """" & sNotDanYinShuangYinStr & valueStr & """ "                          '׷�� ���ǵ����ź�˫�����ַ�            ��ȫ
                else
                    c = c & yinghaoLabel & valueStr & yinghaoLabel                                  '׷�� ���ǵ����ź�˫�����ַ�
                end if 
                yinghaoLabel = "" 
                sNotDanYinShuangYinStr = ""                                                       '���ǵ����ź�˫�����ַ� ���
            else
                valueStr = valueStr & s 
            end if 
        'Ϊ �ָ�
        elseIf s = " " then
            '�ݴ����ݲ�Ϊ��
            if temp <> "" then
                if left(trim(behindStr) & " ", 1) = "=" then
                '����һ���ַ�����=������
                else
                    'Ϊ��ǩ
                    if isValue = false then
                        temp = LCase(temp) & " "                                  '��ǩ��������תСд
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
'׷�ӻ��滻��ַ(20150922) �������   addToOrAddHttpUrl("http://127.0.0.1/aa/","http://127.0.0.1/4.asp","�滻") = http://127.0.0.1/aa/4.asp
function addToOrAddHttpUrl(httpurl, byVal url, action)
    dim s 
    action = "|" & action & "|" 
    if inStr(action, "|�滻|") > 0 then
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

'���HTML��ǩ�� call rwend(getHtmlLableName("<img src><a href=>",0))    ���  img
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
                endStr = right(endStr, len(endStr) - len(s) - 2)                                '����ַ���ȥ��ǰ��ǩ  -2����Ϊ����<>�����ַ�
                'ע��֮ǰ����labelName����
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

'ɾ��html����� ��ķ��� ɾ������
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
'���html����ָ��ֵ20160520 call echo("webtitle",getHtmlValue(content,"webtitle"))
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
                endStr = right(endStr, len(endStr) - len(s) - 2)                                '����ַ���ȥ��ǰ��ǩ  -2����Ϊ����<>�����ַ�
                'ע��֮ǰ����labelName����
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
                    if "web" & paramName = sType then
                        getHtmlValue = getParamValue(s, "content") 
                        exit function 
                    end if 


                end if 

            end if 
        end if 
    next 
    getHtmlValue = "" 
end function 

'��ò���ֵ20160520  call rwend(getParamValue("meta name=""keywords"" content=""{$web_keywords$}""","name"))
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

%>     


