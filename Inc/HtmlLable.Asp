<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-02-27
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'HTML ��ҳ��ǩ����

'���������url
function batchGetKuHuoValue(content)
	dim splstr,s,c,url
	splstr=split(content,",")
	for each s in splstr
		url=getKuHuoValue(s)
		if url<>"" and instr(","& c &",",","& url &",")=false then
			if c<>"" then c=c & vbcrlf
			c=c & url
		end if
	next
	batchGetKuHuoValue=c
end function


'���������url
function getKuHuoValue(content)
	dim url,tempUrl
	url=getStrCut(content,"(",")",2)
	tempUrl=phptrim(url)
	if left(tempUrl,1)="'" then
		url=mid(url,instr(url,"'")+1)
	elseif left(tempUrl,1)="""" then
		url=mid(url,instr(url,"""")+1)
	end if
	if right(tempUrl,1)="'" then
		url=mid(url,1,instrrev(url,"'")-1)
	elseif right(tempUrl,1)="""" then
		url=mid(url,1,instrrev(url,"""")-1)
	end if
	getKuHuoValue=url
end function

 
 '����
'http://www.baidu.com/aaa, :html_*.html   =   html_aaa.html
'http://www.baidu.com/aaa, :11*22.jpg   =   11aaa22.jpg
'http://www.baidu.com/aaa/, :11*22.jpg   =   11aaa22.jpg
'http://www.baidu.com/aaa.gif, :11*22.jpg   =   11aaa22.gif
'http://www.baidu.com/, :11*22.jpg   =   1122.jpg
'http://www.baidu.com/yzm.asp?v1, :*.jpg   =   yzm.jpg
'�����ַ��������ļ�����    �������и�������Ǵ������ļ���������ʹ�������ļ����ƣ����Ժ�İ�
function getUrlToFileName(url,sType) 
 	dim content,splstr,fileName,fileType,splxx
	dim filePrefixName,fileSuffixName		'�ļ���ǰ������󲿷�
	dim customFilePrefixName,customFileSuffixName		'�Զ����ļ���ǰ������󲿷�
	 
	'Ĭ���ļ���ģ�壬���ļ���׺    aa*bb.jpg  ��ǰ��*ΪҪ�滻���ļ�ǰ��������
	if instr(stype,":")>0 then
		stype=mid(sType,instr(stype,":")+1)
		splxx=split(stype & ".",".")
		customFilePrefixName=splxx(0)				
		customFileSuffixName=splxx(1)
	end if 
	
	content=replace(url,"\","/")
	splstr=split(content,"/")
	fileName=splstr(ubound(splstr))
	'�ų�?�����ַ�
	if instr(fileName,"?")>0 then
		fileName=mid(fileName,1,instr(fileName,"?")-1)
	end if
	'�ļ���Ϊ�� ��ʹ����һ����Ŀ¼��
	if fileName="" and ubound(splstr)-1>=0 then
		fileName=splstr(ubound(splstr)-1)
		'�ų����������ļ�����
		if instr(fileName,".")>0 then
			fileName=""
		end if
	end if
	if instr(fileName,".")>0 then
		filePrefixName=getFileAttr(fileName,"3")
		fileSuffixName=lcase(getFileAttr(fileName,"4"))
	else
		filePrefixName=fileName
		fileSuffixName=customFileSuffixName
	end if 
	 
	'�ų�ͼƬ���Զ����ű�   ��֤�� 
	if instr("|asp|aspx|php|jsp|cgi|action|","|"& fileSuffixName &"|")>0 and instr("|asp|aspx|php|jsp|cgi|","|"& customFileSuffixName &"|")=false then
		fileSuffixName=customFileSuffixName		'��.asp  ת���Զ�����.html
	end if 
	'aa_bb*.jpg  �滻���� 
	getUrlToFileName=replace(customFilePrefixName,"*",filePrefixName) & "." & fileSuffixName  
end function
 
 
 
 
 
 
 
 
 
 
 
 
'=========================================================================== ֮ǰ�ɰ�



'����������ַ���� labelType=|*|link|img|a|script|embed|param|meta|
function handleConentUrl(httpurl, content, labelType, byref pubAHrefList, byref pubATitleList)
    dim i, s, yuanStr, tempS, lalType, nLen, lalStr, c 
    labelType = "|" & labelType & "|" 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        if s = "<" then
            yuanStr = mid(content, i) 
            tempS = LCase(yuanStr) 
            tempS = replace(replace(tempS, chr(10), " " & vbCrLf), chr(13), " " & vbCrLf) '�ô���ͼƬ�زĸ�����  ����  <img����src=""  Ҳ���Ի�� 20150714
            lalStr = mid(yuanStr, 1, inStr(yuanStr, ">")) 
            lalType = LCase(mid(tempS, 1, inStr(tempS, " "))) 
            if lalType = "<link " and(inStr(labelType, "|link|") > 0 or inStr(labelType, "|*|") > 0) then
                nLen = len(lalStr) - 1 
                c = c & handleLink(httpurl, lalStr, "href", "", "", pubAHrefList, pubATitleList) 
                i = i + nLen 
            elseIf lalType = "<img " and(inStr(labelType, "|img|") > 0 or inStr(labelType, "|*|") > 0) then
                nLen = len(lalStr) - 1 
                c = c & handleLink(httpurl, lalStr, "src", "", "", pubAHrefList, pubATitleList) 
                i = i + nLen 
            elseIf lalType = "<a " and(inStr(labelType, "|a|") > 0 or inStr(labelType, "|*|") > 0) then
                nLen = len(lalStr) - 1 
                'û��javascript�����У����ǻ����в���֮��
                if inStr(LCase(lalStr), "javascript:") = 0 then
                    c = c & handleLink(httpurl, lalStr, "href", "", "", pubAHrefList, pubATitleList) 
                else
                    c = c & lalStr 
                end if 
                i = i + nLen 
            elseIf lalType = "<script " and(inStr(labelType, "|script|") > 0 or inStr(labelType, "|*|") > 0) then
                nLen = len(lalStr) - 1 
                if inStr(LCase(lalStr), "src") > 0 then
                    c = c & handleLink(httpurl, lalStr, "src", "", "", pubAHrefList, pubATitleList) 
                else
                    c = c & lalStr 
                end if 
                i = i + nLen 
            elseIf lalType = "<embed " and(inStr(labelType, "|embed|") > 0 or inStr(labelType, "|*|") > 0) then
                nLen = len(lalStr) - 1 
                c = c & handleLink(httpurl, lalStr, "src", "", "", pubAHrefList, pubATitleList) 
                i = i + nLen 
            elseIf lalType = "<param " and(inStr(labelType, "|param|") > 0 or inStr(labelType, "|*|") > 0) then
                nLen = len(lalStr) - 1 
                if inStr(LCase(lalStr), "movie") > 0 then
                    c = c & handleLink(httpurl, lalStr, "value", "", "", pubAHrefList, pubATitleList) 
                else
                    c = c & lalStr 
                end if 
                i = i + nLen 
            elseIf lalType = "<meta " and(inStr(labelType, "|meta|") > 0 or inStr(labelType, "|*|") > 0) then
                nLen = len(lalStr) - 1 
                '�滻�ؼ���
                if inStr(LCase(lalStr), "keywords") > 0 then
                    c = c & handleLink(httpurl, lalStr, "content", cfg_webKeywords, "", pubAHrefList, pubATitleList) 
                '�滻��վ����
                elseIf inStr(LCase(lalStr), "description") > 0 then
                    c = c & handleLink(httpurl, lalStr, "content", cfg_webDescription, "", pubAHrefList, pubATitleList) 
                else
                    c = c & lalStr 
                end if 
                i = i + nLen 

            elseIf lalType = "<input " and(inStr(labelType, "|src|") > 0 or inStr(labelType, "|*|") > 0) and instr(lcase(lalStr), " src=") > 0 then
                nLen = len(lalStr) - 1 
                c = c & handleLink(httpurl, lalStr, "src", "", "", pubAHrefList, pubATitleList) 
                i = i + nLen 

            elseIf inStr(labelType, "|imgstyle|") > 0 or inStr(labelType, "|*|") > 0 then
                nLen = len(lalStr) - 1 
                if instr(lcase(lalStr), "url") > 0 and instr(lcase(lalStr), ":") > 0 and instr(lcase(lalStr), "(") > 0 then
                    c = c & handleLink(httpurl, lalStr, "style", "", "", pubAHrefList, pubATitleList) 
                    i = i + nLen 
                else
                    c = c & s 
                end if 
            else
                c = c & s 
            end if 
        else
            c = c & s 
        end if 
        doEvents 
    next 
    handleConentUrl = c 
end function 




'�������ӵ�ַ HttpUrl=׷����ַ��Content=����  SType=����
'�滻Ŀ¼����  call rw(HandleLink("Js/", "111111<script src=""js/Jquery.Min.js""></"& "script>","src", "newurl", "replaceDir","",""))
function handleLink(httpurl, content, sType, setStr, action, byref pubAHrefList, byref pubATitleList)
    'on error resume next
    dim i, s, c, tempContent, findUrl, handleUrl, startStr, endStr, s1, s2, tempHttpUrl, nLen 
    tempHttpUrl = httpurl 
    action = LCase(action) 
    content = replace(replace(content, "= ", "="), "= ", "=") 
    content = replace(replace(content, " =", "="), " =", "=") 
    tempContent = LCase(content) 
    handleLink = "" 
    'û�������˳�
    if inStr(tempContent, " href=") = 0 and inStr(tempContent, " src=") = 0 and sType <> "style" then
        handleLink = "" 
        exit function 
    elseIf inStr(tempContent, " href=\""") > 0 then
        content = replace(content, "\""", """") : tempContent = LCase(content) 
    end if 
    '˫����
    startStr = sType & "=""" 
    endStr = """" 
    if inStr(tempContent, startStr) > 0 and inStr(tempContent, endStr) > 0 then
        'call echo("��ʾ","1")
        findUrl = strCut(content, startStr, endStr, 2) 
        if setStr <> "" then
            handleUrl = setStr 
        else
            '�滻Ŀ¼
            handleUrl = handleLinkAction(httpurl, findUrl, action) 

            pubAHrefList = pubAHrefList & hanldeStyleBackgroundUrl(handleUrl, "", "") & vbCrLf 
            '���ӱ���
            nLen = inStr(content, ">") 
            s2 = right(content, len(content) - nLen) 
            if instr(s2, "</") > 0 then
                s2 = mid(s2, 1, inStrRev(s2, "</") - 1) 
            end if 
            s2 = replace(s2, vbCrLf, "�����С�") 
            pubATitleList = pubATitleList & s2 & vbCrLf 
        end if 
        if findUrl <> handleUrl then
            'ǿǿǿ���滻
            nLen = inStr(content, startStr) - 1 + len(startStr)                             '��������TempContent������Content��Ϊ�д�Сд������20140726
            s2 = right(content, len(content) - nLen) 
            s2 = mid(s2, inStr(s2, endStr)) 
            s1 = left(content, nLen) 
            content = s1 & handleUrl & s2 
        end if 
        if action = "url" then
            handleLink = handleUrl 
        else
            handleLink = content 
        end if 
        exit function 
    end if 
    '������
    startStr = sType & "='" 
    endStr = "'" 
    if inStr(tempContent, startStr) > 0 and inStr(tempContent, endStr) > 0 then
        'call echo("��ʾ","2")
        findUrl = strCut(tempContent, startStr, endStr, 2) 
        if setStr <> "" then
            handleUrl = setStr 
        else
            '�滻Ŀ¼
            handleUrl = handleLinkAction(httpurl, findUrl, action) 

            pubAHrefList = pubAHrefList & hanldeStyleBackgroundUrl(handleUrl, "", "") & vbCrLf 
            '���ӱ���
            nLen = inStr(content, ">") 
            s2 = right(content, len(content) - nLen) 
            if instr(s2, "</") > 0 then
                s2 = mid(s2, 1, inStrRev(s2, "</") - 1) 
            end if 
            s2 = replace(s2, vbCrLf, "�����С�") 
            pubATitleList = pubATitleList & s2 & vbCrLf 
        end if 
        if findUrl <> handleUrl then
            'ǿǿǿ���滻
            nLen = inStr(content, startStr) - 1 + len(startStr) 
            s2 = right(content, len(content) - nLen) 
            s2 = mid(s2, inStr(s2, endStr)) 
            s1 = left(content, nLen) 
            content = s1 & handleUrl & s2 
        end if 
        if action = "url" then
            handleLink = handleUrl 
        else
            handleLink = content 
        end if 
        exit function 
    end if 
    '������
    startStr = sType & "=" 
    endStr = ">"                                                                    '�������֮�ҵ� �ո񻻳�>
    if inStr(tempContent, startStr) > 0 and inStr(tempContent, endStr) > 0 then
        findUrl = strCut(tempContent, startStr, endStr, 2) 

        if setStr <> "" then
            handleUrl = setStr 
        else
            '�滻Ŀ¼
            handleUrl = handleLinkAction(httpurl, findUrl, action) 

            pubAHrefList = pubAHrefList & hanldeStyleBackgroundUrl(handleHttpUrl(handleUrl), "", "") & vbCrLf 
            '���ӱ���
            nLen = inStr(content, ">") 
            s2 = right(content, len(content) - nLen) 
            if instr(s2, "</") > 0 then
                s2 = mid(s2, 1, inStrRev(s2, "</") - 1) 
            end if 
            s2 = replace(s2, vbCrLf, "�����С�") 
            pubATitleList = pubATitleList & s2 & vbCrLf 
        end if 
        if findUrl <> handleUrl then
            'ǿǿǿ���滻
            nLen = inStr(content, startStr) - 1 + len(startStr) 
            s2 = right(content, len(content) - nLen) 
            s2 = mid(s2, inStr(s2, endStr)) 
            s1 = left(content, nLen) 
            content = s1 & handleUrl & s2 
        end if 
        if action = "url" then
            handleLink = handleUrl 
        else
            handleLink = content 
        end if 
        exit function 
    end if 

    if action <> "url" then handleLink = content 
    call createAddFile("���������б�.txt", httpurl & vbCrLf & content & vbCrLf & sType & vbCrLf & setStr & vbCrLf & action & vbCrLf & "----------------------" & vbCrLf) 
end function 
'�������Ӷ���
function handleLinkAction(httpurl, findUrl, action)
    dim tempAction, handleUrl, fileName 
    tempAction = action 
    action = lcase(action) 
    handleUrl = fullHttpUrl(httpurl, findUrl) 
    '�滻Ŀ¼
    if action = "replacedir" then
        handleUrl = tempHttpUrl & getFileAttr(handleUrl,"2")			'����phpд�� 
    '�����ļ�����  ���÷�վ��
    elseif action = "handlefilename" then
        fileName = getHandleUrlValue(findUrl, "����") 
        handleUrl = "index_" & fileName & ".html" 
    end if 
    handleLinkAction = handleUrl 
end function



'�����ļ��������ַ������滻
function hRFileName(url)
	hRFileName="[##"& url &"##]"
	'hRFileName="[##"& md5(url,2) &"##]"			
end function 
%> 
