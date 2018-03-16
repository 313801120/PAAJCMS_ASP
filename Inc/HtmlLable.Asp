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
'HTML 网页标签处理

'获得括弧里url
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


'获得括弧里url
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

 
 '案例
'http://www.baidu.com/aaa, :html_*.html   =   html_aaa.html
'http://www.baidu.com/aaa, :11*22.jpg   =   11aaa22.jpg
'http://www.baidu.com/aaa/, :11*22.jpg   =   11aaa22.jpg
'http://www.baidu.com/aaa.gif, :11*22.jpg   =   11aaa22.gif
'http://www.baidu.com/, :11*22.jpg   =   1122.jpg
'http://www.baidu.com/yzm.asp?v1, :*.jpg   =   yzm.jpg
'获得网址处理后的文件名称    这里面有个问题就是创建的文件名不可以使用特殊文件名称，待以后改吧
function getUrlToFileName(url,sType) 
 	dim content,splstr,fileName,fileType,splxx
	dim filePrefixName,fileSuffixName		'文件名前部分与后部分
	dim customFilePrefixName,customFileSuffixName		'自定义文件名前部分与后部分
	 
	'默认文件名模板，与文件后缀    aa*bb.jpg  分前后，*为要替换的文件前部分名称
	if instr(stype,":")>0 then
		stype=mid(sType,instr(stype,":")+1)
		splxx=split(stype & ".",".")
		customFilePrefixName=splxx(0)				
		customFileSuffixName=splxx(1)
	end if 
	
	content=replace(url,"\","/")
	splstr=split(content,"/")
	fileName=splstr(ubound(splstr))
	'排除?后面字符
	if instr(fileName,"?")>0 then
		fileName=mid(fileName,1,instr(fileName,"?")-1)
	end if
	'文件名为空 则使用上一级的目录名
	if fileName="" and ubound(splstr)-1>=0 then
		fileName=splstr(ubound(splstr)-1)
		'排除域名作用文件名称
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
	 
	'排除图片是以动作脚本   验证码 
	if instr("|asp|aspx|php|jsp|cgi|action|","|"& fileSuffixName &"|")>0 and instr("|asp|aspx|php|jsp|cgi|","|"& customFileSuffixName &"|")=false then
		fileSuffixName=customFileSuffixName		'让.asp  转成自定义如.html
	end if 
	'aa_bb*.jpg  替换名称 
	getUrlToFileName=replace(customFilePrefixName,"*",filePrefixName) & "." & fileSuffixName  
end function
 
 
 
 
 
 
 
 
 
 
 
 
'=========================================================================== 之前旧版



'让内容中网址完整 labelType=|*|link|img|a|script|embed|param|meta|
function handleConentUrl(httpurl, content, labelType, byref pubAHrefList, byref pubATitleList)
    dim i, s, yuanStr, tempS, lalType, nLen, lalStr, c 
    labelType = "|" & labelType & "|" 
    for i = 1 to len(content)
        s = mid(content, i, 1) 
        if s = "<" then
            yuanStr = mid(content, i) 
            tempS = LCase(yuanStr) 
            tempS = replace(replace(tempS, chr(10), " " & vbCrLf), chr(13), " " & vbCrLf) '让处理图片素材更完整  比如  <img换行src=""  也可以获得 20150714
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
                '没有javascript就运行，但是还是有不足之处
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
                '替换关键词
                if inStr(LCase(lalStr), "keywords") > 0 then
                    c = c & handleLink(httpurl, lalStr, "content", cfg_webKeywords, "", pubAHrefList, pubATitleList) 
                '替换网站描述
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




'处理链接地址 HttpUrl=追加网址，Content=内容  SType=类型
'替换目录方法  call rw(HandleLink("Js/", "111111<script src=""js/Jquery.Min.js""></"& "script>","src", "newurl", "replaceDir","",""))
function handleLink(httpurl, content, sType, setStr, action, byref pubAHrefList, byref pubATitleList)
    'on error resume next
    dim i, s, c, tempContent, findUrl, handleUrl, startStr, endStr, s1, s2, tempHttpUrl, nLen 
    tempHttpUrl = httpurl 
    action = LCase(action) 
    content = replace(replace(content, "= ", "="), "= ", "=") 
    content = replace(replace(content, " =", "="), " =", "=") 
    tempContent = LCase(content) 
    handleLink = "" 
    '没有链接退出
    if inStr(tempContent, " href=") = 0 and inStr(tempContent, " src=") = 0 and sType <> "style" then
        handleLink = "" 
        exit function 
    elseIf inStr(tempContent, " href=\""") > 0 then
        content = replace(content, "\""", """") : tempContent = LCase(content) 
    end if 
    '双引号
    startStr = sType & "=""" 
    endStr = """" 
    if inStr(tempContent, startStr) > 0 and inStr(tempContent, endStr) > 0 then
        'call echo("提示","1")
        findUrl = strCut(content, startStr, endStr, 2) 
        if setStr <> "" then
            handleUrl = setStr 
        else
            '替换目录
            handleUrl = handleLinkAction(httpurl, findUrl, action) 

            pubAHrefList = pubAHrefList & hanldeStyleBackgroundUrl(handleUrl, "", "") & vbCrLf 
            '链接标题
            nLen = inStr(content, ">") 
            s2 = right(content, len(content) - nLen) 
            if instr(s2, "</") > 0 then
                s2 = mid(s2, 1, inStrRev(s2, "</") - 1) 
            end if 
            s2 = replace(s2, vbCrLf, "【换行】") 
            pubATitleList = pubATitleList & s2 & vbCrLf 
        end if 
        if findUrl <> handleUrl then
            '强强强旱替换
            nLen = inStr(content, startStr) - 1 + len(startStr)                             '这里面用TempContent而不用Content因为有大小写在里面20140726
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
    '单引号
    startStr = sType & "='" 
    endStr = "'" 
    if inStr(tempContent, startStr) > 0 and inStr(tempContent, endStr) > 0 then
        'call echo("提示","2")
        findUrl = strCut(tempContent, startStr, endStr, 2) 
        if setStr <> "" then
            handleUrl = setStr 
        else
            '替换目录
            handleUrl = handleLinkAction(httpurl, findUrl, action) 

            pubAHrefList = pubAHrefList & hanldeStyleBackgroundUrl(handleUrl, "", "") & vbCrLf 
            '链接标题
            nLen = inStr(content, ">") 
            s2 = right(content, len(content) - nLen) 
            if instr(s2, "</") > 0 then
                s2 = mid(s2, 1, inStrRev(s2, "</") - 1) 
            end if 
            s2 = replace(s2, vbCrLf, "【换行】") 
            pubATitleList = pubATitleList & s2 & vbCrLf 
        end if 
        if findUrl <> handleUrl then
            '强强强旱替换
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
    '无引号
    startStr = sType & "=" 
    endStr = ">"                                                                    '这里面把之家的 空格换成>
    if inStr(tempContent, startStr) > 0 and inStr(tempContent, endStr) > 0 then
        findUrl = strCut(tempContent, startStr, endStr, 2) 

        if setStr <> "" then
            handleUrl = setStr 
        else
            '替换目录
            handleUrl = handleLinkAction(httpurl, findUrl, action) 

            pubAHrefList = pubAHrefList & hanldeStyleBackgroundUrl(handleHttpUrl(handleUrl), "", "") & vbCrLf 
            '链接标题
            nLen = inStr(content, ">") 
            s2 = right(content, len(content) - nLen) 
            if instr(s2, "</") > 0 then
                s2 = mid(s2, 1, inStrRev(s2, "</") - 1) 
            end if 
            s2 = replace(s2, vbCrLf, "【换行】") 
            pubATitleList = pubATitleList & s2 & vbCrLf 
        end if 
        if findUrl <> handleUrl then
            '强强强旱替换
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
    call createAddFile("出错内容列表.txt", httpurl & vbCrLf & content & vbCrLf & sType & vbCrLf & setStr & vbCrLf & action & vbCrLf & "----------------------" & vbCrLf) 
end function 
'处理链接动作
function handleLinkAction(httpurl, findUrl, action)
    dim tempAction, handleUrl, fileName 
    tempAction = action 
    action = lcase(action) 
    handleUrl = fullHttpUrl(httpurl, findUrl) 
    '替换目录
    if action = "replacedir" then
        handleUrl = tempHttpUrl & getFileAttr(handleUrl,"2")			'兼容php写法 
    '处理文件名称  配置仿站用
    elseif action = "handlefilename" then
        fileName = getHandleUrlValue(findUrl, "名称") 
        handleUrl = "index_" & fileName & ".html" 
    end if 
    handleLinkAction = handleUrl 
end function



'处理文件加特殊字符，好替换
function hRFileName(url)
	hRFileName="[##"& url &"##]"
	'hRFileName="[##"& md5(url,2) &"##]"			
end function 
%> 

