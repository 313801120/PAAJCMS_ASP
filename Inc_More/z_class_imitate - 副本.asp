<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-03-15
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 

class imitateWeb
    '������ַ�б�.txt   ���ļ���Ҫɾ���������������ʱ�� �ظ�����ͼƬ���ļ�
	dim version																		'�汾
	dim defaultIEContent															'Ĭ��IE��ʾ����
    dim templateName                                                                'ģ������
	dim templateDir																	'ģ��·��
	dim toTemplateDir																'����ģ�嵽����Ŀ¼
    dim isWebToTemplateDir                                                          '�ļ��Ƿ�ŵ�ģ���ļ�����
    
	dim serverHttpUrl																'��������ַ
    dim httpurl 
    dim newWebDir                                                                   '����վĿ¼
	dim debugWebDir																	'������վĿ¼
	dim createHtmlWebDir															'����html��վĿ¼
	dim cacheFolderPath																'�����ļ���·��
    dim customCharSet                                                               '�Զ������
    dim isGetHttpUrl                                                                'getHttpUrl��ʽ�������
    dim isMakeWeb                                                                   '����WEB�ļ���
    dim isMakeTemplate                                                              '�Ƿ�����ģ���ļ���
    dim isPackWeb                                                                   '�Ƿ���WEB�ļ��� xml
    dim isPackWebZip                                                                '�Ƿ���WEB�ļ��� zip
    dim isPackTemplate                                                              '�Ƿ���ģ���ļ��� xml
    dim isPackTemplateZip                                                           '�Ƿ���ģ���ļ��� zip
    dim isFormatting                                                                '�Ƿ��ʽ��HTML
    dim isUniformCoding                                                             '�Ƿ�ͳһ����
    dim isAddWebTitleKeyword                                                        '׷����վ������ؼ���
    dim isOnMsg                                                                     '�Ƿ���������Ϣ
    dim pubAHrefList 
    dim isUTF8Down                                                                  '��utf8��ʽ����
    dim downUrlList                                                                 '���ص�URL�б�
    dim fileNameList 
    dim saveFolderName                                                              '�����ļ�������
    dim downCountSize                                                               '�����ܴ�С
    dim downFileTypeList                                                            '�����ļ������б�
    dim isDownServer                                                                '�Ƿ�ӷ�������ʼ����
    dim isDownCache                                                                 '�Ƿ���û������ط�ʽ
    dim downImagesList                                                              '����ͼƬ�б�
    dim downImageAddMyInfoType                                                      '����ͼƬ�����Լ�����Ϣ
	dim isDeleteErrorImages															'�Ƿ�ɾ��ɾ��ͼƬ
	dim deleteErrorImagesList														'ɾ������ͼƬ�б�
	dim isStop																		'ֹͣ 

    dim jsListC, cssListC, imgListC, swfListC, otherListC, urlListC, errInfoC, downC, nDownCount '��¼���ػ�����Ϣ


    '���캯�� ��ʼ��
    sub class_Initialize()
        call loadDefaultConfig() 
    end sub 


    '����Ĭ������
    sub loadDefaultConfig()
		version = "beta version"																	'����汾 v1.1
		serverHttpUrl="http://aa/server_FangZan.Asp"									'�ȸĳ�sharembweb.com
        httpurl = "http://sharembweb.com/" 	
        newWebDir = "newweb/"                                                           '����վĿ¼
		debugWebDir="dubug/"															'������վĿ¼
		createHtmlWebDir="web/"													'������վĿ¼ 
        customCharSet = "�Զ����"                                                      '����
        isGetHttpUrl = false                                                            'getHttpUrl��ʽ�������
        isMakeWeb = true                                                                '����WEB�ļ���
        isMakeTemplate = false                                                          '�Ƿ�����ģ���ļ���
        isPackWeb = false                                                               '�Ƿ���WEB�ļ��� xml
        isPackWebZip = false                                                            '�Ƿ���WEB�ļ��� zip
        isWebToTemplateDir = true                                                       '�ļ��Ƿ�ŵ�ģ���ļ�����
        toTemplateDir = "/templates/"                                                   'ģ�帴�Ƶ�Ŀ¼
        isPackTemplate = false                                                          '�Ƿ���ģ���ļ��� xml
        isPackTemplateZip = false                                                       '�Ƿ���ģ���ļ��� zip

        isFormatting = false                                                            '�Ƿ��ʽ��HTML
        isUniformCoding = true                                                          '�Ƿ�ͳһ����
        isAddWebTitleKeyword = true                                                     '׷����վ������ؼ���
        isOnMsg = false                                                                 '�Ƿ���������Ϣ
        templateName = getWebSiteName(httpurl)                                          'ģ�����ƴ���ַ�л��
        isUTF8Down = false                                                              'utf8��ʽ����Ϊ��
        downCountSize = 0                                                               '�����ܴ�С
        downFileTypeList = "|*|"                                                        '�����ļ������б�
        isDownServer = true                                                            '�Ƿ���Դӷ�����������
        isDownCache = true                                                             'Ĭ�ϲ��������ػ���
        downImageAddMyInfoType = "2"                                                    '����ͼƬ�����Լ�����Ϣ
		cacheFolderPath=""																'����Ŀ¼ Ϊ�����ڵ�ǰĿ¼�´���һ��cacheĿ¼
		call createDirFolder(cacheFolderPath)
		defaultIEContent="[defaultIEContent]"											'Ĭ��IE��ʾ����
		isDeleteErrorImages=true														'ɾ������ͼƬ(webĿ¼��templates)
		isStop=false																	'Ĭ��ֹͣΪ��
    end sub 


    '������վ
    sub downweb()
        dim s,  websiteFilePath, websiteFileContent
        dim filePath, toFilePath, fileList, splList, fileName, splStr, startTime, splUrl 
        startTime = now()                                                               '��ʼʱ��


        if phptrim(httpurl) = "" then
            call eerr("ֹͣ", "��ַΪ��") 
        end if 

		isStop=false			'ֹͣΪ��

        if saveFolderName <> "" then
            newWebDir = newWebDir & saveFolderName & "/" 
        else
            newWebDir = newWebDir & getWebSiteCleanName(httpurl) & "/"                  '������ setfilename(getwebsite(httpurl))
        end if 
        call createfolder(newWebDir) 		
		 

        '��������Ŀ¼
		debugWebDir=newWebDir & debugWebDir & "/"
        call createDirFolder(debugWebDir) 
        '��������Ŀ¼ ǰ���ǿ������� �� Ĭ�ϻ����ļ���Ϊ��
		if cacheFolderPath="" and isDownCache=true then 
			cacheFolderPath=newWebDir & "cache/"
        	call createDirFolder(cacheFolderPath) 
		end if
		 
		'����html��վĿ¼
		if isMakeWeb=true then
			createHtmlWebDir=newWebDir & "/web/"
			call createDirFolder(createHtmlWebDir)  
 		end if
		'ģ��Ŀ¼
		if isWebToTemplateDir=true then
			templateDir=newWebDir & "templates/" & templateName & "/"			'���ɵ�ǰģ��Ŀ¼
			toTemplateDir=toTemplateDir & "/" & templateName & "/"				'���Ƶ�ָ��ģ��Ŀ¼
        	call createDirFolder(templateDir)  
		end if 
		

        downUrlList = getftext(newWebDir & "/������ַ�б�.txt") 
        fileNameList = getftext(newWebDir & "/�����б�.txt") 
        call echo("fileNameList", fileNameList) 



        '��������
        splUrl = split(httpurl, vbCrLf) 
        for each httpurl in splUrl
            httpurl = phptrim(httpurl) 
            if getwebsite(httpurl) <> "" then
                call imitateWeb()
				'ֹͣ
				if isStop=true then
					exit sub
				end if
                if downUrlList <> "" then
                    downUrlList = downUrlList & vbCrLf 
                    pubAHrefList = pubAHrefList & vbCrLf 
                end if 
            end if 
        next 
		
		if isDeleteErrorImages=true then
			deleteErrorImagesList= getErrorImagesList()
		end if
		'call echo("deleteErrorImagesList",deleteErrorImagesList)

        '����ģ��Ŀ¼
        if isMakeTemplate = true then 
            call webBeautify(templateDir, true)                  '����
            '��ģ�帴�Ƶ�/template�ļ�����
            if isWebToTemplateDir = true then
                if checkFolder(toTemplateDir) = true then
                    call echored("ģ���ļ��д��ڣ�����ģ���ﲻ���ڵ��ļ���", toTemplateDir) 

                    splStr = split("|images|js|css", "|") 
                    for each s in splStr
                        if s <> "" then
                            s = "/" & s 
                        end if 
                        fileList = getDirFileNameList(debugWebDir & "/" & templateName & "/" & s, "") 
                        splList = split(fileList, vbCrLf) 
                        for each fileName in splList
                            if fileName <> "" then
                                filePath = debugWebDir & toTemplateDir & s & "/" & fileName 
                                toFilePath = toTemplateDir & s & "/" & fileName 
                                if checkFile(toFilePath) = false then
                                    call copyFile(filePath, toFilePath) 
                                    call echored("�����ļ��ɹ�", toFilePath) 
                                end if 
                            end if 
                        next 
                    next 
                else
                    'call deleteFolder(toTemplateDir)        '��Ϊ�����Ѿ����ж���
                    call copyFolder(templateDir, toTemplateDir) 
                    call copyFolder("/Data/WebData", toTemplateDir & "/WebData") 
                    websiteFilePath = toTemplateDir & "/WebData/website.txt" 
                    websiteFileContent = getftext(websiteFilePath) 
                    s = getStrCut(websiteFileContent, "��webtemplate��", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "��webtemplate��" & templateDir & templateName & vbCrLf) 
                    s = getStrCut(websiteFileContent, "��webimages��", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "��webimages��" & templateDir & templateName & "/Images" & vbCrLf) 
                    s = getStrCut(websiteFileContent, "��webcss��", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "��webcss��" & templateDir & templateName & "/Css" & vbCrLf) 
                    s = getStrCut(websiteFileContent, "��webjs��", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "��webjs��" & templateDir & templateName & "/Js" & vbCrLf) 

                    s = getStrCut(websiteFileContent, "��webtitle��", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "��webtitle��" & templateName & "����" & vbCrLf) 
                    s = getStrCut(websiteFileContent, "��webkeywords��", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "��webkeywords��" & templateName & "�ؼ���" & vbCrLf) 
                    s = getStrCut(websiteFileContent, "��webdescription��", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "��webdescription��" & templateName & "����" & vbCrLf) 

                    call createfile(websiteFilePath, websiteFileContent) 
                    s = debugWebDir & toTemplateDir & "   ==>>  " & toTemplateDir 
                    call echoYellowB("�ŵ�ģ���ļ�����", s) 
                end if 
            end if 
        end if 
        '����WEBĿ¼
        if isMakeWeb = true then
            call webBeautify(createHtmlWebDir, false) 
        end if 
        s = vBRunTimer(startTime) & "�������ܴ�С��" & printSpaceValue(downCountSize) & "��" 
        call rw("<title>" & s & "</title>") 
        call echo("�ɹ�", "��վ���" & vBRunTimer(startTime) & "�������ܴ�С��" & printSpaceValue(downCountSize) & "��") 
    end sub 

    '��վ
    sub imitateWeb()
        dim content, fileName, c 
        dim htmlFilePath, htmlBaseFilePath, createFilePath, htmlFileSize,char_Set,s
        if httpurl = "" then
            call echo("����", "��ַΪ��" & httpurl) 
            exit sub 
        end if 
        '�����ļ���
        call createfolder(debugWebDir) 

        call hr() 
        call echored("ģ����ַ", httpurl) 
		'����htmlҳ��
        htmlFilePath = handleDown("", httpurl, getUrlToFileName(httpurl, ":html_*.html"), "") 'Դ�ļ�·��
        htmlBaseFilePath = debugWebDir & "/" & getUrlToFileName(httpurl, ":base_*.html")  'base�ļ�
        createFilePath = debugWebDir & "/" & getUrlToFileName(httpurl, ":index_*.html")   '�����ļ�
		
		if getFSize(htmlFilePath)=0 then
			call echoRedB("�ļ�����Ϊ��",htmlFilePath)
			exit sub
		end if
		
        content = readFile(htmlFilePath, char_Set)
		char_Set=customCharSet
		'�ǳ���Ҫ���Ժ󿴵������ס��Ҫɾ���ˣ�20161011
		if customCharSet="�Զ����" then
			doevents
			char_Set=getFileCharset(htmlFilePath)
			content = readFile(htmlFilePath, char_Set)		'�������Ϊ���Զ��ұ���
			call echo(htmlFilePath,char_Set)
			if char_Set="gb2312" then
				s=pubHtmlObj.handleHtmlLabel(content,"meta","content","**http-equiv=Content-Type","get","lcase(*)")
				if instr(s,"utf-8")>0 then  
					content=pubHtmlObj.handleHtmlLabel(content,"meta","content","**http-equiv=Content-Type","set","text/html; charset=gb2312")
					call echo("����","���������ļ����벻һ��1")
				end if
			elseif char_Set="utf-8" then
				s=pubHtmlObj.handleHtmlLabel(content,"meta","content","**http-equiv=Content-Type","get","lcase(*)")
				if instr(s,"gb2312")>0 then  
					content=pubHtmlObj.handleHtmlLabel(content,"meta","content","**http-equiv=Content-Type","set","text/html; charset=utf-8") 
					call echo("����","���������ļ����벻һ��2")
				end if
			end if
		end if
		
        '����base�ļ�
        if checkFile(htmlBaseFilePath) = false then
            c = regExp_Replace(content, "<head>", "<head>" & vbCrLf & "<base href=""" & getWebsite(httpurl) & """ />" & vbCrLf)
            call WriteToFile(htmlBaseFilePath, c, char_Set) 
        end if 
        htmlFileSize = getFSize(htmlFilePath) 
        call echored("html�ļ�·��", htmlFilePath & "   ��С��" & printSpaceValue(htmlFileSize) & "��" & htmlFileSize & "��") 
        content = replaceTemplateContent(content)                                       '�滻ģ������������



		'call rwend( pubHtmlObj.handleHtmlLabel(content,"img","src","","get","trim(*)isimg(*)") )
		'call rwend( pubHtmlObj.handleHtmlLabel(content,"*","style","style>background","get","") )
		'call rwend( pubHtmlObj.getHtmlLabelParam(content, "link", "href", "rel=shortcut icon, ||type=image/ico")  )
		'call rwend( pubCssObj.handleCssContent("html",content,"*","background||background-image","*","","images") )
		'call rwend( pubHtmlObj.getSwfSrcList(content) )  
		'ͼƬ�б�
		content = handleDownList(httpurl, content, pubHtmlObj.handleHtmlLabel(content,"img","src","","get","trim(*)isimg(*)"), "img")       'ͼƬ<img
		content = handleDownList(httpurl, content, pubHtmlObj.handleHtmlLabel(content,"*","style","style>background","get",""), "style") '��ǩ����<div style
		content = handleDownList(httpurl, content, pubHtmlObj.getHtmlLabelParam(content, "link", "href", "rel=shortcut icon, ||type=image/ico"), "link_image") '���Link ��*.ico
		content = handleDownList(httpurl, content,pubCssObj.handleCssContent("html",content,"*","background||background-image","*","","images"), "style_image") '���<style>

        'Css�б�
        content = handleDownList(httpurl, content, pubHtmlObj.getLinkHrefList(content), "css") 
        'JS�б�
        content = handleDownList(httpurl, content, pubHtmlObj.getScriptSrcList(content), "js") 
        'SWF
        content = handleDownList(httpurl, content, pubHtmlObj.getSwfSrcList(content), "swf") 
        '���ɴ�����ļ�
        call WriteToFile(createFilePath, phptrim(content), char_Set)                    '����ģ���ļ�(ȥ���߿ո�)

        urlListC = urlListC & batchFullHttpUrl(httpurl, pubHtmlObj.getAHrefList(content)) & vbCrLf 'url��ַ�ۼ�
        'A href
        pubAHrefList = pubAHrefList & urlListC 

        call createfile(debugWebDir & "/������ַ�б�.txt", downUrlList) 
        call createfile(debugWebDir & "/�����б�.txt", fileNameList) 
        call createfile(debugWebDir & "/ͼƬ�����б�.txt", downImagesList) 

        call getErrorInfoUrlList(httpurl) 
        nDownCount = nDownCount + 1 
        downC = downC & getImitate_print(nDownCount, htmlFilePath, jsListC, cssListC, imgListC, swfListC, otherListC, urlListC, errInfoC) & vbCrLf 
        jsListC = "" : cssListC = "" : imgListC = "" : swfListC = "" : otherListC = "" : urlListC = "" 
        call createfile(debugWebDir & "/Z��ӡ��վ��־.html", imitate_config(downC)) 		'������־�ļ�

    end sub 
    '��ô�����Ϣ��ַ�б� 20161002
    function getErrorInfoUrlList(byval httpurl)
        dim content, splStr, splxx, s, url, fileName, downStatus, goToUrl 
        content = getftext(debugWebDir & "/������ַ����.txt") 
        splStr = split(content, vbCrLf) 
		errInfoC=""
        for each s in splStr
            splxx = split(s, "��|��") 
            if uBound(splxx) >= 4 then
                goToUrl = splxx(1) 
                url = splxx(2) 
                fileName = splxx(3) 
                downStatus = splxx(4) 
                if splxx(0) = httpurl and downStatus <> "200" then 
                    errInfoC = errInfoC & goToUrl & "��>>��" & url & "��>>��" & fileName & "��>>��" & downStatus & vbCrLf '������Ϣ�ۼ�
                end if 
            end if 
        next 
		getErrorInfoUrlList=errInfoC
    end function 
	'�б����ͼƬ�б�
	function getErrorImagesList()
        dim content, splStr, splxx, s, url, fileName, downStatus, goToUrl ,fileNameList
        content = getftext(debugWebDir & "/������ַ����.txt")  
        splStr = split(content, vbCrLf) 
		fileNameList="|"
        for each s in splStr
            splxx = split(s, "��|��") 
            if uBound(splxx) >= 4 then
                goToUrl = splxx(1) 
                url = splxx(2) 
                fileName = splxx(3) 
                downStatus = splxx(4) 
                if  downStatus <> "200" then 
                    fileNameList=fileNameList & fileName & "|"
                end if 
            end if 
        next
		getErrorImagesList=fileNameList
	end function
    '��������ļ�
    function checkResetTypeFileName(byval debugWebDir, filePath, fileName)
        dim splStr, content, s, fileSuffixName, newFilePath 
        fileSuffixName = lcase(getFileAttr(filePath, "4")) 
        content = getDirFileNameList(debugWebDir, fileSuffixName) 
        'call echo(filePath,content)
        splStr = split(content, vbCrLf) 
        for each s in splStr
            if s <> "" and s <> fileName then
                newFilePath = debugWebDir & "/" & s 
                if checkFile(newFilePath) = true then

                    'call echo(getFSize(newFilePath),getFSize(filePath)   & "   " &  (getFText(newFilePath)=getFText(filePath)) )

                    '��CSS�ļ������ã���ΪCSS�����������ļ���û�д������Աȵ�CSS�Ǵ�����ģ����Բ��У������ڰ��ļ��ŵ�WEB�ļ����¿�����
                    if getFSize(newFilePath) = getFSize(filePath) and getFText(newFilePath) = getFText(filePath) then 
                        call deleteFile(filePath)                                                       'ɾ����ǰ�������Ҫ���ļ�
                        filePath = newFilePath 
                        fileName = s 
                        checkResetTypeFileName = true 
                        exit function 
                    end if 
                end if 
            end if 
        next
        checkResetTypeFileName = false
    end function 
    '���CSS���ݱ�ǩ�Ƿ�ɶԳ���
    function checkCssLabelSuccess_temp(content)
        dim spl1, spl2 
        spl1 = split(content, "{") 
        spl2 = split(content, "}") 
        checkCssLabelSuccess_temp = uBound(spl1) - uBound(spl2) 

    end function 
    '���������б�
    function handleDownList(byval httpurl, byval content, urlList, sType)
        dim splStr, s, c, url, fileName, filePath, cssContent, tempC, tempSFileName, isEchoDown,char_Set
        splStr = split(urlList, vbCrLf) 
        for each s in splStr
			'ֹͣʱ�˳�
			if isStop=true then
				exit function
			end if
            tempSFileName = getFileAttr(phptrim(s), "2") 
			'call echo(s,tempSFileName)
            '�ų�����#��
            if phptrim(s) <> "" and getwebsite(httpurl)<>s  then		'instr("#\/",left(tempSFileName, 1))=false  ��Ҫ����
                url = fullHttpUrl(httpurl, s)
                'img
                if sType = "img" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
                    content = pubHtmlObj.handleHtmlLabel(content, "img", "src", "src=" & s, "set", fileName) 
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              'ͼƬ��ַ�ۼ�

                elseif sType = "style" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
                    content = pubHtmlObj.handleHtmlLabel(content, "*", "style", "style>background", "set", s & "[#check#]" & fileName) 
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              'ͼƬ��ַ�ۼ�

                elseif sType = "link_image" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
					call echo(s,fileName)
                    content = pubHtmlObj.handleHtmlLabel(content, "link", "href", "rel=shortcut icon,||type=image/ico", "set", s & "[#check#]" & fileName) 
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              'ͼƬ��ַ�ۼ�

                elseif sType = "style_image" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
                    content = pubCssObj.handleCssContent("html",content,"*","background||background-image",s,fileName,"set") 
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              'ͼƬ��ַ�ۼ�

 

                '�滻css�ļ������ݱ���
                elseif sType = "css_content" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
                    content = pubCssObj.handleCssContent("css",content, "*", "background||background-image", s, fileName, "set")  
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              'ͼƬ��ַ�ۼ�

                'swf
                elseif sType = "swf" then
                    fileName = getUrlToFileName(url, ":*.swf")  
                    call handleDown(httpurl, url, fileName, isEchoDown)   
					content=pubHtmlObj.handleHtmlLabel(content,"embed","src","type=application/x-shockwave-flash","set", s & "[#check#]trim(*)" & fileName)
                    content = pubHtmlObj.handleHtmlLabel(content, "param", "value", "name=movie", "set", s & "[#check#]trim(*)" & fileName)		'trim(*)����check����Ϊ��Ҫ�ָ��
                    downImagesList = downImagesList & url & vbCrLf 
                    swfListC = swfListC & url & vbCrLf 

                'css
                elseif sType = "css" then
                    fileName = getUrlToFileName(url, ":*.css") 
                    filePath = handleDown(httpurl, url, fileName, isEchoDown) 
                    '�ļ���һ���´�����
                    if isEchoDown = true then 
						char_Set=customCharSet 
						if customCharSet="�Զ����" then 
							char_Set=getFileCharset(filePath)
						end if
			
                        cssContent = readFile(filePath, char_Set) : tempC = cssContent 
                        call echoB("����CSS�ļ�", filePath) 
                        cssContent = handleDownList(url, cssContent, pubCssObj.getCssContentUrlList(cssContent), "css_content") 
                        call WriteToFile(filePath, phptrim(cssContent), char_Set)                       '����ģ���ļ�(ȥ���߿ո�)
                        if checkCssLabelSuccess_temp(cssContent) <> 0 then
                            call echoRedB("CSS�ļ���ǩ���ǳɶԳ���" & filePath, checkCssLabelSuccess_temp(cssContent)) 
                        end if 
                        '�����ļ��Ա�
                        call checkResetTypeFileName(debugWebDir, filePath, fileName) 
                    end if  
                    content = pubHtmlObj.handleHtmlLabel(content, "link", "href", "rel=stylesheet", "set", s & "[#check#]" & fileName)   
                    cssListC = cssListC & url & vbCrLf                                              'css��ַ�ۼ�
					
                'js
                elseif sType = "js" then
                    fileName = getUrlToFileName(url, ":*.js") 
                    call handleDown(httpurl, url, fileName, isEchoDown)  
                    content = pubHtmlObj.handleHtmlLabel(content, "script", "src", "", "set", s & "[#check#]" & fileName) 
					
                    jsListC = jsListC & url & vbCrLf                                                'js��ַ�ۼ�
                end if 
            'call echo(s,url)
            end if 
        next 
        handleDownList = content 
    end function 

    '��������
    function handleDown(goToUrl, byval url, fileName, isEchoDown)
        dim filePath, downStatus, echoS, repeatNameS, s, splxx, nLen 
        dim filePrefixName, fileSuffixName, fileType, tempFileName, tempFilePath, s1, s2 
        isEchoDown = false                                                              'Ĭ��Ϊû����  ���ж��ļ��Ƿ���������

        s = url & "��|��" 
        '���� �������б����ȡ
        nLen = instr(vbCrLf & downUrlList, vbCrLf & s) 
        if nLen > 0 then
            fileName = mid(vbCrLf & downUrlList, nLen + len(vbCrLf & s)) 
            if instr(fileName, vbCrLf) > 0 then
                fileName = mid(fileName & vbCrLf, 1, instr(fileName, vbCrLf) - 1) 
            end if 
        end if 
        '20180921
        if isUTF8Down = true then
            url = toUTFChar(url)                                                            'תutf������ַ
        end if 


        filePath = debugWebDir & "\" & fileName 

        '�ļ����ڣ����ļ����Ѿ������ع��ˣ�����������
        if checkFile(filePath) = true and instr("|" & fileNameList & "|", "|" & fileName & "|") > 0 and instr(vbCrLf & downUrlList & "��|��", vbCrLf & url & "��|��") = false then
            tempFileName = fileName                                                         '��ʱ�����ļ����� �����غ��жϣ��Ƿ�һ������һ�������
            tempFilePath = filePath 
            fileName = "z_" & md5(url, 2) & "." & getFileAttr(fileName, "2") 
            filePath = debugWebDir & "\" & fileName 
        end if 
        '�ļ������ڣ�������
        if checkFile(filePath) = false then
            isEchoDown = true                                                               'Ϊ�ļ�������
            downStatus = newDownUrl(url, filePath) 
            '����ͼƬ�ļ�����
            if checkFile(filePath) = true then
                fileType = lcase(getImageType(filePath))                                        'ͼƬ��ʵ����
                if fileType <> "" then
                    filePrefixName = getFileAttr(filePath, "3") 
                    fileSuffixName = lcase(getFileAttr(filePath, "4")) 
                    if fileType <> fileSuffixName then
                        s1 = fileName 
                        s2 = filePath 
                        fileName = filePrefixName & "." & fileType 
                        filePath = debugWebDir & "\" & fileName 
                        'ת�����ļ�������MD5������
                        if checkFile(filePath) = true then
                            fileName = "z_" & md5(url, 2) & "." & fileType 
                            filePath = debugWebDir & "\" & fileName 
                        end if 
                        call moveFile(s2, filePath) 
						call deleteFile(s2)			'����ļ������Ѿ����ڣ���Ҫɾ���� ���������ν����Ϊ������ַ��һ����
                        call echoBlue("�滻�ļ�����", s1 & "=>>" & fileName) 
                    end if 
                end if 
                '�ж��Ƿ�׷���ҵ���Ϣ
                if instr("|bmp|jpg|gif|png|", "|" & fileType & "|") > 0 and downImageAddMyInfoType <> "" then
                    if getFSize(filePath) < 30240 then                                              'С��30K
                        call decSaveBinary(filePath, readBinary(filePath, 0) & "|" & imageAddMyInfo(downImageAddMyInfoType), 0) 
                        call echo("׷���ҵ���Ϣ", filePath) 
                    end if 
                end if 
            end if 
			
			
			
            '��ͬ�ļ���ͬ�ļ�����
            if tempFilePath <> "" and getFSize(filePath) = getFSize(tempFilePath) and getFText(filePath) = getFText(tempFilePath) then
                call deleteFile(filePath)                                                       '���ļ�һ����ɾ����ǰ����ļ�
                fileName = tempFileName 
                filePath = tempFilePath 
            '��������ļ�  ����Ҫ�ų�CSS����Ϊ�����ж�CSS���ԣ��ȴ�����CSS�����ж�
            elseif checkResetTypeFileName(debugWebDir, filePath, fileName) = true then
                call echoRedB("�滻�ļ���", fileName) 
				call echo(debugWebDir,filePath)
            else
                tempFileName = "" 
                tempFilePath = "" 
            end if 
			doevents

            downCountSize = downCountSize + getFSize(filePath)                              '���ش�С�ۼ�

            '�����б��ۼ�
            s = url & "��|��" & fileName & vbCrLf 
            if instr(vbCrLf & downUrlList & vbCrLf, vbCrLf & s) = false then
                downUrlList = downUrlList & url & "��|��" & fileName & vbCrLf                   '׷������ͼƬ��ַ
                call createfile(debugWebDir & "/������ַ�б�.txt", downUrlList)                   'ʱʱ���棬��ʱ��Ϊ�����ر�ʱ�м�¼
            end if 
            '�ļ����ۼ�
            if instr("|" & fileNameList & "|", "|" & fileName & "|") = false then
                fileNameList = fileNameList & fileName & "|" 
            end if 

            '���ػ���
            echoS = httpurl & "��|��" & goToUrl & "��|��" & url & "��|��" & fileName 
            'if downStatus <> 200 then
                echoS = echoS & "��|��" & downStatus 
            'end if 
            if tempFileName <> "" then
                repeatNameS = "��|�������ظ�" 
            end if 
            call createAddFile(debugWebDir & "/������ַ����.txt", echoS & repeatNameS) 
            call echo("���ԡ�" & downStatus & " ����" & printSpaceValue(getFSize(filePath)) & "��", "���������ز� " & url) 

        end if 
        handleDown = filePath 
    end function 

    '��������ַ   ��IE����
    function newDownUrl(httpurl, filePath)
        dim ieFilePath, fileType 
        fileType = lcase(getFileAttr(httpurl, "4")) 
        if fileType = "" then
            fileType = "html" 
        elseif instr(fileType, "?") > 0 then
            fileType = mid(fileType, 1, instr(fileType, "?") - 1) 
        end if 
        newDownUrl = 999 
        ieFilePath = handlePath(cacheFolderPath & "/" & md5(httpurl, 2) & "." & fileType) 


        '��������  Ҫ�ж��Ƿ���Դӷ�����������
        if(checkFile(ieFilePath) = false and isDownCache = true) or isDownServer = true then
            filePath = handlePath(filePath) 
            newDownUrl = saveRemoteFile(httpurl, filePath) 


            call copyFile(filePath, ieFilePath)                                             '���ļ����Ƶ�ie������
        end if 


        '���������˳�
        if checkFile(ieFilePath) = false then
            'call eerr( md5(httpurl,2), httpurl)
            call createAddFile(debugWebDir & "/888.txt", httpurl & "��|��" & md5(httpurl, 2)) 
            newDownUrl = 888 
            exit function 
        end if 
        if isDownCache = true and isDownServer=false then
            call copyFile(ieFilePath, filePath) 
            newDownUrl = 200 
        end if 
    end function 

    '������վ ����js,css���ļ���  ��ʱ���ţ��������ٸİɣ��е�����
    sub webBeautify(toDir, isTemplate)
        dim content, splStr, s, c, fileName, fileExt, filePath, toFilePath, webBeautifyDir, imagesDir, cssDir, jsDir, indexFilePath, toIndexFilePath, indexFileContent 
        dim cssFileList, splxx, url, findStr, char_Set 
        dim htmlFileList, splHtmlFile, htmlFileNameS 
        webBeautifyDir = toDir
        imagesDir = webBeautifyDir & "/images/" 
        cssDir = webBeautifyDir & "/css/" 

        char_Set = customCharSet 
		
		call echo("webBeautifyDir",webBeautifyDir)
		
        jsDir = webBeautifyDir & "/js/" 
        call deleteFolder(webBeautifyDir) 
        call createDirFolder(webBeautifyDir)                                            '���������ļ���
        call createfolder(imagesDir) 
        call createfolder(cssDir) 
        call createfolder(jsDir) 

        call echo("webBeautifyDir", webBeautifyDir)  
        call echo("jsDir", jsDir) 

        htmlFileList = getDirHtmlList(debugWebDir)  
        splHtmlFile = split(htmlFileList, vbCrLf) 
        for each indexFilePath in splHtmlFile
			fileName=getFileAttr(indexFilePath,"name")  
            if left(fileName,6)="index_" then
                'indexFilePath = debugWebDir & "/index.html"
                htmlFileNameS = getFileAttr(indexFilePath, 3) 
                if right(htmlFileNameS, 1) = "_" then
                    htmlFileNameS = left(htmlFileNameS, len(htmlFileNameS) - 1) 
                end if 
                call echo("htmlFileNameS", htmlFileNameS) 

                'Ϊģ��
                if isTemplate = true then
                    toIndexFilePath = webBeautifyDir & "/" & htmlFileNameS & "_Model.html" 
                else
                    toIndexFilePath = webBeautifyDir & "/" & htmlFileNameS & ".html" 
                end if 
				'�Զ����֮��Ͳ�Ҫ���ж�html�ļ��Ƿ��ǵ�ǰ���룬��Ϊ��֮ǰ�Ѿ������
				if customCharSet="�Զ����" then 
					char_Set=getFileCharset(indexFilePath)
				end if
                indexFileContent = readFile(indexFilePath, char_Set)                         '���ļ� 
                'Ϊģ��
                if isTemplate = true then
                    'link���и�favicon.ico  λ�û������⣬������20160302
                    indexFileContent = handleConentUrl("[$cfg_webcss$]/", indexFileContent, "|link|", "", "") 
                    indexFileContent = handleConentUrl("[$cfg_webimages$]/", indexFileContent, "|img|embed|param|meta|src|imgstyle|", "", "") 
                    indexFileContent = handleConentUrl("[$cfg_webjs$]/", indexFileContent, "|script|", "", "") 
                    indexFileContent = handleHtmlStyleCss(indexFileContent, "�滻·��", "[$cfg_webimages$]/") '����html���style css
                else
                    'link���и�favicon.ico  λ�û������⣬������20160302
                    indexFileContent = handleConentUrl("css/", indexFileContent, "|link|", "", "") 
                    indexFileContent = handleConentUrl("images/", indexFileContent, "|img|embed|param|meta|src|imgstyle", "", "") 
                    indexFileContent = handleConentUrl("js/", indexFileContent, "|script|", "", "") 
                    indexFileContent = handleHtmlStyleCss(indexFileContent, "�滻·��", "images/")  '����html���style css

                end if 
                'Ϊģ�壬���滻������ؼ���������
                if isTemplate = true then
                    indexFileContent = replaceWebTitle(indexFileContent, "{$Web_Title$}") 
                    indexFileContent = replaceWebMate(indexFileContent, "name", "keywords", " content=""{$Web_KeyWords$}"" ", isAddWebTitleKeyword) 
                    indexFileContent = replaceWebMate(indexFileContent, "name", "description", " content=""{$Web_Description$}"" ", isAddWebTitleKeyword) 
                    indexFileContent = replaceWebMate(indexFileContent, "http-equiv", "Content-Type", " content=""text/html; charset=gb2312"" ", isAddWebTitleKeyword) 
 
                end if 
                '��ʽ��HTMLΪ��
                if isFormatting = true then
                    indexFileContent = formatting(indexFileContent, "") 
                    indexFileContent = htmlFormatting(indexFileContent) 
                end if 
                call WriteToFile(toIndexFilePath, phptrim(indexFileContent), char_Set) 			'�����ļ�ʱ�������ǿ����Զ���ģ�������Ժ�������

            end if 
        next 
		'�б�Ŀ¼ȫ���ļ�   �жϰ�ָ���ļ��ŵ���Ӧ�ļ�����
		content = getFileFolderList(debugWebDir, true, "ȫ��", "����", "", "", "") 
		splStr = split(content, vbCrLf) 
		for each fileName in splStr
			if fileName <> "" then
				filePath = debugWebDir & "/" & fileName 
				fileExt = lCase(getFileExt(fileName)) 
				if fileExt = "css" then
					toFilePath = cssDir & fileName 
					cssFileList = cssFileList & toFilePath & vbCrLf 
				elseIf fileExt = "js" then
					toFilePath = jsDir & fileName 
				elseIf fileExt = "txt" or fileExt = "html" or fileExt = "htm" then
					toFilePath = "" 
				elseIf inStr(fileName, ".") > 0 then
					toFilePath = imagesDir & fileName 
				else
					toFilePath = "" 
				end if 
				'call echo(filePath,tofilepath)
				if toFilePath <> "" then
					'�ж��ļ���С��Ϊ��0    20160612
					if getFSize(filePath) <> 0 and instr(deleteErrorImagesList, "|" & fileName & "|")=false then 
						call copyFile(filePath, toFilePath) 
					end if 
				end if 
				call newEcho(fileName,  "images/" & fileName) 
			end if 
		next 
        'css�б�Ϊ��
        if cssFileList <> "" then
            splxx = split(cssFileList, vbCrLf) 
            for each filePath in splxx
                if filePath <> "" then
                    content = readFile(filePath, "") 
                    'Ϊģ��
                    if isTemplate = true then
                        content = pubCssObj.handleCssContent("css",content,"*","background||background-image","*","isimg(*)trim(*)dir(*)" & templateDir & templateName & "/images/","set") 
                    else
                        content = pubCssObj.handleCssContent("css",content,"*","background||background-image","*","isimg(*)trim(*)dir(*)../images/","set")
                    end if 

                    call WriteToFile(filePath, phptrim(content), "") 
                end if 
            next 
        end if 
        'NOVBNet start
		
		content = getFileFolderList(webBeautifyDir, true, "ȫ��", "", "ȫ���ļ���", "", "") 
        c = replace(content, vbCrLf, "|") 
        'call rw(content)
        'call eerr(getThisUrlNoFileName(),getthisurl())
        c = c & "|||||" 
		'Ҫ�õ�����ļ��У���ǰ������  ��vb.net�ﲻ�����
		call createfolder("./htmlweb")		
        url = getThisUrlNoFileName() & "/myZIP.php?webFolderName=web_&zipDir=htmlweb/&isPrintEchoMsg=0" 
		
        'PHP���� ��վ
        if isPackWebZip = true or isPackTemplateZip = true then
            call echo(isPackWebZip, isPackTemplateZip) 
            if isPackWebZip = true then
                s = xMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(debugWebDir)) & "\") 
            else
                'PHP���� ģ����վ
                s = xMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(isWebToTemplateDir)) & "\") 
            end if 
            findStr = strCut(s, "downfile.asp?downfile=", """", 2) 

            '��session��ʽ������ֵ
            session("downfilePath") = "/" & findStr 
            s = replace(s, "downfile.asp?downfile=" & findStr, "/tools/downfile.asp?downfile=session") 
            call echo("", s) 
        end if 
        if(isPackWeb = true and isTemplate = false) or(isPackTemplate = true and isTemplate = true) then
            dim xmlFileName, xmlSize 
            xmlFileName = getIP() & "_update.xml" 

            dim objXmlZIP : set objXmlZIP = new xmlZIP
                call objXmlZIP.callRun(handlePath(debugWebDir), handlePath(debugWebDir & rootDir), false, xmlFileName) 
                call hr() 
                'call echo("1111111111111", "ע��")
                call echo(handlePath(debugWebDir), handlePath(debugWebDir & rootDir)) 
            set objXmlZIP = nothing 
            doevents 
            xmlSize = getFSize(xmlFileName) 
            xmlSize = printSpaceValue(xmlSize) 
            call echo("����xml����ļ�", "<a href=?act=download&downfile=" & xorEnc(xmlFileName, 31380) & " title='�������'>�������" & xmlFileName & "(" & xmlSize & ")</a>") 
        end if 
		
		'NOVBNet end
    end sub 

    '�滻��վ����
    function replaceWebTitle(content, webTitle)
        dim tempContent, startStr, endStr, nLen, bodyStart, bodyEnd 
        tempContent = lCase(content) 
        startStr = "<title>" 
        endStr = "</title>" 
        nLen = inStr(tempContent, startStr) 
        if nLen > 0 then
            bodyStart = mid(content, 1, nLen - 1)                                           '��ʼ����
            tempContent = mid(tempContent, nLen) 
            content = mid(content, nLen) 
        else
            replaceWebTitle = content 
            exit function 
        end if 
        nLen = inStr(tempContent, endStr) 
        if nLen > 0 then
            bodyEnd = mid(content, nLen + len(endStr))                                      '��������
        end if 
        replaceWebTitle = bodyStart & startStr & webTitle & endStr & bodyEnd 
    end function 
    '�滻��վ�ؼ���  ����Ҳ�����滻
    function replaceWebMate(content, sArrt, sTypeName, valueStr, isToAdd)
        dim tempContent, sourceContent, startStr, endStr, nLen, bodyStart, bodyEnd, s 
        sourceContent = content 
        tempContent = lCase(content) 
        startStr = " " & sArrt & "=""" & sTypeName & """" 
        endStr = "/>" 
        if inStr(tempContent, startStr) = false then
            startStr = " " & sArrt & "='" & sTypeName & "'" 
        end if 
        nLen = inStr(tempContent, startStr) 
        if nLen > 0 then
            bodyStart = mid(content, 1, nLen - 1)                                           '��ʼ����
            tempContent = mid(tempContent, nLen) 
            content = mid(content, nLen) 
        end if 
        nLen = inStr(tempContent, endStr) 
        if nLen > 0 then
            bodyEnd = mid(content, nLen + len(endStr))                                      '��������
        end if 
        if bodyStart = "" or bodyEnd = "" then
            s = "<meta" & startStr & valueStr & endStr 
            nLen = instr(lcase(sourceContent), "<head>") 
            if nLen > 0 then
                startStr = mid(content, 1, nLen + 5) 
                endStr = mid(content, nLen + 6) 
                replaceWebMate = startStr & s & endStr 
            else
                replaceWebMate = sourceContent 
            end if 
        else
            replaceWebMate = bodyStart & startStr & valueStr & endStr & bodyEnd 
        end if 
    end function
	
	
	
	
	

    '�滻ԴHTML���ݣ���Բ�ͬ��վ����
    function replaceTemplateContent(content)
        dim findS, replaceS, startStr, endStr, jsList, splStr, s, c, i 
        startStr = "minify?f=" 
        endStr = """" 
        for i = 1 to 9
            if instr(content, startStr) > 0 then
                findS = getStrCut(content, startStr, endStr, 1) 
                jsList = getStrCut(content, startStr, endStr, 0) 
                splStr = split(jsList, ",") 
                for each s in splStr
                    if c = "" then
                        c = s 
                    else
                        c = c & """></script>" & vbCrLf & "<script src=""" & s & "" 
                    end if 
                    c = c & """" 
                next 
                call createfile("1.txt", c) 
                content = replace(content, findS, c) 
                c = "" 
            end if 

        next 
 
        replaceTemplateContent = content 
    end function 



    '����echo
    function newEcho(title, str)
        if isOnMsg = true then
            call echo(title, str) 
        end if 
    end function 
end class 



%> 

