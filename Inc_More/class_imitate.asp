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
'��վ
class imitateWeb
    'dubug/������ַ�б�.txt   ���ļ���Ҫɾ���������������ʱ�� �ظ�����ͼƬ���ļ�
    dim thisFormatObj                                                               '��ʽHTML��
    dim version                                                                     '�汾
    dim defaultIEContent                                                            'Ĭ��IE��ʾ����
    dim templateName                                                                'ģ������
    dim templateDir                                                                 'ģ��·��
    dim toTemplateDir                                                               '����ģ�嵽����Ŀ¼
    dim isWebToTemplateDir                                                          '�ļ��Ƿ�ŵ�ģ���ļ�����

    dim serverHttpUrl                                                               '��������ַ
    dim httpurl 																	'�ɼ���ַ
    dim newWebDir                                                                   '����վĿ¼
    dim debugWebDir                                                                 '������վĿ¼
    dim createHtmlWebDir                                                            '����html��վĿ¼
    dim cacheFolderPath                                                             '�����ļ���·��
    dim cacheConfigPath                                                             '��������·������¼URL״̬��
    dim customCharSet                                                               '�Զ������
    dim isGetHttpUrl                                                                'getHttpUrl��ʽ�������
    dim isMakeWeb                                                                   '����WEB�ļ���
    dim isMakeTemplate                                                              '�Ƿ�����ģ���ļ���
    dim isPackWeb                                                                   '�Ƿ���WEB�ļ��� xml
    dim isPackWebZip                                                                '�Ƿ���WEB�ļ��� zip
    dim isPackTemplate                                                              '�Ƿ���ģ���ļ��� xml
    dim isPackTemplateZip                                                           '�Ƿ���ģ���ļ��� zip
    dim isAddWebTitleKeyword                                                        '׷����վ������ؼ���
    dim isOnMsg                                                                     '�Ƿ���������Ϣ
    dim pubAHrefList                                                                'A�����б�
    dim isUTF8Down                                                                  '��utf8��ʽ����
    dim downUrlList                                                                 '���ص�URL�б�
    dim fileNameList                                                                '�ļ������б�
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
	dim isEditImageType																'�޸�ͼƬ����
	dim isWithNameHeBin																'ͬ��ͼƬ�ϲ�
	dim isWithContentHeBin															'ͬ�����ļ��ϲ�
	dim isAddMyInfo																	'׷���ҵ���Ϣ
	dim isHandleJsInHtml															'�Ƿ���js���html 20171114
	dim replaceSourcePageConfig														'�滻ԭ��ҳ���� 20171109
	dim createHelpFileContent														'���ɰ����ĵ����� 20171109
	

    dim jsListC, cssListC, imgListC, swfListC,htmlListC, otherListC, urlListC, errInfoC, downC, nDownCount '��¼���ػ�����Ϣ
	dim cacheFilepath,htmlFilePath, htmlBaseFilePath, createFilePath			'�����ļ�·����htmlԴ�ļ�·����htmlԴ�ļ�+base·��������html�ļ�
	dim isDownCssImage,isDownStyleImage,isHtmlImg				'�Ƿ�����CSSͼƬ����ʽͼƬ,ImgͼƬ
	dim htmlformat,jsformat,cssformat								'html,js,css��ʽ������
	dim useUniformCoding															'ѡ��ͳһ����

    '���캯�� ��ʼ��
    sub class_Initialize()
        call loadDefaultConfig() 
    end sub 

    '����Ĭ������
    sub loadDefaultConfig()
        set thisFormatObj = new class_formatting                                        '��ʼ����ʽ��HTML��
            version = "��վ v5.6"                                                        '����汾 v1.1 beta version
            serverHttpUrl = "http://aa/server_FangZan.Asp"                                  '�ȸĳ�sharembweb.com
            httpurl = "http://sharembweb.com/" 
            newWebDir = "newweb"                                                            '����վĿ¼
            debugWebDir = "dubug"                                                           '������վĿ¼
            createHtmlWebDir = "web"                                                        '������վĿ¼
            customCharSet = "�Զ����"                                                      '����
            isGetHttpUrl = false                                                            'getHttpUrl��ʽ�������
            isMakeWeb = true                                                                '����WEB�ļ���
            isMakeTemplate = false                                                          '�Ƿ�����ģ���ļ���
            isPackWeb = false                                                               '�Ƿ���WEB�ļ��� xml
            isPackWebZip = false                                                            '�Ƿ���WEB�ļ��� zip
            isWebToTemplateDir = true                                                       '�ļ��Ƿ�ŵ�ģ���ļ�����
            toTemplateDir = "/templates"                                                    'ģ�帴�Ƶ�Ŀ¼
            isPackTemplate = false                                                          '�Ƿ���ģ���ļ��� xml
            isPackTemplateZip = false                                                       '�Ƿ���ģ���ļ��� zip

            isAddWebTitleKeyword = true                                                     '׷����վ������ؼ���
            isOnMsg = false                                                                 '�Ƿ���������Ϣ
            templateName = getWebSiteName(httpurl)                                          'ģ�����ƴ���ַ�л��
            isUTF8Down = false                                                              'utf8��ʽ����Ϊ��
            downCountSize = 0                                                               '�����ܴ�С
            downFileTypeList = "|*|"                                                        '�����ļ������б�
            isDownServer = true                                                             '�Ƿ���Դӷ�����������
            isDownCache = true                                                              'Ĭ�ϲ��������ػ���
            downImageAddMyInfoType = "2"                                                    '����ͼƬ�����Լ�����Ϣ
            cacheFolderPath = ""                                                            '����Ŀ¼ Ϊ�����ڵ�ǰĿ¼�´���һ��cacheĿ¼
            call createDirFolder(cacheFolderPath) 
            defaultIEContent = "[defaultIEContent]"                                    		'Ĭ��IE��ʾ����
            isDeleteErrorImages = true                                                      'ɾ������ͼƬ(webĿ¼��templates)
            isStop = false                                                                  'Ĭ��ֹͣΪ��
            isDownCssImage = true : isDownStyleImage = true : isHtmlImg = true 				'Ĭ������ͼƬΪ��

            isEditImageType = true                                                          '�޸�ͼƬ����
            isWithNameHeBin = true                                                          'ͬ��ͼƬ�ϲ�
            isWithContentHeBin = false                                                      'ͬ�����ļ��ϲ�
            isAddMyInfo = true                                                              '׷���ҵ���Ϣ
			isHandleJsInHtml=false															'�Ƿ���js���html Ϊ��
    end sub


    '������վ
    function downweb()
        dim s, websiteFilePath, websiteFileContent 
        dim filePath, toFilePath, fileList, splList, fileName, splStr, startTime, splUrl,splxx
        startTime = now() 
        httpurl = phptrim(httpurl) 
        if httpurl = "" then
            call eerr("ֹͣ", "��ַΪ��") 
        elseif len(httpurl)<=3 then
            call eerr("ֹͣ", "��ַ���Ϸ�") 
        end if
        isStop = false                                                                  'ֹͣΪ��

        '��÷�վĿ¼
        newWebDir = newWebDir & "/" & IIf(saveFolderName <> "", saveFolderName, getWebSiteCleanName(httpurl)) 
        '���������ļ���
        call createfolder(newWebDir) 
        '��������Ŀ¼
        debugWebDir = newWebDir & "/" & debugWebDir 
        call createDirFolder(debugWebDir) 
        '��������Ŀ¼ ǰ���ǿ������� �� Ĭ�ϻ����ļ���Ϊ��
        if cacheFolderPath = "" and isDownCache = true then
            cacheFolderPath = newWebDir & "/cache" 
            call createDirFolder(cacheFolderPath) 
        end if 

        '����html��վĿ¼
        if isMakeWeb = true then
            createHtmlWebDir = newWebDir & "/web" 
            call createDirFolder(createHtmlWebDir) 
        end if 
        'ģ��Ŀ¼
        if isWebToTemplateDir = true then
            templateDir = newWebDir & "/templates/" & templateName                          '���ɵ�ǰģ��Ŀ¼
            toTemplateDir = toTemplateDir & "/" & templateName                              '���Ƶ�ָ��ģ��Ŀ¼
            call createDirFolder(templateDir) 												'����ģ��Ŀ¼
        end if 

        downUrlList = readFile(debugWebDir & "/������ַ�б�.txt", "")                   '��������ַ�б�
		
		if instr(httpurl,"|")>0 then
			splxx=split(httpurl,"|")
			httpurl=splxx(0)
			fileName=splxx(1)
		end if 
		if fileName="" then
        	fileName = getUrlFileName(httpurl, "", getUrlToFileName(httpurl, ":*.html"))    '���html�ļ���  ���ڱ��������ļ������޸�
		else
        	fileName = getUrlFileName(httpurl, "", fileName)   
		end if

        cacheFilePath = cacheFolderPath & "/" & md5(httpurl, 2) & ".htm"                '�����ļ�·��
        htmlFilePath = debugWebDir & "/html_" & fileName                                'Դ�ļ�·��
        htmlBaseFilePath = debugWebDir & "/base_" & fileName                            'base�ļ�
        createFilePath = debugWebDir & "/index_" & fileName                             '�����ļ�
        cacheConfigPath = cacheFolderPath & "/cacheConfig.txt"                          '��������·��
 

        call echo("�����ļ�·��(cacheFilePath)", cacheFilepath) 
        call echo("HTML�ļ�·��(htmlFilePath)", htmlFilePath) 
        call echo("HTML�����ļ�·��(htmlBaseFilePath)", htmlBaseFilePath) 
        call echo("����HTML�ļ�·��(createFilePath)", createFilePath) 
        if checkFile(htmlFilePath) = false then
            if checkFile(cacheFilePath) = true then
                call copyFile(cacheFilepath, htmlFilePath) 
                call echo(httpurl, getCacheState(httpurl)) 
                call editHttpUrlState(httpurl, getCacheState(httpurl))                          '׷��״̬��
                call echo("downurllist", downUrlList)  
            else
                downweb = false						'����Ϊ����ִ��ģ��
                exit function 
            end if 
        end if 
        call imitateWeb() 			'��վ���ĳ���
        downweb = true				'����Ϊ��

    end function 
    '��վ
    function imitateWeb()
        dim content, fileName 
        dim htmlFileSize, char_Set, splStr, splxx, s, c, htmlFPath, cacheFPath, isAdd 
        fileName = getFileAttr(createFilePath, "2")                                     '����ļ�����
        content = readFile(htmlFilePath, char_Set) 
        char_Set = customCharSet 
        '�ǳ���Ҫ���Ժ󿴵������ס��Ҫɾ���ˣ�20161011
        if customCharSet = "�Զ����" then
            doevents 
            char_Set = getFileCharset(htmlFilePath) 
            content = readFile(htmlFilePath, char_Set)                                      '�������Ϊ���Զ��ұ���
            call echo("�Զ���� ����HTML�ļ�����", htmlFilePath & " = "& char_Set) 
            if char_Set = "gb2312" then
                s = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "get", "lcase(*)") 
                if instr(s, "utf-8") > 0 then
                    content = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "set", "text/html; charset=gb2312") 
                    call echo("����", "���������ļ����벻һ��1") 
                end if 
            elseif char_Set = "utf-8" then
                s = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "get", "lcase(*)") 
                if instr(s, "gb2312") > 0 then
                    content = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "set", "text/html; charset=utf-8") 
                    call echo("����", "���������ļ����벻һ��2") 
                end if 
            end if 
            call WriteToFile(createFilePath, phptrim(content), char_Set)                    '����ģ���ļ�(ȥ���߿ո�)
        end if
		
		content=handleReplaceSourcePageConfig(content)										'����Դ��ҳ�滻20171109
		if char_Set = "gb2312" then
			content=pubHtmlObj.handleHtmlLabel(content,"meta","content","**content=charset","set","replace[**utf-8=gb2312]isreplace(*)")
			content=pubHtmlObj.handleHtmlLabel(content,"meta","charset","","set","replace[**utf-8=gb2312]isreplace(*)")
		end if

        '����base�ļ�
        if checkFile(htmlBaseFilePath) = false then
            c = regExp_Replace(content, "<head>", "<head>" & vbCrLf & "<base href=""" & getWebsite(httpurl) & """ />" & vbCrLf) 
            call WriteToFile(htmlBaseFilePath, c, char_Set) 
        end if 
        content = replaceTemplateContent(content)                                       '����ģ�����滻���� urllistnoenc(*)setaddurlenc(*)
        '��ʽ��HTML
        call thisFormatObj.handleFormatting(content) 
        call thisFormatObj.handleLabelContent("*", "*", "setaddurlenc(*)isimg(*)fullurl(*)" & httpurl, "handleimg") '����CSS����ͼƬ lcase(*)�õ� ��Ҫ��linuxϵͳ�Ĵ�Сд������������� 
        call handleDownList(httpurl, fileName, content, thisFormatObj.imgList, "img") 
        call handleDownList(httpurl, fileName, content, thisFormatObj.swfList & vbcrlf & thisFormatObj.sourceList & vbcrlf & thisFormatObj.videoList, "swf")
        call handleDownList(httpurl, fileName, content, thisFormatObj.jsList, "js") 
        call handleDownList(httpurl, fileName, content, thisFormatObj.cssList, "css") 
		'call echoyellow("thisFormatObj.videoList",thisFormatObj.videoList)

        '��Ҫ��ѹ�����ʱ�ٴδ���html�ļ�����
        if htmlformat <> "" or jsformat <> "" or cssformat <> "" then
            call thisFormatObj.updateHtml() 
        end if 
        call echoredb("cssformat", cssformat) 
        '��ø�ʽ����HTML
        content = thisFormatObj.getFormattingHtml(htmlformat, jsformat, cssformat) 

        'ʹ��ͳһ����
        if useUniformCoding <> "" and char_Set <> useUniformCoding then
            char_Set = useUniformCoding 
            content = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "set", "text/html; charset=" & char_Set) '�ظ�����
        end if 
		content=deleteHandleHtmlWeb(content)											'ɾ���������������ҳ����
		if isHandleJsInHtml=true then
			content=handleContentInScriptHtml(content,httpurl,fileName)										'�����������script js�ű�
		end if
        '���ɴ�����ļ�
        call WriteToFile(createFilePath, phptrim(content), char_Set)                    '����ģ���ļ�(ȥ���߿ո�)
        call WriteToFile(createFilePath & ".txt", phptrim(content), char_Set)           '����ģ���ļ�(ȥ���߿ո�)

    end function 
    '���������б�
    function handleDownList(byval httpurl, localFilePath, byval content, urlList, sType)
        dim splStr, s, c, url, fileName, parentFileName, filePath, cssContent, tempC, isEchoDown, char_Set 
        splStr = split(urlList, vbCrLf) 
        for each s in splStr
            'ֹͣʱ�˳�
            if isStop = true then
                exit function 
            end if 
            '�ų�����#��
            if phptrim(s) <> "" and getwebsite(httpurl) <> s then
                url = fullHttpUrl(httpurl, s) 
                'img
                if sType = "img" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.jpg")) 
                    imgListC = imgListC & url & "��|��" & fileName & vbCrLf 

                'swf
                elseif sType = "swf" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.swf")) 
                    swfListC = swfListC & url & "��|��" & fileName & vbCrLf 

                'css
                elseif sType = "css" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.css")) 
                    cssListC = cssListC & url & "��|��" & fileName & vbCrLf 

                'js
                elseif sType = "js" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.js")) 
                    jsListC = jsListC & url & "��|��" & fileName & vbCrLf 

                'html
                elseif sType = "html" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.html")) 
                    htmlListC = htmlListC & url & "��|��" & fileName & vbCrLf 
                end if 
            end if 
        next 
        handleDownList = content 
    end function 

    '����ļ����� ���ۼ�������ַ�б�
    function getUrlFileName(byval url, localFilePath, fileName)
        dim splStr, splxx, s, parentFileName, isAdd 
        parentFileName = fileNameToParentNameMD5(url, fileName) 
        if parentFileName <> "" then
            fileName = md5(url, 2) & "_" & fileName 
        end if 
        splStr = split(downUrlList, vbCrLf) 
        getUrlFileName = fileName 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                'call echoyellowb(url,splxx(0))
                if splxx(0) = url and splxx(1) = localFilePath then    '��ַ�뱾�ر����ļ���ͬ ��20171110
                    getUrlFileName = splxx(2) 
                    call echoyellowb("����URL�ҵ� filename=", splxx(2)) 
                    exit function 
                end if 
            end if 
        next 
        'call hr()
		if right(downUrlList,2)<>vbcrlf then
			downUrlList=downUrlList & vbcrlf
		end if
        downUrlList = downUrlList & url & "��|��" & localFilePath & "��|��" & fileName & "��|��" & parentFileName & "��|��State״̬��"
    end function 
    '����CSS�ļ�����
    function handleCssFileContent()
        dim splStr, splxx, s, url, fileName, filePath, content, imgList, char_Set 
        splStr = split(cssListC, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                fileName = splxx(1) 
                filePath = debugWebDir & "/" & fileName 
                content = readFile(cacheFolderPath & "/" & md5(url, 2) & "." & getFileAttr(splxx(1), "4"), "") 
                call echoB("����CSS�ļ�", filePath):doevents
                imgList = pubCssObj.handleCssContent("css", content, "*", "background||background-image||src", "*", "isimg(*)lcase(*)fullurl(*)" & s, "images") 
                content = pubCssObj.handleCssContent("css", content, "*", "background||background-image||src", "*", "setaddurlenc(*)isimg(*)lcase(*)fullurl(*)" & s, "set") 
                'call echoredb(httpurl,imgList)
                call handleDownList(httpurl, fileName, content, imgList, "img")   '����
                '����ѹ�����ѹCSS		
                if cssformat <> "" then
                    content = pubCssObj.handleCssContent("css", content, "", "", "", "", cssformat) 
                end if 
                char_Set = "" 
                'ʹ��ͳһ����
                if useUniformCoding <> "" then
                    char_Set = useUniformCoding 
                end if 
                call WriteToFile(filePath, phptrim(content), char_Set) 
            end if 
        next 

    end function 
    '����JS�ļ�����
    function handleJSFileContent()
        dim splStr, splxx, s, url, fileName, filePath, content, imgList, char_Set 
		dim thisJsObj	
		if isHandleJsInHtml=false and jsformat="" and useUniformCoding="" then
			call echoYellow("��ʾ","����js�ļ�δ����ѡ��˳�����")
			exit function
		else
			call echo("isHandleJsInHtml",isHandleJsInHtml)
			call echo("jsformat",jsformat)
			call echo("useUniformCoding",useUniformCoding)
		end if
		
		
        splStr = split(jsListC, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                fileName = splxx(1) 
                filePath = debugWebDir & "/" & fileName 
				if checkFile(filePath)=true then		'�ļ���������
					content = readFile(cacheFolderPath & "/" & md5(url, 2) & "." & getFileAttr(splxx(1), "4"), "") 
					
					'�ų�Jquery�ļ�����Ϊ������   Ҫ��������js���html
					if instr(content,".jQuery=")=false and isHandleJsInHtml=true then 
						call echoB("����JS�ļ���html", filePath):doevents
						set thisJsObj=new class_js							'js��
						thisJsObj.httpurl=httpurl		'����js��ַ
						content=  thisJsObj.handleJsContent(content,"source|html")		'����js������ݣ���html��ͼƬ����
						'������ҳ
						call handleDownList(httpurl, fileName, content, thisJsObj.imgList, "img") 
						call handleDownList(httpurl, fileName, content, thisJsObj.swfList & vbcrlf & thisJsObj.sourceList & vbcrlf & thisJsObj.videoList, "swf")
						call handleDownList(httpurl, fileName, content, thisJsObj.jsList, "js") 
						call handleDownList(httpurl, fileName, content, thisJsObj.cssList, "css")  
					end if
					
					
					'����ѹ�����ѹJS
					if jsformat <> "" then
						call echoB("����JS�ļ�", filePath):doevents
						content = pubJsObj.handleJsContent(content, jsformat) 
					end if 
					char_Set = "" 
					'ʹ��ͳһ����
					if useUniformCoding <> "" then
						char_Set = useUniformCoding 
					end if 
					call WriteToFile(filePath, phptrim(content), char_Set) 
					call WriteToFile(filePath & ".txt", phptrim(content), char_Set)  '2017114 ��js�����ݣ����㴦�������html��ͼƬcss/js·��
				end if
            end if 
        next 

    end function 

    '׷���ҵ���Ϣ
    function addMyInfo()
        dim splStr, content, filePath, fileType, imageType, c, s 
        if downImageAddMyInfoType = "" then
            exit function 
        end if 
        content = getDirFileList(debugWebDir, "|jpg|gif|bmp|png|") 
        splStr = split(content, vbCrLf) 
        for each filePath in splStr
            if filePath <> "" and getFSize(filePath) > 0 then
                fileType = lcase(getFileAttr(filePath, "4")) 
                imageType = lcase(getImageType(filePath)) 
                '�ж��Ƿ�׷���ҵ���Ϣ
                if instr("|jpg|gif|bmp|png|", "|" & imageType & "|") > 0 then

                    if getFSize(filePath) < 30240 then                                              'С��30K
                        c = readBinary(filePath, 0) 
                        s = imageAddMyInfo(downImageAddMyInfoType) 
                        if instr(c, s) = false then
                            call decSaveBinary(filePath, c & "|" & s, 0) 
                            call echo("׷���ҵ���Ϣ", filePath) 
                        else
                            call echoGay("����׷���ҵ���Ϣ���Ѵ���", filePath) 
                        end if 
                    end if 
                end if 
            end if 
        next 
    end function 
    '�޸�ͼƬ����
    function editImageType()
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, imageType, parentFileName, localFileName, tempS 
        dim urlList : urlList = "," 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) : tempFileName = fileName 
                parentFileName = splxx(3) 
                filePath = debugWebDir & "/" & fileName 
                fileType = lcase(getFileAttr(fileName, "4")) 
                '�ļ��������ж�
                if checkFile(filePath) = true then
                    imageType = lcase(getImageType(filePath))                                       'ͼƬ��ʵ����

                    if instr("|bmp|jpg|gif|png|", "|" & fileType & "|") > 0 and fileType <> imageType and imageType <> "" then
                        fileName = getFileAttr(fileName, "3") & "." & imageType 
                        call moveFile(debugWebDir & "/" & tempFileName, debugWebDir & "/" & fileName) 
                        urlList = urlList & url & "|" & fileName & "," 
                        call echo("�ļ�����", url & " �ļ��� " & tempFileName & " ���� " & fileType & " >> " & imageType) 
                    end if 
                elseif instr(urlList, "," & url & "|") > 0 then
                    tempS = mid(urlList, instr(urlList, "," & url & "|") + len("," & url & "|")) 
                    'call echo("tempS",tempS)
                    fileName = mid(tempS, 1, instr(tempS, ",") - 1) 

                    'call echo("tempS",tempS)
                    'call echoblueb("urlList",urlList)
                    'call echoredb(url & "("& checkFile(filePath) &")",tempFileName & ">>" & fileName)
                    call echo("�ļ�����(����)", url & " �ļ��� " & tempFileName & " ���� " & fileType & " >> " & imageType) 

                end if 
                s = url & "��|��" & localFileName & "��|��" & fileName & "��|��" & parentFileName & "��|��" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 
    end function 
    'ͬ���ϲ�
    function withNameHeBin()
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, imageType, parentFileName, parentFilePath, localFileName 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) : tempFileName = fileName 
                parentFileName = fileNameToParentName(splxx(3)) 
                parentFilePath = debugWebDir & "/" & parentFileName 
                filePath = debugWebDir & "/" & fileName 
                fileType = lcase(getFileAttr(fileName, "4")) 

                if parentFileName <> "" then
                    if checkFile(filePath) = checkFile(parentFilePath) then
                        if getFSize(filePath) = getFSize(parentFilePath) then
                            if getFText(filePath) = getFText(parentFilePath) then                           '���ݶԱ�
                                call echoyellowb("ͬ���ϲ�", fileName & " >> " & parentFileName) 
                                fileName = parentFileName                                        '����һ���ļ���
                                call deleteFile(filePath)                                        'ɾ����ǰ�ļ�·��
                            end if 
                        end if 
                    end if 
                end if 
                s = url & "��|��" & localFileName & "��|��" & fileName & "��|��" & parentFileName & "��|��" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 
    end function 

    'ɾ����һ���ļ���    ��ǰ�ļ�������һ���ļ������ʱ
    function delParentFileName()
        dim splStr, splxx, s, c, url, filePath, toFilePath, fileName, tempFileName, fileType, imageType, parentFileName, localFileName, leftS, rightS 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) : tempFileName = fileName 
                parentFileName = splxx(3) 
                if fileName = parentFileName then
                'parentFileName=""
                end if 
                if instr(fileName, "_") > 0 then
                    leftS = mid(fileName, 1, instr(fileName, "_") - 1) 
                    rightS = mid(fileName, instr(fileName, "_") + 1) 
                    if len(leftS) = 8 then
                        filePath = debugWebDir & "/" & fileName 
                        toFilePath = debugWebDir & "/" & rightS 

                        if checkFile(toFilePath) = false then
                            call echoBlueB("�ļ�������С", filePath & "==>>" & toFilePath) 
                            call moveFile(filePath, toFilePath) 
                            fileName = rightS 
                        end if 
                    end if 
                end if 

                s = url & "��|��" & localFileName & "��|��" & fileName & "��|��" & parentFileName & "��|��" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 
    end function 


    '���������ļ��Ƿ��൱����ͬ��ϲ�
    function batchHandleImageWith()
        dim i, c 
        c = downUrlList 
        for i = 1 to 22
            call handleBatchHandleImageWith() 
            if c = downUrlList then
                exit for 
            end if 
        next 
    end function 
    '����������ͬͼƬ
    function handleBatchHandleImageWith()
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, parentFileName, newFileName, isTrue, findFileName, toFileName, localFileName 
        splStr = split(downUrlList, vbCrLf) 
        isTrue = true 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) : tempFileName = fileName 
                parentFileName = splxx(3) 
                filePath = debugWebDir & "/" & fileName 
                fileType = lcase(getFileAttr(fileName, "4")) 
                if isTrue = true then
                    newFileName = findWithContentFile(url, fileName, fileType) 
                    if newFileName <> "" and phpTrim(lcase(tempFileName)) <> phpTrim(lcase(newFileName)) and instr("|htm|html|", "|" & fileType & "|") = false then
                        findFileName = fileName 
                        toFileName = newFileName 
                        fileName = newFileName 
                        call echoBlueB("ͬ����", tempFileName & " >> " & newFileName) 
                        call deleteFile(filePath) 
                        isTrue = false                                                                  'Ϊ���򲻴����滻�ˣ���Ϊ��ǰ�滻���������һ���ļ�����Ժ�����һ���ļ�ͬ�࣬���滻���ǵ�ǰ����ļ��Ͳ�����
                    end if 
                end if 

                s = url & "��|��" & localFileName & "��|��" & fileName & "��|��" & parentFileName & "��|��" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 

        c = "" 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) : tempFileName = fileName 
                parentFileName = splxx(3) 
                if fileName = findFileName then
                    fileName = toFileName 
                end if 
                s = url & "��|��" & localFileName & "��|��" & fileName & "��|��" & parentFileName & "��|��" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 

    end function 

    '�滻html/css��ͼƬ·��
    function handleFileImagesPath()
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, imageType, parentFileName, content, localFileName, char_Set 
        dim splFile, fileList,parentFileExt,fileExt
        fileList = createFilePath & vbCrLf & getDirCssList(debugWebDir)'��ǰhtml��+CSS�б�
		if isHandleJsInHtml=true then
			fileList=fileList  & vbCrLf & getDirJSList(debugWebDir) '+JS�б�(0171113) ,�����js�ļ��жϱ���ʱ̫��
		end if
        splFile = split(fileList, vbCrLf) 
        splStr = split(downUrlList, vbCrLf) 
        for each filePath in splFile
            if filePath <> "" then
                char_Set = checkCode(filePath) 
				parentFileExt=getFileAttr(fileName,4)
                content = readFile(filePath, char_Set)
                call echo(filePath, char_Set) 
                for each s in splStr
                    if instr(s, "��|��") > 0 then
                        splxx = split(s, "��|��") 
                        url = splxx(0) 
                        fileName = splxx(2)
						fileExt=getFileAttr(fileName,4) 
                        content = replace(content, hRFileName(url), fileName) 
                    end if 
                next 
                call WriteToFile(filePath, content, char_Set) 
            end if 
        next 

    end function 

    '����ͬ�����ļ���
    function findWithContentFile(url, fileName, fileType)
        dim splStr, s, c, thisFilePath, filePath 
        splStr = split(getDirFileNameList(debugWebDir, fileType), vbCrLf) 
        for each s in splStr
            if s <> "" and s <> fileName then
                thisFilePath = debugWebDir & "/" & s 
                filePath = debugWebDir & "/" & fileName 
                if checkFile(thisFilePath) = true and checkFile(filePath) = true then
                    if getFSize(filePath) = getFSize(thisFilePath) then
                        if getFText(filePath) = getFText(thisFilePath) then
                            findWithContentFile = s 
                            exit function 
                        end if 
                    end if 
                end if 
            end if 
        next 
        findWithContentFile = "" 
    end function 
    '���ͬ����һ������
    function fileNameToParentNameMD5(byval findUrl, byval findFName)
        dim splStr, splxx, s, url, fileName 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
				if ubound(splxx)>=2 then
					url = splxx(0) 
					fileName = splxx(2) 
					if url <> findUrl and fileName = findFName then
						fileNameToParentNameMD5 = md5(url, 2) 
						exit function 
					end if 
				end if
            end if 
        next 
        fileNameToParentNameMD5 = "" 
    end function 
    '�ļ��� �� ��һ���ļ���
    function fileNameToParentName(parentFNameMD5)
        dim splStr, splxx, s, url, fileName 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                fileName = splxx(2) 
                if md5(url, 2) = parentFNameMD5 then
                    fileNameToParentName = fileName 
                    exit function 
                end if 
            end if 
        next 
        fileNameToParentName = "" 
    end function 
    '��ַ �� �ļ���
    function urlToFileName(findUrl)
        dim splStr, splxx, s, c, url, fileName, tempFileName, localFileName 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) : tempFileName = fileName 
                if findUrl = url then
                    urlToFileName = fileName 
                    exit function 
                end if 
            end if 
        next 
        urlToFileName = "" 
    end function 


    '�޸�httpurl״̬
    function editHttpUrlState(findUrl, stateIDValue)
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, imageType, parentFileName, parentFilePath, localFileName, stateID 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0)
				if ubound(splxx)>=4 then
					localFileName = splxx(1) 
					fileName = splxx(2) 
					parentFileName = splxx(3) 
					stateID = splxx(4) 
					if url = findUrl then
						stateID = stateIDValue 
					end if 
					s = url & "��|��" & localFileName & "��|��" & fileName & "��|��" & parentFileName & "��|��" & stateID 
	
					c = c & s & vbCrLf 
				end if
            end if 
        next 
        downUrlList = c 

        dim content 
        content = readFile(cacheConfigPath, "") 
        s = findUrl & "��|��" & stateIDValue 
        if instr(vbCrLf & content & vbCrLf, vbCrLf & s & vbCrLf) = false then
            content = content & s & vbCrLf 
            call WriteToFile(cacheConfigPath, content, "") 
        end if 
    end function 
    '�ӻ�������״̬��
    function getCacheState(findUrl)
        dim s, nLen, content 
        getCacheState = "" 
        content = vbCrLf & getftext(cacheConfigPath) & vbCrLf 
        s = vbCrLf & findUrl & "��|��" 
        nLen = instr(content, s) 
        if nLen > 0 then
            content = mid(content, nLen + len(s)) 
            content = mid(content, 1, instr(s, vbCrLf) + 2) 
            getCacheState = content 
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


    '����HTML����վ
    function handleHtmlToWeb()
        dim imagesDir, cssDir, jsDir 
        imagesDir = createHtmlWebDir & "/images"                                        'images���Ŀ¼
        cssDir = createHtmlWebDir & "/css"                                              'css���Ŀ¼
        jsDir = createHtmlWebDir & "/js"                                                'js���Ŀ¼
        call createDirFolder(imagesDir) 
        call createDirFolder(cssDir) 
        call createDirFolder(jsDir) 
		if createHelpFileContent<>"" then
			call createFile(createHtmlWebDir & "/README.txt" ,createHelpFileContent)
		end if

        dim splStr, splxx, s, c, url, filePath, toFilePath, fileName, tempFileName, fileType, imageType, parentFileName, localFileName, tempS 
        dim stateID, content 
        splStr = split(downUrlList, vbCrLf)  
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) 
                parentFileName = splxx(3) 
                stateID = splxx(4) 
                'ֻ���������ļ�����
                if stateID = "200"  or stateID="" then		'����Ϊ��20171218
                    filePath = debugWebDir & "/" & fileName 
                    fileType = lcase(getFileAttr(fileName, "4")) 
                    if fileType = "js" then
                        toFilePath = jsDir & "/" & fileName 
                        call deleteFile(toFilePath)                                                     'ɾ��css�ļ�����Ϊ��Ҫ���´���
                    elseif instr("|css|woff|woff2|eot|ttf|", "|" & fileType & "|") > 0 then
                        toFilePath = cssDir & "/" & fileName 
                        call deleteFile(toFilePath)                                                     'ɾ��css�ļ�����Ϊ��Ҫ���´���
                    elseif instr("|htm|html|", "|" & fileType & "|") > 0 then
                        filePath = debugWebDir & "/index_" & fileName 
                        toFilePath = createHtmlWebDir & "/index_" & fileName 
                        call deleteFile(toFilePath)                                                     'ɾ��html�ļ�����Ϊ��Ҫ���´���
                    else
                        toFilePath = imagesDir & "/" & fileName 										'Ĭ�Ϸŵ�images�ļ�����
					end if 
                    '�ļ��������ж�
                    if checkFile(filePath) = true and checkFile(toFilePath) = false then
                        call copyFile(filePath, toFilePath) 
                        if fileType = "css" then
                            content = readFile(toFilePath, "") 
                            content = pubCssObj.handleCssContent("css", content, "*", "background||background-image", "*", "isimg(*)trim(*)dir(*)../images/", "set") 
                            call WriteToFile(toFilePath, content, "") 
						
						'����js�ļ�ʱҪ���� ����js���html
                        elseif fileType = "html" or fileType = "htm" or (fileType = "js" and isHandleJsInHtml=true) then
                            call handleHtmlContentPath(filePath & ".txt", toFilePath, "") 
                        end if 
                    end if 
                end if 
            end if 
        next 
    end function 

    '����HTML��ģ��
    function handleHtmlToTemplate()
        dim imagesDir, cssDir, jsDir 
        imagesDir = templateDir & "/images"                                             'images���Ŀ¼ /Templates2015/sharembweb/Images/
        cssDir = templateDir & "/css"                                                   'css���Ŀ¼
        jsDir = templateDir & "/js"                                                     'js���Ŀ¼
        call createDirFolder(imagesDir) 
        call createDirFolder(cssDir) 
        call createDirFolder(jsDir) 

        dim splStr, splxx, s, c, url, filePath, toFilePath, fileName, tempFileName, fileType, imageType, parentFileName, localFileName, tempS 
        dim stateID, content, websiteFilePath, websiteFileContent 
        splStr = split(downUrlList, vbCrLf)  
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) 
                parentFileName = splxx(3) 
                stateID = splxx(4) 
                'ֻ���������ļ�����
                if stateID = "200" or stateID="" then		'����Ϊ��20171218
                    filePath = debugWebDir & "/" & fileName 
                    fileType = lcase(getFileAttr(fileName, "4")) 
                    if fileType = "js" then
                        toFilePath = jsDir & "/" & fileName 
                        call deleteFile(toFilePath)                                                     'ɾ��css�ļ�����Ϊ��Ҫ���´���
                    elseif instr("|css|woff|woff2|eot|ttf|", "|" & fileType & "|") > 0 then
                        toFilePath = cssDir & "/" & fileName 
                        call deleteFile(toFilePath)                                                     'ɾ��css�ļ�����Ϊ��Ҫ���´���
                    elseif instr("|htm|html|", "|" & fileType & "|") > 0 then
                        filePath = debugWebDir & "/index_" & fileName 
                        toFilePath = templateDir & "/index_Model" & fileName 
                        call deleteFile(toFilePath)                                                     'ɾ��html�ļ�����Ϊ��Ҫ���´���
                    else
						 toFilePath = imagesDir & "/" & fileName 
					end if 
                    '�ļ��������ж�
                    if checkFile(filePath) = true and checkFile(toFilePath) = false then
                        call copyFile(filePath, toFilePath) 
                        if fileType = "css" then
                            content = readFile(toFilePath, "") 
                            content = pubCssObj.handleCssContent("css", content, "*", "background||background-image", "*", "isimg(*)trim(*)dir(*)" & "/templates/" & templateName & "/images/", "set") 
                            call WriteToFile(toFilePath, content, "") 

                        elseif fileType = "html" or fileType = "htm" then
                            call handleHtmlContentPath(filePath & ".txt", toFilePath, "template") 
                        end if 
                    end if 
                end if 
            end if 
        next 
        '����ģ�嵽ָ���ļ�����
        call deleteFolder(toTemplateDir) 
        call copyFolder(templateDir, toTemplateDir) 

        call copyFolder("/Data/WebData", toTemplateDir & "/WebData") 
        websiteFilePath = toTemplateDir & "/WebData/website.txt" 
        websiteFileContent = getftext(websiteFilePath) 
        s = getStrCut(websiteFileContent, "��webtemplate��", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "��webtemplate��" & toTemplateDir & vbCrLf) 
        s = getStrCut(websiteFileContent, "��webimages��", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "��webimages��" & toTemplateDir & "/Images" & vbCrLf) 
        s = getStrCut(websiteFileContent, "��webcss��", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "��webcss��" & toTemplateDir & "/Css" & vbCrLf) 
        s = getStrCut(websiteFileContent, "��webjs��", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "��webjs��" & toTemplateDir & "/Js" & vbCrLf) 

        s = getStrCut(websiteFileContent, "��webtitle��", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "��webtitle��" & templateName & "����" & vbCrLf) 
        s = getStrCut(websiteFileContent, "��webkeywords��", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "��webkeywords��" & templateName & "�ؼ���" & vbCrLf) 
        s = getStrCut(websiteFileContent, "��webdescription��", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "��webdescription��" & templateName & "����" & vbCrLf) 

        call createfile(websiteFilePath, websiteFileContent) 



    end function 

    '����HTML�ļ�·��
    function handleHtmlContentPath(filePath, toFilePath, sType)
        dim splStr, splxx, s, c, url, fileName, tempFileName, fileType, imageType, parentFileName, content, localFileName, char_Set 
        dim splFile, fileList, navC, replaceFileName 
        fileList = createFilePath & vbCrLf & getDirCssList(debugWebDir)                 '��ǰhtml��+CSS�б�
        splFile = split(fileList, vbCrLf) 
        splStr = split(downUrlList, vbCrLf) 

        char_Set = checkCode(filePath) 
        content = readFile(filePath, char_Set) 
        for each s in splStr
            if instr(s, "��|��") > 0 then
                splxx = split(s, "��|��") 
                url = splxx(0) 
                fileName = splxx(2) 
                fileType = lcase(getFileAttr(fileName, "4")) 

                if fileType = "js" then
                    replaceFileName = "js/" & fileName 
                    if sType = "template" then
                        replaceFileName = "[$WebJs$]/" & fileName 
                    end if 
                elseif fileType = "css" then
                    replaceFileName = "css/" & fileName 
                    if sType = "template" then
                        replaceFileName = "[$WebCss$]/" & fileName 
                    end if 
                else
                    replaceFileName = "images/" & fileName 
                    if sType = "template" then
                        replaceFileName = "[$WebImages$]/" & fileName 
                    end if 
                end if 
                content = replace(content, hRFileName(url), replaceFileName) 
            end if 
        next 
        '��������ģ��
        if sType = "template" then
            dim fObj : set fObj = new class_formatting
                call fObj.handleFormatting(content) 
                call fObj.delLabelHtml("title") 
                call fObj.delLabelHtml("meta[**name=keywords]") 
                call fObj.delLabelHtml("meta[**name=description]") 

                c = vbCrLf & "<title>{$Web_Title$}</title>" & vbCrLf 
                c = c & "<meta name=""keywords"" content=""{$Web_KeyWords$}"" />" & vbCrLf 
                c = c & "<meta name=""description"" content=""{$Web_Description$}"" />" 
                call fObj.handleLabelContent("head", "*", c, "prepend") 

                content = fObj.getFormattingHtml("", "", "") 
                '���������ģ��
                navC = fObj.handleLabelContent("ul[**class=naV]", "*", "", "get+label") 
                call echo(typeName(content), typeName(navC)) 
                content = pubTemplateObj.batchHandleWebNavLayout(content, navC) 

                navC = fObj.handleLabelContent("div[**class=columnlist2]", "*", "", "get+label") 
                content = pubTemplateObj.batchHandleWebNewsLayout(content, navC) 

        end if
        call WriteToFile(toFilePath, content, char_Set) 
        handleHtmlContentPath = content 
    end function 
    '������־
    function saveLog()
        dim url, filePath, fileName, splStr, errS 

        downC = downC & getImitate_print(md5(httpurl, 2), htmlFilePath, jsListC, cssListC, imgListC, swfListC, otherListC, urlListC, getErrorInfoUrlList(fileName)) & vbCrLf 

        splStr = split(thisFormatObj.cssList, vbCrLf) 
        for each url in splStr
            if url <> "" then
                fileName = urlToFileName(url) 
                filePath = debugWebDir & "/" & fileName 
                errS = getErrorInfoUrlList(fileName) 
                if errS <> "" then
                    downC = downC & getImitate_print(md5(url, 2), filePath, "", "", "", "", "", "", errS) & vbCrLf 
                end if 
            end if 
        next 
        call createfile(debugWebDir & "/Z��ӡ��վ��־.html", imitate_config(downC)) 
    end function 


    '**************************************************
    '��ô�����Ϣ��ַ�б� 20161002
    function getErrorInfoUrlList(byval findFileName)
        dim content, splStr, splxx, s, url, fileName, downStatus, localSaveFileName 
        content = getftext(debugWebDir & "/������ַ�б�.txt") 
        splStr = split(content, vbCrLf) 
        errInfoC = "" 
        for each s in splStr
            splxx = split(s, "��|��") 
            if uBound(splxx) >= 4 then
                url = splxx(0) 
                fileName = splxx(1) 
                localSaveFileName = splxx(2) 
                downStatus = splxx(4) 
                '�����������Ϊ���β��������µģ���������ȥ��������һ�¾Ϳ����ˡ�OK
                if downStatus = "State״̬��" then
                    downStatus = getCacheState(url) 
                end if 

                if findFileName = fileName and downStatus <> "200" then
                    errInfoC = errInfoC & url & "��>>��" & fileName & "��>>��" & localSaveFileName & "��>>��" & downStatus & vbCrLf '������Ϣ�ۼ�
                end if 
            end if 
        next 
        getErrorInfoUrlList = errInfoC 
    end function 
    '�б����ͼƬ�б�
    function getErrorImagesList()
        dim content, splStr, splxx, s, url, fileName, downStatus, goToUrl, fileNameList 
        content = getftext(debugWebDir & "/������ַ�б�.txt") 
        splStr = split(content, vbCrLf) 
        fileNameList = "|" 
        for each s in splStr
            splxx = split(s, "��|��") 
            if uBound(splxx) >= 4 then
                goToUrl = splxx(1) 
                url = splxx(2) 
                fileName = splxx(3) 
                downStatus = splxx(4) 
                if downStatus <> "200" then
                    fileNameList = fileNameList & fileName & "|" 
                end if 
            end if 
        next 
        getErrorImagesList = fileNameList 
    end function 
 
	'�����滻Դ��ҳ�������� 20111109
	function handleReplaceSourcePageConfig(byval content) 
		dim splstr,s,splxx,sLeft,sRight,startStr,endStr,sStr
		splstr=split(replaceSourcePageConfig,"====================") 
		for each s in splstr
			if instr(s,"��|��")>0 and left(phptrim(s),2)<>"##" then
				splxx=split(s,"��|��")
				sLeft=phptrim(splxx(0))
				sRight=phptrim(splxx(1))
				if lcase(left(sLeft,10))=lcase("getStrCut(") then
					splxx=split(sLeft,"��,��")
					startStr=mid(splxx(0),11):endStr=left(splxx(1),len(splxx(1))-1)
					'call echo("sLeft",replace(sLeft,"<","&lt;"))
					'call echo("startStr",replace(startStr,"<","&lt;"))
					'call echo("endStr",replace(endStr,"<","&lt;"))
					sStr=getStrCut(content,startStr,endStr,1)
					'call echo("sStr",replace(sStr,"<","&lt;"))
					content=replace(content,sStr,sRight)
				else
					content=replace(content,sLeft,sRight)
				
				end if
				
				'call echo("�滻","")
			end if
		next
		handleReplaceSourcePageConfig=content
	end function
	'ɾ���������������ҳ���� 20171111
	function deleteHandleHtmlWeb(byval content)
		dim startStr,endStr,s,i
		startStr="<DiV class=hide>"
		endStr="</DiV>"
		for i=1 to 99
			if instr(content,startStr)>0 and instr(content,endStr)>0 then
				s=getStrCut(content,startStr,endStr,1)
				'call echo("s",s)
				content=replace(content,s,"")
			else
				exit for
			end if
		next
		deleteHandleHtmlWeb=content
	end function
	'�����������script js�ű� 20171114 
	function handleContentInScriptHtml(byval content,byval httpurl, byval fileName)										
		dim c,splstr,s,replaceS,thisJsObj
		dim thisFormatObj : set thisFormatObj=new class_formatting			'��ʽ����
		thisFormatObj.handleFormatting(content) 
		c=thisFormatObj.handleLabelContent("script","*","","get+label")
		splstr=split(c,"$Array$")
		for each s in splstr
			if s <>"" and instr(s,"[##")=false then
					replaceS=s
					'�ų�Jquery�ļ�����Ϊ������   Ҫ��������js���html
					if instr(replaceS,".jQuery=")=false and isHandleJsInHtml=true then 
						set thisJsObj=new class_js							'js��
						thisJsObj.httpurl=httpurl		'����js��ַ
						replaceS=  thisJsObj.handleJsContent(replaceS,"source|html")		'����js������ݣ���html��ͼƬ����
						'������ҳ
						call handleDownList(httpurl, fileName, replaceS, thisJsObj.imgList, "img") 
						call handleDownList(httpurl, fileName, replaceS, thisJsObj.swfList & vbcrlf & thisJsObj.sourceList & vbcrlf & thisJsObj.videoList, "swf")
						call handleDownList(httpurl, fileName, replaceS, thisJsObj.jsList, "js") 
						call handleDownList(httpurl, fileName, replaceS, thisJsObj.cssList, "css")  
					end if
			
			
				content=replace(content,s, replaceS)
			end if
		next 
		handleContentInScriptHtml=content
	end function
end class 



%>  



