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
'仿站
class imitateWeb
    'dubug/下载网址列表.txt   此文件不要删除，否则二次下载时会 重复下载图片等文件
    dim thisFormatObj                                                               '格式HTML类
    dim version                                                                     '版本
    dim defaultIEContent                                                            '默认IE显示内容
    dim templateName                                                                '模板名称
    dim templateDir                                                                 '模板路径
    dim toTemplateDir                                                               '复制模板到其它目录
    dim isWebToTemplateDir                                                          '文件是否放到模板文件夹里

    dim serverHttpUrl                                                               '服务器网址
    dim httpurl 																	'采集网址
    dim newWebDir                                                                   '新网站目录
    dim debugWebDir                                                                 '调试网站目录
    dim createHtmlWebDir                                                            '创建html网站目录
    dim cacheFolderPath                                                             '缓冲文件夹路径
    dim cacheConfigPath                                                             '缓冲配置路径，记录URL状态码
    dim customCharSet                                                               '自定义编码
    dim isGetHttpUrl                                                                'getHttpUrl方式获得内容
    dim isMakeWeb                                                                   '生成WEB文件夹
    dim isMakeTemplate                                                              '是否生成模板文件夹
    dim isPackWeb                                                                   '是否打包WEB文件夹 xml
    dim isPackWebZip                                                                '是否打包WEB文件夹 zip
    dim isPackTemplate                                                              '是否打包模板文件夹 xml
    dim isPackTemplateZip                                                           '是否打包模板文件夹 zip
    dim isAddWebTitleKeyword                                                        '追加网站标题描关键词
    dim isOnMsg                                                                     '是否开启回显信息
    dim pubAHrefList                                                                'A链接列表
    dim isUTF8Down                                                                  '以utf8方式下载
    dim downUrlList                                                                 '下载的URL列表
    dim fileNameList                                                                '文件名称列表
    dim saveFolderName                                                              '保存文件夹名称
    dim downCountSize                                                               '下载总大小
    dim downFileTypeList                                                            '下载文件类型列表
    dim isDownServer                                                                '是否从服务器开始下载
    dim isDownCache                                                                 '是否雇用缓冲下载方式
    dim downImagesList                                                              '下载图片列表
    dim downImageAddMyInfoType                                                      '下载图片加我自己的信息
	dim isDeleteErrorImages															'是否删除删除图片
	dim deleteErrorImagesList														'删除错误图片列表
	dim isStop																		'停止 
	dim isEditImageType																'修改图片类型
	dim isWithNameHeBin																'同名图片合并
	dim isWithContentHeBin															'同内容文件合并
	dim isAddMyInfo																	'追加我的信息
	dim isHandleJsInHtml															'是否处理js里的html 20171114
	dim replaceSourcePageConfig														'替换原网页内容 20171109
	dim createHelpFileContent														'生成帮助文档内容 20171109
	

    dim jsListC, cssListC, imgListC, swfListC,htmlListC, otherListC, urlListC, errInfoC, downC, nDownCount '记录下载回显信息
	dim cacheFilepath,htmlFilePath, htmlBaseFilePath, createFilePath			'缓冲文件路径，html源文件路径，html源文件+base路径，创建html文件
	dim isDownCssImage,isDownStyleImage,isHtmlImg				'是否下载CSS图片，样式图片,Img图片
	dim htmlformat,jsformat,cssformat								'html,js,css格式化类型
	dim useUniformCoding															'选择统一编码

    '构造函数 初始化
    sub class_Initialize()
        call loadDefaultConfig() 
    end sub 

    '加载默认配置
    sub loadDefaultConfig()
        set thisFormatObj = new class_formatting                                        '初始化格式化HTML类
            version = "仿站 v5.6"                                                        '软件版本 v1.1 beta version
            serverHttpUrl = "http://aa/server_FangZan.Asp"                                  '等改成sharembweb.com
            httpurl = "http://sharembweb.com/" 
            newWebDir = "newweb"                                                            '新网站目录
            debugWebDir = "dubug"                                                           '调试网站目录
            createHtmlWebDir = "web"                                                        '生成网站目录
            customCharSet = "自动检测"                                                      '编码
            isGetHttpUrl = false                                                            'getHttpUrl方式获得内容
            isMakeWeb = true                                                                '生成WEB文件夹
            isMakeTemplate = false                                                          '是否生成模板文件夹
            isPackWeb = false                                                               '是否打包WEB文件夹 xml
            isPackWebZip = false                                                            '是否打包WEB文件夹 zip
            isWebToTemplateDir = true                                                       '文件是否放到模板文件夹里
            toTemplateDir = "/templates"                                                    '模板复制到目录
            isPackTemplate = false                                                          '是否打包模板文件夹 xml
            isPackTemplateZip = false                                                       '是否打包模板文件夹 zip

            isAddWebTitleKeyword = true                                                     '追加网站标题描关键词
            isOnMsg = false                                                                 '是否开启回显信息
            templateName = getWebSiteName(httpurl)                                          '模板名称从网址中获得
            isUTF8Down = false                                                              'utf8方式下载为假
            downCountSize = 0                                                               '下载总大小
            downFileTypeList = "|*|"                                                        '下载文件类型列表
            isDownServer = true                                                             '是否可以从服务器端下载
            isDownCache = true                                                              '默认不启用下载缓冲
            downImageAddMyInfoType = "2"                                                    '下载图片加我自己的信息
            cacheFolderPath = ""                                                            '缓冲目录 为空则在当前目录下创建一个cache目录
            call createDirFolder(cacheFolderPath) 
            defaultIEContent = "[defaultIEContent]"                                    		'默认IE显示内容
            isDeleteErrorImages = true                                                      '删除错误图片(web目录与templates)
            isStop = false                                                                  '默认停止为假
            isDownCssImage = true : isDownStyleImage = true : isHtmlImg = true 				'默认下载图片为真

            isEditImageType = true                                                          '修改图片类型
            isWithNameHeBin = true                                                          '同名图片合并
            isWithContentHeBin = false                                                      '同内容文件合并
            isAddMyInfo = true                                                              '追加我的信息
			isHandleJsInHtml=false															'是否处理js里的html 为假
    end sub


    '下载网站
    function downweb()
        dim s, websiteFilePath, websiteFileContent 
        dim filePath, toFilePath, fileList, splList, fileName, splStr, startTime, splUrl,splxx
        startTime = now() 
        httpurl = phptrim(httpurl) 
        if httpurl = "" then
            call eerr("停止", "网址为空") 
        elseif len(httpurl)<=3 then
            call eerr("停止", "网址不合法") 
        end if
        isStop = false                                                                  '停止为假

        '获得仿站目录
        newWebDir = newWebDir & "/" & IIf(saveFolderName <> "", saveFolderName, getWebSiteCleanName(httpurl)) 
        '创建基础文件夹
        call createfolder(newWebDir) 
        '创建调试目录
        debugWebDir = newWebDir & "/" & debugWebDir 
        call createDirFolder(debugWebDir) 
        '创建缓冲目录 前提是开启缓冲 和 默认缓冲文件夹为空
        if cacheFolderPath = "" and isDownCache = true then
            cacheFolderPath = newWebDir & "/cache" 
            call createDirFolder(cacheFolderPath) 
        end if 

        '创建html网站目录
        if isMakeWeb = true then
            createHtmlWebDir = newWebDir & "/web" 
            call createDirFolder(createHtmlWebDir) 
        end if 
        '模板目录
        if isWebToTemplateDir = true then
            templateDir = newWebDir & "/templates/" & templateName                          '生成当前模板目录
            toTemplateDir = toTemplateDir & "/" & templateName                              '复制到指定模板目录
            call createDirFolder(templateDir) 												'创建模板目录
        end if 

        downUrlList = readFile(debugWebDir & "/下载网址列表.txt", "")                   '读下载网址列表
		
		if instr(httpurl,"|")>0 then
			splxx=split(httpurl,"|")
			httpurl=splxx(0)
			fileName=splxx(1)
		end if 
		if fileName="" then
        	fileName = getUrlFileName(httpurl, "", getUrlToFileName(httpurl, ":*.html"))    '获得html文件名  可在本地配置文件夹里修改
		else
        	fileName = getUrlFileName(httpurl, "", fileName)   
		end if

        cacheFilePath = cacheFolderPath & "/" & md5(httpurl, 2) & ".htm"                '缓冲文件路径
        htmlFilePath = debugWebDir & "/html_" & fileName                                '源文件路径
        htmlBaseFilePath = debugWebDir & "/base_" & fileName                            'base文件
        createFilePath = debugWebDir & "/index_" & fileName                             '生成文件
        cacheConfigPath = cacheFolderPath & "/cacheConfig.txt"                          '缓冲配置路径
 

        call echo("缓冲文件路径(cacheFilePath)", cacheFilepath) 
        call echo("HTML文件路径(htmlFilePath)", htmlFilePath) 
        call echo("HTML基础文件路径(htmlBaseFilePath)", htmlBaseFilePath) 
        call echo("生成HTML文件路径(createFilePath)", createFilePath) 
        if checkFile(htmlFilePath) = false then
            if checkFile(cacheFilePath) = true then
                call copyFile(cacheFilepath, htmlFilePath) 
                call echo(httpurl, getCacheState(httpurl)) 
                call editHttpUrlState(httpurl, getCacheState(httpurl))                          '追加状态码
                call echo("downurllist", downUrlList)  
            else
                downweb = false						'下载为假则不执行模仿
                exit function 
            end if 
        end if 
        call imitateWeb() 			'仿站核心程序
        downweb = true				'返回为真

    end function 
    '仿站
    function imitateWeb()
        dim content, fileName 
        dim htmlFileSize, char_Set, splStr, splxx, s, c, htmlFPath, cacheFPath, isAdd 
        fileName = getFileAttr(createFilePath, "2")                                     '获得文件名称
        content = readFile(htmlFilePath, char_Set) 
        char_Set = customCharSet 
        '非常重要，以后看到这里，记住不要删除了，20161011
        if customCharSet = "自动检测" then
            doevents 
            char_Set = getFileCharset(htmlFilePath) 
            content = readFile(htmlFilePath, char_Set)                                      '这里读是为了自定找编码
            call echo("自动检测 生成HTML文件编码", htmlFilePath & " = "& char_Set) 
            if char_Set = "gb2312" then
                s = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "get", "lcase(*)") 
                if instr(s, "utf-8") > 0 then
                    content = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "set", "text/html; charset=gb2312") 
                    call echo("见鬼", "检测编码与文件编码不一样1") 
                end if 
            elseif char_Set = "utf-8" then
                s = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "get", "lcase(*)") 
                if instr(s, "gb2312") > 0 then
                    content = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "set", "text/html; charset=utf-8") 
                    call echo("见鬼", "检测编码与文件编码不一样2") 
                end if 
            end if 
            call WriteToFile(createFilePath, phptrim(content), char_Set)                    '保存模板文件(去两边空格)
        end if
		
		content=handleReplaceSourcePageConfig(content)										'处理源网页替换20171109
		if char_Set = "gb2312" then
			content=pubHtmlObj.handleHtmlLabel(content,"meta","content","**content=charset","set","replace[**utf-8=gb2312]isreplace(*)")
			content=pubHtmlObj.handleHtmlLabel(content,"meta","charset","","set","replace[**utf-8=gb2312]isreplace(*)")
		end if

        '保存base文件
        if checkFile(htmlBaseFilePath) = false then
            c = regExp_Replace(content, "<head>", "<head>" & vbCrLf & "<base href=""" & getWebsite(httpurl) & """ />" & vbCrLf) 
            call WriteToFile(htmlBaseFilePath, c, char_Set) 
        end if 
        content = replaceTemplateContent(content)                                       '处理模板里替换内容 urllistnoenc(*)setaddurlenc(*)
        '格式化HTML
        call thisFormatObj.handleFormatting(content) 
        call thisFormatObj.handleLabelContent("*", "*", "setaddurlenc(*)isimg(*)fullurl(*)" & httpurl, "handleimg") '处理CSS背景图片 lcase(*)拿掉 不要用linux系统的大小写区别就有问题了 
        call handleDownList(httpurl, fileName, content, thisFormatObj.imgList, "img") 
        call handleDownList(httpurl, fileName, content, thisFormatObj.swfList & vbcrlf & thisFormatObj.sourceList & vbcrlf & thisFormatObj.videoList, "swf")
        call handleDownList(httpurl, fileName, content, thisFormatObj.jsList, "js") 
        call handleDownList(httpurl, fileName, content, thisFormatObj.cssList, "css") 
		'call echoyellow("thisFormatObj.videoList",thisFormatObj.videoList)

        '需要解压或加密时再次处理html文件内容
        if htmlformat <> "" or jsformat <> "" or cssformat <> "" then
            call thisFormatObj.updateHtml() 
        end if 
        call echoredb("cssformat", cssformat) 
        '获得格式化的HTML
        content = thisFormatObj.getFormattingHtml(htmlformat, jsformat, cssformat) 

        '使用统一编码
        if useUniformCoding <> "" and char_Set <> useUniformCoding then
            char_Set = useUniformCoding 
            content = pubHtmlObj.handleHtmlLabel(content, "meta", "content", "**http-equiv=Content-Type", "set", "text/html; charset=" & char_Set) '重复设置
        end if 
		content=deleteHandleHtmlWeb(content)											'删除处理后生成完网页内容
		if isHandleJsInHtml=true then
			content=handleContentInScriptHtml(content,httpurl,fileName)										'处理内容里的script js脚本
		end if
        '生成处理后文件
        call WriteToFile(createFilePath, phptrim(content), char_Set)                    '保存模板文件(去两边空格)
        call WriteToFile(createFilePath & ".txt", phptrim(content), char_Set)           '保存模板文件(去两边空格)

    end function 
    '处理下载列表
    function handleDownList(byval httpurl, localFilePath, byval content, urlList, sType)
        dim splStr, s, c, url, fileName, parentFileName, filePath, cssContent, tempC, isEchoDown, char_Set 
        splStr = split(urlList, vbCrLf) 
        for each s in splStr
            '停止时退出
            if isStop = true then
                exit function 
            end if 
            '排除空与#号
            if phptrim(s) <> "" and getwebsite(httpurl) <> s then
                url = fullHttpUrl(httpurl, s) 
                'img
                if sType = "img" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.jpg")) 
                    imgListC = imgListC & url & "【|】" & fileName & vbCrLf 

                'swf
                elseif sType = "swf" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.swf")) 
                    swfListC = swfListC & url & "【|】" & fileName & vbCrLf 

                'css
                elseif sType = "css" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.css")) 
                    cssListC = cssListC & url & "【|】" & fileName & vbCrLf 

                'js
                elseif sType = "js" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.js")) 
                    jsListC = jsListC & url & "【|】" & fileName & vbCrLf 

                'html
                elseif sType = "html" then
                    fileName = getUrlFileName(url, localFilePath, getUrlToFileName(url, ":*.html")) 
                    htmlListC = htmlListC & url & "【|】" & fileName & vbCrLf 
                end if 
            end if 
        next 
        handleDownList = content 
    end function 

    '获得文件名称 并累加下载网址列表
    function getUrlFileName(byval url, localFilePath, fileName)
        dim splStr, splxx, s, parentFileName, isAdd 
        parentFileName = fileNameToParentNameMD5(url, fileName) 
        if parentFileName <> "" then
            fileName = md5(url, 2) & "_" & fileName 
        end if 
        splStr = split(downUrlList, vbCrLf) 
        getUrlFileName = fileName 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
                'call echoyellowb(url,splxx(0))
                if splxx(0) = url and splxx(1) = localFilePath then    '网址与本地保存文件相同 修20171110
                    getUrlFileName = splxx(2) 
                    call echoyellowb("根据URL找到 filename=", splxx(2)) 
                    exit function 
                end if 
            end if 
        next 
        'call hr()
		if right(downUrlList,2)<>vbcrlf then
			downUrlList=downUrlList & vbcrlf
		end if
        downUrlList = downUrlList & url & "【|】" & localFilePath & "【|】" & fileName & "【|】" & parentFileName & "【|】State状态码"
    end function 
    '处理CSS文件内容
    function handleCssFileContent()
        dim splStr, splxx, s, url, fileName, filePath, content, imgList, char_Set 
        splStr = split(cssListC, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
                url = splxx(0) 
                fileName = splxx(1) 
                filePath = debugWebDir & "/" & fileName 
                content = readFile(cacheFolderPath & "/" & md5(url, 2) & "." & getFileAttr(splxx(1), "4"), "") 
                call echoB("处理CSS文件", filePath):doevents
                imgList = pubCssObj.handleCssContent("css", content, "*", "background||background-image||src", "*", "isimg(*)lcase(*)fullurl(*)" & s, "images") 
                content = pubCssObj.handleCssContent("css", content, "*", "background||background-image||src", "*", "setaddurlenc(*)isimg(*)lcase(*)fullurl(*)" & s, "set") 
                'call echoredb(httpurl,imgList)
                call handleDownList(httpurl, fileName, content, imgList, "img")   '下载
                '处理压缩与解压CSS		
                if cssformat <> "" then
                    content = pubCssObj.handleCssContent("css", content, "", "", "", "", cssformat) 
                end if 
                char_Set = "" 
                '使用统一编码
                if useUniformCoding <> "" then
                    char_Set = useUniformCoding 
                end if 
                call WriteToFile(filePath, phptrim(content), char_Set) 
            end if 
        next 

    end function 
    '处理JS文件内容
    function handleJSFileContent()
        dim splStr, splxx, s, url, fileName, filePath, content, imgList, char_Set 
		dim thisJsObj	
		if isHandleJsInHtml=false and jsformat="" and useUniformCoding="" then
			call echoYellow("提示","处理js文件未设置选项，退出操作")
			exit function
		else
			call echo("isHandleJsInHtml",isHandleJsInHtml)
			call echo("jsformat",jsformat)
			call echo("useUniformCoding",useUniformCoding)
		end if
		
		
        splStr = split(jsListC, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
                url = splxx(0) 
                fileName = splxx(1) 
                filePath = debugWebDir & "/" & fileName 
				if checkFile(filePath)=true then		'文件存在则处理
					content = readFile(cacheFolderPath & "/" & md5(url, 2) & "." & getFileAttr(splxx(1), "4"), "") 
					
					'排除Jquery文件，因为处理不了   要开启处理js里的html
					if instr(content,".jQuery=")=false and isHandleJsInHtml=true then 
						call echoB("处理JS文件里html", filePath):doevents
						set thisJsObj=new class_js							'js类
						thisJsObj.httpurl=httpurl		'设置js网址
						content=  thisJsObj.handleJsContent(content,"source|html")		'处理js里的内容，让html里图片完整
						'下载网页
						call handleDownList(httpurl, fileName, content, thisJsObj.imgList, "img") 
						call handleDownList(httpurl, fileName, content, thisJsObj.swfList & vbcrlf & thisJsObj.sourceList & vbcrlf & thisJsObj.videoList, "swf")
						call handleDownList(httpurl, fileName, content, thisJsObj.jsList, "js") 
						call handleDownList(httpurl, fileName, content, thisJsObj.cssList, "css")  
					end if
					
					
					'处理压缩与解压JS
					if jsformat <> "" then
						call echoB("处理JS文件", filePath):doevents
						content = pubJsObj.handleJsContent(content, jsformat) 
					end if 
					char_Set = "" 
					'使用统一编码
					if useUniformCoding <> "" then
						char_Set = useUniformCoding 
					end if 
					call WriteToFile(filePath, phptrim(content), char_Set) 
					call WriteToFile(filePath & ".txt", phptrim(content), char_Set)  '2017114 对js做备份，方便处理里面的html及图片css/js路径
				end if
            end if 
        next 

    end function 

    '追加我的信息
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
                '判断是否追加我的信息
                if instr("|jpg|gif|bmp|png|", "|" & imageType & "|") > 0 then

                    if getFSize(filePath) < 30240 then                                              '小于30K
                        c = readBinary(filePath, 0) 
                        s = imageAddMyInfo(downImageAddMyInfoType) 
                        if instr(c, s) = false then
                            call decSaveBinary(filePath, c & "|" & s, 0) 
                            call echo("追加我的信息", filePath) 
                        else
                            call echoGay("无需追加我的信息，已存在", filePath) 
                        end if 
                    end if 
                end if 
            end if 
        next 
    end function 
    '修改图片类型
    function editImageType()
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, imageType, parentFileName, localFileName, tempS 
        dim urlList : urlList = "," 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) : tempFileName = fileName 
                parentFileName = splxx(3) 
                filePath = debugWebDir & "/" & fileName 
                fileType = lcase(getFileAttr(fileName, "4")) 
                '文件存在则判断
                if checkFile(filePath) = true then
                    imageType = lcase(getImageType(filePath))                                       '图片真实类型

                    if instr("|bmp|jpg|gif|png|", "|" & fileType & "|") > 0 and fileType <> imageType and imageType <> "" then
                        fileName = getFileAttr(fileName, "3") & "." & imageType 
                        call moveFile(debugWebDir & "/" & tempFileName, debugWebDir & "/" & fileName) 
                        urlList = urlList & url & "|" & fileName & "," 
                        call echo("文件类型", url & " 文件名 " & tempFileName & " 类型 " & fileType & " >> " & imageType) 
                    end if 
                elseif instr(urlList, "," & url & "|") > 0 then
                    tempS = mid(urlList, instr(urlList, "," & url & "|") + len("," & url & "|")) 
                    'call echo("tempS",tempS)
                    fileName = mid(tempS, 1, instr(tempS, ",") - 1) 

                    'call echo("tempS",tempS)
                    'call echoblueb("urlList",urlList)
                    'call echoredb(url & "("& checkFile(filePath) &")",tempFileName & ">>" & fileName)
                    call echo("文件类型(被动)", url & " 文件名 " & tempFileName & " 类型 " & fileType & " >> " & imageType) 

                end if 
                s = url & "【|】" & localFileName & "【|】" & fileName & "【|】" & parentFileName & "【|】" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 
    end function 
    '同名合并
    function withNameHeBin()
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, imageType, parentFileName, parentFilePath, localFileName 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
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
                            if getFText(filePath) = getFText(parentFilePath) then                           '内容对比
                                call echoyellowb("同名合并", fileName & " >> " & parentFileName) 
                                fileName = parentFileName                                        '用上一级文件名
                                call deleteFile(filePath)                                        '删除当前文件路径
                            end if 
                        end if 
                    end if 
                end if 
                s = url & "【|】" & localFileName & "【|】" & fileName & "【|】" & parentFileName & "【|】" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 
    end function 

    '删除上一级文件名    当前文件名与上一经文件名相等时
    function delParentFileName()
        dim splStr, splxx, s, c, url, filePath, toFilePath, fileName, tempFileName, fileType, imageType, parentFileName, localFileName, leftS, rightS 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
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
                            call echoBlueB("文件名称缩小", filePath & "==>>" & toFilePath) 
                            call moveFile(filePath, toFilePath) 
                            fileName = rightS 
                        end if 
                    end if 
                end if 

                s = url & "【|】" & localFileName & "【|】" & fileName & "【|】" & parentFileName & "【|】" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 
    end function 


    '批量搜索文件是否相当，相同则合并
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
    '批量处理相同图片
    function handleBatchHandleImageWith()
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, parentFileName, newFileName, isTrue, findFileName, toFileName, localFileName 
        splStr = split(downUrlList, vbCrLf) 
        isTrue = true 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
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
                        call echoBlueB("同内容", tempFileName & " >> " & newFileName) 
                        call deleteFile(filePath) 
                        isTrue = false                                                                  '为假则不处理替换了，因为当前替换的这个在下一个文件里可以和下下一个文件同相，并替换，那当前这个文件就不存在
                    end if 
                end if 

                s = url & "【|】" & localFileName & "【|】" & fileName & "【|】" & parentFileName & "【|】" & splxx(4) 
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
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) : tempFileName = fileName 
                parentFileName = splxx(3) 
                if fileName = findFileName then
                    fileName = toFileName 
                end if 
                s = url & "【|】" & localFileName & "【|】" & fileName & "【|】" & parentFileName & "【|】" & splxx(4) 
                if c <> "" then
                    c = c & vbCrLf 
                end if 
                c = c & s 
            end if 
        next 
        downUrlList = c 

    end function 

    '替换html/css里图片路径
    function handleFileImagesPath()
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, imageType, parentFileName, content, localFileName, char_Set 
        dim splFile, fileList,parentFileExt,fileExt
        fileList = createFilePath & vbCrLf & getDirCssList(debugWebDir)'当前html加+CSS列表
		if isHandleJsInHtml=true then
			fileList=fileList  & vbCrLf & getDirJSList(debugWebDir) '+JS列表(0171113) ,避免大js文件判断编码时太慢
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
                    if instr(s, "【|】") > 0 then
                        splxx = split(s, "【|】") 
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

    '搜索同内容文件名
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
    '获得同名上一级名称
    function fileNameToParentNameMD5(byval findUrl, byval findFName)
        dim splStr, splxx, s, url, fileName 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
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
    '文件名 找 上一级文件名
    function fileNameToParentName(parentFNameMD5)
        dim splStr, splxx, s, url, fileName 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
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
    '网址 找 文件名
    function urlToFileName(findUrl)
        dim splStr, splxx, s, c, url, fileName, tempFileName, localFileName 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
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


    '修改httpurl状态
    function editHttpUrlState(findUrl, stateIDValue)
        dim splStr, splxx, s, c, url, filePath, fileName, tempFileName, fileType, imageType, parentFileName, parentFilePath, localFileName, stateID 
        splStr = split(downUrlList, vbCrLf) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
                url = splxx(0)
				if ubound(splxx)>=4 then
					localFileName = splxx(1) 
					fileName = splxx(2) 
					parentFileName = splxx(3) 
					stateID = splxx(4) 
					if url = findUrl then
						stateID = stateIDValue 
					end if 
					s = url & "【|】" & localFileName & "【|】" & fileName & "【|】" & parentFileName & "【|】" & stateID 
	
					c = c & s & vbCrLf 
				end if
            end if 
        next 
        downUrlList = c 

        dim content 
        content = readFile(cacheConfigPath, "") 
        s = findUrl & "【|】" & stateIDValue 
        if instr(vbCrLf & content & vbCrLf, vbCrLf & s & vbCrLf) = false then
            content = content & s & vbCrLf 
            call WriteToFile(cacheConfigPath, content, "") 
        end if 
    end function 
    '从缓冲里获得状态码
    function getCacheState(findUrl)
        dim s, nLen, content 
        getCacheState = "" 
        content = vbCrLf & getftext(cacheConfigPath) & vbCrLf 
        s = vbCrLf & findUrl & "【|】" 
        nLen = instr(content, s) 
        if nLen > 0 then
            content = mid(content, nLen + len(s)) 
            content = mid(content, 1, instr(s, vbCrLf) + 2) 
            getCacheState = content 
        end if 
    end function 


    '替换源HTML内容，针对不同网站处理
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


    '创建HTML到网站
    function handleHtmlToWeb()
        dim imagesDir, cssDir, jsDir 
        imagesDir = createHtmlWebDir & "/images"                                        'images存放目录
        cssDir = createHtmlWebDir & "/css"                                              'css存放目录
        jsDir = createHtmlWebDir & "/js"                                                'js存放目录
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
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) 
                parentFileName = splxx(3) 
                stateID = splxx(4) 
                '只对正常的文件处理
                if stateID = "200"  or stateID="" then		'或者为空20171218
                    filePath = debugWebDir & "/" & fileName 
                    fileType = lcase(getFileAttr(fileName, "4")) 
                    if fileType = "js" then
                        toFilePath = jsDir & "/" & fileName 
                        call deleteFile(toFilePath)                                                     '删除css文件，因为它要重新处理
                    elseif instr("|css|woff|woff2|eot|ttf|", "|" & fileType & "|") > 0 then
                        toFilePath = cssDir & "/" & fileName 
                        call deleteFile(toFilePath)                                                     '删除css文件，因为它要重新处理
                    elseif instr("|htm|html|", "|" & fileType & "|") > 0 then
                        filePath = debugWebDir & "/index_" & fileName 
                        toFilePath = createHtmlWebDir & "/index_" & fileName 
                        call deleteFile(toFilePath)                                                     '删除html文件，因为它要重新处理
                    else
                        toFilePath = imagesDir & "/" & fileName 										'默认放到images文件夹里
					end if 
                    '文件存在则判断
                    if checkFile(filePath) = true and checkFile(toFilePath) = false then
                        call copyFile(filePath, toFilePath) 
                        if fileType = "css" then
                            content = readFile(toFilePath, "") 
                            content = pubCssObj.handleCssContent("css", content, "*", "background||background-image", "*", "isimg(*)trim(*)dir(*)../images/", "set") 
                            call WriteToFile(toFilePath, content, "") 
						
						'处理js文件时要开户 处理js里的html
                        elseif fileType = "html" or fileType = "htm" or (fileType = "js" and isHandleJsInHtml=true) then
                            call handleHtmlContentPath(filePath & ".txt", toFilePath, "") 
                        end if 
                    end if 
                end if 
            end if 
        next 
    end function 

    '创建HTML到模板
    function handleHtmlToTemplate()
        dim imagesDir, cssDir, jsDir 
        imagesDir = templateDir & "/images"                                             'images存放目录 /Templates2015/sharembweb/Images/
        cssDir = templateDir & "/css"                                                   'css存放目录
        jsDir = templateDir & "/js"                                                     'js存放目录
        call createDirFolder(imagesDir) 
        call createDirFolder(cssDir) 
        call createDirFolder(jsDir) 

        dim splStr, splxx, s, c, url, filePath, toFilePath, fileName, tempFileName, fileType, imageType, parentFileName, localFileName, tempS 
        dim stateID, content, websiteFilePath, websiteFileContent 
        splStr = split(downUrlList, vbCrLf)  
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
                url = splxx(0) 
                localFileName = splxx(1) 
                fileName = splxx(2) 
                parentFileName = splxx(3) 
                stateID = splxx(4) 
                '只对正常的文件处理
                if stateID = "200" or stateID="" then		'或者为空20171218
                    filePath = debugWebDir & "/" & fileName 
                    fileType = lcase(getFileAttr(fileName, "4")) 
                    if fileType = "js" then
                        toFilePath = jsDir & "/" & fileName 
                        call deleteFile(toFilePath)                                                     '删除css文件，因为它要重新处理
                    elseif instr("|css|woff|woff2|eot|ttf|", "|" & fileType & "|") > 0 then
                        toFilePath = cssDir & "/" & fileName 
                        call deleteFile(toFilePath)                                                     '删除css文件，因为它要重新处理
                    elseif instr("|htm|html|", "|" & fileType & "|") > 0 then
                        filePath = debugWebDir & "/index_" & fileName 
                        toFilePath = templateDir & "/index_Model" & fileName 
                        call deleteFile(toFilePath)                                                     '删除html文件，因为它要重新处理
                    else
						 toFilePath = imagesDir & "/" & fileName 
					end if 
                    '文件存在则判断
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
        '复制模板到指定文件夹里
        call deleteFolder(toTemplateDir) 
        call copyFolder(templateDir, toTemplateDir) 

        call copyFolder("/Data/WebData", toTemplateDir & "/WebData") 
        websiteFilePath = toTemplateDir & "/WebData/website.txt" 
        websiteFileContent = getftext(websiteFilePath) 
        s = getStrCut(websiteFileContent, "【webtemplate】", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "【webtemplate】" & toTemplateDir & vbCrLf) 
        s = getStrCut(websiteFileContent, "【webimages】", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "【webimages】" & toTemplateDir & "/Images" & vbCrLf) 
        s = getStrCut(websiteFileContent, "【webcss】", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "【webcss】" & toTemplateDir & "/Css" & vbCrLf) 
        s = getStrCut(websiteFileContent, "【webjs】", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "【webjs】" & toTemplateDir & "/Js" & vbCrLf) 

        s = getStrCut(websiteFileContent, "【webtitle】", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "【webtitle】" & templateName & "标题" & vbCrLf) 
        s = getStrCut(websiteFileContent, "【webkeywords】", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "【webkeywords】" & templateName & "关键词" & vbCrLf) 
        s = getStrCut(websiteFileContent, "【webdescription】", vbCrLf, 1) 
        websiteFileContent = replace(websiteFileContent, s, "【webdescription】" & templateName & "描述" & vbCrLf) 

        call createfile(websiteFilePath, websiteFileContent) 



    end function 

    '处理HTML文件路径
    function handleHtmlContentPath(filePath, toFilePath, sType)
        dim splStr, splxx, s, c, url, fileName, tempFileName, fileType, imageType, parentFileName, content, localFileName, char_Set 
        dim splFile, fileList, navC, replaceFileName 
        fileList = createFilePath & vbCrLf & getDirCssList(debugWebDir)                 '当前html加+CSS列表
        splFile = split(fileList, vbCrLf) 
        splStr = split(downUrlList, vbCrLf) 

        char_Set = checkCode(filePath) 
        content = readFile(filePath, char_Set) 
        for each s in splStr
            if instr(s, "【|】") > 0 then
                splxx = split(s, "【|】") 
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
        '单独处理模板
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
                '导航处理成模板
                navC = fObj.handleLabelContent("ul[**class=naV]", "*", "", "get+label") 
                call echo(typeName(content), typeName(navC)) 
                content = pubTemplateObj.batchHandleWebNavLayout(content, navC) 

                navC = fObj.handleLabelContent("div[**class=columnlist2]", "*", "", "get+label") 
                content = pubTemplateObj.batchHandleWebNewsLayout(content, navC) 

        end if
        call WriteToFile(toFilePath, content, char_Set) 
        handleHtmlContentPath = content 
    end function 
    '保存日志
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
        call createfile(debugWebDir & "/Z打印仿站日志.html", imitate_config(downC)) 
    end function 


    '**************************************************
    '获得错误信息网址列表 20161002
    function getErrorInfoUrlList(byval findFileName)
        dim content, splStr, splxx, s, url, fileName, downStatus, localSaveFileName 
        content = getftext(debugWebDir & "/下载网址列表.txt") 
        splStr = split(content, vbCrLf) 
        errInfoC = "" 
        for each s in splStr
            splxx = split(s, "【|】") 
            if uBound(splxx) >= 4 then
                url = splxx(0) 
                fileName = splxx(1) 
                localSaveFileName = splxx(2) 
                downStatus = splxx(4) 
                '这种情况是因为二次操作后留下的，那我们再去缓冲里找一下就可以了。OK
                if downStatus = "State状态码" then
                    downStatus = getCacheState(url) 
                end if 

                if findFileName = fileName and downStatus <> "200" then
                    errInfoC = errInfoC & url & "【>>】" & fileName & "【>>】" & localSaveFileName & "【>>】" & downStatus & vbCrLf '错误信息累加
                end if 
            end if 
        next 
        getErrorInfoUrlList = errInfoC 
    end function 
    '列表错误图片列表
    function getErrorImagesList()
        dim content, splStr, splxx, s, url, fileName, downStatus, goToUrl, fileNameList 
        content = getftext(debugWebDir & "/下载网址列表.txt") 
        splStr = split(content, vbCrLf) 
        fileNameList = "|" 
        for each s in splStr
            splxx = split(s, "【|】") 
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
 
	'处理替换源网页配置内容 20111109
	function handleReplaceSourcePageConfig(byval content) 
		dim splstr,s,splxx,sLeft,sRight,startStr,endStr,sStr
		splstr=split(replaceSourcePageConfig,"====================") 
		for each s in splstr
			if instr(s,"【|】")>0 and left(phptrim(s),2)<>"##" then
				splxx=split(s,"【|】")
				sLeft=phptrim(splxx(0))
				sRight=phptrim(splxx(1))
				if lcase(left(sLeft,10))=lcase("getStrCut(") then
					splxx=split(sLeft,"【,】")
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
				
				'call echo("替换","")
			end if
		next
		handleReplaceSourcePageConfig=content
	end function
	'删除处理后生成完网页内容 20171111
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
	'处理内容里的script js脚本 20171114 
	function handleContentInScriptHtml(byval content,byval httpurl, byval fileName)										
		dim c,splstr,s,replaceS,thisJsObj
		dim thisFormatObj : set thisFormatObj=new class_formatting			'格式化类
		thisFormatObj.handleFormatting(content) 
		c=thisFormatObj.handleLabelContent("script","*","","get+label")
		splstr=split(c,"$Array$")
		for each s in splstr
			if s <>"" and instr(s,"[##")=false then
					replaceS=s
					'排除Jquery文件，因为处理不了   要开启处理js里的html
					if instr(replaceS,".jQuery=")=false and isHandleJsInHtml=true then 
						set thisJsObj=new class_js							'js类
						thisJsObj.httpurl=httpurl		'设置js网址
						replaceS=  thisJsObj.handleJsContent(replaceS,"source|html")		'处理js里的内容，让html里图片完整
						'下载网页
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



