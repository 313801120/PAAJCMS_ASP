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

class imitateWeb
    '下载网址列表.txt   此文件不要删除，否则二次下载时会 重复下载图片等文件
	dim version																		'版本
	dim defaultIEContent															'默认IE显示内容
    dim templateName                                                                '模板名称
	dim templateDir																	'模板路径
	dim toTemplateDir																'复制模板到其它目录
    dim isWebToTemplateDir                                                          '文件是否放到模板文件夹里
    
	dim serverHttpUrl																'服务器网址
    dim httpurl 
    dim newWebDir                                                                   '新网站目录
	dim debugWebDir																	'调试网站目录
	dim createHtmlWebDir															'创建html网站目录
	dim cacheFolderPath																'缓冲文件夹路径
    dim customCharSet                                                               '自定义编码
    dim isGetHttpUrl                                                                'getHttpUrl方式获得内容
    dim isMakeWeb                                                                   '生成WEB文件夹
    dim isMakeTemplate                                                              '是否生成模板文件夹
    dim isPackWeb                                                                   '是否打包WEB文件夹 xml
    dim isPackWebZip                                                                '是否打包WEB文件夹 zip
    dim isPackTemplate                                                              '是否打包模板文件夹 xml
    dim isPackTemplateZip                                                           '是否打包模板文件夹 zip
    dim isFormatting                                                                '是否格式化HTML
    dim isUniformCoding                                                             '是否统一编码
    dim isAddWebTitleKeyword                                                        '追加网站标题描关键词
    dim isOnMsg                                                                     '是否开启回显信息
    dim pubAHrefList 
    dim isUTF8Down                                                                  '以utf8方式下载
    dim downUrlList                                                                 '下载的URL列表
    dim fileNameList 
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

    dim jsListC, cssListC, imgListC, swfListC, otherListC, urlListC, errInfoC, downC, nDownCount '记录下载回显信息


    '构造函数 初始化
    sub class_Initialize()
        call loadDefaultConfig() 
    end sub 


    '加载默认配置
    sub loadDefaultConfig()
		version = "beta version"																	'软件版本 v1.1
		serverHttpUrl="http://aa/server_FangZan.Asp"									'等改成sharembweb.com
        httpurl = "http://sharembweb.com/" 	
        newWebDir = "newweb/"                                                           '新网站目录
		debugWebDir="dubug/"															'调试网站目录
		createHtmlWebDir="web/"													'生成网站目录 
        customCharSet = "自动检测"                                                      '编码
        isGetHttpUrl = false                                                            'getHttpUrl方式获得内容
        isMakeWeb = true                                                                '生成WEB文件夹
        isMakeTemplate = false                                                          '是否生成模板文件夹
        isPackWeb = false                                                               '是否打包WEB文件夹 xml
        isPackWebZip = false                                                            '是否打包WEB文件夹 zip
        isWebToTemplateDir = true                                                       '文件是否放到模板文件夹里
        toTemplateDir = "/templates/"                                                   '模板复制到目录
        isPackTemplate = false                                                          '是否打包模板文件夹 xml
        isPackTemplateZip = false                                                       '是否打包模板文件夹 zip

        isFormatting = false                                                            '是否格式化HTML
        isUniformCoding = true                                                          '是否统一编码
        isAddWebTitleKeyword = true                                                     '追加网站标题描关键词
        isOnMsg = false                                                                 '是否开启回显信息
        templateName = getWebSiteName(httpurl)                                          '模板名称从网址中获得
        isUTF8Down = false                                                              'utf8方式下载为假
        downCountSize = 0                                                               '下载总大小
        downFileTypeList = "|*|"                                                        '下载文件类型列表
        isDownServer = true                                                            '是否可以从服务器端下载
        isDownCache = true                                                             '默认不启用下载缓冲
        downImageAddMyInfoType = "2"                                                    '下载图片加我自己的信息
		cacheFolderPath=""																'缓冲目录 为空则在当前目录下创建一个cache目录
		call createDirFolder(cacheFolderPath)
		defaultIEContent="[defaultIEContent]"											'默认IE显示内容
		isDeleteErrorImages=true														'删除错误图片(web目录与templates)
		isStop=false																	'默认停止为假
    end sub 


    '下载网站
    sub downweb()
        dim s,  websiteFilePath, websiteFileContent
        dim filePath, toFilePath, fileList, splList, fileName, splStr, startTime, splUrl 
        startTime = now()                                                               '开始时间


        if phptrim(httpurl) = "" then
            call eerr("停止", "网址为空") 
        end if 

		isStop=false			'停止为假

        if saveFolderName <> "" then
            newWebDir = newWebDir & saveFolderName & "/" 
        else
            newWebDir = newWebDir & getWebSiteCleanName(httpurl) & "/"                  '用域名 setfilename(getwebsite(httpurl))
        end if 
        call createfolder(newWebDir) 		
		 

        '创建调试目录
		debugWebDir=newWebDir & debugWebDir & "/"
        call createDirFolder(debugWebDir) 
        '创建缓冲目录 前提是开启缓冲 和 默认缓冲文件夹为空
		if cacheFolderPath="" and isDownCache=true then 
			cacheFolderPath=newWebDir & "cache/"
        	call createDirFolder(cacheFolderPath) 
		end if
		 
		'创建html网站目录
		if isMakeWeb=true then
			createHtmlWebDir=newWebDir & "/web/"
			call createDirFolder(createHtmlWebDir)  
 		end if
		'模板目录
		if isWebToTemplateDir=true then
			templateDir=newWebDir & "templates/" & templateName & "/"			'生成当前模板目录
			toTemplateDir=toTemplateDir & "/" & templateName & "/"				'复制到指定模板目录
        	call createDirFolder(templateDir)  
		end if 
		

        downUrlList = getftext(newWebDir & "/下载网址列表.txt") 
        fileNameList = getftext(newWebDir & "/名称列表.txt") 
        call echo("fileNameList", fileNameList) 



        '批量下载
        splUrl = split(httpurl, vbCrLf) 
        for each httpurl in splUrl
            httpurl = phptrim(httpurl) 
            if getwebsite(httpurl) <> "" then
                call imitateWeb()
				'停止
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

        '生成模板目录
        if isMakeTemplate = true then 
            call webBeautify(templateDir, true)                  '美化
            '把模板复制到/template文件夹下
            if isWebToTemplateDir = true then
                if checkFolder(toTemplateDir) = true then
                    call echored("模板文件夹存在（拷贝模板里不存在的文件）", toTemplateDir) 

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
                                    call echored("复制文件成功", toFilePath) 
                                end if 
                            end if 
                        next 
                    next 
                else
                    'call deleteFolder(toTemplateDir)        '因为上面已经有判断了
                    call copyFolder(templateDir, toTemplateDir) 
                    call copyFolder("/Data/WebData", toTemplateDir & "/WebData") 
                    websiteFilePath = toTemplateDir & "/WebData/website.txt" 
                    websiteFileContent = getftext(websiteFilePath) 
                    s = getStrCut(websiteFileContent, "【webtemplate】", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "【webtemplate】" & templateDir & templateName & vbCrLf) 
                    s = getStrCut(websiteFileContent, "【webimages】", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "【webimages】" & templateDir & templateName & "/Images" & vbCrLf) 
                    s = getStrCut(websiteFileContent, "【webcss】", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "【webcss】" & templateDir & templateName & "/Css" & vbCrLf) 
                    s = getStrCut(websiteFileContent, "【webjs】", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "【webjs】" & templateDir & templateName & "/Js" & vbCrLf) 

                    s = getStrCut(websiteFileContent, "【webtitle】", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "【webtitle】" & templateName & "标题" & vbCrLf) 
                    s = getStrCut(websiteFileContent, "【webkeywords】", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "【webkeywords】" & templateName & "关键词" & vbCrLf) 
                    s = getStrCut(websiteFileContent, "【webdescription】", vbCrLf, 1) 
                    websiteFileContent = replace(websiteFileContent, s, "【webdescription】" & templateName & "描述" & vbCrLf) 

                    call createfile(websiteFilePath, websiteFileContent) 
                    s = debugWebDir & toTemplateDir & "   ==>>  " & toTemplateDir 
                    call echoYellowB("放到模板文件夹里", s) 
                end if 
            end if 
        end if 
        '生成WEB目录
        if isMakeWeb = true then
            call webBeautify(createHtmlWebDir, false) 
        end if 
        s = vBRunTimer(startTime) & "，下载总大小（" & printSpaceValue(downCountSize) & "）" 
        call rw("<title>" & s & "</title>") 
        call echo("成功", "仿站完成" & vBRunTimer(startTime) & "，下载总大小（" & printSpaceValue(downCountSize) & "）") 
    end sub 

    '仿站
    sub imitateWeb()
        dim content, fileName, c 
        dim htmlFilePath, htmlBaseFilePath, createFilePath, htmlFileSize,char_Set,s
        if httpurl = "" then
            call echo("回显", "网址为空" & httpurl) 
            exit sub 
        end if 
        '创建文件夹
        call createfolder(debugWebDir) 

        call hr() 
        call echored("模仿网址", httpurl) 
		'下载html页面
        htmlFilePath = handleDown("", httpurl, getUrlToFileName(httpurl, ":html_*.html"), "") '源文件路径
        htmlBaseFilePath = debugWebDir & "/" & getUrlToFileName(httpurl, ":base_*.html")  'base文件
        createFilePath = debugWebDir & "/" & getUrlToFileName(httpurl, ":index_*.html")   '生成文件
		
		if getFSize(htmlFilePath)=0 then
			call echoRedB("文件内容为空",htmlFilePath)
			exit sub
		end if
		
        content = readFile(htmlFilePath, char_Set)
		char_Set=customCharSet
		'非常重要，以后看到这里，记住不要删除了，20161011
		if customCharSet="自动检测" then
			doevents
			char_Set=getFileCharset(htmlFilePath)
			content = readFile(htmlFilePath, char_Set)		'这里读是为了自定找编码
			call echo(htmlFilePath,char_Set)
			if char_Set="gb2312" then
				s=pubHtmlObj.handleHtmlLabel(content,"meta","content","**http-equiv=Content-Type","get","lcase(*)")
				if instr(s,"utf-8")>0 then  
					content=pubHtmlObj.handleHtmlLabel(content,"meta","content","**http-equiv=Content-Type","set","text/html; charset=gb2312")
					call echo("见鬼","检测编码与文件编码不一样1")
				end if
			elseif char_Set="utf-8" then
				s=pubHtmlObj.handleHtmlLabel(content,"meta","content","**http-equiv=Content-Type","get","lcase(*)")
				if instr(s,"gb2312")>0 then  
					content=pubHtmlObj.handleHtmlLabel(content,"meta","content","**http-equiv=Content-Type","set","text/html; charset=utf-8") 
					call echo("见鬼","检测编码与文件编码不一样2")
				end if
			end if
		end if
		
        '保存base文件
        if checkFile(htmlBaseFilePath) = false then
            c = regExp_Replace(content, "<head>", "<head>" & vbCrLf & "<base href=""" & getWebsite(httpurl) & """ />" & vbCrLf)
            call WriteToFile(htmlBaseFilePath, c, char_Set) 
        end if 
        htmlFileSize = getFSize(htmlFilePath) 
        call echored("html文件路径", htmlFilePath & "   大小：" & printSpaceValue(htmlFileSize) & "（" & htmlFileSize & "）") 
        content = replaceTemplateContent(content)                                       '替换模板里特殊内容



		'call rwend( pubHtmlObj.handleHtmlLabel(content,"img","src","","get","trim(*)isimg(*)") )
		'call rwend( pubHtmlObj.handleHtmlLabel(content,"*","style","style>background","get","") )
		'call rwend( pubHtmlObj.getHtmlLabelParam(content, "link", "href", "rel=shortcut icon, ||type=image/ico")  )
		'call rwend( pubCssObj.handleCssContent("html",content,"*","background||background-image","*","","images") )
		'call rwend( pubHtmlObj.getSwfSrcList(content) )  
		'图片列表
		content = handleDownList(httpurl, content, pubHtmlObj.handleHtmlLabel(content,"img","src","","get","trim(*)isimg(*)"), "img")       '图片<img
		content = handleDownList(httpurl, content, pubHtmlObj.handleHtmlLabel(content,"*","style","style>background","get",""), "style") '标签背景<div style
		content = handleDownList(httpurl, content, pubHtmlObj.getHtmlLabelParam(content, "link", "href", "rel=shortcut icon, ||type=image/ico"), "link_image") '获得Link 如*.ico
		content = handleDownList(httpurl, content,pubCssObj.handleCssContent("html",content,"*","background||background-image","*","","images"), "style_image") '如果<style>

        'Css列表
        content = handleDownList(httpurl, content, pubHtmlObj.getLinkHrefList(content), "css") 
        'JS列表
        content = handleDownList(httpurl, content, pubHtmlObj.getScriptSrcList(content), "js") 
        'SWF
        content = handleDownList(httpurl, content, pubHtmlObj.getSwfSrcList(content), "swf") 
        '生成处理后文件
        call WriteToFile(createFilePath, phptrim(content), char_Set)                    '保存模板文件(去两边空格)

        urlListC = urlListC & batchFullHttpUrl(httpurl, pubHtmlObj.getAHrefList(content)) & vbCrLf 'url网址累加
        'A href
        pubAHrefList = pubAHrefList & urlListC 

        call createfile(debugWebDir & "/下载网址列表.txt", downUrlList) 
        call createfile(debugWebDir & "/名称列表.txt", fileNameList) 
        call createfile(debugWebDir & "/图片下载列表.txt", downImagesList) 

        call getErrorInfoUrlList(httpurl) 
        nDownCount = nDownCount + 1 
        downC = downC & getImitate_print(nDownCount, htmlFilePath, jsListC, cssListC, imgListC, swfListC, otherListC, urlListC, errInfoC) & vbCrLf 
        jsListC = "" : cssListC = "" : imgListC = "" : swfListC = "" : otherListC = "" : urlListC = "" 
        call createfile(debugWebDir & "/Z打印仿站日志.html", imitate_config(downC)) 		'更新日志文件

    end sub 
    '获得错误信息网址列表 20161002
    function getErrorInfoUrlList(byval httpurl)
        dim content, splStr, splxx, s, url, fileName, downStatus, goToUrl 
        content = getftext(debugWebDir & "/下载网址回显.txt") 
        splStr = split(content, vbCrLf) 
		errInfoC=""
        for each s in splStr
            splxx = split(s, "【|】") 
            if uBound(splxx) >= 4 then
                goToUrl = splxx(1) 
                url = splxx(2) 
                fileName = splxx(3) 
                downStatus = splxx(4) 
                if splxx(0) = httpurl and downStatus <> "200" then 
                    errInfoC = errInfoC & goToUrl & "【>>】" & url & "【>>】" & fileName & "【>>】" & downStatus & vbCrLf '错误信息累加
                end if 
            end if 
        next 
		getErrorInfoUrlList=errInfoC
    end function 
	'列表错误图片列表
	function getErrorImagesList()
        dim content, splStr, splxx, s, url, fileName, downStatus, goToUrl ,fileNameList
        content = getftext(debugWebDir & "/下载网址回显.txt")  
        splStr = split(content, vbCrLf) 
		fileNameList="|"
        for each s in splStr
            splxx = split(s, "【|】") 
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
    '检测重名文件
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

                    '对CSS文件处理不好，因为CSS下载下来的文件，没有处理，而对比的CSS是处理过的，所以不行，但是在把文件放到WEB文件夹下可以做
                    if getFSize(newFilePath) = getFSize(filePath) and getFText(newFilePath) = getFText(filePath) then 
                        call deleteFile(filePath)                                                       '删除当前这个不需要的文件
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
    '检测CSS内容标签是否成对出现
    function checkCssLabelSuccess_temp(content)
        dim spl1, spl2 
        spl1 = split(content, "{") 
        spl2 = split(content, "}") 
        checkCssLabelSuccess_temp = uBound(spl1) - uBound(spl2) 

    end function 
    '处理下载列表
    function handleDownList(byval httpurl, byval content, urlList, sType)
        dim splStr, s, c, url, fileName, filePath, cssContent, tempC, tempSFileName, isEchoDown,char_Set
        splStr = split(urlList, vbCrLf) 
        for each s in splStr
			'停止时退出
			if isStop=true then
				exit function
			end if
            tempSFileName = getFileAttr(phptrim(s), "2") 
			'call echo(s,tempSFileName)
            '排除空与#号
            if phptrim(s) <> "" and getwebsite(httpurl)<>s  then		'instr("#\/",left(tempSFileName, 1))=false  不要这种
                url = fullHttpUrl(httpurl, s)
                'img
                if sType = "img" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
                    content = pubHtmlObj.handleHtmlLabel(content, "img", "src", "src=" & s, "set", fileName) 
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              '图片网址累加

                elseif sType = "style" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
                    content = pubHtmlObj.handleHtmlLabel(content, "*", "style", "style>background", "set", s & "[#check#]" & fileName) 
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              '图片网址累加

                elseif sType = "link_image" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
					call echo(s,fileName)
                    content = pubHtmlObj.handleHtmlLabel(content, "link", "href", "rel=shortcut icon,||type=image/ico", "set", s & "[#check#]" & fileName) 
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              '图片网址累加

                elseif sType = "style_image" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
                    content = pubCssObj.handleCssContent("html",content,"*","background||background-image",s,fileName,"set") 
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              '图片网址累加

 

                '替换css文件里内容背景
                elseif sType = "css_content" then
                    fileName = getUrlToFileName(url, ":*.jpg") 
                    call handleDown(httpurl, url, fileName, isEchoDown) 
                    content = pubCssObj.handleCssContent("css",content, "*", "background||background-image", s, fileName, "set")  
                    downImagesList = downImagesList & url & vbCrLf 
                    imgListC = imgListC & url & vbCrLf                                              '图片网址累加

                'swf
                elseif sType = "swf" then
                    fileName = getUrlToFileName(url, ":*.swf")  
                    call handleDown(httpurl, url, fileName, isEchoDown)   
					content=pubHtmlObj.handleHtmlLabel(content,"embed","src","type=application/x-shockwave-flash","set", s & "[#check#]trim(*)" & fileName)
                    content = pubHtmlObj.handleHtmlLabel(content, "param", "value", "name=movie", "set", s & "[#check#]trim(*)" & fileName)		'trim(*)放在check后，因为它要分割的
                    downImagesList = downImagesList & url & vbCrLf 
                    swfListC = swfListC & url & vbCrLf 

                'css
                elseif sType = "css" then
                    fileName = getUrlToFileName(url, ":*.css") 
                    filePath = handleDown(httpurl, url, fileName, isEchoDown) 
                    '文件第一次下次则处理
                    if isEchoDown = true then 
						char_Set=customCharSet 
						if customCharSet="自动检测" then 
							char_Set=getFileCharset(filePath)
						end if
			
                        cssContent = readFile(filePath, char_Set) : tempC = cssContent 
                        call echoB("处理CSS文件", filePath) 
                        cssContent = handleDownList(url, cssContent, pubCssObj.getCssContentUrlList(cssContent), "css_content") 
                        call WriteToFile(filePath, phptrim(cssContent), char_Set)                       '保存模板文件(去两边空格)
                        if checkCssLabelSuccess_temp(cssContent) <> 0 then
                            call echoRedB("CSS文件标签不是成对出现" & filePath, checkCssLabelSuccess_temp(cssContent)) 
                        end if 
                        '处理文件对比
                        call checkResetTypeFileName(debugWebDir, filePath, fileName) 
                    end if  
                    content = pubHtmlObj.handleHtmlLabel(content, "link", "href", "rel=stylesheet", "set", s & "[#check#]" & fileName)   
                    cssListC = cssListC & url & vbCrLf                                              'css网址累加
					
                'js
                elseif sType = "js" then
                    fileName = getUrlToFileName(url, ":*.js") 
                    call handleDown(httpurl, url, fileName, isEchoDown)  
                    content = pubHtmlObj.handleHtmlLabel(content, "script", "src", "", "set", s & "[#check#]" & fileName) 
					
                    jsListC = jsListC & url & vbCrLf                                                'js网址累加
                end if 
            'call echo(s,url)
            end if 
        next 
        handleDownList = content 
    end function 

    '处理下载
    function handleDown(goToUrl, byval url, fileName, isEchoDown)
        dim filePath, downStatus, echoS, repeatNameS, s, splxx, nLen 
        dim filePrefixName, fileSuffixName, fileType, tempFileName, tempFilePath, s1, s2 
        isEchoDown = false                                                              '默认为没下载  给判断文件是否下载了用

        s = url & "【|】" 
        '采用 从下载列表里获取
        nLen = instr(vbCrLf & downUrlList, vbCrLf & s) 
        if nLen > 0 then
            fileName = mid(vbCrLf & downUrlList, nLen + len(vbCrLf & s)) 
            if instr(fileName, vbCrLf) > 0 then
                fileName = mid(fileName & vbCrLf, 1, instr(fileName, vbCrLf) - 1) 
            end if 
        end if 
        '20180921
        if isUTF8Down = true then
            url = toUTFChar(url)                                                            '转utf链接网址
        end if 


        filePath = debugWebDir & "\" & fileName 

        '文件存在，此文件名已经被下载过了，并重命名了
        if checkFile(filePath) = true and instr("|" & fileNameList & "|", "|" & fileName & "|") > 0 and instr(vbCrLf & downUrlList & "【|】", vbCrLf & url & "【|】") = false then
            tempFileName = fileName                                                         '暂时重名文件名称 待下载后判断，是否一样，不一样则改名
            tempFilePath = filePath 
            fileName = "z_" & md5(url, 2) & "." & getFileAttr(fileName, "2") 
            filePath = debugWebDir & "\" & fileName 
        end if 
        '文件不存在，则下载
        if checkFile(filePath) = false then
            isEchoDown = true                                                               '为文件下载了
            downStatus = newDownUrl(url, filePath) 
            '处理图片文件类型
            if checkFile(filePath) = true then
                fileType = lcase(getImageType(filePath))                                        '图片真实类型
                if fileType <> "" then
                    filePrefixName = getFileAttr(filePath, "3") 
                    fileSuffixName = lcase(getFileAttr(filePath, "4")) 
                    if fileType <> fileSuffixName then
                        s1 = fileName 
                        s2 = filePath 
                        fileName = filePrefixName & "." & fileType 
                        filePath = debugWebDir & "\" & fileName 
                        '转换后文件存在则MD5处理下
                        if checkFile(filePath) = true then
                            fileName = "z_" & md5(url, 2) & "." & fileType 
                            filePath = debugWebDir & "\" & fileName 
                        end if 
                        call moveFile(s2, filePath) 
						call deleteFile(s2)			'如果文件类型已经存在，则要删除的 ，这个无所谓，因为两个网址是一样的
                        call echoBlue("替换文件类型", s1 & "=>>" & fileName) 
                    end if 
                end if 
                '判断是否追加我的信息
                if instr("|bmp|jpg|gif|png|", "|" & fileType & "|") > 0 and downImageAddMyInfoType <> "" then
                    if getFSize(filePath) < 30240 then                                              '小于30K
                        call decSaveBinary(filePath, readBinary(filePath, 0) & "|" & imageAddMyInfo(downImageAddMyInfoType), 0) 
                        call echo("追加我的信息", filePath) 
                    end if 
                end if 
            end if 
			
			
			
            '不同文件夹同文件处理
            if tempFilePath <> "" and getFSize(filePath) = getFSize(tempFilePath) and getFText(filePath) = getFText(tempFilePath) then
                call deleteFile(filePath)                                                       '两文件一样则删除当前这个文件
                fileName = tempFileName 
                filePath = tempFilePath 
            '检测重名文件  这里要排除CSS，因为这里判断CSS不对，等处理了CSS后再判断
            elseif checkResetTypeFileName(debugWebDir, filePath, fileName) = true then
                call echoRedB("替换文件名", fileName) 
				call echo(debugWebDir,filePath)
            else
                tempFileName = "" 
                tempFilePath = "" 
            end if 
			doevents

            downCountSize = downCountSize + getFSize(filePath)                              '下载大小累加

            '下载列表累加
            s = url & "【|】" & fileName & vbCrLf 
            if instr(vbCrLf & downUrlList & vbCrLf, vbCrLf & s) = false then
                downUrlList = downUrlList & url & "【|】" & fileName & vbCrLf                   '追加下面图片地址
                call createfile(debugWebDir & "/下载网址列表.txt", downUrlList)                   '时时保存，那时因为，当关闭时有记录
            end if 
            '文件名累加
            if instr("|" & fileNameList & "|", "|" & fileName & "|") = false then
                fileNameList = fileNameList & fileName & "|" 
            end if 

            '下载回显
            echoS = httpurl & "【|】" & goToUrl & "【|】" & url & "【|】" & fileName 
            'if downStatus <> 200 then
                echoS = echoS & "【|】" & downStatus 
            'end if 
            if tempFileName <> "" then
                repeatNameS = "【|】内容重复" 
            end if 
            call createAddFile(debugWebDir & "/下载网址回显.txt", echoS & repeatNameS) 
            call echo("回显【" & downStatus & " 】（" & printSpaceValue(getFSize(filePath)) & "）", "正在下载素材 " & url) 

        end if 
        handleDown = filePath 
    end function 

    '新下载网址   带IE缓冲
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


        '启动下载  要判断是否可以从服务器端下载
        if(checkFile(ieFilePath) = false and isDownCache = true) or isDownServer = true then
            filePath = handlePath(filePath) 
            newDownUrl = saveRemoteFile(httpurl, filePath) 


            call copyFile(filePath, ieFilePath)                                             '把文件复制到ie缓冲里
        end if 


        '不存在则退出
        if checkFile(ieFilePath) = false then
            'call eerr( md5(httpurl,2), httpurl)
            call createAddFile(debugWebDir & "/888.txt", httpurl & "【|】" & md5(httpurl, 2)) 
            newDownUrl = 888 
            exit function 
        end if 
        if isDownCache = true and isDownServer=false then
            call copyFile(ieFilePath, filePath) 
            newDownUrl = 200 
        end if 
    end function 

    '美化网站 生成js,css分文件夹  暂时留着，待后面再改吧，有点懒了
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
        call createDirFolder(webBeautifyDir)                                            '批量创建文件夹
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

                '为模板
                if isTemplate = true then
                    toIndexFilePath = webBeautifyDir & "/" & htmlFileNameS & "_Model.html" 
                else
                    toIndexFilePath = webBeautifyDir & "/" & htmlFileNameS & ".html" 
                end if 
				'自动检测之后就不要再判断html文件是否是当前编码，因为在之前已经算过了
				if customCharSet="自动检测" then 
					char_Set=getFileCharset(indexFilePath)
				end if
                indexFileContent = readFile(indexFilePath, char_Set)                         '读文件 
                '为模板
                if isTemplate = true then
                    'link里有个favicon.ico  位置会有问题，不管了20160302
                    indexFileContent = handleConentUrl("[$cfg_webcss$]/", indexFileContent, "|link|", "", "") 
                    indexFileContent = handleConentUrl("[$cfg_webimages$]/", indexFileContent, "|img|embed|param|meta|src|imgstyle|", "", "") 
                    indexFileContent = handleConentUrl("[$cfg_webjs$]/", indexFileContent, "|script|", "", "") 
                    indexFileContent = handleHtmlStyleCss(indexFileContent, "替换路径", "[$cfg_webimages$]/") '处理html里的style css
                else
                    'link里有个favicon.ico  位置会有问题，不管了20160302
                    indexFileContent = handleConentUrl("css/", indexFileContent, "|link|", "", "") 
                    indexFileContent = handleConentUrl("images/", indexFileContent, "|img|embed|param|meta|src|imgstyle", "", "") 
                    indexFileContent = handleConentUrl("js/", indexFileContent, "|script|", "", "") 
                    indexFileContent = handleHtmlStyleCss(indexFileContent, "替换路径", "images/")  '处理html里的style css

                end if 
                '为模板，则替换标题与关键词与描述
                if isTemplate = true then
                    indexFileContent = replaceWebTitle(indexFileContent, "{$Web_Title$}") 
                    indexFileContent = replaceWebMate(indexFileContent, "name", "keywords", " content=""{$Web_KeyWords$}"" ", isAddWebTitleKeyword) 
                    indexFileContent = replaceWebMate(indexFileContent, "name", "description", " content=""{$Web_Description$}"" ", isAddWebTitleKeyword) 
                    indexFileContent = replaceWebMate(indexFileContent, "http-equiv", "Content-Type", " content=""text/html; charset=gb2312"" ", isAddWebTitleKeyword) 
 
                end if 
                '格式化HTML为真
                if isFormatting = true then
                    indexFileContent = formatting(indexFileContent, "") 
                    indexFileContent = htmlFormatting(indexFileContent) 
                end if 
                call WriteToFile(toIndexFilePath, phptrim(indexFileContent), char_Set) 			'创建文件时，编码是可以自定义的，这个等以后再做了

            end if 
        next 
		'列表目录全部文件   判断把指定文件放到对应文件夹里
		content = getFileFolderList(debugWebDir, true, "全部", "名称", "", "", "") 
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
					'判断文件大小不为零0    20160612
					if getFSize(filePath) <> 0 and instr(deleteErrorImagesList, "|" & fileName & "|")=false then 
						call copyFile(filePath, toFilePath) 
					end if 
				end if 
				call newEcho(fileName,  "images/" & fileName) 
			end if 
		next 
        'css列表不为空
        if cssFileList <> "" then
            splxx = split(cssFileList, vbCrLf) 
            for each filePath in splxx
                if filePath <> "" then
                    content = readFile(filePath, "") 
                    '为模板
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
		
		content = getFileFolderList(webBeautifyDir, true, "全部", "", "全部文件夹", "", "") 
        c = replace(content, vbCrLf, "|") 
        'call rw(content)
        'call eerr(getThisUrlNoFileName(),getthisurl())
        c = c & "|||||" 
		'要用到这个文件夹，提前创建好  在vb.net里不会出现
		call createfolder("./htmlweb")		
        url = getThisUrlNoFileName() & "/myZIP.php?webFolderName=web_&zipDir=htmlweb/&isPrintEchoMsg=0" 
		
        'PHP版打包 网站
        if isPackWebZip = true or isPackTemplateZip = true then
            call echo(isPackWebZip, isPackTemplateZip) 
            if isPackWebZip = true then
                s = xMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(debugWebDir)) & "\") 
            else
                'PHP版打包 模板网站
                s = xMLPost(url, "content=" & escape(c) & "&replaceZipFilePath=" & escape(handlePath(isWebToTemplateDir)) & "\") 
            end if 
            findStr = strCut(s, "downfile.asp?downfile=", """", 2) 

            '用session方式传下载值
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
                'call echo("1111111111111", "注意")
                call echo(handlePath(debugWebDir), handlePath(debugWebDir & rootDir)) 
            set objXmlZIP = nothing 
            doevents 
            xmlSize = getFSize(xmlFileName) 
            xmlSize = printSpaceValue(xmlSize) 
            call echo("下载xml打包文件", "<a href=?act=download&downfile=" & xorEnc(xmlFileName, 31380) & " title='点击下载'>点击下载" & xmlFileName & "(" & xmlSize & ")</a>") 
        end if 
		
		'NOVBNet end
    end sub 

    '替换网站标题
    function replaceWebTitle(content, webTitle)
        dim tempContent, startStr, endStr, nLen, bodyStart, bodyEnd 
        tempContent = lCase(content) 
        startStr = "<title>" 
        endStr = "</title>" 
        nLen = inStr(tempContent, startStr) 
        if nLen > 0 then
            bodyStart = mid(content, 1, nLen - 1)                                           '开始部分
            tempContent = mid(tempContent, nLen) 
            content = mid(content, nLen) 
        else
            replaceWebTitle = content 
            exit function 
        end if 
        nLen = inStr(tempContent, endStr) 
        if nLen > 0 then
            bodyEnd = mid(content, nLen + len(endStr))                                      '结束部分
        end if 
        replaceWebTitle = bodyStart & startStr & webTitle & endStr & bodyEnd 
    end function 
    '替换网站关键词  描述也可以替换
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
            bodyStart = mid(content, 1, nLen - 1)                                           '开始部分
            tempContent = mid(tempContent, nLen) 
            content = mid(content, nLen) 
        end if 
        nLen = inStr(tempContent, endStr) 
        if nLen > 0 then
            bodyEnd = mid(content, nLen + len(endStr))                                      '结束部分
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



    '操作echo
    function newEcho(title, str)
        if isOnMsg = true then
            call echo(title, str) 
        end if 
    end function 
end class 



%> 

