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

function getAspToVbNetContent_filePath(filePath,language, msgStr)
	getAspToVbNetContent_filePath= getAspToVbNetContent(getftext(filepath),language, msgStr)
end function

'获得ASP转PHP内容
function getAspToVbNetContent(content,language, msgStr)
    '(20150805不完善架构不好)
    dim  c 
    dim obj : set obj = new aspToLanguageClass
        '开启导入函数功能20160115
        if request("isOpenImportFunction") <> "" then
            obj.setIsOpenImportFunction = true 
        end if 
	
	
        '设置为函数列表
        c = "|abs|cLng|uBound|lBound|chr|replace|write|Session|Cookies|QueryString|Form|typeName|Redirect|formatNumber|cStr|cInt|asc|join|len|lCase|uCase|isNumeric|isNull|int|execute|ServerVariables|createObject|hex|ascW|time|isObject|isArray|right|left|lTrim|rTrim|case|timer|dateDiff|fix|now|dateAdd|isDate|year|month|day|hour|minute|second|array|mapPath|request|requestGet|requestPost|datePart|" 
        'php里用到的
        c = c & "|header|echo|aspTrim|PHPEcho|aspEcho|getYMDHMS|isEmpty|" 
        c = c & "|aspTrim|webPageControl|handleAction|XY_getLableValue|replaceValueParam|newGetArray|aspLTrim|aspRTrim|aspDate|aspTime|" 
        c = c & "handleMakeHtmlFileEditTime|XY_admin|XY_test|ManyTimesHandleTemplateAction|handleTemplateAction|XY_executeSQL|XY_NavDidExists|NewRParam|newReplaceValueParam|XY_PHP_GeneralInfoList|XY_PHP_GeneralList||XY_PHP_GetFieldValue|XY_PHP_SinglePage|aspMD5|session_set_cookie_params|getMsg1|myMD5|inputText|inputText2|showSelectcolumn|setCookie|columnList|columnList2|showColumnList|getEditorStr|handleInputHiddenTextArea|displayUploadDialog|handleInputCheckBox|inputCheckBox3|arrayToString|resetAccessData|XY_PHP_NavList|XY_PHP_DetailList|getColumnName|batchDeleteTempStr|createFileUTF|handleRemoteJsCssParam|handleIsWebSite|handleGetContentUrlList|replaceGlobleVariable|handleHttpUrlArray|MainNav|dispalyManage|addEditDisplay|saveAddEdit|del|sortHandle|displayDefaultLayout|loadWebConfig|adminIndex|handleRGV|makeAllHtml|getRsUrl|handleUpDownArticle|upArticle|downArticle|getColumnUrl|handleReplaceValueParam|" 
        c = c & "|displayLayout|saveRobots|saveSiteMap|isOpenTemplate|updateWebsiteStat|contentTranscoding|stat2016|displayTemplatesList|XY_DisplayWrap|XY_GetColumnUrl|XY_GetOnePageUrl|copyHtmlToWeb|makeHtmlWebToZip|getRandArticleId|handleWebUrl|getOnePageUrl|XY_PHP_GetColumnContent|XY_PHP_CommentList|displayEditor|specialStrReplace|unSpecialStrReplace|" 
        c = c & "|getBrType|templateFileList|delTemplateFile|addEditFile|folderSearch|getDirJsList|getDirCssList|getDateilTemplate|websiteDetail|getFieldConfigList|handlePower|handlePower|checkPower|getFormFieldName|handlePower|getHandleFieldList|getHandleDate|strtolower|strtoupper|runScanWebUrl|PHPStrLen|reaFile|scanCheckDomain|callFunction2|" 
        c = c & "|ord|intval|is_array|gettype|xmlZIP|run|callFunction_cai|callfile_setAccess|Request|" 
		
		
		'暂存函数列表 追加上面
		c=c & "getArrayAABB|"

        dim addC 
        addC = getFText("\Config\网站函数总杂205.txt") & getScanFunctionNameList(content) 
        addC = replace(replace(addC, ",", "|"), vbCrLf, "|") 
        obj.setFunctionNameList = obj.getWordList() & c & addC 
		
		
		
        '设置不显示函数名称列表 删除函数		
		c=""
		'.net删除函数
		if language="vb.net" then
			c=""
			obj.set_clearFunVarList=""
		end if
        obj.setNotShowFunctionList = c
		
		
		
		'设置留空函数名称列表 清空函数内容
        c = "IsMail|BytesToBstr|Agent|isNul|AsaiLinkAdd|AsaiLinkDel|Iswww|XY_RelatedInformationList|HandleDispalyArticleList|CheckAccessPass|" 
        c = c & "|CheckSql|CreateMdb|EditTable|GetFieldAlt|GetFieldList|GetDifferentTableFieldList|CompactDB|HandleSqlServer|GetDataTableList|CheckTable|CheckField|GetTableList_Conn|EditAccessPassWord|" 
        c = c & "|ExportToExcel|ImportToDatabase|getFieldValue|getConnFieldValue|setFieldValue|SetTableField|FileBinaryAddAccess|ShowAccessStream|DeleteAction|AddDateViewWebEffect|" 
        c = c & "|CopyDateToWebDate|CheckEval|phpExec|handleGetHttpPage|HandleContentCode|XY_PHP_CustomNavList_1|makeHtmlWebToXmlZip|" 	
		obj.setSpaceFunctionList = c
		
		
		'设置严谨【函数】列表
		c="|getn,int|getvoid,int|getCountPage,int|getPageNumb,int|getCountStep,int|getBL,int[]|val,int|array_Sort,int[]|"
		c=c & "getTimerSet,int|calculationTimer,int|phpPrint,void|phpRand,int|phpRnd,int|toNumber,double|scanUrl,int|"
		c=c & "checkNumberType,string|minMaxBetween,int|phpStrLen,int|strLen,int|handleHaveStr,bool|haveLowerCase,bool|haveUpperCase,bool|haveNumber,bool|"
		c=c & "haveChina,bool|handleHttpUrlArray,string[]|handleIsWebSite,bool|newShowOnOffImg,string|newReplaceValueParam,string|newRParam,string|"
		c=c & "checkContentRunStr,string|handleToArray,string[]|remoteArrayJingHao,string[]|getValueInArrayID,int|deleteRepeatArray,string[]|characterUpset,string|"
		c=c & "arrayToString,string|getArrayCount,int|str_Split,string[]|getSplitCount,int|strToTrueFalse,bool|strTrueFalseToInt,int|"
		c=c & "chkPost,bool|getLen,int|foundInArr,bool|getStrIntContentNumb,int|newGetStrCut,string||"
		obj.set_yjFunctionList=c
		
		'设置严谨【变量】列表   |startTime,System.DateTime|endTime,System.DateTime|dateTime,System.DateTime|
		c="|i,int|j,int|n,int|id,string|n1,int|n2,int|n3,int|splstr,string[]|splxx,string[]|getBL,string[]|arrNSplStr,int[]|"
		c=c & "makeHtmlFileToLCase,bool|HandleisCache,bool|openErrorLog,bool|moduleReplaceArray,string[]|glb_isonhtml,bool|double1,double|double2,double|double3,double|long1,long|long2,long|long3,long|"
		obj.set_yjVarList=c

        '设置默认类名称列表
        obj.setClassNameList = "myclass" 
        '设置全局变量列表
        c = "|@$_REQUEST|@$_POST|@$_GET|@$_COOKIE|@$_SESSION|" 
        c = c & "|WEBCOLUMNTYPE|WEB_VIEWURL|WEB_ADMINURL|EDITORTYPE|ROOT_PATH|" 
        obj.setGlobalDimList = c 

        '设置转换成语言
        obj.setToLanguage = language 

        '设置显示HTML部分
        obj.setIsDisplayHtml = true 
 
        content = deleteAspTempCode(content,"删除此行")       '删除里暂存代码
		
		if language="vb.net" then
	        content = deleteAspTempCode(content,"删除vb.net ")       '删除里暂存代码
		end if
		
		

        content = obj.aspToLanguage(content) 

        '测试用到
        msgStr = obj.getTempActionList() 
    set obj = nothing 
    content = phpTrim(content) 
		 
    getAspToVbNetContent = content 
end function 
'删除里暂存代码   
function deleteAspTempCode(content,lableName)
    dim startStr, endStr, i, s 
    startStr = "'【"& lableName &"start】" : endStr = "'【"& lableName &"end】" 
    for i = 1 to 9
        if inStr(content, startStr) = false then
            exit for 
        end if 
        s = getStrCut(content, startStr, endStr, 1) 
        content = replace(content, s, "") 
    next 
    deleteAspTempCode = content 
end function 
 

'ASP转Language 201504
class aspToLanguageClass
    dim toLanguage                                                                  '转换成什么语言
    dim toLanguageLinkCha                                                           '转换语言变量与变量之间链接字符
    dim toLanguageVarNameCha                                                        '转换语言变量前字符

    dim ifJudge(99)                                                                 'IF判断
    dim ifJudgeRow(99)                                                              'IF判断 是否为一行
    dim ifJudgeRowStr(99)                                                           'IF判断 为一行内容
    dim ifRsObjHandStr                                                              'IF判断Rs句柄内容
	dim isIfElse																	'是If的Else后台的
    dim forJudge                                                                    'For判断
    dim trueJudge                                                                   '真假判断
    dim forFirstVarName                                                             'For第一个变量名
    dim forEachName                                                                 'For的Each名称
    dim isForStep                                                                   '是否为For循环步数
    dim forStep                                                                     'For的Step步骤名称
    dim functionName                                                                '记录函数名称
	dim nFunctionNameShowCount														'函数名称出现总数
    dim isNotShowFunctionName                                                       '记录不显示函数名称
    dim isFunctionDimName                                                           '是否为函数定义名称
	dim isFunctionSelfVar															'函数自身是否为变量
					
					
    dim isSpaceFunctionName                                                         '记录留空函数名称
    dim isSelectCase                                                                '是Select Case判断
    dim isSelectMoreCheck                                                           '为select判断的多个判断如 case 1,2,""  这个在PHP不能用，因为php只能用一个
    dim isCase                                                                      'Case是否开启 只让第一次显示后台:为:
    dim nCaseCount                                                                  'Case总数累加

    dim isDimAraay(99)                                                              '是否为定义数组
    dim nDimAraayLeftKuoHao(99)                                                     '定义数组左括号数
    dim nDimAraayRightKuoHao(99)                                                    '定义数组右括号数
    dim dimAraayList                                                                '定义数组列表
    dim dimAraayStart                                                               '定义数组列表开始  与结束 判断
    dim nDimArrayQianTao                                                            '定义数据嵌套数
	dim isDimArray																	'是否为定义的数组
    dim functionNameList                                                            '函数名称列表
	dim functionVarList																'函数名称列表
    dim notShowFunctionList                                                         '不显示函数名称列表
    dim spaceFunctionList                                                           '留空函数名称列表
    dim globalDimList                                                               '全局变量列表
    dim noteCode                                                                    '注释代码
    dim isToSetCookies                                                              '为设置Cookies
    dim nMidDuanLuo                                                                 'mid分段数  为了处理最后无参数的情况
    dim isWhile                                                                     'while是否为真
    dim ismultidimensionalArray                                                     '是多维数组
    dim isOpenTable                                                                 '打开数据库表
    dim isLoop                                                                      '为循环字符


    dim nVariableSetValue(99)                                                       '变量设置值累加数


    dim upWord                                                                      '上一个变量
    dim upWord2, upWord3, upWord4, upWord5, upWord6                                 '上上一个变量
    dim upWordYuan, upWordYuan2, upWordYuan3, upWordYuan4, upWordYuan5, upWordYuan6 '上一个源变量与上上一个源变量
	dim upRow,upRow2,upRow3,upRow4,upRow5,upRow6									'上一行与上上行

    dim endCode                                                                     '当前行字符往后面代码

    dim beforeStr                                                                   '前一个字符
    dim afterStr                                                                    '后一个字符
    dim rowC                                                                        '每行代码
    dim tempRow                                                                     '每行完整代码
    dim nc                                                                          '为数字累加
    dim arrayVar(20)                                                                '数据变量
    dim arrayVarLeftKuoHao(99)                                                      '数据变量 左括号数
    dim arrayVarRightKuoHao(99)                                                     '数据变量 右括号数

    dim tempActionList                                                              '暂存动作列表
    dim nShowFunction                                                               '函数显示数
    dim functionReturn                                                              '函数返回，给return 后台=用
    dim isFunctionReturn                                                            '判断函数是否有返回
    dim systemFunctonList 
    dim downRow                                                                     '下一行代码
    dim downWord                                                                    '下一行单词
    dim isDisplayHtml                                                               '是否显示HTML代码
    dim tempWord                                                                    '暂存单词

    dim functionInDimList                                                           '函数里定义的Dim变量
    dim isSearverAddKuoHu                                                           'Server里有括弧
    dim isHeader                                                                    '是header函数
    dim dimRsStr                                                                    '定义RS内容 给 If Rs.Eof Then Exit For   用
	dim dimMaoHaoLianStr															'dim aa:aa="aa"  处理成一个，因为在.net里有重复
    dim isOpenRsStr                                                                 '是打开数据记录
    dim onDimRsStr                                                                  '开户定义Rs字符
    dim nRsStrCount                                                                 'Rs定义总数
    dim nThisIndex                                                                  '当前索引
    dim importFunctionList                                                          '需要导入系统函数列表（20160115
    dim isOpenImportFunction                                                        '是否打开导入函数
    dim isClass, className, classNameList, dimClassNameList                         '是否为类，类名称，类名称列表
    dim rowEndStr                                                                   '行最后字符
    dim isZD, zDName, zDNameList                                                    '是否字典，字典名称，字典名称列表
    dim pubRowS                                                                     '一行内容
    dim phpClearLable                                                               'php清除错误标签
	dim strYingHaoFu																	'内容引与符
	dim rsArrayValue(999)															'定义rs名称，内容，是否完成
	dim rsList																		'定义rs列表
	dim rsCompleteList																'定义rs完成列表
	dim rsInsideList																'定义rs在内容列表
	dim rsDeleteList																'定义rs删除列表 因为while是不需要有rs出现的
	dim dotNetC																		'.net不是函数内容
	dim yjFunctionList																'严谨定义函数列表
	dim yjVarList																	'严谨定义变量列表
	dim createTabFileSuffix 														'生成转换文件后缀
	dim splStrRow,nI,toC 																			'分行
	dim clearFunVarList																'清除不显示函数或变量列表
	dim isInstrReplaceFlash 		'替换instr后台的false为零
	dim isRsToString                  			'最后是否转为字符类型
	dim requestFormNoKH							'request的form不要()括弧
	dim requestUpWord							'request 上一个单词
	dim rsUpWord								'rs上一个单词
	dim sessionUpWord							'session上一个单词
	dim chrUpWord								'chr上一个单词
	dim rsDataArray(99,2)															'rs记录
	dim isRowASP								'是否为一行ASP代码

    '处理获得Asp代码
    function aspToLanguage(byVal content)
        dim sx, tempSx, s, ganJingS,  i, isASP, sYHCount, s1, s2, tempZc
        dim noHandleStr                                                                 '不处理字符与长度
        dim wc                                                                          '输入文本存储内容
        dim zc                                                                          '字母文件存储内容
        dim yesOK                                                                       '是否OK
        dim tempS, tempS2, lCaseS 
        dim splxx, tempI 
		dim splTemp1,splTemp2,tempS1,tempS3,tempC1,tempC2,tempC3,tempI1,tempI2,tempI3
		dotNetC=""																		'清空.net函数外内容
        isASP = false                                                                  '是ASP 默认为假
        sYHCount = 0                                                                    '双引号默认为0
        onDimRsStr = false                                                              '开户定义Rs字符
        nRsStrCount = 0                                                                 'Rs定义字符内容默认为0

        '不同语言用不同链接符，与变量前缀
        if toLanguage = "vb.net" then
            toLanguageLinkCha = "&" 
            phpClearLable = "" 							'除错符
			strYingHaoFu="""" 
			createTabFileSuffix=".vb" 
        end if 

        noHandleStr = "[#不处理#]" 

        splStrRow = split(content, vbCrLf)                                                 '分割行
        '循环分行
        for nI = 0 to uBound(splStrRow)
            s = splStrRow(nI) : pubRowS = s                                                    '为一行内容

            s = replace(replace(s, chr(10), ""), chr(13), "") '奇怪为什么 s里会有 chr(10)与chr(13) 呢？  所以就要把 chr(10)与chr(13) 删除
            ganJingS = pHPTrim(s)                                                           '干净s
            lCaseS = lCase(ganJingS)                                                        '小写干净的S内容
            tempS = s                                                                       '暂存s内容
            rowC = "" : tempRow = ""                                                        '清空每行ASP代码和暂存完整行代码
            noteCode = ""                                                                   '默认注释代码为空
            downRow = ""                                                                    '下一行代码
            downWord = ""                                                                   '下一行单词
            if(nI + 1) <= uBound(splStrRow) then
                downRow = splStrRow(nI + 1)                                                        '下一行内容
                downWord = getBeforeWord(downRow)                                               '下一行单词
            'call echo("【下一行内容与单词】" & downRow,downWord)
            end if 
            nc = ""                                                                         '数字为空
            nDimArrayQianTao = 0                                                            '定义数据嵌套数
            isRowASP=false												'是否为一行ASP代码，为假
			'call echo("s",replace(lCaseS,"<","&lt;"))
            '不处理打头为 on 这种的
            if left(pHPTrim(lCase(s)), 3) = "on " and 1=2 then
                s = ""  
			'隐藏代码
			elseif checkHideCode(s)=true and isASP=true then	
				s=""
            end if 

            isLoop = true                                                                   '循环字符为真
            '循环每个字符
            for i = 1 to len(s)
                sx = mid(s, i, 1) : tempSx = sx 
                beforeStr = right(replace(mid(s, 1, i - 1), " ", ""), 1)                        '上一个字符
                afterStr = left(replace(mid(s, i + 1), " ", ""), 1)                             '下一个字符
                endCode = mid(s, i + 1)                                                         '当前字符往后面代码 一行
				
				'call echo(s,i & "、" & sx)

                'Asp开始
                if sx = "<" and isASP = false and wc = "" then
                    if mid(s, i + 1, 1) = "%" then
                        isASP = true                                                                   'ASP为真
 
						sx=getLanguageStartEndLabel(s,"start")
                        upWord = "<" & "%" 

                        i = i + 1 
                        rowC = rowC & sx 
                    elseIf isDisplayHtml = true then
                        rowC = rowC & sx                                                                '给不是ASP代码使用
                    end if 

                'ASP结束
                elseIf sx = "%" and mid(s, i + 1, 1) = ">" and isASP = true and wc = "" then
                    isASP = false                   
					sx=getLanguageStartEndLabel(language,"end") 
                    i = i + 1 
                    rowC = rowC & sx 
                'ASP为真
                elseIf isASP = true then
				
				
                    '输入文本
                    if(sx = """" or wc <> "") then
                        '双引号累加
                        if sx = """" then sYHCount = sYHCount + 1 
                        '判断是否"在最后
                        if sYHCount mod 2 = 0 then
                            if mid(s, i + 1, 1) <> """" then
                                wc = wc & sx 
                                s1 = right(replace(mid(s, 1, i - len(wc)), " ", ""), 1) '必需放在这里，要不会出错

                                '处理输入内容
                                wc = right(wc, len(wc) - 1) 
                                wc = left(wc, len(wc) - 1) 
								
								
								 
                                if 1=2 then
									wc = handleLanguage(wc, "输入")                         'ASP变量转PHP变量
								else
									wc=""""& wc &""""
								end if
								
								if isSpaceFunctionName = "留空函数" or isNotShowFunctionName = "不显示函数" then
									wc = "" 
								end if 
								
                                if wc = "'[为字典不输出内容]'" and isZD = true then
                                    wc = "" 
                                end if 

                                rowC = rowC & wc                                        '函数代码累加
                                sYHCount = 0 : wc = ""                                  '清除
                            else
                                wc = wc & sx 
                            end if 
                        else
                            wc = wc & sx 
                        end if 

                    '+-*\=&   字符特殊处理
                    elseIf inStr(".&=,+-*/:()><[]", sx) > 0 and isASP = true and nc = "" then
                        if zc <> "" then
                            tempZc = zc 
                            zc = handleLanguage(zc, "变量")              'ASP变量转PHP变量 
							
							call recordUpWord(tempZc,zc)				'记录上个单词级上上单词
							
                            rowC = rowC & zc 
                            zc = ""                                      '清空字母内容
                        end if 
                        if sx = "=" then
                            rowC = rTrim(rowC)
                        end if 
						 
                        tempSx=sx
                        rowC = rowC & handleLanguage(sx, "") 
   						call recordUpWord(tempSx,sx)				'记录上个单词级上上单词
						 
                    '注释则退出
                    elseIf sx = "'" then
                        noteCode = "'" & mid(s, i + 1) 
                        if isSpaceFunctionName = "留空函数" or isNotShowFunctionName = "不显示函数" then
                            noteCode = "" 
                        end if 
                        exit for
                    '字母
                    elseIf checkABC(sx) = true or sx = "_" or zc <> "" then
                        zc = zc & sx 
                        yesOK = true 
                        s1 = mid(s & " ", i + 1, 1) 
                        s2 = mid(zc, 1, 1) 
                        if checkABC(s1) <> true and s1 <> "_" then
                            yesOK = false 
                        end if 
                        '允许变量后面是数字
                        if checkNumber(s1) = true and checkABC(s2) = true then
                            yesOK = true 
                        end if 
                        if yesOK = false then
                            'Rem注释
                            if lCase(zc) = "rem" then
                                sx = mid(s, i + 1) : i = i + len(sx) 
                                zc = zc & sx 
                            end if 

                            tempZc = zc 
                            zc = handleLanguage(zc, "变量")              'ASP变量转PHP变量 
   							call recordUpWord(tempZc,zc)				'记录上个单词级上上单词
							 
                            rowC = rowC & zc 
                            zc = ""                                      '清空字母内容
                        end if 
                    '为数字
                    elseIf checkNumber(sx) = true or nc <> "" then
                        nc = nc & sx 
                        if afterStr <> "." and checkNumber(afterStr) = false then
                            nc = handleLanguage(nc, "数字")              'ASP变量转PHP变量
                            rowC = rowC & nc 
                            nc = "" 
                        end if 
                    '其它类型的值如 \ ^
					else
						if isSpaceFunctionName = "留空函数" or isNotShowFunctionName = "不显示函数" then
							sx = "" 
						end if 
                        rowC = rowC & sx 

                    end if 
                    tempRow = tempRow & sx 

                elseIf isDisplayHtml = true then
                    rowC = rowC & sx 
                end if 

                if isLoop = false then exit for 
            next    
		  
            toC = toC &  rowC & noteCode        
            '对不显示函数 与留宿函数 处理  不要换行20151223
            if isNotShowFunctionName <> "不显示函数" and isSpaceFunctionName <> "留空函数" then
                if isASP = true then toC = toC & vbCrLf                                            'Css为真则
            end if 
        next 
     
        aspToLanguage = toC
		

    end function 
	
	'处理语言
	function handleLanguage(varName,sType) 
        dim s, action, tempVarName, temp1,nLen,tempS,tempS2,tempS3,tempS4,tempS5,tempS6,tempS7,tempStart,tempEnd,tempBR,tempLCase,tempWord
		
        tempVarName = varName                                                           '暂时变量名称
        s = lCase(varName)
		'判断函数自身做为变量在函数里出现次数累加
		if functionName=varName and functionName<>"" then 
			nFunctionNameShowCount=nFunctionNameShowCount+1
			'call echo( functionName,nFunctionNameShowCount)
		end if
		
		
		
		'处理不显示函数
        if isNotShowFunctionName = "不显示函数" then
            if inStr("|function|sub|", "|" & getArrayVar() & "|") = false then
                handleLanguage = "" 
                exit function 
            else
                if inStr("|function|sub|", "|" & lCase(varName) & "|") = false then
                    handleLanguage = "" 
                    exit function 
                end if 
            end if 
        elseIf isSpaceFunctionName = "留空函数" then
            if inStr("|function|sub|", "|" & getArrayVar() & "|") = false then
                handleLanguage = "" 
                exit function 
            '真假判断  为假
            elseIf trueJudge = false and s <> "function" then
                handleLanguage = "" 
                exit function 
            end if 
		end if
		
		if s="err" then
			if lcase(getBeforeNStr(endCode, "全部", 1)) = "then" then
                handleLanguage = "Err.Number<>0" 
                exit function 
			ElseIf lcase(getBeforeNStr(endCode, "全部", 1)) = "=" then
                handleLanguage = "Err.Number" 
                exit function 
			else
				varName="Err"
			end if
		'清除
		elseif s="set" or s="let" then
			handleLanguage = "" 
			exit function 
		elseif s="wend" then
			handleLanguage = "end while" 
			exit function 
		elseif s="ascb" then
			varName="asc" 
		elseif s="chrb" then
			varName="chr" 
		elseif s="lenb" then
			varName="len" 
		elseif s="midb" then
			varName="mid" 
		end if
		
		
        '判断是否为系统函数 真为追加变量
        if checkAspSystemFunction(s) = true then
            call AddArrayVar(s)                                                             '追加数组变量
            arrayVarLeftKuoHao(getArrayIndex()) = 0                                         '数据变量 左括号数
            arrayVarRightKuoHao(getArrayIndex()) = 0                                        '数据变量 右括号数
        end if 

        action = getArrayVar()                                                          '获得当前动作
        nThisIndex = getArrayIndex()                                                    '当前索引
		 
		'call echo("action",action)
        if action = "array" then
			if s="array" then 
				varName="" 
         	elseif s="(" then 
				varName="{" 
            elseIf s = ")" then 
	            varName = "}"  
                call DelArrayVar()                                                              '删除数组变量 
			end if

            call addMsg(action, tempVarName, sType, varName)                                '追加回显
			 
        elseIf action = "function" then
            if s = "function" then
                if upWord = "end" then
                    if isNotShowFunctionName = "不显示函数" then
                        varName = "" 
                    '可用可无 (不对有用20151119)
                    elseIf isFunctionReturn = true then					
						'函数最后
						if nFunctionNameShowCount=0 then 
							varName="return """""
						end if
                   end if
 					
					'call echo("functionName",functionName)
                    if isSpaceFunctionName = "留空函数" then
                        varName =  vbcrlf & "    " & functionName & "=""""" & vbCrLf & "end " & varName 
                    end if 

                    functionName = ""                                                               '清空函数名称
                    isNotShowFunctionName = ""                                                      '清除不显示函数
                    isSpaceFunctionName = ""                                                        '清空留空
                    call DelArrayVar()                                                              '删除数组变量
                    call DelArrayVar()                                                              '删除数组变量
                    functionInDimList = ""                                                          '清空函数里定义变量
					functionVarList="|"
                '退出 并删除一次数组
                elseIf upWord = "exit" then
                    call DelArrayVar()                                                              '删除数组变量
                else
                    varName = "function "
					 
                    trueJudge = true                                                                '判断为真
                    functionName = "" 
                    isFunctionDimName = "函数变量名称" 
                    isFunctionReturn = false 
					nFunctionNameShowCount=0														'函数名称出现总数
					isFunctionSelfVar=false															'函数自身是否为变量
					functionVarList="|"																'函数名称列表清空
					 
                    if inStr("|" & notShowFunctionList & "|", "|" & lCase(getBeforeWord(endCode)) & "|") > 0 then 
                        isNotShowFunctionName = "不显示函数" 
                        varName = "" 
                    elseIf inStr("|" & spaceFunctionList & "|", "|" & lCase(getBeforeWord(endCode)) & "|") > 0 then
                        isSpaceFunctionName = "留空函数"  
                    end if 
                end if 
			elseIf s = ")" and trueJudge = true then
				varName = ")" 
				trueJudge = false 
				isFunctionDimName = "关闭函数变量名称" 
			elseIf sType = "变量" then
                if functionName = "" then
                    functionName = varName 
                    functionNameList = functionNameList & varName & "|"                      '函数名称累加 
                    nShowFunction = 1 

                    functionInDimList = replace(getFunctionStrDim(endCode), ",", "|") & "|"         '函数里定义的变量 这里不要累加
     
                end if
			end if
		end if
		handleLanguage=varName
	end function     
    '追加数组变量
    function addArrayVar(varName)
        dim i 
        for i = 0 to uBound(arrayVar)
            if arrayVar(i) = "" then
                arrayVar(i) = varName 
                exit for 
            end if 
        next 
    end function 
    '获得数组变量  从最后往前  (应该是从前向后)
    function getArrayVar()
        dim i 
        for i = 0 to uBound(arrayVar)
            if arrayVar(i) <> "" then
                getArrayVar = arrayVar(i) 
            else
                exit for 
            end if 
        next 
    end function 
    '获得数组索引
    function getArrayIndex()
        dim i 
        getArrayIndex = 0 
        for i = 0 to uBound(arrayVar)
            if arrayVar(i) <> "" then
                getArrayIndex = i 
            else
                exit for 
            end if 
        next 
    end function 
    '删除数组变量 从最后往前  (应该是从前向后)
    function delArrayVar()
        dim i 
        for i = 0 to uBound(arrayVar)
            if arrayVar(i) = "" then
                if i > 0 then
                    arrayVar(i - 1) = "" 
                end if 
                exit for 
            end if 
        next 
    end function 
    '追加回显 call addMsg("if",tempVarName,sType,varName)
    sub addMsg(lableName, tempVarName, sType, varName)
        'getArrayIndex()替换成功 nThisIndex 显示精准
        tempActionList = tempActionList & nThisIndex & "、" & lableName & "(" & tempVarName & ")(" & sType & ")(" & varName & ")<br>" & vbCrLf 
    end sub  
    '获得单词列表，从文本里提取
    function getWordList()
        dim c, splStr, i, s, splxx 
        c = getFText("\VB工程\Config\网站函数总杂205.txt") 
        splStr = split(c, vbCrLf) 
        c = "|" 
        for each s in splStr
            if inStr(s, ",") > 0 then
                splxx = split(s, ",") 
                c = c & splxx(1) & "|" 
            end if 
        next 
        getWordList = c 
    end function 
    '检测是否为ASP系统函数  这个非常重要 对系统变量特殊处理
    function checkAspSystemFunction(byVal functionName)
        if inStr(systemFunctonList, "|" & lCase(functionName) & "|") > 0 then
            checkAspSystemFunction = true 
        else
            checkAspSystemFunction = false 
        end if 
    end function 
	
	'获得前一个字符 GetBeforeStr(EndCode)
    function getBeforeStr(byVal content)
        getBeforeStr = left(trimVbCrlf(content), 1) 
    end function 
    '获得后一个字符
    function getAfterStr(byVal content)
        getAfterStr = right(trimVbCrlf(content), 1) 
    end function 


    '获得前一个单词
    function getBeforeWord(byVal content)
        content = trimVbCrlf(content) 
        content = replace(replace(replace(replace(content, "(", " ( "), ",", "  , "), ")", " ) "), ".", " . ") 
        if inStr(content, " ") then
            content = mid(content, 1, inStr(content, " ") - 1) 
        end if 
        getBeforeWord = content 
    end function 
    '获得前前一个单词
    function getBefore2Word(byVal content)
        content = trimVbCrlf(content) 
        content = replace(replace(replace(replace(content, "(", " ( "), ",", "  , "), ")", " ) "), ".", " . ") 
        if inStr(content, " ") then
            content = mid(content, inStr(content, " ") + 1) 
        end if 
        if inStr(content, " ") then
            content = mid(content, 1, inStr(content, " ") - 1) 
        end if 
        getBefore2Word = content 
    end function 
    '获得后一个单词
    function getAfterWord(byVal content)
        content = trimVbCrlf(content) 
        if inStr(content, " ") then
            content = mid(content, inStr(content, " ") + 1) 
        end if 
        getAfterWord = content 
    end function 


    '初始化
    private sub class_Initialize()
        'systemFunctonList = "|class|if|for|while|select|iif|function|sub|dim|rs|response|request|session|chr||call|left|right|rnd|mid|cint|instr|instrrev|vbtab|vbcrlf|split|mod|server|array|" 
		systemFunctonList = "|array|function|sub|" 
		
        systemFunctonList = lCase(systemFunctonList) 
		
		yjFunctionList="|getN,int|getVoid,void|"										'严谨定义函数列表
		yjFunctionList=lcase(yjFunctionList)
		yjVarList="|i,int|n,int|"												'严谨定义变量列表
		yjVarList=lcase(yjVarList)
		
        isDisplayHtml = false                                                           '设置默认是否显示HTML内容
        isOpenImportFunction = false                                                    '设置默认是否打开导入函数
        isClass = false : className = "" : classNameList = "|" : dimClassNameList = "|" '类初始化
        isZD = false : zDName = "" : zDNameList = "|"                                   '字典初始化
        phpClearLable = "@"                                                             '20160624
		isDimArray=false																'是否为Dim定义的变量
		isInstrReplaceFlash=false		'替换instr后台的false为零
    end sub 


    '获得 暂存动作列表
    public property get getGlobalDimList()
        getGlobalDimList = globalDimList 
    end property 
    '设置函数名称列表
    public property let setGlobalDimList(str)
        globalDimList = "|" & str & "|" 
    end property 
    '获得 暂存动作列表
    public property get getTempActionList()
        getTempActionList = tempActionList 
    end property 
    '设置函数名称列表
    public property let setFunctionNameList(str)
        functionNameList = "|" & str & "|" 
    end property 
    '设置不显示函数名称列表
    public property let setNotShowFunctionList(str)
        notShowFunctionList = lCase("|" & str & "|") 
    end property 
    '设置留空函数名称列表
    public property let setSpaceFunctionList(str)
        spaceFunctionList = lCase("|" & str & "|") 
    end property 
    '设置是否显示HTML部分
    public property let setIsDisplayHtml(str)
        isDisplayHtml = str 
    end property 
    '设置是否导入函数功能
    public property let setIsOpenImportFunction(str)
        isOpenImportFunction = str 
    end property 
    '设置类名称列表
    public property let setClassNameList(str)
        classNameList = "|" & str & "|" 
    end property 
    '设置转换成什么语言
    public property let setToLanguage(str)
        toLanguage = str 
    end property 
    '设置严谨函数列表  如  getn 为int
    public property let set_yjFunctionList(str)
        yjFunctionList = lcase(str)				'转小写 
    end property 
	'设置严谨变量列表
    public property let set_yjVarList(str)
        yjVarList = lcase(str)			'转小写 
    end property  
	'设计清除不显示函数名称与变量名称列表
    public property let set_clearFunVarList(str)
        clearFunVarList = lcase(str)			'转小写 
    end property  
	 
	'检测是否要隐藏当前代码 20160729
	function checkHideCode(thisS)
		dim splstr,s,content,thisRow
		thisRow=lcase(pHPTrim(thisS))
		checkHideCode=false
		if instr(thisRow,"'【@是"& toLanguage &"屏蔽@】")>0 then
			checkHideCode=true
			exit function
		end if
		
		content=replace("|.netc|asp|php|as|jsp|","|"& toLanguage &"|","|")
		splstr=split(content,"|")
		for each s in splstr
			if s <>"" then
				if instr(thisRow,"'【@不是"& s &"屏蔽@】")>0 or instr(ucase(thisRow),"'【"& ucase(s) &"】")>0 then
					checkHideCode=true
					exit function
				end if
			end if
		next
		if instr(thisRow,"'【@是"& toLanguage &"显示@】")>0 then
			toC=toC & replace(thisS, "'【@是"& toLanguage &"显示@】","")
			checkHideCode=true
			exit function
		end if
		 
	end function
	
	'记录上一级单词  yuanWord是转换后的单词
	sub recordUpWord(byval word, byval yuanWord)
		upWordYuan6 = upWordYuan5                    '第6源单词 等于 源第5单词
		upWordYuan5 = upWordYuan4                    '第5源单词 等于 源第4单词
		upWordYuan4 = upWordYuan3                    '第4源单词 等于 源第3单词
		upWordYuan3 = upWordYuan2                    '第3源单词 等于 源第2单词
		upWordYuan2 = upWordYuan                     '第2源单词 等于 源第1单词
		upWordYuan = yuanWord                          '记录上一变量
		
		upWord6 = upWord5
		upWord5 = upWord4 
		upWord4 = upWord3 
		upWord3 = upWord2 
		upWord2 = upWord                             '记录上上一个变量
		upWord = lCase(word)                       '记录上一个变量名称   为小写
	end sub



	'获得语言开始与结果标签值  如PHP里 <?php
	function getLanguageStartEndLabel(language,sType)
		dim s
		if toLanguage = "vb.net" then
			s=""
		else
			if sType="start" then
				s="<" & "%"
			elseif sType="end" then
				s="%" &">"
			end if
		end if
		getLanguageStartEndLabel=s
	end function

end class
	 
 
%>   
