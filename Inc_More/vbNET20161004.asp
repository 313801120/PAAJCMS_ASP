<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-03-16
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%

function getAspToVbNetContent_filePath(filePath,language, msgStr)
	getAspToVbNetContent_filePath= getAspToVbNetContent(getftext(filepath),language, msgStr)
end function

'���ASPתPHP����
function getAspToVbNetContent(content,language, msgStr)
    '(20150805�����Ƽܹ�����)
    dim  c 
    dim obj : set obj = new aspToLanguageClass
        '�������뺯������20160115
        if request("isOpenImportFunction") <> "" then
            obj.setIsOpenImportFunction = true 
        end if 
	
	
        '����Ϊ�����б�
        c = "|abs|cLng|uBound|lBound|chr|replace|write|Session|Cookies|QueryString|Form|typeName|Redirect|formatNumber|cStr|cInt|asc|join|len|lCase|uCase|isNumeric|isNull|int|execute|ServerVariables|createObject|hex|ascW|time|isObject|isArray|right|left|lTrim|rTrim|case|timer|dateDiff|fix|now|dateAdd|isDate|year|month|day|hour|minute|second|array|mapPath|request|requestGet|requestPost|datePart|" 
        'php���õ���
        c = c & "|header|echo|aspTrim|PHPEcho|aspEcho|getYMDHMS|isEmpty|" 
        c = c & "|aspTrim|webPageControl|handleAction|XY_getLableValue|replaceValueParam|newGetArray|aspLTrim|aspRTrim|aspDate|aspTime|" 
        c = c & "handleMakeHtmlFileEditTime|XY_admin|XY_test|ManyTimesHandleTemplateAction|handleTemplateAction|XY_executeSQL|XY_NavDidExists|NewRParam|newReplaceValueParam|XY_PHP_GeneralInfoList|XY_PHP_GeneralList||XY_PHP_GetFieldValue|XY_PHP_SinglePage|aspMD5|session_set_cookie_params|getMsg1|myMD5|inputText|inputText2|showSelectcolumn|setCookie|columnList|columnList2|showColumnList|getEditorStr|handleInputHiddenTextArea|displayUploadDialog|handleInputCheckBox|inputCheckBox3|arrayToString|resetAccessData|XY_PHP_NavList|XY_PHP_DetailList|getColumnName|batchDeleteTempStr|createFileUTF|handleRemoteJsCssParam|handleIsWebSite|handleGetContentUrlList|replaceGlobleVariable|handleHttpUrlArray|MainNav|dispalyManage|addEditDisplay|saveAddEdit|del|sortHandle|displayDefaultLayout|loadWebConfig|adminIndex|handleRGV|makeAllHtml|getRsUrl|handleUpDownArticle|upArticle|downArticle|getColumnUrl|handleReplaceValueParam|" 
        c = c & "|displayLayout|saveRobots|saveSiteMap|isOpenTemplate|updateWebsiteStat|contentTranscoding|stat2016|displayTemplatesList|XY_DisplayWrap|XY_GetColumnUrl|XY_GetOnePageUrl|copyHtmlToWeb|makeHtmlWebToZip|getRandArticleId|handleWebUrl|getOnePageUrl|XY_PHP_GetColumnContent|XY_PHP_CommentList|displayEditor|specialStrReplace|unSpecialStrReplace|" 
        c = c & "|getBrType|templateFileList|delTemplateFile|addEditFile|folderSearch|getDirJsList|getDirCssList|getDateilTemplate|websiteDetail|getFieldConfigList|handlePower|handlePower|checkPower|getFormFieldName|handlePower|getHandleFieldList|getHandleDate|strtolower|strtoupper|runScanWebUrl|PHPStrLen|reaFile|scanCheckDomain|callFunction2|" 
        c = c & "|ord|intval|is_array|gettype|xmlZIP|run|callFunction_cai|callfile_setAccess|Request|" 
		
		
		'�ݴ溯���б� ׷������
		c=c & "getArrayAABB|"

        dim addC 
        addC = getFText("\Config\��վ��������205.txt") & getScanFunctionNameList(content) 
        addC = replace(replace(addC, ",", "|"), vbCrLf, "|") 
        obj.setFunctionNameList = obj.getWordList() & c & addC 
		
		
		
        '���ò���ʾ���������б� ɾ������		
		c=""
		'.netɾ������
		if language="vb.net" then
			c=""
			obj.set_clearFunVarList=""
		end if
        obj.setNotShowFunctionList = c
		
		
		
		'�������պ��������б� ��պ�������
        c = "IsMail|BytesToBstr|Agent|isNul|AsaiLinkAdd|AsaiLinkDel|Iswww|XY_RelatedInformationList|HandleDispalyArticleList|CheckAccessPass|" 
        c = c & "|CheckSql|CreateMdb|EditTable|GetFieldAlt|GetFieldList|GetDifferentTableFieldList|CompactDB|HandleSqlServer|GetDataTableList|CheckTable|CheckField|GetTableList_Conn|EditAccessPassWord|" 
        c = c & "|ExportToExcel|ImportToDatabase|getFieldValue|getConnFieldValue|setFieldValue|SetTableField|FileBinaryAddAccess|ShowAccessStream|DeleteAction|AddDateViewWebEffect|" 
        c = c & "|CopyDateToWebDate|CheckEval|phpExec|handleGetHttpPage|HandleContentCode|XY_PHP_CustomNavList_1|makeHtmlWebToXmlZip|" 	
		obj.setSpaceFunctionList = c
		
		
		'�����Ͻ����������б�
		c="|getn,int|getvoid,int|getCountPage,int|getPageNumb,int|getCountStep,int|getBL,int[]|val,int|array_Sort,int[]|"
		c=c & "getTimerSet,int|calculationTimer,int|phpPrint,void|phpRand,int|phpRnd,int|toNumber,double|scanUrl,int|"
		c=c & "checkNumberType,string|minMaxBetween,int|phpStrLen,int|strLen,int|handleHaveStr,bool|haveLowerCase,bool|haveUpperCase,bool|haveNumber,bool|"
		c=c & "haveChina,bool|handleHttpUrlArray,string[]|handleIsWebSite,bool|newShowOnOffImg,string|newReplaceValueParam,string|newRParam,string|"
		c=c & "checkContentRunStr,string|handleToArray,string[]|remoteArrayJingHao,string[]|getValueInArrayID,int|deleteRepeatArray,string[]|characterUpset,string|"
		c=c & "arrayToString,string|getArrayCount,int|str_Split,string[]|getSplitCount,int|strToTrueFalse,bool|strTrueFalseToInt,int|"
		c=c & "chkPost,bool|getLen,int|foundInArr,bool|getStrIntContentNumb,int|newGetStrCut,string||"
		obj.set_yjFunctionList=c
		
		'�����Ͻ����������б�   |startTime,System.DateTime|endTime,System.DateTime|dateTime,System.DateTime|
		c="|i,int|j,int|n,int|id,string|n1,int|n2,int|n3,int|splstr,string[]|splxx,string[]|getBL,string[]|arrNSplStr,int[]|"
		c=c & "makeHtmlFileToLCase,bool|HandleisCache,bool|openErrorLog,bool|moduleReplaceArray,string[]|glb_isonhtml,bool|double1,double|double2,double|double3,double|long1,long|long2,long|long3,long|"
		obj.set_yjVarList=c

        '����Ĭ���������б�
        obj.setClassNameList = "myclass" 
        '����ȫ�ֱ����б�
        c = "|@$_REQUEST|@$_POST|@$_GET|@$_COOKIE|@$_SESSION|" 
        c = c & "|WEBCOLUMNTYPE|WEB_VIEWURL|WEB_ADMINURL|EDITORTYPE|ROOT_PATH|" 
        obj.setGlobalDimList = c 

        '����ת��������
        obj.setToLanguage = language 

        '������ʾHTML����
        obj.setIsDisplayHtml = true 
 
        content = deleteAspTempCode(content,"ɾ������")       'ɾ�����ݴ����
		
		if language="vb.net" then
	        content = deleteAspTempCode(content,"ɾ��vb.net ")       'ɾ�����ݴ����
		end if
		
		

        content = obj.aspToLanguage(content) 

        '�����õ�
        msgStr = obj.getTempActionList() 
    set obj = nothing 
    content = phpTrim(content) 
		 
    getAspToVbNetContent = content 
end function 
'ɾ�����ݴ����   
function deleteAspTempCode(content,lableName)
    dim startStr, endStr, i, s 
    startStr = "'��"& lableName &"start��" : endStr = "'��"& lableName &"end��" 
    for i = 1 to 9
        if inStr(content, startStr) = false then
            exit for 
        end if 
        s = getStrCut(content, startStr, endStr, 1) 
        content = replace(content, s, "") 
    next 
    deleteAspTempCode = content 
end function 
 

'ASPתLanguage 201504
class aspToLanguageClass
    dim toLanguage                                                                  'ת����ʲô����
    dim toLanguageLinkCha                                                           'ת�����Ա��������֮�������ַ�
    dim toLanguageVarNameCha                                                        'ת�����Ա���ǰ�ַ�

    dim ifJudge(99)                                                                 'IF�ж�
    dim ifJudgeRow(99)                                                              'IF�ж� �Ƿ�Ϊһ��
    dim ifJudgeRowStr(99)                                                           'IF�ж� Ϊһ������
    dim ifRsObjHandStr                                                              'IF�ж�Rs�������
	dim isIfElse																	'��If��Else��̨��
    dim forJudge                                                                    'For�ж�
    dim trueJudge                                                                   '����ж�
    dim forFirstVarName                                                             'For��һ��������
    dim forEachName                                                                 'For��Each����
    dim isForStep                                                                   '�Ƿ�ΪForѭ������
    dim forStep                                                                     'For��Step��������
    dim functionName                                                                '��¼��������
	dim nFunctionNameShowCount														'�������Ƴ�������
    dim isNotShowFunctionName                                                       '��¼����ʾ��������
    dim isFunctionDimName                                                           '�Ƿ�Ϊ������������
	dim isFunctionSelfVar															'���������Ƿ�Ϊ����
					
					
    dim isSpaceFunctionName                                                         '��¼���պ�������
    dim isSelectCase                                                                '��Select Case�ж�
    dim isSelectMoreCheck                                                           'Ϊselect�жϵĶ���ж��� case 1,2,""  �����PHP�����ã���Ϊphpֻ����һ��
    dim isCase                                                                      'Case�Ƿ��� ֻ�õ�һ����ʾ��̨:Ϊ:
    dim nCaseCount                                                                  'Case�����ۼ�

    dim isDimAraay(99)                                                              '�Ƿ�Ϊ��������
    dim nDimAraayLeftKuoHao(99)                                                     '����������������
    dim nDimAraayRightKuoHao(99)                                                    '����������������
    dim dimAraayList                                                                '���������б�
    dim dimAraayStart                                                               '���������б�ʼ  ����� �ж�
    dim nDimArrayQianTao                                                            '��������Ƕ����
	dim isDimArray																	'�Ƿ�Ϊ���������
    dim functionNameList                                                            '���������б�
	dim functionVarList																'���������б�
    dim notShowFunctionList                                                         '����ʾ���������б�
    dim spaceFunctionList                                                           '���պ��������б�
    dim globalDimList                                                               'ȫ�ֱ����б�
    dim noteCode                                                                    'ע�ʹ���
    dim isToSetCookies                                                              'Ϊ����Cookies
    dim nMidDuanLuo                                                                 'mid�ֶ���  Ϊ�˴�������޲��������
    dim isWhile                                                                     'while�Ƿ�Ϊ��
    dim ismultidimensionalArray                                                     '�Ƕ�ά����
    dim isOpenTable                                                                 '�����ݿ��
    dim isLoop                                                                      'Ϊѭ���ַ�


    dim nVariableSetValue(99)                                                       '��������ֵ�ۼ���


    dim upWord                                                                      '��һ������
    dim upWord2, upWord3, upWord4, upWord5, upWord6                                 '����һ������
    dim upWordYuan, upWordYuan2, upWordYuan3, upWordYuan4, upWordYuan5, upWordYuan6 '��һ��Դ����������һ��Դ����
	dim upRow,upRow2,upRow3,upRow4,upRow5,upRow6									'��һ����������

    dim endCode                                                                     '��ǰ���ַ����������

    dim beforeStr                                                                   'ǰһ���ַ�
    dim afterStr                                                                    '��һ���ַ�
    dim rowC                                                                        'ÿ�д���
    dim tempRow                                                                     'ÿ����������
    dim nc                                                                          'Ϊ�����ۼ�
    dim arrayVar(20)                                                                '���ݱ���
    dim arrayVarLeftKuoHao(99)                                                      '���ݱ��� ��������
    dim arrayVarRightKuoHao(99)                                                     '���ݱ��� ��������

    dim tempActionList                                                              '�ݴ涯���б�
    dim nShowFunction                                                               '������ʾ��
    dim functionReturn                                                              '�������أ���return ��̨=��
    dim isFunctionReturn                                                            '�жϺ����Ƿ��з���
    dim systemFunctonList 
    dim downRow                                                                     '��һ�д���
    dim downWord                                                                    '��һ�е���
    dim isDisplayHtml                                                               '�Ƿ���ʾHTML����
    dim tempWord                                                                    '�ݴ浥��

    dim functionInDimList                                                           '�����ﶨ���Dim����
    dim isSearverAddKuoHu                                                           'Server��������
    dim isHeader                                                                    '��header����
    dim dimRsStr                                                                    '����RS���� �� If Rs.Eof Then Exit For   ��
	dim dimMaoHaoLianStr															'dim aa:aa="aa"  �����һ������Ϊ��.net�����ظ�
    dim isOpenRsStr                                                                 '�Ǵ����ݼ�¼
    dim onDimRsStr                                                                  '��������Rs�ַ�
    dim nRsStrCount                                                                 'Rs��������
    dim nThisIndex                                                                  '��ǰ����
    dim importFunctionList                                                          '��Ҫ����ϵͳ�����б�20160115
    dim isOpenImportFunction                                                        '�Ƿ�򿪵��뺯��
    dim isClass, className, classNameList, dimClassNameList                         '�Ƿ�Ϊ�࣬�����ƣ��������б�
    dim rowEndStr                                                                   '������ַ�
    dim isZD, zDName, zDNameList                                                    '�Ƿ��ֵ䣬�ֵ����ƣ��ֵ������б�
    dim pubRowS                                                                     'һ������
    dim phpClearLable                                                               'php��������ǩ
	dim strYingHaoFu																	'���������
	dim rsArrayValue(999)															'����rs���ƣ����ݣ��Ƿ����
	dim rsList																		'����rs�б�
	dim rsCompleteList																'����rs����б�
	dim rsInsideList																'����rs�������б�
	dim rsDeleteList																'����rsɾ���б� ��Ϊwhile�ǲ���Ҫ��rs���ֵ�
	dim dotNetC																		'.net���Ǻ�������
	dim yjFunctionList																'�Ͻ����庯���б�
	dim yjVarList																	'�Ͻ���������б�
	dim createTabFileSuffix 														'����ת���ļ���׺
	dim splStrRow,nI,toC 																			'����
	dim clearFunVarList																'�������ʾ����������б�
	dim isInstrReplaceFlash 		'�滻instr��̨��falseΪ��
	dim isRsToString                  			'����Ƿ�תΪ�ַ�����
	dim requestFormNoKH							'request��form��Ҫ()����
	dim requestUpWord							'request ��һ������
	dim rsUpWord								'rs��һ������
	dim sessionUpWord							'session��һ������
	dim chrUpWord								'chr��һ������
	dim rsDataArray(99,2)															'rs��¼
	dim isRowASP								'�Ƿ�Ϊһ��ASP����

    '������Asp����
    function aspToLanguage(byVal content)
        dim sx, tempSx, s, ganJingS,  i, isASP, sYHCount, s1, s2, tempZc
        dim noHandleStr                                                                 '�������ַ��볤��
        dim wc                                                                          '�����ı��洢����
        dim zc                                                                          '��ĸ�ļ��洢����
        dim yesOK                                                                       '�Ƿ�OK
        dim tempS, tempS2, lCaseS 
        dim splxx, tempI 
		dim splTemp1,splTemp2,tempS1,tempS3,tempC1,tempC2,tempC3,tempI1,tempI2,tempI3
		dotNetC=""																		'���.net����������
        isASP = false                                                                  '��ASP Ĭ��Ϊ��
        sYHCount = 0                                                                    '˫����Ĭ��Ϊ0
        onDimRsStr = false                                                              '��������Rs�ַ�
        nRsStrCount = 0                                                                 'Rs�����ַ�����Ĭ��Ϊ0

        '��ͬ�����ò�ͬ���ӷ��������ǰ׺
        if toLanguage = "vb.net" then
            toLanguageLinkCha = "&" 
            phpClearLable = "" 							'�����
			strYingHaoFu="""" 
			createTabFileSuffix=".vb" 
        end if 

        noHandleStr = "[#������#]" 

        splStrRow = split(content, vbCrLf)                                                 '�ָ���
        'ѭ������
        for nI = 0 to uBound(splStrRow)
            s = splStrRow(nI) : pubRowS = s                                                    'Ϊһ������

            s = replace(replace(s, chr(10), ""), chr(13), "") '���Ϊʲô s����� chr(10)��chr(13) �أ�  ���Ծ�Ҫ�� chr(10)��chr(13) ɾ��
            ganJingS = pHPTrim(s)                                                           '�ɾ�s
            lCaseS = lCase(ganJingS)                                                        'Сд�ɾ���S����
            tempS = s                                                                       '�ݴ�s����
            rowC = "" : tempRow = ""                                                        '���ÿ��ASP������ݴ������д���
            noteCode = ""                                                                   'Ĭ��ע�ʹ���Ϊ��
            downRow = ""                                                                    '��һ�д���
            downWord = ""                                                                   '��һ�е���
            if(nI + 1) <= uBound(splStrRow) then
                downRow = splStrRow(nI + 1)                                                        '��һ������
                downWord = getBeforeWord(downRow)                                               '��һ�е���
            'call echo("����һ�������뵥�ʡ�" & downRow,downWord)
            end if 
            nc = ""                                                                         '����Ϊ��
            nDimArrayQianTao = 0                                                            '��������Ƕ����
            isRowASP=false												'�Ƿ�Ϊһ��ASP���룬Ϊ��
			'call echo("s",replace(lCaseS,"<","&lt;"))
            '�������ͷΪ on ���ֵ�
            if left(pHPTrim(lCase(s)), 3) = "on " and 1=2 then
                s = ""  
			'���ش���
			elseif checkHideCode(s)=true and isASP=true then	
				s=""
            end if 

            isLoop = true                                                                   'ѭ���ַ�Ϊ��
            'ѭ��ÿ���ַ�
            for i = 1 to len(s)
                sx = mid(s, i, 1) : tempSx = sx 
                beforeStr = right(replace(mid(s, 1, i - 1), " ", ""), 1)                        '��һ���ַ�
                afterStr = left(replace(mid(s, i + 1), " ", ""), 1)                             '��һ���ַ�
                endCode = mid(s, i + 1)                                                         '��ǰ�ַ���������� һ��
				
				'call echo(s,i & "��" & sx)

                'Asp��ʼ
                if sx = "<" and isASP = false and wc = "" then
                    if mid(s, i + 1, 1) = "%" then
                        isASP = true                                                                   'ASPΪ��
 
						sx=getLanguageStartEndLabel(s,"start")
                        upWord = "<" & "%" 

                        i = i + 1 
                        rowC = rowC & sx 
                    elseIf isDisplayHtml = true then
                        rowC = rowC & sx                                                                '������ASP����ʹ��
                    end if 

                'ASP����
                elseIf sx = "%" and mid(s, i + 1, 1) = ">" and isASP = true and wc = "" then
                    isASP = false                   
					sx=getLanguageStartEndLabel(language,"end") 
                    i = i + 1 
                    rowC = rowC & sx 
                'ASPΪ��
                elseIf isASP = true then
				
				
                    '�����ı�
                    if(sx = """" or wc <> "") then
                        '˫�����ۼ�
                        if sx = """" then sYHCount = sYHCount + 1 
                        '�ж��Ƿ�"�����
                        if sYHCount mod 2 = 0 then
                            if mid(s, i + 1, 1) <> """" then
                                wc = wc & sx 
                                s1 = right(replace(mid(s, 1, i - len(wc)), " ", ""), 1) '����������Ҫ�������

                                '������������
                                wc = right(wc, len(wc) - 1) 
                                wc = left(wc, len(wc) - 1) 
								
								
								 
                                if 1=2 then
									wc = handleLanguage(wc, "����")                         'ASP����תPHP����
								else
									wc=""""& wc &""""
								end if
								
								if isSpaceFunctionName = "���պ���" or isNotShowFunctionName = "����ʾ����" then
									wc = "" 
								end if 
								
                                if wc = "'[Ϊ�ֵ䲻�������]'" and isZD = true then
                                    wc = "" 
                                end if 

                                rowC = rowC & wc                                        '���������ۼ�
                                sYHCount = 0 : wc = ""                                  '���
                            else
                                wc = wc & sx 
                            end if 
                        else
                            wc = wc & sx 
                        end if 

                    '+-*\=&   �ַ����⴦��
                    elseIf inStr(".&=,+-*/:()><[]", sx) > 0 and isASP = true and nc = "" then
                        if zc <> "" then
                            tempZc = zc 
                            zc = handleLanguage(zc, "����")              'ASP����תPHP���� 
							
							call recordUpWord(tempZc,zc)				'��¼�ϸ����ʼ����ϵ���
							
                            rowC = rowC & zc 
                            zc = ""                                      '�����ĸ����
                        end if 
                        if sx = "=" then
                            rowC = rTrim(rowC)
                        end if 
						 
                        tempSx=sx
                        rowC = rowC & handleLanguage(sx, "") 
   						call recordUpWord(tempSx,sx)				'��¼�ϸ����ʼ����ϵ���
						 
                    'ע�����˳�
                    elseIf sx = "'" then
                        noteCode = "'" & mid(s, i + 1) 
                        if isSpaceFunctionName = "���պ���" or isNotShowFunctionName = "����ʾ����" then
                            noteCode = "" 
                        end if 
                        exit for
                    '��ĸ
                    elseIf checkABC(sx) = true or sx = "_" or zc <> "" then
                        zc = zc & sx 
                        yesOK = true 
                        s1 = mid(s & " ", i + 1, 1) 
                        s2 = mid(zc, 1, 1) 
                        if checkABC(s1) <> true and s1 <> "_" then
                            yesOK = false 
                        end if 
                        '�����������������
                        if checkNumber(s1) = true and checkABC(s2) = true then
                            yesOK = true 
                        end if 
                        if yesOK = false then
                            'Remע��
                            if lCase(zc) = "rem" then
                                sx = mid(s, i + 1) : i = i + len(sx) 
                                zc = zc & sx 
                            end if 

                            tempZc = zc 
                            zc = handleLanguage(zc, "����")              'ASP����תPHP���� 
   							call recordUpWord(tempZc,zc)				'��¼�ϸ����ʼ����ϵ���
							 
                            rowC = rowC & zc 
                            zc = ""                                      '�����ĸ����
                        end if 
                    'Ϊ����
                    elseIf checkNumber(sx) = true or nc <> "" then
                        nc = nc & sx 
                        if afterStr <> "." and checkNumber(afterStr) = false then
                            nc = handleLanguage(nc, "����")              'ASP����תPHP����
                            rowC = rowC & nc 
                            nc = "" 
                        end if 
                    '�������͵�ֵ�� \ ^
					else
						if isSpaceFunctionName = "���պ���" or isNotShowFunctionName = "����ʾ����" then
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
            '�Բ���ʾ���� �����޺��� ����  ��Ҫ����20151223
            if isNotShowFunctionName <> "����ʾ����" and isSpaceFunctionName <> "���պ���" then
                if isASP = true then toC = toC & vbCrLf                                            'CssΪ����
            end if 
        next 
     
        aspToLanguage = toC
		

    end function 
	
	'��������
	function handleLanguage(varName,sType) 
        dim s, action, tempVarName, temp1,nLen,tempS,tempS2,tempS3,tempS4,tempS5,tempS6,tempS7,tempStart,tempEnd,tempBR,tempLCase,tempWord
		
        tempVarName = varName                                                           '��ʱ��������
        s = lCase(varName)
		'�жϺ���������Ϊ�����ں�������ִ����ۼ�
		if functionName=varName and functionName<>"" then 
			nFunctionNameShowCount=nFunctionNameShowCount+1
			'call echo( functionName,nFunctionNameShowCount)
		end if
		
		
		
		'������ʾ����
        if isNotShowFunctionName = "����ʾ����" then
            if inStr("|function|sub|", "|" & getArrayVar() & "|") = false then
                handleLanguage = "" 
                exit function 
            else
                if inStr("|function|sub|", "|" & lCase(varName) & "|") = false then
                    handleLanguage = "" 
                    exit function 
                end if 
            end if 
        elseIf isSpaceFunctionName = "���պ���" then
            if inStr("|function|sub|", "|" & getArrayVar() & "|") = false then
                handleLanguage = "" 
                exit function 
            '����ж�  Ϊ��
            elseIf trueJudge = false and s <> "function" then
                handleLanguage = "" 
                exit function 
            end if 
		end if
		
		if s="err" then
			if lcase(getBeforeNStr(endCode, "ȫ��", 1)) = "then" then
                handleLanguage = "Err.Number<>0" 
                exit function 
			ElseIf lcase(getBeforeNStr(endCode, "ȫ��", 1)) = "=" then
                handleLanguage = "Err.Number" 
                exit function 
			else
				varName="Err"
			end if
		'���
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
		
		
        '�ж��Ƿ�Ϊϵͳ���� ��Ϊ׷�ӱ���
        if checkAspSystemFunction(s) = true then
            call AddArrayVar(s)                                                             '׷���������
            arrayVarLeftKuoHao(getArrayIndex()) = 0                                         '���ݱ��� ��������
            arrayVarRightKuoHao(getArrayIndex()) = 0                                        '���ݱ��� ��������
        end if 

        action = getArrayVar()                                                          '��õ�ǰ����
        nThisIndex = getArrayIndex()                                                    '��ǰ����
		 
		'call echo("action",action)
        if action = "array" then
			if s="array" then 
				varName="" 
         	elseif s="(" then 
				varName="{" 
            elseIf s = ")" then 
	            varName = "}"  
                call DelArrayVar()                                                              'ɾ��������� 
			end if

            call addMsg(action, tempVarName, sType, varName)                                '׷�ӻ���
			 
        elseIf action = "function" then
            if s = "function" then
                if upWord = "end" then
                    if isNotShowFunctionName = "����ʾ����" then
                        varName = "" 
                    '���ÿ��� (��������20151119)
                    elseIf isFunctionReturn = true then					
						'�������
						if nFunctionNameShowCount=0 then 
							varName="return """""
						end if
                   end if
 					
					'call echo("functionName",functionName)
                    if isSpaceFunctionName = "���պ���" then
                        varName =  vbcrlf & "    " & functionName & "=""""" & vbCrLf & "end " & varName 
                    end if 

                    functionName = ""                                                               '��պ�������
                    isNotShowFunctionName = ""                                                      '�������ʾ����
                    isSpaceFunctionName = ""                                                        '�������
                    call DelArrayVar()                                                              'ɾ���������
                    call DelArrayVar()                                                              'ɾ���������
                    functionInDimList = ""                                                          '��պ����ﶨ�����
					functionVarList="|"
                '�˳� ��ɾ��һ������
                elseIf upWord = "exit" then
                    call DelArrayVar()                                                              'ɾ���������
                else
                    varName = "function "
					 
                    trueJudge = true                                                                '�ж�Ϊ��
                    functionName = "" 
                    isFunctionDimName = "������������" 
                    isFunctionReturn = false 
					nFunctionNameShowCount=0														'�������Ƴ�������
					isFunctionSelfVar=false															'���������Ƿ�Ϊ����
					functionVarList="|"																'���������б����
					 
                    if inStr("|" & notShowFunctionList & "|", "|" & lCase(getBeforeWord(endCode)) & "|") > 0 then 
                        isNotShowFunctionName = "����ʾ����" 
                        varName = "" 
                    elseIf inStr("|" & spaceFunctionList & "|", "|" & lCase(getBeforeWord(endCode)) & "|") > 0 then
                        isSpaceFunctionName = "���պ���"  
                    end if 
                end if 
			elseIf s = ")" and trueJudge = true then
				varName = ")" 
				trueJudge = false 
				isFunctionDimName = "�رպ�����������" 
			elseIf sType = "����" then
                if functionName = "" then
                    functionName = varName 
                    functionNameList = functionNameList & varName & "|"                      '���������ۼ� 
                    nShowFunction = 1 

                    functionInDimList = replace(getFunctionStrDim(endCode), ",", "|") & "|"         '�����ﶨ��ı��� ���ﲻҪ�ۼ�
     
                end if
			end if
		end if
		handleLanguage=varName
	end function     
    '׷���������
    function addArrayVar(varName)
        dim i 
        for i = 0 to uBound(arrayVar)
            if arrayVar(i) = "" then
                arrayVar(i) = varName 
                exit for 
            end if 
        next 
    end function 
    '����������  �������ǰ  (Ӧ���Ǵ�ǰ���)
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
    '�����������
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
    'ɾ��������� �������ǰ  (Ӧ���Ǵ�ǰ���)
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
    '׷�ӻ��� call addMsg("if",tempVarName,sType,varName)
    sub addMsg(lableName, tempVarName, sType, varName)
        'getArrayIndex()�滻�ɹ� nThisIndex ��ʾ��׼
        tempActionList = tempActionList & nThisIndex & "��" & lableName & "(" & tempVarName & ")(" & sType & ")(" & varName & ")<br>" & vbCrLf 
    end sub  
    '��õ����б����ı�����ȡ
    function getWordList()
        dim c, splStr, i, s, splxx 
        c = getFText("\VB����\Config\��վ��������205.txt") 
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
    '����Ƿ�ΪASPϵͳ����  ����ǳ���Ҫ ��ϵͳ�������⴦��
    function checkAspSystemFunction(byVal functionName)
        if inStr(systemFunctonList, "|" & lCase(functionName) & "|") > 0 then
            checkAspSystemFunction = true 
        else
            checkAspSystemFunction = false 
        end if 
    end function 
	
	'���ǰһ���ַ� GetBeforeStr(EndCode)
    function getBeforeStr(byVal content)
        getBeforeStr = left(trimVbCrlf(content), 1) 
    end function 
    '��ú�һ���ַ�
    function getAfterStr(byVal content)
        getAfterStr = right(trimVbCrlf(content), 1) 
    end function 


    '���ǰһ������
    function getBeforeWord(byVal content)
        content = trimVbCrlf(content) 
        content = replace(replace(replace(replace(content, "(", " ( "), ",", "  , "), ")", " ) "), ".", " . ") 
        if inStr(content, " ") then
            content = mid(content, 1, inStr(content, " ") - 1) 
        end if 
        getBeforeWord = content 
    end function 
    '���ǰǰһ������
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
    '��ú�һ������
    function getAfterWord(byVal content)
        content = trimVbCrlf(content) 
        if inStr(content, " ") then
            content = mid(content, inStr(content, " ") + 1) 
        end if 
        getAfterWord = content 
    end function 


    '��ʼ��
    private sub class_Initialize()
        'systemFunctonList = "|class|if|for|while|select|iif|function|sub|dim|rs|response|request|session|chr||call|left|right|rnd|mid|cint|instr|instrrev|vbtab|vbcrlf|split|mod|server|array|" 
		systemFunctonList = "|array|function|sub|" 
		
        systemFunctonList = lCase(systemFunctonList) 
		
		yjFunctionList="|getN,int|getVoid,void|"										'�Ͻ����庯���б�
		yjFunctionList=lcase(yjFunctionList)
		yjVarList="|i,int|n,int|"												'�Ͻ���������б�
		yjVarList=lcase(yjVarList)
		
        isDisplayHtml = false                                                           '����Ĭ���Ƿ���ʾHTML����
        isOpenImportFunction = false                                                    '����Ĭ���Ƿ�򿪵��뺯��
        isClass = false : className = "" : classNameList = "|" : dimClassNameList = "|" '���ʼ��
        isZD = false : zDName = "" : zDNameList = "|"                                   '�ֵ��ʼ��
        phpClearLable = "@"                                                             '20160624
		isDimArray=false																'�Ƿ�ΪDim����ı���
		isInstrReplaceFlash=false		'�滻instr��̨��falseΪ��
    end sub 


    '��� �ݴ涯���б�
    public property get getGlobalDimList()
        getGlobalDimList = globalDimList 
    end property 
    '���ú��������б�
    public property let setGlobalDimList(str)
        globalDimList = "|" & str & "|" 
    end property 
    '��� �ݴ涯���б�
    public property get getTempActionList()
        getTempActionList = tempActionList 
    end property 
    '���ú��������б�
    public property let setFunctionNameList(str)
        functionNameList = "|" & str & "|" 
    end property 
    '���ò���ʾ���������б�
    public property let setNotShowFunctionList(str)
        notShowFunctionList = lCase("|" & str & "|") 
    end property 
    '�������պ��������б�
    public property let setSpaceFunctionList(str)
        spaceFunctionList = lCase("|" & str & "|") 
    end property 
    '�����Ƿ���ʾHTML����
    public property let setIsDisplayHtml(str)
        isDisplayHtml = str 
    end property 
    '�����Ƿ��뺯������
    public property let setIsOpenImportFunction(str)
        isOpenImportFunction = str 
    end property 
    '�����������б�
    public property let setClassNameList(str)
        classNameList = "|" & str & "|" 
    end property 
    '����ת����ʲô����
    public property let setToLanguage(str)
        toLanguage = str 
    end property 
    '�����Ͻ������б�  ��  getn Ϊint
    public property let set_yjFunctionList(str)
        yjFunctionList = lcase(str)				'תСд 
    end property 
	'�����Ͻ������б�
    public property let set_yjVarList(str)
        yjVarList = lcase(str)			'תСд 
    end property  
	'����������ʾ������������������б�
    public property let set_clearFunVarList(str)
        clearFunVarList = lcase(str)			'תСд 
    end property  
	 
	'����Ƿ�Ҫ���ص�ǰ���� 20160729
	function checkHideCode(thisS)
		dim splstr,s,content,thisRow
		thisRow=lcase(pHPTrim(thisS))
		checkHideCode=false
		if instr(thisRow,"'��@��"& toLanguage &"����@��")>0 then
			checkHideCode=true
			exit function
		end if
		
		content=replace("|.netc|asp|php|as|jsp|","|"& toLanguage &"|","|")
		splstr=split(content,"|")
		for each s in splstr
			if s <>"" then
				if instr(thisRow,"'��@����"& s &"����@��")>0 or instr(ucase(thisRow),"'��"& ucase(s) &"��")>0 then
					checkHideCode=true
					exit function
				end if
			end if
		next
		if instr(thisRow,"'��@��"& toLanguage &"��ʾ@��")>0 then
			toC=toC & replace(thisS, "'��@��"& toLanguage &"��ʾ@��","")
			checkHideCode=true
			exit function
		end if
		 
	end function
	
	'��¼��һ������  yuanWord��ת����ĵ���
	sub recordUpWord(byval word, byval yuanWord)
		upWordYuan6 = upWordYuan5                    '��6Դ���� ���� Դ��5����
		upWordYuan5 = upWordYuan4                    '��5Դ���� ���� Դ��4����
		upWordYuan4 = upWordYuan3                    '��4Դ���� ���� Դ��3����
		upWordYuan3 = upWordYuan2                    '��3Դ���� ���� Դ��2����
		upWordYuan2 = upWordYuan                     '��2Դ���� ���� Դ��1����
		upWordYuan = yuanWord                          '��¼��һ����
		
		upWord6 = upWord5
		upWord5 = upWord4 
		upWord4 = upWord3 
		upWord3 = upWord2 
		upWord2 = upWord                             '��¼����һ������
		upWord = lCase(word)                       '��¼��һ����������   ΪСд
	end sub



	'������Կ�ʼ������ǩֵ  ��PHP�� <?php
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
