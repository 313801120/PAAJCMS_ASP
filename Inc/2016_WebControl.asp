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
'��վ���� 20160223



'����ģ���滻����
function handleModuleReplaceArray(byVal content)
    dim i, startStr, endStr, s, lableName 
    for i = 1 to uBound(moduleReplaceArray) - 1
        if moduleReplaceArray(i, 0) = "" then
            exit for 
        end if 
        'call echo(ModuleReplaceArray(i,0),ModuleReplaceArray(0,i))
        lableName = moduleReplaceArray(i, 0) 
        s = moduleReplaceArray(0, i) 
        if lableName = "��ɾ����" then
            content = replace(content, s, "") 
        else
            startStr = "<replacestrname " & lableName & ">" : endStr = "</replacestrname " & lableName & ">" 
            if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
                content = replaceContentModule(content, startStr, endStr, s, "") 
            end if 
            startStr = "<replacestrname " & lableName & "/>" 
            if inStr(content, startStr) > 0 then
                content = replaceContentRowModule(content, "<replacestrname " & lableName & "/>", s, "") 
            end if 
        end if 
    next 
    handleModuleReplaceArray = content 
end function 

'ȥ��ģ���ﲻ��Ҫ��ʾ���� ɾ��ģ�����ҵ�ע�ʹ���
function delTemplateMyNote(code)
    dim startStr, endStr, i, s, nHandleNumb, splStr, block, id 
    dim content, dragSortCssStr, dragSortStart, dragSortEnd, dragSortValue, c 
    dim lableName, lableStartStr, lableEndStr 
    nHandleNumb = 99                                                                '���ﶨ�����Ҫ

    '��ǿ��  �����Ҳ����<!--#aaa start#--><!--#aaa end#-->
    startStr = "<!--#" : endStr = "#-->" 
    for i = 1 to nHandleNumb
        if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
            lableName = strCut(code, startStr, endStr, 2) 
            if instr(lableName, " start") > 0 then
                lableName = mid(lableName, 1, len(lableName) - 6) 
            end if 

            s = startStr & lableName & endStr 
            lableStartStr = startStr & lableName & " start" & endStr 
            lableEndStr = startStr & lableName & " end" & endStr 
            if inStr(code, lableStartStr) > 0 and inStr(code, lableEndStr) > 0 then
                s = strCut(code, lableStartStr, lableEndStr, 1) 
            'call echo(">>",s)
            end if 
            code = replace(code, s, "") 
            'call echo("s",s)
            'call echo("lableName",lableName)
            'call echo("lableStartStr",replace(lableStartStr,"<","&lt;"))
        'call echo("lableEndStr",replace(lableEndStr,"<","&lt;"))
        else
            exit for 
        end if 
    next 



    '���ReadBlockList�������б�����  �����и�����ĵط����������ݿ��Դ��ⲿ�������ݣ�����Ժ���
    'Call Eerr("ReadBlockList",ReadBlockList)
    'д��20141118
    'splStr = Split(ReadBlockList, vbCrLf)                 '�������֣�������
    '�޸���20151230
    for i = 1 to nHandleNumb
        startStr = "<R#��������" : endStr = " start#>" 
        block = strCut(code, startStr, endStr, 2) 
        if block <> "" then
            startStr = "<R#��������" & block & " start#>" : endStr = "<R#��������" & block & " end#>" 
            if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
                s = strCut(code, startStr, endStr, 1) 
                code = replace(code, s, "")                                                     '�Ƴ�
            end if 
        else
            exit for 
        end if 
    next 

    'ɾ����ҳ����20160309
    startStr = "<!--#list start#-->" 
    endStr = "<!--#list end#-->" 
    if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
        s = strCut(code, startStr, endStr, 2) 
        code = replace(code, s, "") 
    end if 

    if request("gl") = "yun" then
        content = getFText("/Jquery/dragsort/Config.html") 
        content = getFText("/Jquery/dragsort/ģ����ק.html") 
        'Css��ʽ
        startStr = "<style>" 
        endStr = "</style>" 
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            dragSortCssStr = strCut(content, startStr, endStr, 1) 
        end if 
        '��ʼ����
        startStr = "<!--#top start#-->" 
        endStr = "<!--#top end#-->" 
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            dragSortStart = strCut(content, startStr, endStr, 2) 
        end if 
        '��������
        startStr = "<!--#foot start#-->" 
        endStr = "<!--#foot end#-->" 
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            dragSortEnd = strCut(content, startStr, endStr, 2) 
        end if 
        '��ʾ������
        startStr = "<!--#value start#-->" 
        endStr = "<!--#value end#-->" 
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            dragSortValue = strCut(content, startStr, endStr, 2) 
        end if 



        '���ƴ���
        startStr = "<dIv datid='" 
        endStr = "</dIv>" 
        content = getArray(code, startStr, endStr, false, false) 
        splStr = split(content, "$Array$") 
        for each s in splStr
            startStr = "��DatId��'" 
            id = mid(s, 1, inStr(s, startStr) - 1) 
            s = mid(s, inStr(s, startStr) + len(startStr)) 
            'C=C & "<li><div title='"& Id &"'>" & vbcrlf & "<div " & S & "</div>"& vbcrlf &"<div class='clear'></div></div><div class='clear'></div></li>"
            s = "<div" & s & "</div>" 
            'Call Die(S)
            c = c & replace(replace(dragSortValue, "{$value$}", s), "{$id$", id) 
        next 
        c = replace(c, "�����С�", vbCrLf) 
        c = dragSortStart & c & dragSortEnd 
        code = mid(code, 1, inStr(code, "<body>") - 1) 
        code = replace(code, "</head>", dragSortCssStr & "</head></body>" & c & "</body></html>") 
    end if 

    'ɾ��VB������ɵ���������
    startStr = "<dIv datid='" : endStr = "��DatId��'" 
    for i = 1 to nHandleNumb
        if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
            id = strCut(code, startStr, endStr, 2) 
            code = replace2(code, startStr & id & endStr, "<div ") 
        else
            exit for 
        end if 
    next 
    code = replace(code, "</dIv>", "</div>")                                  '�滻���������div

    '����Χ���
    startStr = "<!--#dialogteststart#-->" : endStr = "<!--#dialogtestend#-->" 
    code = replace(code, "<!--#dialogtest start#-->", startStr) 
    code = replace(code, "<!--#dialogtest end#-->", endStr) 
    for i = 1 to nHandleNumb
        if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
            s = strCut(code, startStr, endStr, 1) 
            code = replace2(code, s, "") 
        else
            exit for 
        end if 
    next 
    '��ת���
    startStr = "<!--#teststart#-->" : endStr = "<!--#testend#-->" 
    code = replace(code, "<!--#del start#-->", startStr)                         '������һ��
    code = replace(code, "<!--#del end#-->", endStr)                             '������һ�� ����ʽ
    code = replace(code, "<!--#test start#-->", startStr) 
    code = replace(code, "<!--#test end#-->", endStr) 

    for i = 1 to nHandleNumb
        if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
            s = strCut(code, startStr, endStr, 1) 
            code = replace2(code, s, "") 
        else
            exit for 
        end if 
    next 
    'ɾ��ע�͵�span
    code = replace(code, "<sPAn class=""testspan"">", "")                        '����Span
    code = replace(code, "<sPAn class=""testhidde"">", "")                       '����Span
    code = replace(code, "</sPAn>", "") 

    'delTemplateMyNote = Code:Exit Function

    startStr = "<!--#" : endStr = "#-->" 
    for i = 1 to nHandleNumb
        if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
            s = strCut(code, startStr, endStr, 1) 
            code = replace2(code, s, "") 
        else
            exit for 
        end if 
    next 


    delTemplateMyNote = code 
end function 

'�����滻����ֵ 20160114
function handleReplaceValueParam(content, byVal paramName, replaceStr)
    if inStr(content, "[$" & paramName) = false then
        paramName = LCase(paramName) 
    end if 
    handleReplaceValueParam = replaceValueParam(content, paramName, replaceStr) 
end function 

'�滻����ֵ 2014  12 01
function replaceValueParam(content, paramName, replaceStr)
    dim startStr, endStr, labelStr, tempLabelStr, sLen, sNTimeFormat, delHtmlYes, funStr, trimYes, sIsEscape, s, i 
    dim ifStr                                                                       '�ж��ַ�
    dim elseIfStr                                                                   '�ڶ��ж��ַ�
    dim valueStr                                                                    '��ʾ�ַ�
    dim elseStr                                                                     '�����ַ�
    dim elseIfValue, elseValue                                                      '�ڶ��ж�ֵ
    dim instrStr, instr2Str                                                         '�����ַ�
    dim tempReplaceStr                                                              '�ݴ�
	dim stype																		'����
	dim nLeft,nRight
	dim isIf,isElseIf:isIf=false:isElseIf=false										'Ϊif�ж�
	dim isDefault:isDefault=true													'�Ƿ�ΪĬ��ֵ 
	dim tableName																	'����������
    'ReplaceStr = ReplaceStr & "�����������������ʱ̼ѽ��"
    'ReplaceStr = CStr(ReplaceStr)            'ת���ַ�����
    if isNul(replaceStr) = true then replaceStr = "" 
    tempReplaceStr = replaceStr 
	dim ssubnav												'�����ӵ�������

    '��ദ��99��  20160225
    for i = 1 to 999
        replaceStr = tempReplaceStr                                                     '�ָ�
        startStr = "[$" & paramName : endStr = "$]" 
        '�ֶ������ϸ��ж� 20160226
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 and(inStr(content, startStr & " ") > 0 or inStr(content, startStr & endStr) > 0) then
            '��ö�Ӧ�ֶμ�ǿ��20151231
            if inStr(content, startStr & endStr) > 0 then
                labelStr = startStr & endStr 
            elseIf inStr(content, startStr & " ") > 0 then
                labelStr = strCut(content, startStr & " ", endStr, 1) 
            else
                labelStr = strCut(content, startStr, endStr, 1) 
            end if 

            tempLabelStr = labelStr 
            labelStr = handleInModule(labelStr, "start") 
            'ɾ��Html
            delHtmlYes = rParam(labelStr, "delHtml")                                        '�Ƿ�ɾ��Html
            if delHtmlYes = "true" then replaceStr = replace(delHtml(replaceStr), "<", "&lt;") 'HTML����
			
			'ɾ��url��̨Ϊindex.html 20180106
			dim delIndexHtml			        
			delIndexHtml = rParam(labelStr, "delIndexHtml")            
			if delIndexHtml=true or delIndexHtml="1" then 
				if lcase(right(replaceStr,10))="index.html" then
					replaceStr=left(replaceStr,len(replaceStr)-10)
				end if
			end if
			
            'ɾ�����߿ո�
            trimYes = rParam(labelStr, "trim")                                              '�Ƿ�ɾ�����߿ո�
            if trimYes = "true" then replaceStr = trimVbCrlf(replaceStr) 

            '��ȡ�ַ�����
            sLen = rParam(labelStr, "len")                                                  '�ַ�����ֵ
            sLen = handleNumber(sLen) 
            'If sLen<>"" Then ReplaceStr = CutStr(ReplaceStr,sLen,"null")' Left(ReplaceStr,sLen)
            if sLen <> "" then  
				replaceStr = cutStr(replaceStr, cint(sLen), "...")           'Left(ReplaceStr,nLen)
			end if
			'���� �ӵ���20180116
			ssubnav=rParam(labelStr, "subnav")     
			if ssubnav<>"" then
				replaceStr=subnav(replaceStr,ssubnav)
			end if 
			
			
            'ʱ�䴦��
            sNTimeFormat = rParam(labelStr, "format_time")                                   'ʱ�䴦��ֵ
			tableName = rParam(labelStr, "tablename")										'������ 
            if sNTimeFormat <> "" then
				if instr("qmydwhms",sNTimeFormat)>0 then
                	replaceStr = getFormatYMD(replaceStr, sNTimeFormat) 
				else
                	replaceStr = format_Time(replaceStr, cint(sNTimeFormat)) 
            	end if
			end if 

            '�����Ŀ����
            s = rParam(labelStr, "getcolumnname") 
            if s <> "" then
                if s = "@ME" then
                    s = replaceStr 
                end if 
				if EDITORTYPE = "jsp" then
					replaceStr=s
				else 	
                	replaceStr =handleGetColumnName(tableName,s) 'getcolumnname(s) 
				end if
            end if 
            '�����ĿURL
            s = rParam(labelStr, "getcolumnurl") 
            if s <> "" then
                if s = "@ME" then
                    s = replaceStr 
                end if 
                replaceStr = getcolumnurl(s, "id") 
            end if 
            '�Ƿ�Ϊ��������  ����ʲô��˼������
            s = rParam(labelStr, "password") 
            if s <> "" then
                if s <> "" then
                    replaceStr = s 
                end if 
            end if 

            ifStr = rParam(labelStr, "if") 
            elseIfStr = rParam(labelStr, "elseif") 
            valueStr = rParam(labelStr, "value") 
			 
            elseIfValue = rParam(labelStr, "elseifvalue") 
            elseValue = rParam(labelStr, "elsevalue") 
            instrStr = rParam(labelStr, "instr") 
            instr2Str = rParam(labelStr, "instr2") 
            nLeft = rParam(labelStr, "left") 
            nRight = rParam(labelStr, "right")  
			
            stype = trim(lcase(rParam(labelStr, "stype")))
			
			'Ϊif�жϼ��
			if instr(labelStr," if='")>0 then
				isIf=true
			end if
			if instr(labelStr," elseif='")>0 then
				isElseIf=true
			end if
			
			'����ID�����Ŀ����
			if stype=lcase("ParentColumnName") then
				replaceStr=getParentColumnName(replaceStr)
			'����ID�����Ŀ��ַ
			elseif stype=lcase("ColumnNameUrl") then
				replaceStr=getColumnUrl(replaceStr,"")
			'����ID�����Ŀ��ַ
			elseif stype=lcase("ColumnNameToNameUrl") then
				replaceStr=getColumnUrl(replaceStr,"name")
			end if
			
			'׷����20170603
			if nLeft<>"" then
				replaceStr=left(replaceStr,cint(nLeft))
			elseif nRight<>"" then
				replaceStr=right(replaceStr,cint(nRight))
			end if

            'call echo("ifStr",ifStr)
            'call echo("valueStr",valueStr)
            'call echo("elseStr",elseStr)
            'call echo("elseIfStr",elseIfStr)
            'call echo("replaceStr",replaceStr)
            if isIf=true then 
				isDefault=false 						'Ĭ��ֵΪ��
                if ifStr = CStr(replaceStr)   then		'and ifStr <> "" 
                    replaceStr = valueStr 
                elseif elseIfStr = CStr(replaceStr) and isElseIf=true then	'and elseIfStr <> ""
                    replaceStr = valueStr 
                    if elseIfValue <> "" then
                        replaceStr = elseIfValue 
                    end if
				else 
					replaceStr=elseValue				'IFĬ��ֵ
					isDefault=true 						'Ĭ��ֵΪ��
					'call echo(ifStr , CStr(replaceStr))
				end if
			end if
			if instrStr <> "" then
                if inStr(CStr(replaceStr), instrStr) > 0 and instrStr <> "" then
                    replaceStr = valueStr 
                elseIf inStr(CStr(replaceStr), instr2Str) > 0 and instr2Str <> "" then
				
                    replaceStr = valueStr 
                    if elseIfValue <> "" then
                        replaceStr = elseIfValue 
                    end if 
                else
                    if elseValue <> "@ME" then
                        replaceStr = elseValue 
                    end if 
                end if 
            end if 

            '��������20151231    [$title  function='left(@ME,40)'$]
            funStr = rParam(labelStr, "function")                                           '����
            if funStr <> "" then
                funStr = replace(funStr, "@ME", replaceStr) 
                replaceStr = handleContentCode(funStr, "") 
            end if 

            'Ĭ��ֵ
            s = rParam(labelStr, "default")  
            if s <> "" and isDefault=true then   ' and s <> "@ME"  
            	replaceStr = s  
            end if
            'escapeת��
            sIsEscape = lcase(rParam(labelStr, "escape")) 
            if sIsEscape = "1" or sIsEscape = "true" then
                replaceStr = escape(replaceStr) 
            end if 

            '�ı���ɫ
            s = rParam(labelStr, "fontcolor")                                               '����
            if s <> "" then
                replaceStr = "<font color=""" & s & """>" & replaceStr & "</font>" 
            end if 
			 
			replaceStr=replace(replaceStr,"@ME",tempReplaceStr)			'�滻@ME����20170707

            'call echo(tempLabelStr,replaceStr)
            content = replace(content, tempLabelStr, replaceStr) 
        else
            exit for 
        end if 
    next 
    replaceValueParam = content 
end function 


'��ʾ�༭��20160115
function displayEditor(action)
    dim c 
    c = c & "<script type=""text/javascript"" src=""\Jquery\syntaxhighlighter\scripts/shCore.js""></" & "script> " & vbCrLf 
    c = c & "<script type=""text/javascript"" src=""\Jquery\syntaxhighlighter\scripts/shBrushJScript.js""></" & "script>" & vbCrLf 
    c = c & "<script type=""text/javascript"" src=""\Jquery\syntaxhighlighter\scripts/shBrushPhp.js""></" & "script> " & vbCrLf 
    c = c & "<script type=""text/javascript"" src=""\Jquery\syntaxhighlighter\scripts/shBrushVb.js""></" & "script> " & vbCrLf 
    c = c & "<link type=""text/css"" rel=""stylesheet"" href=""\Jquery\syntaxhighlighter\styles/shCore.css""/>" & vbCrLf 
    c = c & "<link type=""text/css"" rel=""stylesheet"" href=""\Jquery\syntaxhighlighter\styles/shThemeDefault.css""/>" & vbCrLf 
    c = c & "<script type=""text/javascript"">" & vbCrLf 
    c = c & "    SyntaxHighlighter.config.clipboardSwf = '\Jquery\syntaxhighlighter\scripts/clipboard.swf';" & vbCrLf 
    c = c & "    SyntaxHighlighter.all();" & vbCrLf 
    c = c & "</" & "script>" & vbCrLf 

    displayEditor = c 
end function 
'������վurl20160202 
function handleWebUrl(url)
	'�ų�����htmlҳ20180313
	if isMakeHtml = false then
		if request("gl") <> "" then
			url = getUrlAddToParam(url, "&gl=" & request("gl"), "replace") 
		end if 
		if request("skin") <> "" then		'׷����20180313 ��վƤ��
			url = getUrlAddToParam(url, "&skin=" & request("skin"), "replace") 
		end if 
		if request("templatedir") <> "" then
			url = getUrlAddToParam(url, "&templatedir=" & request("templatedir"), "replace") 
		end if 
	end if
    handleWebUrl = url 
end function 

'
'���������޸�
'MainContent = HandleDisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainContent,"style='float:right;padding:0 4px;'")
function handleDisplayOnlineEditDialog(url, content, cssStyle, replaceStr)
    dim controlStr, splStr, s, isAdd 
    if request("gl") = "edit" then
        if inStr(url, "&") > 0 then
            url = url & "&vbgl=true" 
        end if 
        isAdd = false                                                                   '���Ĭ��Ϊ��
        controlStr = getControlStr(url) & """" & cssStyle 
        if replaceStr <> "" then
            splStr = split(replaceStr, "|") 
            for each s in splStr
                if s <> "" and inStr(content, s) > 0 then
                    content = replace2(content, s, s & controlStr) 
                    isAdd = true 
                    exit for 
                end if 
            next 
        end if 
        if isAdd = false then
            '��һ��
            'C = "<div "& ControlStr &">" & vbCrlf
            'C=C & Content & vbCrlf
            'C = C & "</div>" & vbCrlf
            'Content = C
            '�ڶ���
            content = htmlAddAction(content, controlStr) 

        'Content = "<div "& ControlStr &">" & Content & "</div>"
        end if 
    end if 
    handleDisplayOnlineEditDialog = content 
end function 
'��ÿ�������
function getControlStr(url)
	getControlStr=""
    if request("gl") = "edit" then
        getControlStr = " onMouseMove=""onColor(this,'#FDFAC6','red')"" onMouseOut=""offColor(this,'','')"" onDblClick=""window1('" & url & "','��Ϣ�޸�')"" title='˫�����Ҽ�ѡ�����޸�' oncontextmenu=""CommonMenu(event,this,'')" 'ɾ����ַΪ��
    end if 
end function 

'html�Ӷ���(20151103)  call rw(htmlAddAction("  <a href=""javascript:;"">222222</a>", "onclick=""javascript:alert(111);"" "))
function htmlAddAction(content, jsAction)
    dim s, startStr, endStr, isHandle, lableName 
    s = content 
    s = phptrim(s) 
    startStr = mid(s, 1, inStr(s, " ")) 
    endStr = ">" 
    isHandle = true 

    lableName = trim(LCase(replace(startStr, "<", ""))) 
    if inStr(s, startStr) = false or inStr(s, endStr) = false or inStr("|a|div|span|font|h1|h2|h3|h4|h5|h6|dt|dd|dl|li|ul|table|tr|td|", "|" & lableName & "|") = false then
        isHandle = false 
    end if 

    if isHandle = true then
        content = startStr & jsAction & right(s, len(s) - len(startStr)) 
    end if 
    htmlAddAction = content 
end function 


%>  

