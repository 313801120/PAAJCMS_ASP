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
'网站控制 20160223



'处理模块替换数组
function handleModuleReplaceArray(byVal content)
    dim i, startStr, endStr, s, lableName 
    for i = 1 to uBound(moduleReplaceArray) - 1
        if moduleReplaceArray(i, 0) = "" then
            exit for 
        end if 
        'call echo(ModuleReplaceArray(i,0),ModuleReplaceArray(0,i))
        lableName = moduleReplaceArray(i, 0) 
        s = moduleReplaceArray(0, i) 
        if lableName = "【删除】" then
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

'去掉模板里不需要显示内容 删除模板中我的注释代码
function delTemplateMyNote(code)
    dim startStr, endStr, i, s, nHandleNumb, splStr, block, id 
    dim content, dragSortCssStr, dragSortStart, dragSortEnd, dragSortValue, c 
    dim lableName, lableStartStr, lableEndStr 
    nHandleNumb = 99                                                                '这里定义很重要

    '加强版  对这个也可以<!--#aaa start#--><!--#aaa end#-->
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



    '清除ReadBlockList读出块列表内容  不过有个不足的地方，读出内容可以从外部读出内容，这个以后考虑
    'Call Eerr("ReadBlockList",ReadBlockList)
    '写于20141118
    'splStr = Split(ReadBlockList, vbCrLf)                 '不用这种，复杂了
    '修改于20151230
    for i = 1 to nHandleNumb
        startStr = "<R#读出内容" : endStr = " start#>" 
        block = strCut(code, startStr, endStr, 2) 
        if block <> "" then
            startStr = "<R#读出内容" & block & " start#>" : endStr = "<R#读出内容" & block & " end#>" 
            if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
                s = strCut(code, startStr, endStr, 1) 
                code = replace(code, s, "")                                                     '移除
            end if 
        else
            exit for 
        end if 
    next 

    '删除翻页配置20160309
    startStr = "<!--#list start#-->" 
    endStr = "<!--#list end#-->" 
    if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
        s = strCut(code, startStr, endStr, 2) 
        code = replace(code, s, "") 
    end if 

    if request("gl") = "yun" then
        content = getFText("/Jquery/dragsort/Config.html") 
        content = getFText("/Jquery/dragsort/模块拖拽.html") 
        'Css样式
        startStr = "<style>" 
        endStr = "</style>" 
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            dragSortCssStr = strCut(content, startStr, endStr, 1) 
        end if 
        '开始部分
        startStr = "<!--#top start#-->" 
        endStr = "<!--#top end#-->" 
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            dragSortStart = strCut(content, startStr, endStr, 2) 
        end if 
        '结束部分
        startStr = "<!--#foot start#-->" 
        endStr = "<!--#foot end#-->" 
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            dragSortEnd = strCut(content, startStr, endStr, 2) 
        end if 
        '显示块内容
        startStr = "<!--#value start#-->" 
        endStr = "<!--#value end#-->" 
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 then
            dragSortValue = strCut(content, startStr, endStr, 2) 
        end if 



        '控制处理
        startStr = "<dIv datid='" 
        endStr = "</dIv>" 
        content = getArray(code, startStr, endStr, false, false) 
        splStr = split(content, "$Array$") 
        for each s in splStr
            startStr = "【DatId】'" 
            id = mid(s, 1, inStr(s, startStr) - 1) 
            s = mid(s, inStr(s, startStr) + len(startStr)) 
            'C=C & "<li><div title='"& Id &"'>" & vbcrlf & "<div " & S & "</div>"& vbcrlf &"<div class='clear'></div></div><div class='clear'></div></li>"
            s = "<div" & s & "</div>" 
            'Call Die(S)
            c = c & replace(replace(dragSortValue, "{$value$}", s), "{$id$", id) 
        next 
        c = replace(c, "【换行】", vbCrLf) 
        c = dragSortStart & c & dragSortEnd 
        code = mid(code, 1, inStr(code, "<body>") - 1) 
        code = replace(code, "</head>", dragSortCssStr & "</head></body>" & c & "</body></html>") 
    end if 

    '删除VB软件生成的垃圾代码
    startStr = "<dIv datid='" : endStr = "【DatId】'" 
    for i = 1 to nHandleNumb
        if inStr(code, startStr) > 0 and inStr(code, endStr) > 0 then
            id = strCut(code, startStr, endStr, 2) 
            code = replace2(code, startStr & id & endStr, "<div ") 
        else
            exit for 
        end if 
    next 
    code = replace(code, "</dIv>", "</div>")                                  '替换成这个结束div

    '最外围清除
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
    '内转清除
    startStr = "<!--#teststart#-->" : endStr = "<!--#testend#-->" 
    code = replace(code, "<!--#del start#-->", startStr)                         '与下面一样
    code = replace(code, "<!--#del end#-->", endStr)                             '与下面一样 多样式
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
    '删除注释的span
    code = replace(code, "<sPAn class=""testspan"">", "")                        '测试Span
    code = replace(code, "<sPAn class=""testhidde"">", "")                       '隐藏Span
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

'处理替换参数值 20160114
function handleReplaceValueParam(content, byVal paramName, replaceStr)
    if inStr(content, "[$" & paramName) = false then
        paramName = LCase(paramName) 
    end if 
    handleReplaceValueParam = replaceValueParam(content, paramName, replaceStr) 
end function 

'替换参数值 2014  12 01
function replaceValueParam(content, paramName, replaceStr)
    dim startStr, endStr, labelStr, tempLabelStr, sLen, sNTimeFormat, delHtmlYes, funStr, trimYes, sIsEscape, s, i 
    dim ifStr                                                                       '判断字符
    dim elseIfStr                                                                   '第二判断字符
    dim valueStr                                                                    '显示字符
    dim elseStr                                                                     '否则字符
    dim elseIfValue, elseValue                                                      '第二判断值
    dim instrStr, instr2Str                                                         '查找字符
    dim tempReplaceStr                                                              '暂存
	dim stype																		'类型
	dim nLeft,nRight
	dim isIf,isElseIf:isIf=false:isElseIf=false										'为if判断
	dim isDefault:isDefault=true													'是否为默认值 
	dim tableName																	'操作表名称
    'ReplaceStr = ReplaceStr & "这里面放上内容在这时碳呀。"
    'ReplaceStr = CStr(ReplaceStr)            '转成字符类型
    if isNul(replaceStr) = true then replaceStr = "" 
    tempReplaceStr = replaceStr 
	dim ssubnav												'调用子导航动作

    '最多处理99个  20160225
    for i = 1 to 999
        replaceStr = tempReplaceStr                                                     '恢复
        startStr = "[$" & paramName : endStr = "$]" 
        '字段名称严格判断 20160226
        if inStr(content, startStr) > 0 and inStr(content, endStr) > 0 and(inStr(content, startStr & " ") > 0 or inStr(content, startStr & endStr) > 0) then
            '获得对应字段加强版20151231
            if inStr(content, startStr & endStr) > 0 then
                labelStr = startStr & endStr 
            elseIf inStr(content, startStr & " ") > 0 then
                labelStr = strCut(content, startStr & " ", endStr, 1) 
            else
                labelStr = strCut(content, startStr, endStr, 1) 
            end if 

            tempLabelStr = labelStr 
            labelStr = handleInModule(labelStr, "start") 
            '删除Html
            delHtmlYes = rParam(labelStr, "delHtml")                                        '是否删除Html
            if delHtmlYes = "true" then replaceStr = replace(delHtml(replaceStr), "<", "&lt;") 'HTML处理
			
			'删除url后台为index.html 20180106
			dim delIndexHtml			        
			delIndexHtml = rParam(labelStr, "delIndexHtml")            
			if delIndexHtml=true or delIndexHtml="1" then 
				if lcase(right(replaceStr,10))="index.html" then
					replaceStr=left(replaceStr,len(replaceStr)-10)
				end if
			end if
			
            '删除两边空格
            trimYes = rParam(labelStr, "trim")                                              '是否删除两边空格
            if trimYes = "true" then replaceStr = trimVbCrlf(replaceStr) 

            '截取字符处理
            sLen = rParam(labelStr, "len")                                                  '字符长度值
            sLen = handleNumber(sLen) 
            'If sLen<>"" Then ReplaceStr = CutStr(ReplaceStr,sLen,"null")' Left(ReplaceStr,sLen)
            if sLen <> "" then  
				replaceStr = cutStr(replaceStr, cint(sLen), "...")           'Left(ReplaceStr,nLen)
			end if
			'调用 子导航20180116
			ssubnav=rParam(labelStr, "subnav")     
			if ssubnav<>"" then
				replaceStr=subnav(replaceStr,ssubnav)
			end if 
			
			
            '时间处理
            sNTimeFormat = rParam(labelStr, "format_time")                                   '时间处理值
			tableName = rParam(labelStr, "tablename")										'表名称 
            if sNTimeFormat <> "" then
				if instr("qmydwhms",sNTimeFormat)>0 then
                	replaceStr = getFormatYMD(replaceStr, sNTimeFormat) 
				else
                	replaceStr = format_Time(replaceStr, cint(sNTimeFormat)) 
            	end if
			end if 

            '获得栏目名称
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
            '获得栏目URL
            s = rParam(labelStr, "getcolumnurl") 
            if s <> "" then
                if s = "@ME" then
                    s = replaceStr 
                end if 
                replaceStr = getcolumnurl(s, "id") 
            end if 
            '是否为密码类型  这是什么意思？？？
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
			
			'为if判断检测
			if instr(labelStr," if='")>0 then
				isIf=true
			end if
			if instr(labelStr," elseif='")>0 then
				isElseIf=true
			end if
			
			'根据ID获得栏目名称
			if stype=lcase("ParentColumnName") then
				replaceStr=getParentColumnName(replaceStr)
			'根据ID获得栏目网址
			elseif stype=lcase("ColumnNameUrl") then
				replaceStr=getColumnUrl(replaceStr,"")
			'根据ID获得栏目网址
			elseif stype=lcase("ColumnNameToNameUrl") then
				replaceStr=getColumnUrl(replaceStr,"name")
			end if
			
			'追加于20170603
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
				isDefault=false 						'默认值为假
                if ifStr = CStr(replaceStr)   then		'and ifStr <> "" 
                    replaceStr = valueStr 
                elseif elseIfStr = CStr(replaceStr) and isElseIf=true then	'and elseIfStr <> ""
                    replaceStr = valueStr 
                    if elseIfValue <> "" then
                        replaceStr = elseIfValue 
                    end if
				else 
					replaceStr=elseValue				'IF默认值
					isDefault=true 						'默认值为假
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

            '函数处理20151231    [$title  function='left(@ME,40)'$]
            funStr = rParam(labelStr, "function")                                           '函数
            if funStr <> "" then
                funStr = replace(funStr, "@ME", replaceStr) 
                replaceStr = handleContentCode(funStr, "") 
            end if 

            '默认值
            s = rParam(labelStr, "default")  
            if s <> "" and isDefault=true then   ' and s <> "@ME"  
            	replaceStr = s  
            end if
            'escape转码
            sIsEscape = lcase(rParam(labelStr, "escape")) 
            if sIsEscape = "1" or sIsEscape = "true" then
                replaceStr = escape(replaceStr) 
            end if 

            '文本颜色
            s = rParam(labelStr, "fontcolor")                                               '函数
            if s <> "" then
                replaceStr = "<font color=""" & s & """>" & replaceStr & "</font>" 
            end if 
			 
			replaceStr=replace(replaceStr,"@ME",tempReplaceStr)			'替换@ME这种20170707

            'call echo(tempLabelStr,replaceStr)
            content = replace(content, tempLabelStr, replaceStr) 
        else
            exit for 
        end if 
    next 
    replaceValueParam = content 
end function 


'显示编辑器20160115
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
'处理网站url20160202 
function handleWebUrl(url)
	'排除生成html页20180313
	if isMakeHtml = false then
		if request("gl") <> "" then
			url = getUrlAddToParam(url, "&gl=" & request("gl"), "replace") 
		end if 
		if request("skin") <> "" then		'追加于20180313 网站皮肤
			url = getUrlAddToParam(url, "&skin=" & request("skin"), "replace") 
		end if 
		if request("templatedir") <> "" then
			url = getUrlAddToParam(url, "&templatedir=" & request("templatedir"), "replace") 
		end if 
	end if
    handleWebUrl = url 
end function 

'
'处理在线修改
'MainContent = HandleDisplayOnlineEditDialog(""& adminDir &"NavManage.Asp?act=EditNavBig&Id=" & TempRs("Id") & "&n=" & GetRnd(11), MainContent,"style='float:right;padding:0 4px;'")
function handleDisplayOnlineEditDialog(url, content, cssStyle, replaceStr)
    dim controlStr, splStr, s, isAdd 
    if request("gl") = "edit" then
        if inStr(url, "&") > 0 then
            url = url & "&vbgl=true" 
        end if 
        isAdd = false                                                                   '添加默认为假
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
            '第一种
            'C = "<div "& ControlStr &">" & vbCrlf
            'C=C & Content & vbCrlf
            'C = C & "</div>" & vbCrlf
            'Content = C
            '第二种
            content = htmlAddAction(content, controlStr) 

        'Content = "<div "& ControlStr &">" & Content & "</div>"
        end if 
    end if 
    handleDisplayOnlineEditDialog = content 
end function 
'获得控制内容
function getControlStr(url)
	getControlStr=""
    if request("gl") = "edit" then
        getControlStr = " onMouseMove=""onColor(this,'#FDFAC6','red')"" onMouseOut=""offColor(this,'','')"" onDblClick=""window1('" & url & "','信息修改')"" title='双击或右键选在线修改' oncontextmenu=""CommonMenu(event,this,'')" '删除网址为空
    end if 
end function 

'html加动作(20151103)  call rw(htmlAddAction("  <a href=""javascript:;"">222222</a>", "onclick=""javascript:alert(111);"" "))
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

